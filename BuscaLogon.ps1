# ========================================
# SCRIPT COMPLETO - BUSCA LOGONS (ALVO ESPECÍFICO OU DOMÍNIO)
# Data da Versão: 2025-07-08
# INCLUI LOGONS OFFLINE (CACHED INTERACTIVE - TIPO 11)
# ========================================

# --- Pré-requisitos ---
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    Write-Error "ERRO: Módulo 'ActiveDirectory' ou 'ImportExcel' não encontrado."
    Write-Error "Este script DEVE ser executado em uma máquina com as Ferramentas de Administração de Domínio (RSAT) instaladas."
    Read-Host "Pressione Enter para sair."
    return
}

# Verifica Privilégios de Administrador
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "ERRO: Este script precisa ser executado com privilégios de Administrador."
    Read-Host "Pressione Enter para sair."
    return
}

# --- Configuração ---
$outputDirectory = "C:\Temp\EventLog_Exports"
$commandTimeout = 60  # Timeout em segundos para comandos remotos

if (-NOT (Test-Path $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory | Out-Null }

# --- Entrada do Usuário (ALVO)---
Write-Host "=== BUSCA DE EVENTOS DE LOGON NO DOMÍNIO ===" -ForegroundColor Cyan
$targetComputerName = $null
while ([string]::IsNullOrWhiteSpace($targetComputerName)) {
    $targetComputerName = Read-Host "Digite o NOME do computador alvo ou 'T' para buscar em TODO o domínio"
}

# --- Busca de Computadores (Condicional) ---
$reportNamePart = ""
if ($targetComputerName -eq 'T') {
    Write-Host "`nBuscando todos os computadores no domínio..." -ForegroundColor Yellow
    $reportNamePart = "Dominio"
    try {
        $domainComputers = Get-ADComputer -Filter {OperatingSystem -like "*Windows*"} -Properties Name, OperatingSystem, LastLogonDate |
                           Where-Object { $_.Enabled -eq $true } |
                           Sort-Object Name
        Write-Host "Encontrados $($domainComputers.Count) computadores Windows ativos no domínio." -ForegroundColor Green
    }
    catch {
        Write-Error "Erro ao buscar computadores no domínio: $($_.Exception.Message)"
        Read-Host "Pressione Enter para sair."
        return
    }
}
else {
    Write-Host "`nVerificando o computador alvo '$targetComputerName' no Active Directory..." -ForegroundColor Yellow
    $reportNamePart = $targetComputerName
    try {
        $singleComputer = Get-ADComputer -Identity $targetComputerName -Properties Name, OperatingSystem, LastLogonDate -ErrorAction Stop
        if ($singleComputer.Enabled -ne $true) {
            Write-Error "O computador '$targetComputerName' está desabilitado no Active Directory."
            Read-Host "Pressione Enter para sair."
            return
        }
        $domainComputers = @($singleComputer) # Coloca o objeto único em um array
        Write-Host "Alvo definido com sucesso para: $($singleComputer.Name)" -ForegroundColor Green
    }
    catch {
        Write-Error "Erro: Computador '$targetComputerName' não encontrado no Active Directory ou falha na consulta."
        Read-Host "Pressione Enter para sair."
        return
    }
}

# --- Entrada do Usuário (DATAS) ---
Write-Host "`nFormato de data: AAAA-MM-DD (exemplo: 2025-01-15)" -ForegroundColor Yellow

$inputStartDate = $null
$inputEndDate = $null

while ($null -eq $inputStartDate -or $null -eq $inputEndDate) {
    try {
        if ($null -eq $inputStartDate) {
            $input = Read-Host "Digite a DATA INICIAL da busca (AAAA-MM-DD)"
            $inputStartDate = [datetime]$input
        }
        if ($null -eq $inputEndDate) {
            $input = Read-Host "Digite a DATA FINAL da busca (AAAA-MM-DD)"
            $inputEndDate = [datetime]$input
        }
        if ($inputEndDate -lt $inputStartDate) {
            Write-Warning "A data final não pode ser anterior à data inicial."
            $inputStartDate = $null
            $inputEndDate = $null
        }
    }
    catch {
        Write-Warning "Formato de data inválido. Use AAAA-MM-DD."
        $inputStartDate = $null
        $inputEndDate = $null
    }
}

# --- Preparação da Query ---
$startDateFormatted = $inputStartDate.Date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z')
$endDateObject = $inputEndDate.Date.AddDays(1).AddTicks(-1)
$endDateFormatted = $endDateObject.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z')

# Query para buscar logons humanos: Interativo(2), Remote(10), Desbloqueio(7) e Offline/Cache(11)
$xPathQuery = "*[System[EventID=4624 and TimeCreated[@SystemTime>='$startDateFormatted' and @SystemTime<='$endDateFormatted']] and EventData[Data[@Name='LogonType']='2' or Data[@Name='LogonType']='10' or Data[@Name='LogonType']='7' or Data[@Name='LogonType']='11']]"

Write-Host "`nPeríodo: $($inputStartDate.ToString('dd/MM/yyyy')) até $($inputEndDate.ToString('dd/MM/yyyy'))" -ForegroundColor Yellow

# --- Coleta de Eventos ---
$allFoundEvents = New-Object System.Collections.ArrayList
$successCount = 0
$failCount = 0
$offlineCount = 0
$timeoutCount = 0
$totalComputers = $domainComputers.Count

Write-Host "`nIniciando coleta de eventos..." -ForegroundColor Yellow
Write-Host "Buscando logons: Interativo(2), Remote(10), Desbloqueio(7) e Offline/Cache(11)" -ForegroundColor Yellow

foreach ($computer in $domainComputers) {
    $computerName = $computer.Name
    $currentIndex = $domainComputers.IndexOf($computer) + 1
   
    Write-Host "[$currentIndex/$totalComputers] Verificando: $computerName" -ForegroundColor White
   
    if (Test-Connection -ComputerName $computerName -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        try {
            # Usa Job com timeout para evitar travamento
            $job = Start-Job -ScriptBlock {
                param($computerName, $xPathQuery)
                try {
                    Invoke-Command -ComputerName $computerName -ScriptBlock {
                        param($xPathQuery)
                        try {
                            $events = Get-WinEvent -LogName Security -FilterXPath $xPathQuery -ErrorAction Stop
                           
                            $results = @()
                            foreach ($event in $events) {
                                $xml = [xml]$event.ToXml()
                                $results += [PSCustomObject]@{
                                    LogonTime    = $event.TimeCreated
                                    ComputerName = $event.MachineName
                                    LogonType    = $xml.Event.EventData.Data[8].'#text'
                                    UserName     = $xml.Event.EventData.Data[5].'#text'
                                    UserDomain   = $xml.Event.EventData.Data[6].'#text'
                                    IpAddress    = if ($xml.Event.EventData.Data[18]) { $xml.Event.EventData.Data[18].'#text' } else { "N/A" }
                                    Workstation  = if ($xml.Event.EventData.Data[11]) { $xml.Event.EventData.Data[11].'#text' } else { "N/A" }
                                }
                            }
                            return $results
                        }
                        catch {
                            return @()
                        }
                    } -ArgumentList $xPathQuery -ErrorAction Stop
                }
                catch {
                    throw $_
                }
            } -ArgumentList $computerName, $xPathQuery
            
            # Espera o job completar com timeout
            $completed = $job | Wait-Job -Timeout $commandTimeout
            
            if ($completed) {
                $remoteEvents = Receive-Job -Job $job
                Remove-Job -Job $job -Force
                
                if ($null -ne $remoteEvents -and $remoteEvents.Count -gt 0) {
                    $remoteEvents | ForEach-Object { [void]$allFoundEvents.Add($_) }
                    Write-Host "   -> SUCESSO: Encontrados $($remoteEvents.Count) eventos em $computerName" -ForegroundColor Green
                } else {
                    Write-Host "   -> SEM EVENTOS: Nenhum logon correspondente encontrado em $computerName" -ForegroundColor Gray
                }
                $successCount++
            } else {
                # Timeout - mata o job e continua
                Write-Host "   -> TIMEOUT: Máquina $computerName demorou mais que ${commandTimeout}s - continuando..." -ForegroundColor Yellow
                Stop-Job -Job $job -Force
                Remove-Job -Job $job -Force
                $timeoutCount++
            }
        }
        catch {
            Write-Warning "   -> FALHA: Erro ao acessar logs de '$computerName': $($_.Exception.Message)"
            $failCount++
        }
    }
    else {
        Write-Host "   -> OFFLINE: Máquina $computerName não está respondendo" -ForegroundColor DarkYellow
        $offlineCount++
    }
}

# --- Relatório da Coleta ---
Write-Host "`n=== RESUMO DA COLETA ===" -ForegroundColor Cyan
Write-Host "Máquinas processadas com sucesso: $successCount" -ForegroundColor Green
Write-Host "Máquinas offline: $offlineCount" -ForegroundColor Yellow
Write-Host "Máquinas com timeout: $timeoutCount" -ForegroundColor Yellow
Write-Host "Máquinas com falha: $failCount" -ForegroundColor Red
Write-Host "Total de eventos de logon coletados: $($allFoundEvents.Count)" -ForegroundColor White

# --- Processamento e Exportação ---
if ($allFoundEvents.Count -gt 0) {
    Write-Host "`nProcessando eventos para exportação..." -ForegroundColor Yellow
   
    $detailedResults = $allFoundEvents | ForEach-Object {
        $properties = [ordered]@{
            LogonTime    = $_.LogonTime
            ComputerName = $_.ComputerName
            LogonType    = switch ($_.LogonType) {
                2  { "Interativo" }
                7  { "Desbloqueio" }
                10 { "RemoteInterativo" }
                11 { "Offline (Cache)" }
                default { "Outro ($($_.LogonType))" }
            }
            UserName     = $_.UserName
            UserDomain   = $_.UserDomain
            IpAddress    = $_.IpAddress
            Workstation  = $_.Workstation
        }
        New-Object -TypeName PSObject -Property $properties
    }

    $outputPath = Join-Path -Path $outputDirectory -ChildPath "RELATORIO_Logons_$($reportNamePart)_$($inputStartDate.ToString('yyyyMMdd'))-a-$($inputEndDate.ToString('yyyyMMdd')).xlsx"
   
    try {
        $detailedResults | Export-Excel -Path $outputPath -AutoSize -TableName "Logons" -TableStyle Medium9 -ClearSheet
        Write-Host "`nSUCESSO! Relatório Excel salvo em:" -ForegroundColor Green
        Write-Host "$outputPath" -ForegroundColor White
    }
    catch {
        Write-Warning "Erro ao exportar Excel. Salvando em CSV..."
        $csvPath = $outputPath -replace '\.xlsx$', '.csv'
        $detailedResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Relatório CSV salvo em: $csvPath" -ForegroundColor Yellow
    }
   
    Write-Host "`nPreview dos primeiros 10 eventos:" -ForegroundColor Cyan
    $detailedResults | Select-Object -First 10 | Format-Table -AutoSize
   
}
else {
    Write-Host "`nNenhum evento de logon correspondente encontrado no período especificado para o alvo definido." -ForegroundColor Yellow
}

Write-Host "`nScript finalizado!" -ForegroundColor Green
Read-Host "Pressione Enter para sair"
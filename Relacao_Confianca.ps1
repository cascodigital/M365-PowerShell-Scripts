<#
.SYNOPSIS
    Auditoria completa de relacao de confianca entre computadores e dominio Active Directory.

.DESCRIPTION    
    Script unificado que combina verificacao rapida em lote com diagnostico detalhado de falhas.
    Oferece quatro modos de operacao:
    1. Verificacao rapida: Testa um computador especifico com nltest.
    2. Diagnostico completo: Analise profunda de problemas de confianca.
    3. Lote ATIVO: Verifica apenas computadores ativos e gera relatorio Excel.
    4. Lote COMPLETO: Verifica todos os computadores do dominio.

.PARAMETER None
    O script solicita a selecao do modo de operacao no inicio.

.EXAMPLE
    BuscaRelacaoConfianca.ps1
    # Apresenta um menu para escolher entre verificar uma maquina ou todas.

.OUTPUTS
    - Console: Resultado para verificacao de maquina unica ou progresso da verificacao em lote.
    - Arquivo Excel (.xlsx): C:\Temp\Relatorio_RelacaoDeConfianca_[timestamp].xlsx (apenas no modo lote).

.NOTES
    Versao Unificada: 4.0 (Corrigida - filtro de computadores ativos)
    Combina verificacao rapida + diagnostico detalhado + relatorio inteligente
    Autor         : Andre Kittler
    Versao        :  3.1 (Corrigida logica de sincronizacao) - Combina verificacao rapida + diagnostico detalhado + relatorio em lote
    Compatibilidade: PowerShell 5.1+, Windows Server/Client

    Requisitos obrigatorios:
    - Modulo ActiveDirectory (parte das RSAT Tools).
    - Modulo ImportExcel (Execute: Install-Module ImportExcel -Scope CurrentUser).
    - Privilegios de usuario de dominio ou superior.
    - Conectividade de rede com os controladores de dominio.
#>


















# === CONFIGURACAO ===
$reportFolder = "C:\Temp"
$daysRecentActivity = 90  # Apenas computadores com atividade nos ultimos X dias

# === PREPARACAO DO AMBIENTE ===
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    # ImportExcel opcional para modo lote
    try { Import-Module ImportExcel -ErrorAction Stop } catch { $noExcel = $true }
}
catch {
    Write-Host "ERRO: Modulo ActiveDirectory nao encontrado." -ForegroundColor Red
    return
}

$domainName = (Get-ADDomain).NetBIOSName
if (-not $domainName) {
    Write-Host "ERRO: Nao foi possivel obter o nome do dominio." -ForegroundColor Red
    return
}

if (-not (Test-Path -Path $reportFolder)) {
    New-Item -Path $reportFolder -ItemType Directory | Out-Null
}

# === FUNCAO DE VERIFICACAO RAPIDA MELHORADA ===
function Test-TrustRelationship {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [Parameter(Mandatory = $true)]
        [string]$DomainName,
        [Parameter(Mandatory = $true)]
        [object]$ADComputerObject
    )

    $statusObject = [PSCustomObject]@{
        ComputerName  = $ComputerName
        Status        = ''
        Detalhes      = ''
        LastLogonDate = $ADComputerObject.LastLogonDate
        DaysOffline   = $null
        OperatingSystem = $ADComputerObject.OperatingSystem
        PingOK        = $false
        TrustOK       = $false
    }

    # Calcular dias offline de forma mais precisa
    if ($ADComputerObject.LastLogonDate -and $ADComputerObject.LastLogonDate -is [datetime]) {
        $timeSpan = New-TimeSpan -Start $ADComputerObject.LastLogonDate -End (Get-Date)
        $statusObject.DaysOffline = [math]::Round($timeSpan.TotalDays)
    }
    else {
        $statusObject.DaysOffline = "Desconhecido"
    }

    # Teste de conectividade
    if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
        $statusObject.PingOK = $true

        # Usa nltest para verificar canal seguro
        $nltestOutput = & nltest /sc_query:$domainName /server:$ComputerName 2>&1

        if ($nltestOutput -match "Success") {
            $statusObject.Status = "OK"
            $statusObject.TrustOK = $true
            $statusObject.Detalhes = "Online e relacao de confianca valida."
        }
        else {
            $statusObject.Status = "FALHA (Confianca)"
            $statusObject.TrustOK = $false
            $failureReason = ($nltestOutput | Select-Object -Last 1).ToString().Trim()
            $statusObject.Detalhes = "Relacao de confianca quebrada. Detalhe: $failureReason"
        }
    }
    else {
        $statusObject.Status = "Offline"
        $statusObject.PingOK = $false

        if ($statusObject.DaysOffline -ne "Desconhecido") {
            $statusObject.Detalhes = "Aproximadamente $($statusObject.DaysOffline) dias offline."
        }
        else {
            $statusObject.Detalhes = "Nunca fez logon ou data de logon invalida."
        }
    }

    return $statusObject
}

# === FUNCAO DE DIAGNOSTICO COMPLETO ===
function Invoke-CompleteDiagnostic {
    param(
        [string]$ComputerName,
        [string]$DomainName
    )

    Write-Host ""
    Write-Host "=====================================================" -ForegroundColor Magenta
    Write-Host "           DIAGNOSTICO COMPLETO DE FALHA" -ForegroundColor Magenta
    Write-Host "=====================================================" -ForegroundColor Magenta
    Write-Host ""

    $diagnosticResults = @{
        ComputerInAD = $false
        ConnectivityOK = $false
        TrustEvents = 0
        TimeSyncOK = $false
        OverallStatus = ""
        Recommendations = @()
    }

    # 1. Verificar computador no AD
    Write-Host "=== 1. VERIFICACAO NO ACTIVE DIRECTORY ===" -ForegroundColor Cyan
    try {
        $adComputer = Get-ADComputer -Identity $ComputerName -Properties LastLogonDate, PasswordLastSet, Enabled, OperatingSystem -ErrorAction Stop

        Write-Host "Computador encontrado no AD: SIM" -ForegroundColor Green
        Write-Host "Status: $($adComputer.Enabled)" -ForegroundColor White
        Write-Host "Ultimo Logon: $($adComputer.LastLogonDate)" -ForegroundColor White
        Write-Host "Senha alterada em: $($adComputer.PasswordLastSet)" -ForegroundColor White
        Write-Host "Sistema: $($adComputer.OperatingSystem)" -ForegroundColor White

        $diagnosticResults.ComputerInAD = $true

        # Verificar idade da senha do computador
        if ($adComputer.PasswordLastSet) {
            $daysSincePasswordChange = (Get-Date) - $adComputer.PasswordLastSet
            Write-Host "Dias desde alteracao da senha: $([Math]::Round($daysSincePasswordChange.TotalDays, 1))" -ForegroundColor White

            if ($daysSincePasswordChange.TotalDays -gt 30) {
                Write-Host "ALERTA: Senha do computador muito antiga (>30 dias)" -ForegroundColor Red
                $diagnosticResults.Recommendations += "Senha do computador precisa ser renovada"
            }
        }

    } catch {
        Write-Host "Computador NAO encontrado no AD" -ForegroundColor Red
        Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red
        $diagnosticResults.Recommendations += "Computador nao encontrado no Active Directory"
        return $diagnosticResults
    }

    # 2. Teste de conectividade detalhado
    Write-Host ""
    Write-Host "=== 2. TESTE DE CONECTIVIDADE DETALHADO ===" -ForegroundColor Cyan

    $pingOK = $false
    $portsOK = 0

    if (Test-Connection -ComputerName $ComputerName -Count 2 -Quiet) {
        Write-Host "Ping: OK" -ForegroundColor Green
        $pingOK = $true

        # Teste de portas importantes
        $ports = @(
            @{Port=135; Name="RPC"},
            @{Port=445; Name="SMB"},
            @{Port=5985; Name="WinRM"}
        )

        foreach ($portInfo in $ports) {
            try {
                $tcpTest = Test-NetConnection -ComputerName $ComputerName -Port $portInfo.Port -InformationLevel Quiet -ErrorAction SilentlyContinue
                $status = if ($tcpTest) { "OK"; $portsOK++ } else { "FALHOU" }
                $color = if ($tcpTest) { "Green" } else { "Red" }
                Write-Host "Porta $($portInfo.Port) ($($portInfo.Name)): $status" -ForegroundColor $color
            } catch {
                Write-Host "Porta $($portInfo.Port) ($($portInfo.Name)): ERRO" -ForegroundColor Red
            }
        }

        if ($pingOK -and $portsOK -ge 2) {
            $diagnosticResults.ConnectivityOK = $true
        } elseif ($pingOK) {
            $diagnosticResults.Recommendations += "Conectividade parcial - algumas portas bloqueadas"
        }

    } else {
        Write-Host "Ping: FALHOU" -ForegroundColor Red
        $diagnosticResults.Recommendations += "Computador esta offline ou com problemas de rede"
    }

    # 3. Verificar eventos de falha de confianca
    Write-Host ""
    Write-Host "=== 3. EVENTOS DE FALHA DE CONFIANCA (ULTIMOS 7 DIAS) ===" -ForegroundColor Cyan

    $totalEvents = 0
    $events = @(
        @{ID=4776; Desc="Falha de autenticacao do computador"; Log="Security"},
        @{ID=5722; Desc="Falha de confianca NetLogon"; Log="System"},
        @{ID=3210; Desc="Falha de canal seguro"; Log="System"}
    )

    foreach ($event in $events) {
        try {
            $logs = Get-WinEvent -FilterHashtable @{
                LogName=$event.Log
                ID=$event.ID
                StartTime=(Get-Date).AddDays(-7)
            } -MaxEvents 5 -ErrorAction SilentlyContinue |
                Where-Object { $_.Message -like "*$ComputerName*" }

            if ($logs) {
                $count = $logs.Count
                $totalEvents += $count
                Write-Host "Event ID $($event.ID) - $($event.Desc): $count eventos encontrados" -ForegroundColor Red
                $recent = $logs | Select-Object -First 1
                Write-Host "  Mais recente: $($recent.TimeCreated)" -ForegroundColor Yellow
            } else {
                Write-Host "Event ID $($event.ID) - $($event.Desc): Nenhum evento encontrado" -ForegroundColor Green
            }
        } catch {
            Write-Host "Erro ao verificar Event ID $($event.ID)" -ForegroundColor Yellow
        }
    }

    $diagnosticResults.TrustEvents = $totalEvents
    if ($totalEvents -gt 0) {
        $diagnosticResults.Recommendations += "Eventos de falha de confianca detectados ($totalEvents eventos)"
    }

    # 4. Verificar sincronizacao de tempo (MELHORADO)
    Write-Host ""
    Write-Host "=== 4. TESTE DE CANAL SEGURO E SINCRONIZACAO ===" -ForegroundColor Cyan

    # Primeiro, testar nltest para validar se a confianca esta realmente OK
    Write-Host "Testando canal seguro com nltest..." -ForegroundColor White
    $nltestOutput = & nltest /sc_query:$DomainName /server:$ComputerName 2>&1
    $nltestOK = $nltestOutput -match "Success"

    if ($nltestOK) {
        Write-Host "nltest: Canal seguro OK" -ForegroundColor Green
        $diagnosticResults.TimeSyncOK = $true  # Se nltest OK, assumir time sync OK
    } else {
        Write-Host "nltest: Canal seguro com problema" -ForegroundColor Red
        Write-Host "Detalhe: $($nltestOutput | Select-Object -Last 1)" -ForegroundColor Gray
        $diagnosticResults.Recommendations += "Canal seguro falhando no nltest"
        $diagnosticResults.TimeSyncOK = $false
    }

    # Depois verificar w32tm (informativo, nao critico se nltest estiver OK)
    try {
        Write-Host ""
        Write-Host "Verificando w32tm monitor (informativo)..." -ForegroundColor White
        $w32tmResult = cmd /c "w32tm /monitor /domain:$DomainName 2>&1"

        $computerFound = $false
        $timeOffset = $null

        # Mostrar apenas as linhas relevantes do w32tm
        $relevantLines = $w32tmResult | Where-Object { 
            $_ -match $ComputerName -or 
            $_ -match "NTP:" -or 
            $_ -match "error" -or
            $_ -match "failed"
        }

        if ($relevantLines) {
            foreach ($line in $relevantLines) {
                Write-Host $line -ForegroundColor Gray

                if ($line -match $ComputerName) {
                    $computerFound = $true
                    Write-Host ">>> Computador $ComputerName encontrado na sincronizacao!" -ForegroundColor Green

                    if ($line -match "(\+|\-)\d+\.\d+s") {
                        $timeOffset = $matches[0]
                        Write-Host ">>> Offset de tempo: $timeOffset" -ForegroundColor Yellow

                        # So marcar como problema se offset for muito grande E nltest falhar
                        if ([Math]::Abs([float]$timeOffset.TrimEnd('s')) -gt 300 -and -not $nltestOK) {
                            $diagnosticResults.Recommendations += "Diferenca de horario maior que 5 minutos + canal seguro com problema"
                            $diagnosticResults.TimeSyncOK = $false
                        }
                    }
                }
            }
        } else {
            Write-Host "Nenhuma informacao relevante no w32tm monitor" -ForegroundColor Gray
        }

        if (-not $computerFound -and $nltestOK) {
            Write-Host ""
            Write-Host "INFO: Computador nao aparece no w32tm monitor, mas nltest OK" -ForegroundColor Yellow
            Write-Host "Isso e NORMAL - computador pode sincronizar diretamente com DC" -ForegroundColor Yellow
        } elseif (-not $computerFound -and -not $nltestOK) {
            Write-Host ""
            Write-Host "ALERTA: Problema tanto no nltest quanto no w32tm" -ForegroundColor Red
            $diagnosticResults.Recommendations += "Computador com problemas de canal seguro e sincronizacao"
        }

    } catch {
        Write-Host "Erro ao verificar w32tm: $($_.Exception.Message)" -ForegroundColor Red
        # Se nltest OK, nao considerar erro critico
        if (-not $nltestOK) {
            $diagnosticResults.Recommendations += "Erro ao verificar sincronizacao de tempo"
        }
    }

    # Determinar status geral (AJUSTADO)
    if ($diagnosticResults.ComputerInAD -and $diagnosticResults.ConnectivityOK -and $diagnosticResults.TrustEvents -eq 0 -and $diagnosticResults.TimeSyncOK) {
        $diagnosticResults.OverallStatus = "SAUDAVEL"
    } elseif ($diagnosticResults.ComputerInAD -and $diagnosticResults.ConnectivityOK -and $diagnosticResults.TimeSyncOK) {
        # Se nltest OK mas teve alguns eventos, ainda considerar saudavel
        $diagnosticResults.OverallStatus = "SAUDAVEL (com alertas menores)"
    } elseif ($diagnosticResults.ComputerInAD -and $diagnosticResults.ConnectivityOK) {
        $diagnosticResults.OverallStatus = "PROBLEMAS DETECTADOS"
    } elseif ($diagnosticResults.ComputerInAD) {
        $diagnosticResults.OverallStatus = "CONECTIVIDADE COMPROMETIDA"
    } else {
        $diagnosticResults.OverallStatus = "CRITICO"
    }

    # Resumo do diagnostico
    Write-Host ""
    Write-Host "=====================================================" -ForegroundColor Magenta
    Write-Host "                RESUMO DO DIAGNOSTICO" -ForegroundColor Magenta
    Write-Host "=====================================================" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "Computador: $ComputerName" -ForegroundColor White
    Write-Host "Dominio: $DomainName" -ForegroundColor White
    Write-Host "Data/Hora: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor White
    Write-Host ""

    # Status individual
    Write-Host "STATUS INDIVIDUAL:" -ForegroundColor Cyan
    $checkAD = if ($diagnosticResults.ComputerInAD) { "[OK] PASS" } else { "[X] FAIL" }
    $checkConn = if ($diagnosticResults.ConnectivityOK) { "[OK] PASS" } else { "[X] FAIL" }
    $checkEvents = if ($diagnosticResults.TrustEvents -eq 0) { "[OK] PASS" } else { "[!] $($diagnosticResults.TrustEvents) eventos" }
    $checkTime = if ($diagnosticResults.TimeSyncOK) { "[OK] PASS" } else { "[X] FAIL" }

    Write-Host "  Computador no AD: $checkAD" -ForegroundColor $(if ($diagnosticResults.ComputerInAD) { "Green" } else { "Red" })
    Write-Host "  Conectividade: $checkConn" -ForegroundColor $(if ($diagnosticResults.ConnectivityOK) { "Green" } else { "Red" })
    Write-Host "  Eventos de Falha: $checkEvents" -ForegroundColor $(if ($diagnosticResults.TrustEvents -eq 0) { "Green" } else { "Yellow" })
    Write-Host "  Canal Seguro (nltest): $checkTime" -ForegroundColor $(if ($diagnosticResults.TimeSyncOK) { "Green" } else { "Red" })
    Write-Host ""

    Write-Host "STATUS GERAL: $($diagnosticResults.OverallStatus)" -ForegroundColor $(
        switch ($diagnosticResults.OverallStatus) {
            "SAUDAVEL" { "Green" }
            "SAUDAVEL (com alertas menores)" { "Green" }
            "PROBLEMAS DETECTADOS" { "Yellow" }
            "CONECTIVIDADE COMPROMETIDA" { "DarkYellow" }
            "CRITICO" { "Red" }
        }
    )

    # Recomendacoes
    if ($diagnosticResults.Recommendations.Count -gt 0) {
        Write-Host ""
        Write-Host "PROBLEMAS IDENTIFICADOS:" -ForegroundColor Red
        for ($i = 0; $i -lt $diagnosticResults.Recommendations.Count; $i++) {
            Write-Host "  $($i + 1). $($diagnosticResults.Recommendations[$i])" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "COMANDOS DE CORRECAO SUGERIDOS:" -ForegroundColor Cyan
    switch ($diagnosticResults.OverallStatus) {
        "SAUDAVEL" {
            Write-Host "Nenhuma acao necessaria. Sistema funcionando corretamente." -ForegroundColor Green
        }

        "SAUDAVEL (com alertas menores)" {
            Write-Host "Sistema funcionando, mas monitore os eventos mencionados." -ForegroundColor Green
        }

        "PROBLEMAS DETECTADOS" {
            Write-Host ""
            Write-Host "1. RESETAR CANAL SEGURO (se tiver acesso ao computador):" -ForegroundColor White
            Write-Host "   Test-ComputerSecureChannel -Repair -Credential (Get-Credential)" -ForegroundColor Gray
            Write-Host ""
            Write-Host "2. RESETAR SENHA DO COMPUTADOR NO AD:" -ForegroundColor White
            Write-Host "   Reset-ComputerMachinePassword -Server $DomainName -Credential (Get-Credential)" -ForegroundColor Gray
            Write-Host ""
            Write-Host "3. SE PROBLEMA PERSISTIR, RESETAR CONTA NO AD:" -ForegroundColor White
            Write-Host "   `$senha = ConvertTo-SecureString 'NovaSenh@123!' -AsPlainText -Force" -ForegroundColor Gray
            Write-Host "   Set-ADAccountPassword -Identity '$ComputerName$' -NewPassword `$senha -Reset" -ForegroundColor Gray
        }

        "CONECTIVIDADE COMPROMETIDA" {
            Write-Host ""
            Write-Host "1. VERIFICAR SE O COMPUTADOR ESTA LIGADO" -ForegroundColor White
            Write-Host "2. VERIFICAR CONECTIVIDADE DE REDE" -ForegroundColor White
            Write-Host "3. QUANDO VOLTAR ONLINE, EXECUTAR:" -ForegroundColor White
            Write-Host "   Test-ComputerSecureChannel -Repair -Credential (Get-Credential)" -ForegroundColor Gray
        }

        "CRITICO" {
            Write-Host ""
            Write-Host "PROBLEMA CRITICO DETECTADO!" -ForegroundColor Red
            Write-Host "1. VERIFICAR SE O NOME DO COMPUTADOR ESTA CORRETO" -ForegroundColor White
            Write-Host "2. READICIONAR AO DOMINIO:" -ForegroundColor White
            Write-Host "   Add-Computer -DomainName $DomainName -Restart -Credential (Get-Credential)" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "=====================================================" -ForegroundColor Magenta

    return $diagnosticResults
}

# === MENU PRINCIPAL ===
Clear-Host
Write-Host "Auditoria Unificada de Relacao de Confianca - Dominio: $domainName" -ForegroundColor Yellow
Write-Host ("=" * 80)
Write-Host "Configuracao atual: Incluir apenas computadores ativos nos ultimos $daysRecentActivity dias" -ForegroundColor Cyan
Write-Host ""
Write-Host "Selecione uma opcao:"
Write-Host "1. Verificacao RAPIDA de um computador (nltest)."
Write-Host "2. Diagnostico COMPLETO de um computador (analise detalhada)."
Write-Host "3. Verificar computadores ATIVOS e gerar relatorio Excel. [RECOMENDADO]"
Write-Host "4. Verificar TODOS os computadores (incluindo muito antigos)."

$choice = Read-Host "Digite sua opcao (1, 2, 3 ou 4)"

switch ($choice) {
    '1' {
        # === MODO: VERIFICACAO RAPIDA ===
        do {
            $targetComputer = Read-Host "`nDigite o nome do computador"
            if ([string]::IsNullOrWhiteSpace($targetComputer)) {
                Write-Host "Nome nao pode ser vazio." -ForegroundColor Yellow
                continue
            }

            Write-Host "Buscando '$targetComputer' no Active Directory..."
            $adComp = Get-ADComputer -Identity $targetComputer -Properties LastLogonDate, OperatingSystem -ErrorAction SilentlyContinue

            if (-not $adComp) {
                Write-Host "ERRO: Computador '$targetComputer' nao encontrado no AD." -ForegroundColor Red
            } else {
                Write-Host "Verificando relacao de confianca para: $($adComp.Name)..."
                $result = Test-TrustRelationship -ComputerName $adComp.Name -DomainName $domainName -ADComputerObject $adComp

                Write-Host "`n--- Resultado Rapido para $($result.ComputerName) ---" -ForegroundColor Green
                $statusColor = if ($result.Status -eq "OK") {"Green"} elseif ($result.Status -like "FALHA*") {"Red"} else {"Yellow"}
                Write-Host "Status        :" -NoNewline; Write-Host " $($result.Status)" -ForegroundColor $statusColor
                Write-Host "Detalhes      : $($result.Detalhes)"
                Write-Host "Ultimo Logon  : $($result.LastLogonDate)"
                Write-Host "Sistema       : $($result.OperatingSystem)"
                Write-Host ("=" * 60)

                # Oferecer diagnostico completo se houver falha
                if ($result.Status -like "FALHA*") {
                    $diagnostic = Read-Host "`nDeseja executar DIAGNOSTICO COMPLETO? (s/n)"
                    if ($diagnostic -eq 's') {
                        Invoke-CompleteDiagnostic -ComputerName $adComp.Name -DomainName $domainName
                    }
                }
            }

            $another = Read-Host "`nDeseja verificar outro computador? (s/n)"
        } while ($another -eq 's')
    }

    '2' {
        # === MODO: DIAGNOSTICO COMPLETO ===
        $targetComputer = Read-Host "`nDigite o nome do computador para diagnostico completo"
        if (-not [string]::IsNullOrWhiteSpace($targetComputer)) {
            $adComp = Get-ADComputer -Identity $targetComputer -Properties LastLogonDate -ErrorAction SilentlyContinue
            if ($adComp) {
                Invoke-CompleteDiagnostic -ComputerName $adComp.Name -DomainName $domainName
            } else {
                Write-Host "ERRO: Computador nao encontrado no AD." -ForegroundColor Red
            }
        }
    }

    '3' {
        # === MODO: VERIFICAR COMPUTADORES ATIVOS ===
        if ($noExcel) {
            Write-Host "AVISO: Modulo ImportExcel nao instalado. Relatorio sera gerado em CSV." -ForegroundColor Yellow
        }

        $timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
        $reportPath = Join-Path -Path $reportFolder -ChildPath "Relatorio_RelacaoDeConfianca_Ativos_$timestamp.xlsx"
        $csvPath = Join-Path -Path $reportFolder -ChildPath "Relatorio_RelacaoDeConfianca_Ativos_$timestamp.csv"

        Write-Host "`nBuscando computadores ATIVOS (ultimos $daysRecentActivity dias) no AD..."

        # FILTRO MELHORADO - apenas computadores com atividade recente
        $recentDate = (Get-Date).AddDays(-$daysRecentActivity)

        $allComputers = Get-ADComputer -Filter {
            Enabled -eq $true -and 
            LastLogonDate -gt $recentDate
        } -Properties LastLogonDate, OperatingSystem, whenCreated

        $results = @()

        if (-not $allComputers) {
            Write-Host "Nenhum computador ativo nos ultimos $daysRecentActivity dias encontrado." -ForegroundColor Yellow
            Write-Host "Tente a opcao 4 para ver todos os computadores." -ForegroundColor Cyan
            return
        }

        Write-Host "Encontrados $($allComputers.Count) computadores ativos."
        Write-Host "Iniciando verificacao..."

        $processedCount = 0
        foreach ($computer in $allComputers) {
            $processedCount++
            Write-Progress -Activity "Verificando Computadores Ativos" -Status "Processando $($computer.Name)" -PercentComplete (($processedCount / $allComputers.Count) * 100)
            Write-Host "[$processedCount/$($allComputers.Count)] Verificando: $($computer.Name)"

            $result = Test-TrustRelationship -ComputerName $computer.Name -DomainName $domainName -ADComputerObject $computer

            # Adicionar informacoes extras ao resultado
            $result | Add-Member -NotePropertyName "WhenCreated" -NotePropertyValue $computer.whenCreated -Force

            $results += $result
        }

        Write-Progress -Activity "Verificando Computadores Ativos" -Completed

        if ($results) {
            try {
                if (-not $noExcel) {
                    Write-Host "`nExportando relatorio Excel para $reportPath..."
                    $results | Export-Excel -Path $reportPath -AutoSize -WorksheetName "Status de Confianca Ativos" -TableStyle Medium9 -FreezeTopRow -ErrorAction Stop
                    Write-Host "Relatorio Excel gerado: $reportPath" -ForegroundColor Green
                } else {
                    Write-Host "`nExportando relatorio CSV para $csvPath..."
                    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                    Write-Host "Relatorio CSV gerado: $csvPath" -ForegroundColor Green
                }

                # Resumo melhorado no console
                $summary = $results | Group-Object Status | Select-Object Name, Count
                Write-Host "`nResumo dos Resultados (Computadores Ativos - ultimos $daysRecentActivity dias):" -ForegroundColor Yellow
                $summary | ForEach-Object { 
                    $color = switch ($_.Name) {
                        "OK" { "Green" }
                        "Offline" { "Yellow" }
                        default { "Red" }
                    }
                    Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor $color
                }

                # Estatisticas adicionais
                $avgDaysOffline = ($results | Where-Object {$_.DaysOffline -ne "Desconhecido" -and $_.DaysOffline -is [int]} | Measure-Object -Property DaysOffline -Average).Average
                if ($avgDaysOffline) {
                    Write-Host "`nMedia de dias desde ultimo logon: $([math]::Round($avgDaysOffline, 1)) dias" -ForegroundColor Cyan
                }

                # Mostrar computadores com falha se houver
                $failures = $results | Where-Object { $_.Status -like "FALHA*" }
                if ($failures) {
                    Write-Host "`nComputadores ATIVOS com FALHA DE CONFIANCA:" -ForegroundColor Red
                    $failures | ForEach-Object { Write-Host "  - $($_.ComputerName) (ultimo logon: $($_.LastLogonDate))" -ForegroundColor Yellow }
                    Write-Host "`nUse a opcao 2 (Diagnostico Completo) para analisar cada um." -ForegroundColor Cyan
                }

                # Mostrar computadores offline mas recentes
                $offlineRecent = $results | Where-Object { $_.Status -eq "Offline" }
                if ($offlineRecent) {
                    Write-Host "`nComputadores offline (mas com atividade recente):" -ForegroundColor Yellow
                    $offlineRecent | ForEach-Object { Write-Host "  - $($_.ComputerName) ($($_.DaysOffline) dias offline)" -ForegroundColor Gray }
                }

            } catch {
                Write-Host "ERRO ao exportar: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    '4' {
        # === MODO: VERIFICAR TODOS OS COMPUTADORES (MODO ANTIGO) ===
        Write-Host "`nAVISO: Esta opcao incluira computadores muito antigos!" -ForegroundColor Yellow
        $confirm = Read-Host "Tem certeza? Pode incluir contas com 2000+ dias offline (s/n)"

        if ($confirm -eq 's') {
            Write-Host "Executando verificacao de TODOS os computadores habilitados..." -ForegroundColor Yellow

            if ($noExcel) {
                Write-Host "AVISO: Modulo ImportExcel nao instalado. Relatorio sera gerado em CSV." -ForegroundColor Yellow
            }

            $timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
            $reportPath = Join-Path -Path $reportFolder -ChildPath "Relatorio_RelacaoDeConfianca_TODOS_$timestamp.xlsx"
            $csvPath = Join-Path -Path $reportFolder -ChildPath "Relatorio_RelacaoDeConfianca_TODOS_$timestamp.csv"

            Write-Host "`nBuscando TODOS os computadores habilitados no AD..."
            $allComputers = Get-ADComputer -Filter {Enabled -eq $true} -Properties LastLogonDate, OperatingSystem
            $results = @()

            if (-not $allComputers) {
                Write-Host "Nenhum computador ativo encontrado." -ForegroundColor Yellow
                return
            }

            Write-Host "Iniciando verificacao em $($allComputers.Count) computadores..."

            $processedCount = 0
            foreach ($computer in $allComputers) {
                $processedCount++
                Write-Progress -Activity "Verificando Todos os Computadores" -Status "Processando $($computer.Name)" -PercentComplete (($processedCount / $allComputers.Count) * 100)
                Write-Host "[$processedCount/$($allComputers.Count)] Verificando: $($computer.Name)"
                $results += Test-TrustRelationship -ComputerName $computer.Name -DomainName $domainName -ADComputerObject $computer
            }

            Write-Progress -Activity "Verificando Todos os Computadores" -Completed

            if ($results) {
                try {
                    if (-not $noExcel) {
                        Write-Host "`nExportando relatorio Excel para $reportPath..."
                        $results | Export-Excel -Path $reportPath -AutoSize -WorksheetName "Status de Confianca TODOS" -TableStyle Medium9 -FreezeTopRow -ErrorAction Stop
                        Write-Host "Relatorio Excel gerado: $reportPath" -ForegroundColor Green
                    } else {
                        Write-Host "`nExportando relatorio CSV para $csvPath..."
                        $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                        Write-Host "Relatorio CSV gerado: $csvPath" -ForegroundColor Green
                    }

                    # Resumo no console
                    $summary = $results | Group-Object Status | Select-Object Name, Count
                    Write-Host "`nResumo dos Resultados (TODOS os computadores habilitados):" -ForegroundColor Yellow
                    $summary | ForEach-Object { 
                        $color = switch ($_.Name) {
                            "OK" { "Green" }
                            "Offline" { "Yellow" }
                            default { "Red" }
                        }
                        Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor $color
                    }

                    # Mostrar computadores muito antigos
                    $veryOld = $results | Where-Object { $_.DaysOffline -is [int] -and $_.DaysOffline -gt 365 }
                    if ($veryOld) {
                        Write-Host "`nComputadores com mais de 1 ano offline: $($veryOld.Count)" -ForegroundColor Yellow
                        $over2000 = ($veryOld | Where-Object { $_.DaysOffline -gt 2000 }).Count
                        if ($over2000 -gt 0) {
                            Write-Host "Computadores com mais de 2000 dias offline: $over2000" -ForegroundColor Red
                            Write-Host "RECOMENDACAO: Use a opcao 3 para focar apenas em computadores ativos." -ForegroundColor Cyan
                        }
                    }

                    # Mostrar computadores com falha
                    $failures = $results | Where-Object { $_.Status -like "FALHA*" }
                    if ($failures) {
                        Write-Host "`nComputadores com FALHA DE CONFIANCA:" -ForegroundColor Red
                        $failures | ForEach-Object { Write-Host "  - $($_.ComputerName)" -ForegroundColor Yellow }
                        Write-Host "`nUse a opcao 2 (Diagnostico Completo) para analisar cada um." -ForegroundColor Cyan
                    }

                } catch {
                    Write-Host "ERRO ao exportar: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "Operacao cancelada. Use a opcao 3 para computadores ativos." -ForegroundColor Cyan
        }
    }

    default {
        Write-Host "Opcao invalida. Use 1-4." -ForegroundColor Red
    }
}

Write-Host "`nScript finalizado." -ForegroundColor Green

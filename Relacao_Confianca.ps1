# --- Inicio do Script ---

# 1. ===== CONFIGURACAO =====
# Caminho da pasta onde o relatorio sera salvo.
$reportFolder = "C:\Temp"

# --- Fim da Configuracao ---

# 2. Preparar ambiente e nome do arquivo
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    # O modulo ImportExcel pode ser instalado com: Install-Module -Name ImportExcel
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    Write-Host "ERRO: Verifique se os modulos 'ActiveDirectory' e 'ImportExcel' estao instalados." -ForegroundColor Red
    Write-Host "Para instalar o ImportExcel, execute: Install-Module -Name ImportExcel -Scope CurrentUser"
    return
}

# Obter o nome NetBIOS do dominio atual automaticamente
$domainName = (Get-ADDomain).NetBIOSName
if (-not $domainName) {
    Write-Host "ERRO: Nao foi possivel obter o nome do dominio do Active Directory." -ForegroundColor Red
    return
}

Write-Host "Verificando para o dominio: $domainName" -ForegroundColor Cyan

# Criar a pasta de relatorios se ela nao existir
if (-not (Test-Path -Path $reportFolder)) {
    New-Item -Path $reportFolder -ItemType Directory | Out-Null
}

# Adiciona data e hora ao nome do arquivo
$timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$reportPath = Join-Path -Path $reportFolder -ChildPath "Relatorio_RelacaoDeConfianca_$timestamp.xlsx"

# 3. Obter TODOS os computadores ativos do Active Directory
Write-Host "Buscando TODOS os computadores ativos no Active Directory..."
$allComputers = Get-ADComputer -Filter {Enabled -eq $true} -Properties LastLogonDate

# Array para armazenar os resultados
$results = @()

# Verifica se algum computador foi encontrado
if (-not $allComputers) {
    Write-Host "Nenhum computador ativo foi encontrado no Active Directory. Saindo do script." -ForegroundColor Yellow
    return
}

Write-Host "Iniciando verificacao em $($allComputers.Count) computadores..."

# 4. Loop para verificar cada computador
$processedCount = 0
foreach ($computer in $allComputers) {
    $computerName = $computer.Name
    $processedCount++
    
    Write-Host "[$processedCount/$($allComputers.Count)] Verificando: $computerName"

    $statusObject = [PSCustomObject]@{
        ComputerName  = $computerName
        Status        = ''
        Detalhes      = ''
        LastLogonDate = $computer.LastLogonDate
    }

    # Teste de conectividade (ping)
    if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {
        # Usa o utilitario nltest, que nao depende de WinRM/PSRemoting
        # O '&' e o operador de chamada, necessario para executar comandos com argumentos
        $nltestOutput = & nltest /sc_query:$domainName /server:$computerName 2>&1
        
        # Verifica a saida do comando
        if ($nltestOutput -match "Success") {
            $statusObject.Status = "OK"
            $statusObject.Detalhes = "Online e relacao de confianca valida."
        }
        else {
            $statusObject.Status = "FALHA (Confianca)"
            # Captura a mensagem de erro do nltest para diagnostico
            $failureReason = ($nltestOutput | Select-Object -Last 1).ToString().Trim()
            $statusObject.Detalhes = "Relacao de confianca quebrada ou erro de comunicacao. Detalhe: $failureReason"
        }
    }
    else {
        # A maquina nao respondeu ao ping
        $statusObject.Status = "Offline"
        
        if ($computer.LastLogonDate -is [datetime]) {
            $timeSpan = New-TimeSpan -Start $computer.LastLogonDate -End (Get-Date)
            $daysOffline = [math]::Round($timeSpan.TotalDays)
            $statusObject.Detalhes = "Aproximadamente $daysOffline dias offline."
        }
        else {
            $statusObject.Detalhes = "Data de ultimo logon invalida ou inexistente no AD."
        }
    }

    $results += $statusObject
}

# 5. Exportar os resultados para um arquivo Excel
if ($results) {
    Write-Host "`nExportando relatorio para $reportPath..."
    try {
        $results | Export-Excel -Path $reportPath -AutoSize -WorksheetName "Status de Confianca" -TableStyle Medium9 -FreezeTopRow
        Write-Host "Relatorio gerado com sucesso em $reportPath" -ForegroundColor Green
        
        # Mostra um resumo
        $summary = $results | Group-Object Status | Select-Object Name, Count
        Write-Host "`nResumo:" -ForegroundColor Yellow
        $summary | ForEach-Object { Write-Host "  $($_.Name): $($_.Count)" }
    }
    catch {
        Write-Host "ERRO ao exportar para Excel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Os resultados estao na variavel `$results`. Voce pode exporta-los para CSV com:"
        Write-Host "`$results | Export-Csv -Path 'C:\Temp\relatorio.csv' -NoTypeInformation"
    }
}
else {
    Write-Host "Nenhum computador foi processado." -ForegroundColor Yellow
}

# --- Fim do Script ---

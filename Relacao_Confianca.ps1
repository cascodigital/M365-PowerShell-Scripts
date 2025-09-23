<#
.SYNOPSIS
    Auditoria automatizada de relacao de confianca entre computadores e dominio Active Directory

.DESCRIPTION
    Script completo para diagnostico e auditoria da integridade de relacoes de confianca entre
    computadores Windows e controladores de dominio Active Directory. Utiliza comando nltest
    nativo para verificacao precisa do canal seguro sem dependencias de WinRM ou PSRemoting.
    Gera relatorio Excel detalhado com status, diagnosticos e estatisticas de conectividade.
    
    Funcionalidades principais:
    - Descoberta automatica de todos os computadores ativos no AD
    - Verificacao de conectividade de rede (ping) pre-validacao
    - Teste de canal seguro usando nltest nativo do Windows
    - Calculo automatico de tempo offline baseado em LastLogonDate
    - Classificacao automatica: OK, FALHA (Confianca), Offline
    - Exportacao Excel formatada com tabelas e freeze de cabecalho
    - Relatorio de resumo com contadores por status
    - Tratamento robusto de erros com fallback para CSV
    
    Cenarios de uso corporativo:
    - Preparacao para migracao de dominio
    - Auditoria de seguranca e compliance
    - Troubleshooting de problemas de autenticacao
    - Limpeza de objetos obsoletos no Active Directory
    - Monitoramento proativo da saude do dominio

.PARAMETER None
    Script automatico - utiliza dominio atual e processa todos computadores ativos

.EXAMPLE
    .\Test-DomainTrustRelationship.ps1
    # Executa auditoria completa no dominio atual
    # Resultado: Relatorio Excel em C:\Temp\ com status de 250+ computadores

.INPUTS
    None - Script automatico usa dominio atual do computador de execucao

.OUTPUTS
    - Arquivo Excel (.xlsx): C:\Temp\Relatorio_RelacaoDeConfianca_[timestamp].xlsx
    - Planilha formatada: Status de Confianca com colunas organizadas
    - Console: Progresso detalhado e resumo estatistico final
    - Fallback CSV: Se exportacao Excel falhar

.NOTES
    Autor         : Andre Kittler
    Versao        : 1.0
    Compatibilidade: PowerShell 5.1+, Windows Server/Client
    
    Requisitos obrigatorios:
    - Modulo ActiveDirectory (RSAT Tools instalado)
    - Modulo ImportExcel (Install-Module ImportExcel)
    - Privilegios Domain User ou superior
    - Conectividade de rede com controladores de dominio
    - Firewall liberado para ICMP (ping) e RPC
    
    Tecnologias utilizadas:
    - Get-ADComputer para descoberta automatica
    - Test-Connection para validacao de conectividade
    - nltest /sc_query para verificacao de canal seguro
    - Export-Excel para relatorio formatado profissionalmente
    
    Interpretacao de resultados:
    - OK: Canal seguro funcionando corretamente
    - FALHA (Confianca): Relacao quebrada - requer netdom resetpwd
    - Offline: Computador nao responde - possivel desligado/removido
    
    Comandos de correcao comuns:
    - netdom resetpwd /server:[DC] /userd:[user] /passwordd:*
    - Remove-ADComputer [computer] (para objetos obsoletos)
    - Test-ComputerSecureChannel -Repair (PowerShell local)
    
    Consideracoes de performance:
    - Processamento sequencial para evitar sobrecarga de rede
    - Timeout automatico em computadores offline
    - Progress indicator para ambientes com muitos computadores
    
    Limitacoes conhecidas:
    - Nao funciona atraves de NAT ou firewalls restritivos
    - Requer resolucao DNS bidirecional funcional
    - Computadores em modo sleep podem aparecer como offline

.LINK
    https://docs.microsoft.com/en-us/windows-server/identity/ad-ds/manage/troubleshoot/troubleshooting-active-directory-replication-problems
#>


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

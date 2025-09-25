<#
.SYNOPSIS
    Auditoria completa de relacao de confianca entre computadores e dominio Active Directory.

.DESCRIPTION
    Script unificado que combina verificacao rapida em lote com diagnostico detalhado de falhas.
    Oferece tres modos de operacao:
    1. Verificacao rapida: Testa um computador especifico com nltest.
    2. Diagnostico completo: Analise profunda de problemas de confianca.
    3. Lote: Verifica todos os computadores do dominio e gera relatorio Excel.

.PARAMETER None
    O script solicita a selecao do modo de operacao no inicio.

.EXAMPLE
    BuscaRelacaoConfianca.ps1
    # Apresenta um menu para escolher entre verificar uma maquina ou todas.

.OUTPUTS
    - Console: Resultado para verificacao de maquina unica ou progresso da verificacao em lote.
    - Arquivo Excel (.xlsx): C:\Temp\Relatorio_RelacaoDeConfianca_[timestamp].xlsx (apenas no modo lote).

.NOTES
    Autor         : Andre Kittler
    Versao        :  3.1 (Corrigida logica de sincronizacao) - Combina verificacao rapida + diagnostico detalhado + relatorio em lote
    Compatibilidade: PowerShell 5.1+, Windows Server/Client

    Requisitos obrigatorios:
    - Modulo ActiveDirectory (parte das RSAT Tools).
    - Modulo ImportExcel (Execute: Install-Module ImportExcel -Scope CurrentUser).
    - Privilegios de usuario de dominio ou superior.
    - Conectividade de rede com os controladores de dominio.
#>

# --- Inicio do Script ---

# 1. ===== CONFIGURACAO =====
# Caminho da pasta onde o relatorio sera salvo (usado no modo de verificacao de todos).
$reportFolder = "C:\Temp"
# --- Fim da Configuracao ---


# 2. ===== PREPARACAO DO AMBIENTE =====
try {
    Import-Module ActiveDirectory -ErrorAction Stop
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

# Criar a pasta de relatorios se ela nao existir
if (-not (Test-Path -Path $reportFolder)) {
    New-Item -Path $reportFolder -ItemType Directory | Out-Null
}

# 3. ===== FUNCAO DE VERIFICACAO =====
# Centraliza a logica de teste para ser reutilizada
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
    }

    # Teste de conectividade (ping)
    if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
        # Usa o utilitario nltest
        $nltestOutput = & nltest /sc_query:$domainName /server:$ComputerName 2>&1
        
        # Verifica a saida do comando
        if ($nltestOutput -match "Success") {
            $statusObject.Status = "OK"
            $statusObject.Detalhes = "Online e relacao de confianca valida."
        }
        else {
            $statusObject.Status = "FALHA (Confianca)"
            $failureReason = ($nltestOutput | Select-Object -Last 1).ToString().Trim()
            $statusObject.Detalhes = "Relacao de confianca quebrada ou erro de comunicacao. Detalhe: $failureReason"
        }
    }
    else {
        # A maquina nao respondeu ao ping
        $statusObject.Status = "Offline"
        
        if ($ADComputerObject.LastLogonDate -is [datetime]) {
            $timeSpan = New-TimeSpan -Start $ADComputerObject.LastLogonDate -End (Get-Date)
            $daysOffline = [math]::Round($timeSpan.TotalDays)
            $statusObject.Detalhes = "Aproximadamente $daysOffline dias offline."
        }
        else {
            $statusObject.Detalhes = "Data de ultimo logon invalida ou inexistente no AD."
        }
    }

    return $statusObject
}


# 4. ===== MENU PRINCIPAL E SELECAO DE MODO =====
Clear-Host
Write-Host "Auditoria de Relacao de Confianca com o Dominio: $domainName" -ForegroundColor Yellow
Write-Host ("-" * 60)
Write-Host "Selecione uma opcao:"
Write-Host "1. Verificar um UNICO computador."
Write-Host "2. Verificar TODOS os computadores do dominio e gerar relatorio."
$choice = Read-Host "Digite sua opcao (1 ou 2)"

switch ($choice) {
    '1' {
        # --- MODO: VERIFICACAO UNICA ---
        do {
            $targetComputer = Read-Host "`nDigite o nome do computador a ser verificado"
            if ([string]::IsNullOrWhiteSpace($targetComputer)) {
                Write-Host "Nome do computador nao pode ser vazio." -ForegroundColor Yellow
                continue
            }

            Write-Host "Buscando '$targetComputer' no Active Directory..."
            $adComp = Get-ADComputer -Identity $targetComputer -Properties LastLogonDate -ErrorAction SilentlyContinue
            
            if (-not $adComp) {
                 Write-Host "ERRO: Computador '$targetComputer' nao encontrado no Active Directory." -ForegroundColor Red
            } else {
                Write-Host "Verificando relacao de confianca para: $($adComp.Name)..."
                $result = Test-TrustRelationship -ComputerName $adComp.Name -DomainName $domainName -ADComputerObject $adComp
                
                # Exibe o resultado formatado
                Write-Host "`n--- Resultado para $($result.ComputerName) ---" -ForegroundColor Green
                $statusColor = if ($result.Status -eq "OK") {"Green"} elseif ($result.Status -like "FALHA*") {"Red"} else {"Yellow"}
                Write-Host "Status        :" -NoNewline; Write-Host " $($result.Status)" -ForegroundColor $statusColor
                Write-Host "Detalhes      : $($result.Detalhes)"
                Write-Host "Ultimo Logon  : $($result.LastLogonDate)"
                Write-Host ("-" * 35)
            }

            $another = Read-Host "`nDeseja verificar outro computador? (s/n)"
        } while ($another -eq 's')
        Write-Host "Verificacao finalizada."
    }
    '2' {
        # --- MODO: VERIFICAR TODOS E GERAR RELATORIO ---
        $timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
        $reportPath = Join-Path -Path $reportFolder -ChildPath "Relatorio_RelacaoDeConfianca_$timestamp.xlsx"

        Write-Host "`nBuscando TODOS os computadores ativos no Active Directory..."
        $allComputers = Get-ADComputer -Filter {Enabled -eq $true} -Properties LastLogonDate
        $results = @()

        if (-not $allComputers) {
            Write-Host "Nenhum computador ativo foi encontrado no Active Directory. Saindo do script." -ForegroundColor Yellow
            return
        }

        Write-Host "Iniciando verificacao em $($allComputers.Count) computadores..."
        
        $processedCount = 0
        foreach ($computer in $allComputers) {
            $processedCount++
            Write-Progress -Activity "Verificando Computadores" -Status "Processando $($computer.Name)" -PercentComplete (($processedCount / $allComputers.Count) * 100)
            Write-Host "[$processedCount/$($allComputers.Count)] Verificando: $($computer.Name)"

            $results += Test-TrustRelationship -ComputerName $computer.Name -DomainName $domainName -ADComputerObject $computer
        }
        Write-Progress -Activity "Verificando Computadores" -Completed

        if ($results) {
            Write-Host "`nExportando relatorio para $reportPath..."
            try {
                $results | Export-Excel -Path $reportPath -AutoSize -WorksheetName "Status de Confianca" -TableStyle Medium9 -FreezeTopRow -ErrorAction Stop
                Write-Host "Relatorio gerado com sucesso em $reportPath" -ForegroundColor Green
                
                # Mostra um resumo no console
                $summary = $results | Group-Object Status | Select-Object Name, Count
                Write-Host "`nResumo:" -ForegroundColor Yellow
                $summary | ForEach-Object { Write-Host "  $($_.Name): $($_.Count)" }
            }
            catch {
                Write-Host "ERRO ao exportar para Excel: $($_.Exception.Message)" -ForegroundColor Red
                $csvPath = $reportPath -replace '\.xlsx$', '.csv'
                Write-Host "Como alternativa, exportando para CSV em: $csvPath"
                $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            }
        }
        else {
            Write-Host "Nenhum computador foi processado." -ForegroundColor Yellow
        }
    }
    default {
        Write-Host "Opcao invalida. O script sera encerrado." -ForegroundColor Red
    }
}

# --- Fim do Script ---

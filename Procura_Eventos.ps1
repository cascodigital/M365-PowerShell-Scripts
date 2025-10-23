<#
.SYNOPSIS
    Analisador abrangente de logs de eventos Windows com busca multi-ID e exportacao Excel avancada

.DESCRIPTION
    Script robusto para analise forense e investigacao de eventos Windows em ambiente corporativo.
    Realiza busca simultanea de multiplos Event IDs em todos os arquivos de log (.evtx) do sistema,
    incluindo logs arquivados e rotacionados. Exporta resultados estruturados para Excel com abas
    separadas por tipo de evento e formatacao profissional para analise detalhada.
    
    Funcionalidades principais:
    - Busca simultanea de multiplos Event IDs com operador OR logico
    - Varredura completa em todos os logs do sistema (Application, System, Security, etc.)
    - Filtragem precisa por intervalo de datas com timezone UTC
    - Processamento de logs arquivados e rotacionados automaticamente
    - Exportacao Excel multi-aba com agrupamento por Event ID
    - Extracao completa de propriedades XML dos eventos
    - Otimizacao de performance com skip de arquivos antigos
    - Tratamento robusto de arquivos corrompidos ou em uso
    
    Casos de uso tipicos:
    - Investigacao de incidentes de seguranca
    - Auditoria de logons e acessos 
    - Analise de falhas de sistema
    - Monitoramento de atividades suspeitas
    - Compliance e relatoria regulatoria

.PARAMETER None
    Script interativo - solicita Event IDs e intervalo de datas durante execucao

.EXAMPLE
    .\Analyze-WindowsEventLogs.ps1
    # Event IDs: 4624,4625,4648
    # Data inicial: 2024-12-01  
    # Data final: 2024-12-31
    # Resultado: Analise completa de logons, falhas e elevacao de privilegios

.EXAMPLE
    .\Analyze-WindowsEventLogs.ps1
    # Event IDs: 1074,1076,6005,6006
    # Data inicial: 2024-11-15
    # Data final: 2024-11-30
    # Resultado: Historico completo de shutdowns, reboots e inicializacoes do sistema

.EXAMPLE
    # Investigacao de incidente de seguranca
    .\Analyze-WindowsEventLogs.ps1
    # Event IDs: 4720,4726,4738,4740
    # Data inicial: 2024-10-01
    # Data final: 2024-10-07  
    # Resultado: Analise de criacao, exclusao, alteracao e bloqueio de contas

.INPUTS
    String - Lista de Event IDs separados por virgula (ex: 4624,4625,4648)
    DateTime - Data inicial da busca no formato AAAA-MM-DD
    DateTime - Data final da busca no formato AAAA-MM-DD

.OUTPUTS
    - Arquivo Excel (.xlsx) em C:\Temp\EventLog_Exports\
    - Abas separadas por Event ID para analise organizada
    - Aba consolidada "Todos os Eventos" com visao geral
    - Console: Progresso detalhado e estatisticas de coleta

.NOTES
    Autor         : Andre Kittler
    Versao        : 7.0
    Compatibilidade: PowerShell 5.1+, Windows Server/Client
    
    Requisitos obrigatorios:
    - Privilegios de Administrador local
    - Modulo ImportExcel instalado (Install-Module ImportExcel)
    - Acesso de leitura aos logs de eventos do Windows
        
    Configuracoes tecnicas:
    - Busca recursiva em C:\Windows\System32\winevt\Logs\
    - Extracao completa de propriedades XML EventData
    - Skip automatico de arquivos fora do intervalo temporal
    
    Logs pesquisados:
    - Application.evtx (eventos de aplicacoes)
    - System.evtx (eventos do sistema)  
    - Security.evtx (eventos de seguranca)
    - Todos os logs adicionais (.evtx) encontrados
    - Logs arquivados e rotacionados historicos
    
    Performance e otimizacao:
    - Verificacao de LastWriteTime para skip de arquivos antigos
    - Tratamento silencioso de arquivos corrompidos ou bloqueados
    - Processamento em lote para reduzir overhead de memoria
    
    Limitacoes conhecidas:
    - Requer privilegios administrativos para logs de Security
    - Performance varia com quantidade de logs e periodo pesquisado
    - Arquivos em uso podem ser ignorados

.LINK
    https://docs.microsoft.com/en-us/windows/win32/wes/consuming-events

.LINK
    https://docs.microsoft.com/en-us/windows/security/threat-protection/auditing/advanced-security-auditing
#>

# --- Verificação de Privilégios de Administrador ---
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "ERRO: Este script precisa ser executado com privilégios de Administrador."
    Write-Warning "Por favor, feche esta janela, clique com o botão direito no PowerShell e selecione 'Executar como Administrador'."
    if ($Host.Name -eq "ConsoleHost") {
        Read-Host "Pressione Enter para sair."
    }
    return # Encerra o script
}

# --- Início da Configuração e Entrada do Usuário ---

# Tenta importar o módulo ImportExcel
try {
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    Write-Error "ERRO: O módulo 'ImportExcel' não foi encontrado ou não pôde ser carregado."
    Write-Error "Execute 'Install-Module -Name ImportExcel -Scope CurrentUser' e tente novamente."
    if ($Host.Name -eq "ConsoleHost") {
        Read-Host "Pressione Enter para sair."
    }
    return
}

# Define o diretório para salvar os arquivos exportados
$exportDirectory = "C:\Temp\EventLog_Exports"
if (-NOT (Test-Path -Path $exportDirectory)) {
    New-Item -ItemType Directory -Path $exportDirectory | Out-Null
    Write-Host "Diretório de exportação criado em '$exportDirectory'" -ForegroundColor Cyan
}

# --- NOVA LÓGICA PARA MÚLTIPLOS EVENT IDs ---
$eventIds = $null
while ($null -eq $eventIds) {
    try {
        $input = Read-Host "Digite o(s) Event ID(s) que deseja buscar, separados por vírgula (ex: 4624,4625,4648)"
        
        # Remove espaços e divide por vírgula
        $eventIdStrings = $input -replace '\s+', '' -split ','
        
        # Converte cada string para inteiro e valida
        $eventIds = @()
        foreach ($idString in $eventIdStrings) {
            if ([string]::IsNullOrWhiteSpace($idString)) {
                continue
            }
            $eventIds += [int]$idString
        }
        
        if ($eventIds.Count -eq 0) {
            throw "Nenhum Event ID válido foi fornecido"
        }
        
        Write-Host "Event IDs selecionados: $($eventIds -join ', ')" -ForegroundColor Green
        
    } catch { 
        Write-Warning "Entrada inválida. Por favor, digite apenas números separados por vírgula para os Event IDs."
        $eventIds = $null
    }
}

$inputStartDate = $null
$inputEndDate = $null
while ($null -eq $inputStartDate -or $null -eq $inputEndDate) {
    try {
        if ($null -eq $inputStartDate) { $input = Read-Host "Digite a DATA INICIAL da busca no formato AAAA-MM-DD"; $inputStartDate = [datetime]$input }
        if ($null -eq $inputEndDate) { $input = Read-Host "Digite a DATA FINAL da busca no formato AAAA-MM-DD"; $inputEndDate = [datetime]$input }
        if ($inputEndDate -lt $inputStartDate) { Write-Warning "A data final não pode ser anterior à data inicial."; $inputStartDate = $null; $inputEndDate = $null }
    } catch { Write-Warning "Formato de data inválido. Use AAAA-MM-DD."; $inputStartDate = $null; $inputEndDate = $null }
}

$startDateFormatted = $inputStartDate.Date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z')
$endDateObject = $inputEndDate.Date.AddDays(1).AddTicks(-1)
$endDateFormatted = $endDateObject.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z')

# --- NOVA XPATH QUERY PARA MÚLTIPLOS EVENT IDs ---
# Constrói a condição OR para múltiplos Event IDs
$eventIdConditions = @()
foreach ($id in $eventIds) {
    $eventIdConditions += "EventID=$id"
}
$eventIdQuery = "(" + ($eventIdConditions -join " or ") + ")"

$xPathQuery = "*[System[$eventIdQuery and TimeCreated[@SystemTime>='$startDateFormatted' and @SystemTime<='$endDateFormatted']]]"

$logsPath = "C:\Windows\System32\winevt\Logs"
$allFoundEvents = @()

Write-Host "`nIniciando busca pelos Event IDs '$($eventIds -join ', ')' para o período de $($inputStartDate.ToString('yyyy-MM-dd')) até $($inputEndDate.ToString('yyyy-MM-dd'))" -ForegroundColor Green
Write-Warning "AVISO: A busca será realizada em TODOS os arquivos de log em '$logsPath'. Este processo pode ser demorado."

# --- Lógica de Busca Unificada ---

Write-Host "[1/2] Procurando em todos os arquivos de log (.evtx)..." -ForegroundColor Cyan
$allLogFiles = Get-ChildItem -Path $logsPath -Filter "*.evtx" -File -Recurse -ErrorAction SilentlyContinue

if ($null -ne $allLogFiles) {
    Write-Host "    -> Encontrados $($allLogFiles.Count) arquivos de log (.evtx) para verificar."
    
    foreach ($file in $allLogFiles) {
        # Performance: Pula arquivos que estão completamente fora do intervalo de datas
        if ($file.LastWriteTime -lt $inputStartDate.Date) {
            continue
        }

        Write-Host "    Verificando arquivo: $($file.Name)" -ForegroundColor Gray
        try {
            $archivedEvents = Get-WinEvent -Path $file.FullName -FilterXPath $xPathQuery -ErrorAction Stop
            if ($null -ne $archivedEvents) {
                $allFoundEvents += $archivedEvents
                Write-Host "        -> Encontrados $($archivedEvents.Count) eventos neste arquivo." -ForegroundColor White
            }
        }
        catch {
            # Ignora arquivos que não podem ser processados (em uso, corrompidos, etc)
        }
    }
}
else {
    Write-Host "    -> Nenhum arquivo de log (.evtx) encontrado no diretório especificado." -ForegroundColor Gray
}

# --- Processamento e Exportação dos Resultados ---

Write-Host "[2/2] Processando e exportando resultados..." -ForegroundColor Cyan

if ($allFoundEvents.Count -gt 0) {
    Write-Host "Total de $($allFoundEvents.Count) eventos brutos encontrados. Extraindo detalhes..." -ForegroundColor Green
    
    $detailedResults = @()

    foreach ($event in $allFoundEvents) {
        $xml = [xml]$event.ToXml()
        $properties = [ordered]@{
            LogName      = $event.LogName
            TimeCreated  = $event.TimeCreated
            EventID      = $event.Id
            ProviderName = $event.ProviderName
            ComputerName = $event.MachineName
            Level        = $event.LevelDisplayName
            UserID       = $event.UserId
        }
        
        $xml.Event.EventData.Data | ForEach-Object {
            $properties[$_.Name] = $_.'#text'
        }
        
        $detailedResults += New-Object -TypeName PSObject -Property $properties
    }

    # --- NOVO NOME DE ARQUIVO PARA MÚLTIPLOS EVENT IDs ---
    $eventIdsList = $eventIds -join '-'
    $fileName = "EventLog_IDs-$($eventIdsList)_$($inputStartDate.ToString('yyyyMMdd'))-a-$($inputEndDate.ToString('yyyyMMdd')).xlsx"
    $outputPath = Join-Path -Path $exportDirectory -ChildPath $fileName

    Write-Host "Exportando $($detailedResults.Count) eventos detalhados para '$outputPath'..." -ForegroundColor Cyan
    
    # Agrupa os resultados por Event ID para criar abas separadas
    $groupedResults = $detailedResults | Group-Object -Property EventID
    
    # Se houver múltiplos Event IDs, cria uma aba para cada um
    if ($groupedResults.Count -gt 1) {
        foreach ($group in $groupedResults) {
            $worksheetName = "Event ID $($group.Name)"
            $group.Group | Export-Excel -Path $outputPath -AutoSize -TableName "EventData_$($group.Name)" -TableStyle Medium9 -WorksheetName $worksheetName
        }
        
        # Cria também uma aba com todos os resultados consolidados
        $detailedResults | Export-Excel -Path $outputPath -AutoSize -TableName "AllEvents" -TableStyle Medium9 -WorksheetName "Todos os Eventos"
    }
    else {
        # Se houver apenas um Event ID, usa o formato original
        $detailedResults | Export-Excel -Path $outputPath -AutoSize -TableName "EventData" -TableStyle Medium9 -WorksheetName "Event ID $($eventIds[0])" -ClearSheet
    }
    
    Write-Host "`nExportação concluída com sucesso!" -ForegroundColor Green
    Write-Host "Arquivo salvo em: $outputPath"
    
    # Exibe resumo por Event ID
    Write-Host "`nResumo dos eventos encontrados:" -ForegroundColor Cyan
    foreach ($group in $groupedResults) {
        Write-Host "  Event ID $($group.Name): $($group.Count) eventos" -ForegroundColor White
    }
}
else {
    Write-Host "`nBusca finalizada. Nenhum evento com os IDs '$($eventIds -join ', ')' foi encontrado no período especificado." -ForegroundColor Yellow
}

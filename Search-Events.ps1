<#
.SYNOPSIS
    Busca múltiplos Event IDs em um intervalo de datas em TODOS os logs de eventos e exporta os detalhes para um arquivo Excel.

.DESCRIPTION
    Este script solicita um ou mais IDs de Evento (separados por vírgula) e um intervalo de datas. 
    Ele pesquisa em TODOS os arquivos de log .evtx (Application, System, Security, etc.) e, 
    se encontrar resultados, extrai todas as propriedades de cada evento e salva em um arquivo .xlsx detalhado.
    Versão 7.0 permite busca por múltiplos Event IDs simultaneamente.

.NOTES
    Autor: Gemini (Assistente do Google) - Modificado
    Versão: 7.0
    REQUER EXECUÇÃO COMO ADMINISTRADOR e o módulo ImportExcel instalado.
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

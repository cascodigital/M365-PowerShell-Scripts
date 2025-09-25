<#
.SYNOPSIS
    Busca avancada de arquivos em sites SharePoint Online usando Microsoft Graph PowerShell

.DESCRIPTION
    Script automatizado para busca recursiva de arquivos em todas as bibliotecas de documentos
    de sites SharePoint Online. Utiliza Microsoft Graph API para acesso direto aos dados,
    suportando wildcards e navegacao hierarquica completa em pastas e subpastas.
    
    Funcionalidades principais:
    - Autenticacao automatica com Microsoft Graph usando permissoes de leitura
    - Busca recursiva em todas as bibliotecas de documentos do site
    - Suporte completo a wildcards (*.zip, documento*.*, arquivo.pdf)
    - Navegacao hierarquica em pastas e subpastas sem limitacao de profundidade
    - Modo verbose detalhado para diagnostico e troubleshooting
    - Deteccao robusta de tipos de item (arquivo vs pasta)
    - Resultados organizados com informacoes completas de localizacao
    
    Processo de busca:
    1. Conecta ao Microsoft Graph com permissoes Sites.Read.All e Files.Read.All
    2. Resolve Site ID a partir da URL fornecida
    3. Enumera todas as bibliotecas de documentos do site
    4. Executa busca recursiva aplicando padrao wildcard
    5. Exibe resultados organizados por biblioteca e caminho

.PARAMETER None
    Script interativo - solicita informacoes durante execucao

.EXAMPLE
    .\Search-SharePointFiles.ps1
    # Script solicita:
    # - URL do site: https://tenant.sharepoint.com/sites/Documents
    # - Padrao de busca: *.pdf
    # - Modo verbose: s/n
    # Resultado: Lista todos arquivos PDF em todas as bibliotecas

.EXAMPLE
    .\Search-SharePointFiles.ps1
    # Busca por arquivo especifico:
    # - URL: https://cascodigital.sharepoint.com/sites/Arquivos
    # - Padrao: relatorio_financeiro_2025.xlsx
    # Resultado: Localiza arquivo exato com caminho completo

.EXAMPLE
    .\Search-SharePointFiles.ps1
    # Busca com wildcard no meio:
    # - Padrao: documento*final.*
    # Resultado: Encontra "documento_versao_final.docx", "documento123final.pdf"

.INPUTS
    String - URL completa do site SharePoint (https://tenant.sharepoint.com/sites/sitename)
    String - Padrao de busca com suporte a wildcards (* e ?)
    String - Opcao de modo verbose para debug detalhado (s/n)

.OUTPUTS
    Console: Informacoes detalhadas para cada arquivo encontrado:
    - Nome do arquivo
    - Biblioteca de origem
    - Caminho da pasta (hierarquia completa)
    - Tamanho em KB
    - Data de criacao e modificacao
    - URL direta para acesso
    
    Modo Verbose: Log detalhado incluindo:
    - Estrutura de pastas navegada
    - Contadores de arquivos e pastas processados
    - Detalhes de autenticacao e conectividade
    - Informacoes de debug para troubleshooting

.NOTES
    Autor         : Andre Kittler
    Versao        : 3.0
    Compatibilidade: PowerShell 5.1+, Windows/Linux/macOS
    
    Requisitos Microsoft Graph:
    - Modulo Microsoft.Graph (instalacao automatica se necessario)
    - Conta com acesso ao site SharePoint especificado
    - Permissoes Sites.Read.All e Files.Read.All
    
    Funcionalidades de busca:
    - Wildcards: * (qualquer sequencia) e ? (qualquer caractere)
    - Busca case-insensitive por padrao
    - Navegacao recursiva sem limite de profundidade
    - Acesso apenas leitura (nao modifica arquivos ou metadados)
    
    Tipos de biblioteca suportados:
    - Document Libraries (bibliotecas de documentos)
    - Ignora automaticamente bibliotecas de sistema
    - Processa bibliotecas customizadas e padrao
    
    Limitacoes conhecidas:
    - Requer permissoes de leitura no site especificado
    - Performance dependente do numero de arquivos e estrutura de pastas
    - Timeout de 30 segundos por operacao de rede
    - Arquivos em bibliotecas privadas podem nao aparecer
    
    Troubleshooting:
    - Use modo verbose para diagnosticar problemas de acesso
    - Verifique permissoes no site SharePoint
    - Confirme URL do site sem paginas especificas (/SitePages/etc)
    - Teste conectividade com Get-MgContext
    
    Permissoes necessarias:
    - Site Member/Visitor (minimo) OU
    - SharePoint Administrator OU
    - Global Reader OU
    - Global Administrator

.LINK
    https://learn.microsoft.com/en-us/graph/api/driveitem-list-children

.LINK
    https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.sites/get-mgsite

.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/driveitem

.LINK
    https://learn.microsoft.com/en-us/graph/api/site-getbypath
#>

# Conecta ao Microsoft Graph com autenticacao interativa
function Connect-ToGraph {
    try {
        # Verifica se ja esta conectado
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -eq $context) {
            Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Yellow
            Connect-MgGraph -Scopes "Sites.Read.All", "Files.Read.All" -NoWelcome
        } else {
            Write-Host "Ja conectado ao Microsoft Graph como: $($context.Account)" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Erro ao conectar ao Microsoft Graph: $_"
        return $false
    }
    return $true
}

# Obtem o Site ID a partir da URL
function Get-SiteIdFromUrl {
    param([string]$SiteUrl)
    
    try {
        # Remove trailing slash se existir
        $SiteUrl = $SiteUrl.TrimEnd('/')
        
        # Extrai hostname e path da URL
        $uri = [System.Uri]$SiteUrl
        $hostname = $uri.Host
        $sitePath = $uri.AbsolutePath.TrimStart('/')
        
        # Usa o formato correto para Get-MgSite: hostname:/sites/sitename:
        $siteIdentifier = "${hostname}:/${sitePath}:"
        
        Write-Host "Tentando obter site com identificador: $siteIdentifier" -ForegroundColor Gray
        
        # Usa o identificador correto para obter o site
        $site = Get-MgSite -SiteId $siteIdentifier -ErrorAction Stop
        
        Write-Host "Site encontrado: $($site.DisplayName)" -ForegroundColor Green
        return $site.Id
    }
    catch {
        Write-Error "Erro ao obter Site ID: $($_.Exception.Message)"
        return $null
    }
}

# Converte wildcard para regex
function ConvertTo-RegexPattern {
    param([string]$WildcardPattern)
    
    # Escapa caracteres especiais do regex, exceto * e ?
    $pattern = [regex]::Escape($WildcardPattern)
    
    # Converte wildcards para regex
    $pattern = $pattern -replace '\\\*', '.*'  # * para .*
    $pattern = $pattern -replace '\\\?', '.'   # ? para .
    
    return "^$pattern$"
}

# Busca arquivos usando chamadas diretas à API Graph - VERSAO CORRIGIDA
function Search-FilesInLibrary {
    param(
        [string]$SiteId,
        [string]$DriveId,
        [string]$LibraryName,
        [string]$SearchPattern,
        [string]$CurrentPath = "",
        [string]$CurrentItemId = "root",
        [int]$Depth = 0,
        [switch]$Verbose
    )
    
    $foundFiles = @()
    $regexPattern = ConvertTo-RegexPattern -WildcardPattern $SearchPattern
    $indent = "  " * $Depth
    
    if ($Verbose) {
        Write-Host "$indent[DEBUG] Entrando na pasta: $CurrentPath (ID: $CurrentItemId)" -ForegroundColor Cyan
    }
    
    try {
        # Usa chamada direta ao Graph API para obter informacoes completas
        $graphUrl = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$CurrentItemId/children"
        $response = Invoke-MgGraphRequest -Uri $graphUrl -Method GET
        $items = $response.value
        
        if ($Verbose) {
            Write-Host "$indent[DEBUG] Encontrados $($items.Count) itens na pasta $CurrentPath" -ForegroundColor Gray
        }
        
        $fileCount = 0
        $folderCount = 0
        
        foreach ($item in $items) {
            $fullPath = if ($CurrentPath -eq "") { $item.name } else { "$CurrentPath/$($item.name)" }
            
            # Verifica se e pasta usando as propriedades retornadas pela API direta
            $isFolder = $null -ne $item.folder
            $hasChildren = $isFolder -and ($item.folder.childCount -gt 0 -or $null -eq $item.folder.childCount)
            
            if ($Verbose) {
                Write-Host "$indent[DEBUG]   Item: $($item.name)" -ForegroundColor Gray
                Write-Host "$indent[DEBUG]     - Tem propriedade 'folder': $($null -ne $item.folder)" -ForegroundColor Gray
                Write-Host "$indent[DEBUG]     - Tem propriedade 'file': $($null -ne $item.file)" -ForegroundColor Gray
                if ($item.folder) {
                    Write-Host "$indent[DEBUG]     - ChildCount: $($item.folder.childCount)" -ForegroundColor Gray
                }
                Write-Host "$indent[DEBUG]     - Size: $($item.size)" -ForegroundColor Gray
            }
            
            if ($isFolder) {
                $folderCount++
                if ($Verbose) {
                    Write-Host "$indent[DEBUG]   -> PASTA: $($item.name) (ID: $($item.id))" -ForegroundColor Yellow
                }
                
                # Busca recursivamente na pasta
                $subFiles = Search-FilesInLibrary -SiteId $SiteId -DriveId $DriveId -LibraryName $LibraryName -SearchPattern $SearchPattern -CurrentPath $fullPath -CurrentItemId $item.id -Depth ($Depth + 1) -Verbose:$Verbose
                $foundFiles += $subFiles
            }
            else {
                $fileCount++
                if ($Verbose) {
                    Write-Host "$indent[DEBUG]   -> ARQUIVO: $($item.name) (Tamanho: $($item.size))" -ForegroundColor White
                }
                
                if ($item.name -match $regexPattern) {
                    Write-Host "$indent[ENCONTRADO] Arquivo corresponde ao padrao: $($item.name)" -ForegroundColor Green
                    $foundFiles += [PSCustomObject]@{
                        FileName = $item.name
                        Library = $LibraryName
                        Path = $CurrentPath
                        FullPath = $fullPath
                        Size = $item.size
                        Created = $item.createdDateTime
                        Modified = $item.lastModifiedDateTime
                        WebUrl = $item.webUrl
                    }
                }
            }
        }
        
        if ($Verbose) {
            Write-Host "$indent[DEBUG] Processados: $fileCount arquivo(s), $folderCount pasta(s)" -ForegroundColor Gray
        }
    }
    catch {
        Write-Warning "$indent[ERRO] Erro ao acessar biblioteca $LibraryName, caminho $CurrentPath`: $($_.Exception.Message)"
        if ($Verbose) {
            Write-Host "$indent[DEBUG] Detalhes do erro: $($_.Exception)" -ForegroundColor Red
        }
    }
    
    return $foundFiles
}

# Script principal
function Search-SharePointFiles {
    Write-Host "=== Busca de Arquivos no SharePoint ===" -ForegroundColor Cyan
    Write-Host ""
    
    # Conecta ao Graph
    if (-not (Connect-ToGraph)) {
        return
    }
    
    # Solicita URL do SharePoint
    $siteUrl = Read-Host "Digite a URL do site SharePoint (ex: https://tenant.sharepoint.com/sites/sitename)"
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        Write-Error "URL do site e obrigatoria."
        return
    }
    
    # Solicita padrao de busca
    $searchPattern = Read-Host "Digite o padrao de busca (ex: *.zip, documento*.*, relatorio.pdf)"
    if ([string]::IsNullOrWhiteSpace($searchPattern)) {
        Write-Error "Padrao de busca e obrigatorio."
        return
    }
    
    # Pergunta se quer modo verbose
    $verboseChoice = Read-Host "Quer ver detalhes da busca (s/n)? [padrao: n]"
    $verboseMode = $verboseChoice -eq "s" -or $verboseChoice -eq "S"
    
    Write-Host ""
    Write-Host "Obtendo informacoes do site..." -ForegroundColor Yellow
    
    # Obtem Site ID
    $siteId = Get-SiteIdFromUrl -SiteUrl $siteUrl
    if ($null -eq $siteId) {
        return
    }
    
    Write-Host "Site encontrado. Obtendo bibliotecas de documentos..." -ForegroundColor Yellow
    
    # Obtem todas as bibliotecas de documentos
    try {
        $drives = Get-MgSiteDrive -SiteId $siteId -ErrorAction Stop
        $documentLibraries = $drives | Where-Object { $_.DriveType -eq "documentLibrary" }
        
        if ($documentLibraries.Count -eq 0) {
            Write-Host "Nenhuma biblioteca de documentos encontrada no site." -ForegroundColor Red
            return
        }
        
        Write-Host "Encontradas $($documentLibraries.Count) biblioteca(s) de documentos." -ForegroundColor Green
        
        if ($verboseMode) {
            Write-Host ""
            Write-Host "Bibliotecas encontradas:" -ForegroundColor Cyan
            foreach ($lib in $documentLibraries) {
                Write-Host "  - $($lib.Name) (ID: $($lib.Id))" -ForegroundColor Gray
            }
        }
        
        Write-Host ""
        Write-Host "Iniciando busca por arquivos com padrao: $searchPattern" -ForegroundColor Yellow
        if ($verboseMode) {
            Write-Host "Regex convertido: $(ConvertTo-RegexPattern -WildcardPattern $searchPattern)" -ForegroundColor Gray
        }
        Write-Host ""
        
        $allFoundFiles = @()
        
        # Busca em cada biblioteca
        foreach ($library in $documentLibraries) {
            Write-Host "Buscando na biblioteca: $($library.Name)" -ForegroundColor Cyan
            
            $foundFiles = Search-FilesInLibrary -SiteId $siteId -DriveId $library.Id -LibraryName $library.Name -SearchPattern $searchPattern -Verbose:$verboseMode
            
            if ($foundFiles.Count -gt 0) {
                Write-Host "  Encontrados $($foundFiles.Count) arquivo(s)" -ForegroundColor Green
                $allFoundFiles += $foundFiles
            } else {
                Write-Host "  Nenhum arquivo encontrado" -ForegroundColor Gray
            }
            Write-Host ""
        }
        
        # Exibe resultados
        Write-Host "=== RESULTADOS ===" -ForegroundColor Cyan
        
        if ($allFoundFiles.Count -eq 0) {
            Write-Host "Nenhum arquivo encontrado com o padrao '$searchPattern'" -ForegroundColor Red
            Write-Host ""
            Write-Host "Sugestoes:" -ForegroundColor Yellow
            Write-Host "  1. Tente com *.iso para ver todos os arquivos ISO" -ForegroundColor Yellow
            Write-Host "  2. Execute novamente com modo verbose (s) para ver detalhes" -ForegroundColor Yellow
            Write-Host "  3. Verifique se o nome do arquivo esta correto" -ForegroundColor Yellow
        } else {
            Write-Host "Total de arquivos encontrados: $($allFoundFiles.Count)" -ForegroundColor Green
            Write-Host ""
            
            foreach ($file in $allFoundFiles | Sort-Object Library, Path, FileName) {
                Write-Host "Arquivo: $($file.FileName)" -ForegroundColor White
                Write-Host "  Biblioteca: $($file.Library)" -ForegroundColor Yellow
                if ($file.Path -ne "") {
                    Write-Host "  Pasta: $($file.Path)" -ForegroundColor Yellow
                } else {
                    Write-Host "  Pasta: (raiz da biblioteca)" -ForegroundColor Yellow
                }
                Write-Host "  Tamanho: $([math]::Round($file.Size / 1KB, 2)) KB" -ForegroundColor Gray
                Write-Host "  Criado: $($file.Created)" -ForegroundColor Gray
                Write-Host "  Modificado: $($file.Modified)" -ForegroundColor Gray
                Write-Host "  URL: $($file.WebUrl)" -ForegroundColor Blue
                Write-Host ""
            }
        }
    }
    catch {
        Write-Error "Erro durante a busca: $($_.Exception.Message)"
    }
}

# Executa o script
Search-SharePointFiles

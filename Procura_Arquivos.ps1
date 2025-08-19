<#
.SYNOPSIS
    Localiza arquivos com nome especifico no OneDrive de um usuario.

.DESCRIPTION
    Script interativo que solicita email do usuario e nome do arquivo para busca.
    Suporta wildcards para busca flexivel.

.NOTES
    Autor: Gemini & Claude
    Versao: 3.0 - Versao interativa
    Requer: Modulo Microsoft.Graph.PowerShell
    Permissoes de API necessarias: User.Read.All, Files.Read.All
#>

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "           LOCALIZADOR DE ARQUIVOS NO ONEDRIVE" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Solicita o email do usuario
do {
    $usuarioAlvoUPN = Read-Host "Digite o email do usuario para pesquisar"
    if ([string]::IsNullOrWhiteSpace($usuarioAlvoUPN)) {
        Write-Host "Email nao pode estar vazio. Tente novamente." -ForegroundColor Red
    }
} while ([string]::IsNullOrWhiteSpace($usuarioAlvoUPN))

Write-Host ""

# Solicita o nome do arquivo com explicacao dos filtros
Write-Host "OPCOES DE FILTRO PARA BUSCA DE ARQUIVO:" -ForegroundColor Yellow
Write-Host "  Exemplo 1: mikrotik          (encontra qualquer arquivo com 'mikrotik' no nome)"
Write-Host "  Exemplo 2: Mikrotik_V2.docx  (busca pelo nome exato)"
Write-Host "  Exemplo 3: Mikrotik_V2.*     (qualquer Mikrotik_V2 com qualquer extensao)"
Write-Host "  Exemplo 4: *v2*              (qualquer arquivo com 'v2' em qualquer parte do nome)"
Write-Host "  Exemplo 5: *.docx            (todos os arquivos .docx)"
Write-Host ""

do {
    $nomeArquivo = Read-Host "Digite o filtro de busca para o arquivo"
    if ([string]::IsNullOrWhiteSpace($nomeArquivo)) {
        Write-Host "Nome do arquivo nao pode estar vazio. Tente novamente." -ForegroundColor Red
    }
} while ([string]::IsNullOrWhiteSpace($nomeArquivo))

# Adiciona wildcards automaticamente se nao houver
if (-not ($nomeArquivo.Contains("*"))) {
    $nomeArquivoFiltro = "*$nomeArquivo*"
    Write-Host "Filtro aplicado: $nomeArquivoFiltro" -ForegroundColor Gray
} else {
    $nomeArquivoFiltro = $nomeArquivo
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "INICIANDO BUSCA..." -ForegroundColor Cyan
Write-Host "Usuario: $usuarioAlvoUPN" -ForegroundColor White
Write-Host "Filtro: $nomeArquivoFiltro" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

try {
    # Conexao
    $conexao = Get-MgContext
    if (-not $conexao) {
        Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Yellow
        $scopes = @("User.Read.All", "Files.Read.All")
        Connect-MgGraph -Scopes $scopes
    } else {
        Write-Host "Ja conectado ao Microsoft Graph." -ForegroundColor Green
    }
    
    Write-Host "Buscando pelo usuario $($usuarioAlvoUPN)..." -ForegroundColor Yellow
    $usuarioAlvo = Get-MgUser -UserId $usuarioAlvoUPN -ErrorAction Stop
    Write-Host "Usuario encontrado: $($usuarioAlvo.DisplayName)" -ForegroundColor Green
    Write-Host "ID: $($usuarioAlvo.Id)" -ForegroundColor Gray

    Write-Host ""
    Write-Host "Obtendo drives do usuario..." -ForegroundColor Yellow
    $drives = Get-MgUserDrive -UserId $usuarioAlvo.Id -ErrorAction Stop
    
    Write-Host "------------------------------------------------------------"
    Write-Host "DRIVES ENCONTRADOS:" -ForegroundColor Cyan
    foreach ($drive in $drives) {
        Write-Host "  $($drive.Name) - Tipo: $($drive.DriveType)"
    }
    Write-Host "------------------------------------------------------------"
    
    $todosArquivosEncontrados = @()
    $drivesPesquisados = 0
    $drivesComSucesso = 0
    
    # Pesquisa em todos os drives
    foreach ($drive in $drives) {
        $drivesPesquisados++
        Write-Host "[$drivesPesquisados/$($drives.Count)] Pesquisando no drive: $($drive.Name)" -ForegroundColor Yellow
        
        try {
            Write-Host "  Carregando arquivos..." -ForegroundColor Cyan
            $todosItens = Get-MgUserDriveItem -UserId $usuarioAlvo.Id -DriveId $drive.Id -All -Filter "file ne null" -ErrorAction Stop
            
            # Filtra arquivos usando o filtro fornecido pelo usuario
            $arquivosEncontrados = $todosItens | Where-Object { $_.Name -like $nomeArquivoFiltro }
            
            if ($arquivosEncontrados) {
                $drivesComSucesso++
                Write-Host "  ENCONTRADOS: $($arquivosEncontrados.Count) arquivo(s)!" -ForegroundColor Green
                
                foreach ($arquivo in $arquivosEncontrados) {
                    # Monta o caminho completo do arquivo
                    $caminhoCompleto = ""
                    if ($arquivo.ParentReference -and $arquivo.ParentReference.Path) {
                        $caminhoRelativo = $arquivo.ParentReference.Path -replace "^.+:", ""
                        $caminhoCompleto = "$caminhoRelativo/$($arquivo.Name)" -replace "^/", ""
                    } else {
                        $caminhoCompleto = $arquivo.Name
                    }
                    
                    $todosArquivosEncontrados += [PSCustomObject]@{
                        Nome = $arquivo.Name
                        CaminhoCompleto = $caminhoCompleto
                        Drive = $drive.Name
                        Tamanho = $arquivo.Size
                        UltimaModificacao = $arquivo.LastModifiedDateTime
                        ID = $arquivo.Id
                    }
                }
            } else {
                Write-Host "  Nenhum arquivo encontrado neste drive" -ForegroundColor Gray
            }
            
        } catch {
            Write-Host "  Drive nao acessivel (ignorando)" -ForegroundColor Gray
        }
    }
    
    # Resultado final
    Write-Host ""
    Write-Host "============================================================"
    if ($todosArquivosEncontrados.Count -gt 0) {
        Write-Host "BUSCA CONCLUIDA COM SUCESSO!" -ForegroundColor Green
        Write-Host "============================================================" -ForegroundColor Green
        Write-Host "TOTAL ENCONTRADO: $($todosArquivosEncontrados.Count) arquivo(s)" -ForegroundColor Green
        Write-Host "DRIVES PESQUISADOS: $drivesPesquisados"
        Write-Host "DRIVES COM RESULTADOS: $drivesComSucesso"
        
        # Verifica se ha duplicatas
        if ($todosArquivosEncontrados.Count -gt 1) {
            $nomesDuplicados = $todosArquivosEncontrados | Group-Object Nome | Where-Object Count -gt 1
            if ($nomesDuplicados) {
                Write-Host ""
                Write-Host "ALERTA: DUPLICATAS ENCONTRADAS!" -ForegroundColor Red
                foreach ($grupo in $nomesDuplicados) {
                    Write-Host "  Arquivo '$($grupo.Name)' aparece $($grupo.Count) vezes" -ForegroundColor Red
                }
            }
        }
        
        Write-Host "============================================================"
        
        foreach ($arquivo in $todosArquivosEncontrados) {
            Write-Host ""
            Write-Host "ARQUIVO: $($arquivo.Nome)" -ForegroundColor Yellow
            Write-Host "CAMINHO: $($arquivo.CaminhoCompleto)" -ForegroundColor White
            Write-Host "DRIVE: $($arquivo.Drive)" -ForegroundColor Cyan
            Write-Host "TAMANHO: $([math]::Round($arquivo.Tamanho / 1KB, 2)) KB"
            Write-Host "MODIFICADO: $($arquivo.UltimaModificacao)"
            Write-Host "ID: $($arquivo.ID)" -ForegroundColor Gray
        }
        
    } else {
        Write-Host "BUSCA CONCLUIDA - NENHUM ARQUIVO ENCONTRADO" -ForegroundColor Yellow
        Write-Host "============================================================" -ForegroundColor Yellow
        Write-Host "Nenhum arquivo encontrado com o filtro: $nomeArquivoFiltro"
        Write-Host "DRIVES PESQUISADOS: $drivesPesquisados"
        Write-Host ""
        Write-Host "SUGESTOES:" -ForegroundColor Cyan
        Write-Host "1. Verifique se o nome esta correto"
        Write-Host "2. Tente usar um filtro mais amplo (ex: *parte_do_nome*)"
        Write-Host "3. Verifique se tem permissao para acessar o arquivo"
        Write-Host "4. O arquivo pode estar em SharePoint ou Teams"
    }
    
    Write-Host ""
    Write-Host "============================================================"

}
catch {
    Write-Host ""
    Write-Host "ERRO durante a execucao:" -ForegroundColor Red
    Write-Host "$($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.Exception.Message -like "*not found*" -or $_.Exception.Message -like "*does not exist*") {
        Write-Host ""
        Write-Host "DICA: Verifique se o email esta correto e se o usuario existe." -ForegroundColor Yellow
    }
}
finally {
    # Disconnect-MgGraph
}

Write-Host ""
Write-Host "Pressione ENTER para fechar..." -ForegroundColor Gray
Read-Host

# --- SCRIPT COMPLETO DE REMOÇÃO DE E-MAILS ---

Write-Host "=== REMOÇÃO DE E-MAILS - Microsoft 365 ===" -ForegroundColor Cyan
Write-Host ""

# Verifica e conecta ao Security & Compliance Center
Write-Host "Verificando conexão com Security & Compliance Center..." -ForegroundColor White
try {
    Get-ComplianceSearch -ErrorAction Stop | Out-Null
    Write-Host "✓ Já conectado!" -ForegroundColor Green
} catch {
    Write-Host "Conectando ao Security & Compliance Center..." -ForegroundColor Yellow
    try {
        Connect-IPPSSession -ErrorAction Stop
        Write-Host "✓ Conectado com sucesso!" -ForegroundColor Green
    } catch {
        Write-Host "Erro ao conectar: $($_.Exception.Message)" -ForegroundColor Red
        exit
    }
}

Write-Host ""

# Solicita informações do usuário
$sender = Read-Host "Digite o endereço do remetente"
if ([string]::IsNullOrWhiteSpace($sender)) {
    Write-Host "Erro: Remetente não pode estar vazio!" -ForegroundColor Red
    exit
}

$subject = Read-Host "Digite o assunto do e-mail"
if ([string]::IsNullOrWhiteSpace($subject)) {
    Write-Host "Erro: Assunto não pode estar vazio!" -ForegroundColor Red
    exit
}

# Gera nome único para a busca usando timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$searchName = "RemocaoEmail_$timestamp"

# Monta a query de busca
$searchQuery = "(from:$sender) AND (subject:'$subject')"

Write-Host ""
Write-Host "=== RESUMO DA OPERAÇÃO ===" -ForegroundColor Yellow
Write-Host "Remetente: $sender"
Write-Host "Assunto: $subject"
Write-Host "Query: $searchQuery"
Write-Host "Nome da busca: $searchName"
Write-Host ""

# Confirmação final
$confirmacao = Read-Host "Deseja prosseguir com a remoção? (S/N)"
if ($confirmacao -ne 'S' -and $confirmacao -ne 's') {
    Write-Host "Operação cancelada pelo usuário." -ForegroundColor Yellow
    exit
}

Write-Host ""
Write-Host "=== INICIANDO PROCESSO DE REMOÇÃO ===" -ForegroundColor Green

try {
    # Cria e inicia a busca
    Write-Host "1. Criando busca de conformidade..." -ForegroundColor White
    New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery $searchQuery -ErrorAction Stop
    
    Write-Host "2. Iniciando busca..." -ForegroundColor White
    Start-ComplianceSearch -Identity $searchName -ErrorAction Stop

    # Aguarda conclusão da busca
    Write-Host "3. Aguardando conclusão da busca..." -ForegroundColor White
    $contador = 0
    do {
        $status = (Get-ComplianceSearch -Identity $searchName).Status
        $contador++
        Write-Host "   Status: $status (verificação $contador)" -ForegroundColor Gray
        
        if ($status -eq 'Failed') {
            throw "A busca falhou. Verifique os parâmetros informados."
        }
        
        Start-Sleep -Seconds 10
    } while ($status -ne 'Completed')

    # Verifica resultados da busca
    $searchResults = Get-ComplianceSearch -Identity $searchName
    $itemCount = $searchResults.Items
    
    Write-Host "4. Busca concluída!" -ForegroundColor Green
    Write-Host "   Itens encontrados: $itemCount" -ForegroundColor White
    
    if ($itemCount -eq 0) {
        Write-Host "Nenhum e-mail encontrado com os critérios especificados." -ForegroundColor Yellow
        # Remove a busca vazia
        Remove-ComplianceSearch -Identity $searchName -Confirm:$false
        exit
    }

    # Executa a remoção
    Write-Host "5. Executando remoção (SoftDelete)..." -ForegroundColor White
    New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -Confirm:$false -ErrorAction Stop

    Write-Host ""
    Write-Host "=== OPERAÇÃO CONCLUÍDA ===" -ForegroundColor Green
    Write-Host "✓ $itemCount e-mail(s) movido(s) para Itens Recuperáveis"
    Write-Host "✓ Nome da busca: $searchName"
    Write-Host "✓ Os e-mails foram removidos das caixas dos usuários"
    Write-Host ""
    Write-Host "Nota: Para recuperar os e-mails, use o nome da busca: $searchName" -ForegroundColor Cyan

} catch {
    Write-Host ""
    Write-Host "ERRO: $($_.Exception.Message)" -ForegroundColor Red
    
    # Tenta limpar busca em caso de erro
    try {
        if (Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue) {
            Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction SilentlyContinue
            Write-Host "Busca de teste removida." -ForegroundColor Gray
        }
    } catch {}
}

Write-Host ""
Write-Host "Pressione qualquer tecla para sair..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# --- FIM DO SCRIPT ---

<#
.SYNOPSIS
    Remocao automatizada de emails especificos em tenant Microsoft 365 via Security & Compliance Center

.DESCRIPTION
    Script corporativo para remocao controlada e auditavel de emails maliciosos, spam ou conteudo
    inadequado em escala organizacional. Utiliza Compliance Search e Purge Actions do Security &
    Compliance Center para localizacao precisa e remocao segura com SoftDelete, mantendo emails
    em Itens Recuperaveis para auditoria e recuperacao posterior se necessario.
    
    Funcionalidades principais:
    - Busca precisa baseada em remetente e assunto simultaneamente
    - Remocao em todas as caixas postais do tenant automaticamente
    - SoftDelete preserva emails em Itens Recuperaveis por 30 dias
    - Nomenclatura unica com timestamp para rastreabilidade
    - Validacao de resultados antes da execucao da remocao
    - Confirmacao interativa para prevencao de erros operacionais
    - Cleanup automatico de buscas vazias ou com falha
    
    Casos de uso tipicos:
    - Remocao de emails de phishing em surtos de seguranca
    - Eliminacao de spam massivo com bypass de filtros
    - Remocao de conteudo inadequado ou vazamentos de dados
    - Limpeza de emails com malware ou anexos perigosos
    - Acao de resposta a incidentes de seguranca cibernetica

.PARAMETER None
    Script interativo - solicita remetente e assunto durante execucao

.EXAMPLE
    .\Remove-MailboxEmails.ps1
    # Remetente: phishing@malicious-site.com
    # Assunto: Urgent Action Required - Verify Account
    # Resultado: Remove emails de phishing de todas as caixas do tenant

.EXAMPLE
    .\Remove-MailboxEmails.ps1
    # Remetente: noreply@spam-source.net
    # Assunto: Limited Time Offer - Click Now
    # Resultado: Eliminacao de campanha de spam massivo

.EXAMPLE
    # Remocao de email com dados sensiveis vazados
    .\Remove-MailboxEmails.ps1
    # Remetente: insider@empresa.com
    # Assunto: Confidential Customer Database Export
    # Resultado: Remocao imediata com preservacao para investigacao forense

.INPUTS
    String - Endereco email do remetente para filtro de busca
    String - Assunto exato do email para filtro de busca

.OUTPUTS
    - Console: Log detalhado do processo com contadores
    - Compliance Search: Criada com nome timestamp para auditoria  
    - Purge Action: Executada com SoftDelete em todas caixas
    - Itens Recuperaveis: Emails preservados por 30 dias

.NOTES
    Autor         : Andre Kittler
    Versao        : 1.0
    Compatibilidade: PowerShell 5.1+, Windows/Linux/macOS
    
    Requisitos obrigatorios:
    - Conexao ativa ao Security & Compliance Center (Connect-IPPSSession)
    - Privilegios administrativos especificos:
      * Compliance Administrator OU
      * Security Administrator OU
      * Organization Management OU
      * eDiscovery Manager com Purge permissions
    
    Permissoes de API necessarias:
    - Compliance Search (criar e executar buscas)
    - Purge Actions (remover conteudo das caixas)
    - Exchange Online acesso total para todas mailboxes
    
    Processo tecnico detalhado:
    1. New-ComplianceSearch com ExchangeLocation All
    2. Start-ComplianceSearch com query (from: AND subject:)
    3. Monitoramento de status ate Completed
    4. New-ComplianceSearchAction com PurgeType SoftDelete
    5. Cleanup automatico em caso de falha ou busca vazia
    
    Configuracoes de seguranca:
    - SoftDelete: Emails movidos para Itens Recuperaveis (nao deletados permanentemente)
    - Confirmacao obrigatoria antes da execucao
    - Nome unico com timestamp para rastreabilidade completa
    - Preservacao por 30 dias para auditoria e recuperacao
    
    Tipos de Purge disponiveis:
    - SoftDelete: Move para Itens Recuperaveis (recomendado)
    - HardDelete: Remove permanentemente (apenas casos extremos)
    
    Recuperacao de emails removidos:
    - Acessivel via Outlook > Itens Recuperaveis
    - Recuperacao por administrador usando mesmo nome da busca
    - Prazo: 30 dias apos SoftDelete (14 dias padrao + 16 extensao)
    
    Limitacoes e consideracoes:
    - Query exata: remetente E assunto devem coincidir
    - Propagacao pode levar alguns minutos em tenants grandes
    - Nao remove emails ja em Itens Recuperaveis do usuario
    - Log de auditoria mantem registro da acao por compliance

.LINK
    https://docs.microsoft.com/en-us/microsoft-365/compliance/search-for-and-delete-messages-in-your-organization

.LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/new-compliancesearch
#>


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

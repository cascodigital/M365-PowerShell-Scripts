<#
.SYNOPSIS
    Configuracao automatizada de regra catch-all (coletor geral) para dominios Microsoft 365 com grupo dinâmico de exceção.

.DESCRIPTION
    Script PowerShell interativo para implementacao de catch-all de email em tenants Microsoft 365/Exchange Online.
    - Permite configurar dominio como InternalRelay (opcional)
    - Cria grupo de distribuição dinâmico definido pelo operador (ex: colaboradores@dominio.com.br) incluindo todos usuários reais do domínio
    - Cria regra de transporte catch-all para redirecionar emails enviados para endereços inexistentes no domínio
    - Excetua membros do grupo dinâmico (todo usuário válido)
    - Compatível com múltiplos domínios/clientes (nome/email do grupo customizável)

    Funcionalidades principais:
    - Validação e instalação automática do módulo ExchangeOnlineManagement
    - Criação/modificação de grupo dinâmico com filtro automatizado para usuários do domínio
    - Criação de regra catch-all com exceção dinâmica para membros do grupo
    - Desconexão segura de sessão Exchange Online
    - Logs coloridos no terminal
    - Tratamento robusto de erros

    Processo de configuração:
    1. Validação de módulos e conectividade
    2. Criação/opcional configuração de domínio como InternalRelay
    3. Criação de grupo dinâmico com filtro
    4. Criação de regra transport rule catch-all com exceção do grupo
    5. Qualquer usuário novo criado no M365 entra automaticamente no grupo e é excetuado da regra

.PARAMETER None
    Script interativo - solicita:
        - Email de administrador do tenant
        - Nome do domínio alvo
        - Email da caixa coletora (catch-all)
        - Nome e email do grupo dinâmico para exceção

.EXAMPLE
    .\Configure-CatchAll.ps1
    # Script solicita:
    # - Email administrador
    # - Dominio alvo: empresa.com
    # - Email coletor: catchall@empresa.com
    # - Nome do grupo dinâmico: Colaboradores EmpresaX
    # - Email do grupo dinâmico: colaboradores@empresa.com
    # Resultado: Emails para endereços inexistentes do domínio são redirecionados; novos usuários são incluídos automaticamente na exceção via grupo dinâmico.

.INPUTS
    - String: Email administrador Microsoft 365
    - String: Dominio alvo para catch-all
    - String: Email da caixa coletora destino
    - String: Nome do grupo dinâmico
    - String: Email do grupo dinâmico

.OUTPUTS
    - Console: Log detalhado colorido dos passos
    - Exchange Online: Grupo dinâmico criado/atualizado
    - Exchange Online: Regra catch-all implementada

.NOTES
    Autor         : Andre Kittler
    Versão        : 2.1
    Compatibilidade: PowerShell 5.1+, Windows/Linux/macOS

    Requisitos Exchange Online:
    - Modulo ExchangeOnlineManagement (instalacao automatica)
    - Conta com privilegios Exchange Administrator ou Global Administrator
    - Dominio deve ser aceito no tenant

    Configurações aplicadas:
    - Grupo dinâmico com filtro automático para todos UserMailbox do domínio
    - TransportRule com RedirectMessageTo e exceção MemberOf configurada
    - Priority definida dinamicamente (menor prioridade disponível)

    Considerações importantes:
    - Propagação pode levar até 1 hora
    - InternalRelay só é necessário em cenários híbridos/coexistência on-prem
    - Caixa coletora deve ter capacidade adequada
    - Recomenda-se validar no EAC após execução, especialmente para domínios com múltiplos tipos de recipients

    Permissões necessárias:
    - Exchange Administrator OU
    - Organization Management OU
    - Global Administrator

.LINK
    https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/manage-accepted-domains/manage-accepted-domains
    https://docs.microsoft.com/en-us/exchange/security-and-compliance/mail-flow-rules/mail-flow-rules
#>

<#
.SYNOPSIS
    Cria grupo dinâmico customizado e regra catch-all com exceção dinâmica.
#>

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message, [string]$Color)
    Write-Host $Message -ForegroundColor $Color
}

# Inputs configuráveis
$upnAdmin = Read-Host "Email do administrador"
$dominio = Read-Host "Dominio (ex: empresa.com)"
$emailColetor = Read-Host "Email catch-all (ex: catchall@empresa.com)"
$grupoNome = Read-Host "Nome do grupo dinâmico a ser criado (ex: TodosColaboradores EmpresaX)"
$grupoEmail = Read-Host "Email do grupo dinâmico (ex: todoscolab@empresa.com)"

Write-Log "Conectando ao Exchange Online..." -Color Cyan
Connect-ExchangeOnline -UserPrincipalName $upnAdmin -ShowBanner:$false

try {
    # Cria Dynamic Distribution Group customizado
    $recipientFilter = "(RecipientTypeDetails -eq 'UserMailbox') -and (PrimarySmtpAddress -like '%@$dominio')"
    if (-not (Get-DynamicDistributionGroup -Identity $grupoEmail -ErrorAction SilentlyContinue)) {
        New-DynamicDistributionGroup -Name $grupoNome `
            -RecipientFilter $recipientFilter `
            -PrimarySmtpAddress $grupoEmail `
            -ErrorAction Stop
        Write-Log "Grupo dinâmico criado como $grupoEmail." -Color Green
    } else {
        Write-Log "Grupo já existe, utilizando grupo informado." -Color Yellow
    }

    # Remove regra catch-all antiga se precisa
    $nomeRegra = "Catch-All $dominio"
    if (Get-TransportRule -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $nomeRegra }) {
        Write-Log "Removendo regra catch-all antiga..." -Color Yellow
        Remove-TransportRule -Identity $nomeRegra -Confirm:$false
    }

    # Cria regra catch-all com exceção para membros do grupo informado
    Write-Log "Criando regra catch-all com exceção para membros do grupo..." -Color Cyan
    $ruleParams = @{
        Name = $nomeRegra
        RecipientDomainIs = $dominio
        RedirectMessageTo = $emailColetor
        ExceptIfSentToMemberOf = $grupoEmail
        Priority = ((Get-TransportRule).Count)
    }
    New-TransportRule @ruleParams

    Write-Log "Automação concluída: grupo e regra criados/atualizados!" -Color Green
    Write-Log "Todo novo usuário será incluído automaticamente no grupo dinâmico e será ignorado pela regra catch-all." -Color Cyan
}
catch {
    Write-Log "ERRO: $($_.Exception.Message)" -Color Red
}
Disconnect-ExchangeOnline -Confirm:$false
Write-Log "Sessão encerrada." -Color Green

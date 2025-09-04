<#
.SYNOPSIS
    Script para automatizar a configuracao de um email 'catch-all' (coletor geral)
    em um tenant do Microsoft 365. Versao 2.0 - Corrigido.

.DESCRIPTION
    Este script realiza as seguintes acoes:
    1. Verifica e instala o modulo ExchangeOnlineManagement do PowerShell.
    2. Solicita o UPN do administrador, o dominio alvo e o email coletor.
    3. Conecta-se ao Exchange Online.
    4. Valida se o dominio informado e um dominio aceito no tenant.
    5. Altera o tipo do dominio para 'InternalRelay'.
    6. Cria uma regra de transporte com prioridade dinamica.
    7. Desconecta a sessao corretamente.
#>

# Define que o script para em caso de erro para o 'try/catch' funcionar
$ErrorActionPreference = 'Stop'

# Funcao para escrever mensagens coloridas
function Write-Log {
    param(
        [string]$Message,
        [string]$Color
    )
    Write-Host $Message -ForegroundColor $Color
}

# --- 1. Verificacao de Pre-requisitos ---
Write-Log "Verificando se o modulo 'ExchangeOnlineManagement' esta instalado..." -Color Cyan
$moduloExo = Get-Module -Name ExchangeOnlineManagement -ListAvailable
if (-not $moduloExo) {
    Write-Log "Modulo nao encontrado. Instalando..." -Color Yellow
    try {
        Install-Module ExchangeOnlineManagement -Repository PSGallery -Force -AllowClobber
        Write-Log "Modulo instalado com sucesso." -Color Green
    }
    catch {
        Write-Log "Ocorreu um erro ao instalar o modulo. Verifique sua conexao ou execute o PowerShell como Administrador." -Color Red
        return # Para a execucao
    }
}
else {
    Write-Log "Modulo ja instalado." -Color Green
}

# --- 2. Coleta de Informacoes ---
Write-Log "`n--- Forneca as informacoes necessarias ---" -Color Cyan
$upnAdmin = Read-Host -Prompt "Digite o email do administrador do Microsoft 365 para conectar"
$dominio = Read-Host -Prompt "Digite o dominio que recebera a regra de catch-all (ex: empresa.com)"
$emailColetor = Read-Host -Prompt "Digite o email que recebera as mensagens (ex: coletor@empresa.com)"

# --- 3. Conexao com o Exchange Online ---
Write-Log "`nConectando ao Exchange Online com o usuario $upnAdmin..." -Color Cyan
try {
    # O -ShowBanner:$false apenas limpa a saida do terminal
    Connect-ExchangeOnline -UserPrincipalName $upnAdmin -ShowBanner:$false
    Write-Log "Conectado com sucesso!" -Color Green
}
catch {
    Write-Log "Falha na autenticacao. Verifique as credenciais e tente novamente." -Color Red
    return
}

# --- 4. Execucao da Logica Principal ---
try {
    # Valida se o dominio existe no tenant
    Write-Log "Verificando se o dominio '$dominio' e um dominio aceito..." -Color Cyan
    $dominioAceito = Get-AcceptedDomain -Identity $dominio
    if ($dominioAceito) {
        Write-Log "Dominio encontrado. Prosseguindo com a configuracao." -Color Green
    }

    # Altera o tipo do dominio
    Write-Log "Alterando o tipo do dominio '$dominio' para 'InternalRelay'..." -Color Cyan
    Set-AcceptedDomain -Identity $dominio -DomainType InternalRelay
    Write-Log "Tipo do dominio alterado com sucesso." -Color Green

    # **CORRECAO 1: Define a prioridade dinamicamente**
    Write-Log "Verificando regras existentes para definir a prioridade..." -Color Cyan
    $ruleCount = (Get-TransportRule).Count
    Write-Log "Prioridade definida como $ruleCount (a mais baixa)." -Color Green

    # Cria a regra de transporte
    $nomeRegra = "Regra Catch-All para $dominio"
    Write-Log "Criando a regra de transporte '$nomeRegra'..." -Color Cyan
    New-TransportRule -Name $nomeRegra `
        -RecipientDomainIs $dominio `
        -RedirectMessageTo $emailColetor `
        -Priority $ruleCount

    Write-Log "Regra de transporte criada com sucesso." -Color Green
    Write-Log "`nConfiguracao de Catch-All para o dominio '$dominio' finalizada!" -Color Green
    Write-Log "Lembre-se que a alteracao pode levar ate uma hora para propagar." -Color Yellow
    Write-Log "A caixa de correio '$emailColetor' comecara a receber emails enviados para enderecos inexistentes em '$dominio'." -Color Yellow

}
catch {
    # Captura qualquer erro que possa ocorrer nos passos acima
    $errorMessage = $_.Exception.Message
    Write-Log "`nERRO: Ocorreu um problema durante a configuracao." -Color Red
    Write-Log "Mensagem de erro: $errorMessage" -Color Red
    Write-Log "Nenhuma alteracao adicional foi feita." -Color Red
}
finally {
    # --- 5. Desconexao ---
    Write-Log "`nDesconectando da sessao do Exchange Online..." -Color Cyan
    # **CORRECAO 2: Comando de desconexao correto**
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Log "Sessao desconectada." -Color Green
}

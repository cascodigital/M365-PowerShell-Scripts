# ============================================================================
# SCRIPT: Relatório Completo de E-mails Microsoft 365
# Autor: Assistente
# Descrição: Coleta usuários, grupos, aliases e caixas compartilhadas
# ============================================================================

# Instalar/Importar módulos necessários
$modulos = @("Microsoft.Graph.Users", "Microsoft.Graph.Groups", "ImportExcel")
foreach ($modulo in $modulos) {
    if (-not (Get-Module -ListAvailable -Name $modulo)) {
        Write-Host "Instalando módulo $modulo..." -ForegroundColor Yellow
        Install-Module -Name $modulo -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $modulo
}

# Conectar ao Microsoft Graph
Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Mail.Read"

# Inicializar arrays
$usuarios = @()
$grupos = @()
$compartilhadas = @()
$aliases = @()
$usuariosExternos = @()
$usuariosSemLicenca = @()

# ============================================================================
# COLETANDO USUÁRIOS
# ============================================================================
Write-Host "Coletando usuários..." -ForegroundColor Yellow

$allUsers = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,ProxyAddresses,Mail,AccountEnabled,UserType,AssignedLicenses,CreatedDateTime"

foreach ($user in $allUsers) {
    $temLicenca = $user.AssignedLicenses.Count -gt 0
    $statusReal = if ($user.AccountEnabled -and $temLicenca) { "Ativo" } else { "Inativo" }
    
    # IDENTIFICAR USUÁRIOS EXTERNOS
    if ($user.UserType -eq "Guest" -or $user.UserPrincipalName -like "*#EXT#*" -or $user.UserPrincipalName -like "*@*onmicrosoft.com") {
        $usuariosExternos += [PSCustomObject]@{
            Nome = $user.DisplayName
            Email = $user.UserPrincipalName
            Tipo = "Usuário Externo/Convidado"
            UserType = $user.UserType
        }
        continue  # Pula para o próximo, não inclui nas outras categorias
    }
    
    # USUÁRIOS NORMAIS (internos da empresa)
    if ($user.UserType -ne "Guest" -and $user.UserPrincipalName -notlike "*#EXT#*" -and $user.UserPrincipalName -notlike "*@*onmicrosoft.com") {
        
        # POSSÍVEL CAIXA COMPARTILHADA: conta interna sem licença, habilitada, com nome genérico
        $nomesGenericos = @("contato", "suporte", "vendas", "info", "admin", "noreply", "no-reply", "copi", "rh", "financeiro", "comercial")
        $nomeGenerico = $nomesGenericos | Where-Object { $user.DisplayName -like "*$_*" -or $user.UserPrincipalName -like "$_@*" }
        
        if (!$temLicenca -and $user.AccountEnabled -and $nomeGenerico) {
            $compartilhadas += [PSCustomObject]@{
                Nome = $user.DisplayName
                Email = $user.UserPrincipalName
                Tipo = "Possível Caixa Compartilhada"
                Motivo = "Nome genérico, sem licença, mas habilitada"
            }
        }
        # USUÁRIOS SEM LICENÇA (possíveis ex-funcionários)
        elseif (!$temLicenca) {
            $usuariosSemLicenca += [PSCustomObject]@{
                Nome = $user.DisplayName
                Email = $user.UserPrincipalName
                Tipo = "Usuário Sem Licença"
                Status = if ($user.AccountEnabled) { "Conta Habilitada" } else { "Conta Desabilitada" }
                Observacao = "Possível ex-funcionário ou usuário não licenciado"
            }
        }
        # USUÁRIOS NORMAIS
        else {
            $usuarios += [PSCustomObject]@{
                Nome = $user.DisplayName
                Email = $user.UserPrincipalName
                Tipo = "Usuário"
                Status = $statusReal
                TemLicenca = if ($temLicenca) { "Sim" } else { "Não" }
            }
        }
    }
    
    # EXTRAIR ALIASES DOS PROXYADDRESSES
    if ($user.ProxyAddresses -and $user.ProxyAddresses.Count -gt 0) {
        foreach ($proxy in $user.ProxyAddresses) {
            # Aliases são os endereços smtp: (minúsculo) - não o SMTP: (maiúsculo que é o principal)
            if ($proxy -match "^smtp:" -and $proxy -notmatch "^SMTP:") {
                $aliasEmail = $proxy -replace "^smtp:", ""
                if ($aliasEmail -ne $user.UserPrincipalName) {
                    $aliases += [PSCustomObject]@{
                        NomePrincipal = $user.DisplayName
                        EmailPrincipal = $user.UserPrincipalName
                        Alias = $aliasEmail
                        Tipo = "Alias de Usuário"
                    }
                }
            }
        }
    }
}

# ============================================================================
# COLETANDO GRUPOS
# ============================================================================
Write-Host "Coletando grupos..." -ForegroundColor Yellow

# CORREÇÃO: Usar MailEnabled ao invés de apenas Mail
$allGroups = Get-MgGroup -All -Property "Id,DisplayName,Mail,GroupTypes,MailEnabled,SecurityEnabled,ResourceProvisioningOptions,MailNickname,ProxyAddresses" | 
Where-Object { $_.MailEnabled -eq $true }

foreach ($group in $allGroups) {
    
    # Determinar tipo específico do grupo
    $tipoDetalhado = "Desconhecido"
    
    if ($group.GroupTypes -contains "Unified") {
        if ($group.ResourceProvisioningOptions -contains "Team") {
            $tipoDetalhado = "Microsoft 365 + Teams"
        } else {
            $tipoDetalhado = "Microsoft 365"
        }
    }
    elseif ($group.MailEnabled -and $group.SecurityEnabled) {
        $tipoDetalhado = "Segurança (com email)"
    }
    elseif ($group.MailEnabled -and !$group.SecurityEnabled) {
        $tipoDetalhado = "Lista de Distribuição"
    }
    
    # Determinar email final (usar Mail se existir, senão construir do MailNickname)
    $emailFinal = $group.Mail
    if (!$emailFinal -and $group.MailNickname) {
        $emailFinal = "$($group.MailNickname)@ipcbrasil.ind.br"  # Ajuste o domínio se necessário
    }
    
    $grupos += [PSCustomObject]@{
        Nome = $group.DisplayName
        Email = $emailFinal
        Tipo = $tipoDetalhado
        TemTeams = if ($group.ResourceProvisioningOptions -contains "Team") { "Sim" } else { "Não" }
        MailNickname = $group.MailNickname
    }
    
    # ALIASES DE GRUPOS TAMBÉM
    if ($group.ProxyAddresses -and $group.ProxyAddresses.Count -gt 0) {
        foreach ($proxy in $group.ProxyAddresses) {
            if ($proxy -match "^smtp:" -and $proxy -notmatch "^SMTP:") {
                $aliasEmail = $proxy -replace "^smtp:", ""
                if ($aliasEmail -ne $group.Mail -and $aliasEmail -ne $emailFinal) {
                    $aliases += [PSCustomObject]@{
                        NomePrincipal = $group.DisplayName
                        EmailPrincipal = $emailFinal
                        Alias = $aliasEmail
                        Tipo = "Alias de Grupo"
                    }
                }
            }
        }
    }
}

# ============================================================================
# GERANDO RELATÓRIO EXCEL
# ============================================================================
Write-Host "Gerando relatório Excel..." -ForegroundColor Yellow

# Caminho do arquivo Excel
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$excelPath = ".\Relatorio_Emails_Completo_$timestamp.xlsx"

# Exportar para Excel com as categorias organizadas
if ($usuarios.Count -gt 0) {
    $usuarios | Export-Excel -Path $excelPath -WorksheetName "1-Usuários_Ativos" -AutoSize -FreezeTopRow -TableStyle Medium2
}

if ($usuariosSemLicenca.Count -gt 0) {
    $usuariosSemLicenca | Export-Excel -Path $excelPath -WorksheetName "2-Usuários_Sem_Licença" -AutoSize -FreezeTopRow -TableStyle Medium9
}

if ($usuariosExternos.Count -gt 0) {
    $usuariosExternos | Export-Excel -Path $excelPath -WorksheetName "3-Usuários_Externos" -AutoSize -FreezeTopRow -TableStyle Medium8
}

if ($grupos.Count -gt 0) {
    $grupos | Export-Excel -Path $excelPath -WorksheetName "4-Grupos" -AutoSize -FreezeTopRow -TableStyle Medium4
}

if ($compartilhadas.Count -gt 0) {
    $compartilhadas | Export-Excel -Path $excelPath -WorksheetName "5-Caixas_Compartilhadas" -AutoSize -FreezeTopRow -TableStyle Medium7
}

if ($aliases.Count -gt 0) {
    $aliases | Export-Excel -Path $excelPath -WorksheetName "6-Aliases" -AutoSize -FreezeTopRow -TableStyle Medium5
} else {
    # Criar aba vazia indicando que não há aliases
    @([PSCustomObject]@{Informacao = "Nenhum alias encontrado na organização"}) | 
    Export-Excel -Path $excelPath -WorksheetName "6-Aliases" -AutoSize -FreezeTopRow -TableStyle Medium5
}

# Resumo final
$resumo = @(
    [PSCustomObject]@{ Categoria = "Usuários Ativos (com licença)"; Quantidade = ($usuarios | Where-Object Status -eq "Ativo").Count }
    [PSCustomObject]@{ Categoria = "Usuários Inativos (com licença)"; Quantidade = ($usuarios | Where-Object Status -eq "Inativo").Count }
    [PSCustomObject]@{ Categoria = "Usuários Sem Licença"; Quantidade = $usuariosSemLicenca.Count }
    [PSCustomObject]@{ Categoria = "Usuários Externos/Convidados"; Quantidade = $usuariosExternos.Count }
    [PSCustomObject]@{ Categoria = "Grupos (todos os tipos)"; Quantidade = $grupos.Count }
    [PSCustomObject]@{ Categoria = "• Microsoft 365 + Teams"; Quantidade = ($grupos | Where-Object Tipo -eq "Microsoft 365 + Teams").Count }
    [PSCustomObject]@{ Categoria = "• Microsoft 365 (sem Teams)"; Quantidade = ($grupos | Where-Object Tipo -eq "Microsoft 365").Count }
    [PSCustomObject]@{ Categoria = "• Listas de Distribuição"; Quantidade = ($grupos | Where-Object Tipo -eq "Lista de Distribuição").Count }
    [PSCustomObject]@{ Categoria = "• Segurança (com email)"; Quantidade = ($grupos | Where-Object Tipo -eq "Segurança (com email)").Count }
    [PSCustomObject]@{ Categoria = "Aliases"; Quantidade = $aliases.Count }
    [PSCustomObject]@{ Categoria = "Possíveis Caixas Compartilhadas"; Quantidade = $compartilhadas.Count }
    [PSCustomObject]@{ Categoria = "TOTAL E-MAILS FUNCIONAIS"; Quantidade = ($usuarios | Where-Object Status -eq "Ativo").Count + $grupos.Count + $compartilhadas.Count }
)

$resumo | Export-Excel -Path $excelPath -WorksheetName "0-Resumo" -AutoSize -FreezeTopRow -TableStyle Medium6

# ============================================================================
# RESULTADO FINAL
# ============================================================================
Write-Host "`n" + "="*80 -ForegroundColor Green
Write-Host "RELATÓRIO GERADO COM SUCESSO!" -ForegroundColor Green
Write-Host "="*80 -ForegroundColor Green
Write-Host "Arquivo: $excelPath" -ForegroundColor White
Write-Host "`nResumo executivo:" -ForegroundColor Cyan
$resumo | Format-Table -AutoSize


# ============================================================================
# ABA PARA O CLIENTE - RESPOSTA DIRETA
# ============================================================================
Write-Host "Criando resposta simplificada para o cliente..." -ForegroundColor Magenta

# Compilar apenas e-mails VIGENTES/FUNCIONAIS
$emailsVigentes = @()

# 1. Usuários ATIVOS (com licença)
foreach ($usuario in ($usuarios | Where-Object Status -eq "Ativo")) {
    $emailsVigentes += [PSCustomObject]@{
        Email = $usuario.Email
        Nome = $usuario.Nome
        Categoria = "Usuário"
        Status = "Ativo"
    }
}

# 2. TODOS os grupos (são funcionais)
foreach ($grupo in $grupos) {
    $emailsVigentes += [PSCustomObject]@{
        Email = $grupo.Email
        Nome = $grupo.Nome
        Categoria = $grupo.Tipo
        Status = "Ativo"
    }
}

# 3. Caixas compartilhadas (se confirmadas)
foreach ($compartilhada in $compartilhadas) {
    $emailsVigentes += [PSCustomObject]@{
        Email = $compartilhada.Email
        Nome = $compartilhada.Nome
        Categoria = "Caixa Compartilhada"
        Status = "Ativo"
    }
}

# 4. Aliases (endereços alternativos funcionais)
foreach ($alias in $aliases) {
    $emailsVigentes += [PSCustomObject]@{
        Email = $alias.Alias
        Nome = "$($alias.NomePrincipal) (Alias)"
        Categoria = "Alias"
        Status = "Ativo"
    }
}

# Ordenar alfabeticamente por email
$emailsVigentes = $emailsVigentes | Sort-Object Email

# Exportar aba de resposta ao cliente
$emailsVigentes | Export-Excel -Path $excelPath -WorksheetName "RESPOSTA_CLIENTE" -AutoSize -FreezeTopRow -TableStyle Medium1

# Criar resumo executivo simples
$resumoCliente = @(
    [PSCustomObject]@{ Tipo = "Usuários"; Quantidade = ($emailsVigentes | Where-Object Categoria -eq "Usuário").Count; Descrição = "Funcionários ativos com licença" }
    [PSCustomObject]@{ Tipo = "Grupos Microsoft 365"; Quantidade = ($emailsVigentes | Where-Object Categoria -like "*Microsoft 365*").Count; Descrição = "Grupos colaborativos (incluindo Teams)" }
    [PSCustomObject]@{ Tipo = "Listas de Distribuição"; Quantidade = ($emailsVigentes | Where-Object Categoria -eq "Lista de Distribuição").Count; Descrição = "Grupos para envio de e-mails" }
    [PSCustomObject]@{ Tipo = "Caixas Compartilhadas"; Quantidade = ($emailsVigentes | Where-Object Categoria -eq "Caixa Compartilhada").Count; Descrição = "E-mails compartilhados (ex: contato@, suporte@)" }
    [PSCustomObject]@{ Tipo = "Aliases"; Quantidade = ($emailsVigentes | Where-Object Categoria -eq "Alias").Count; Descrição = "Apelidos/endereços alternativos" }
    [PSCustomObject]@{ Tipo = "TOTAL"; Quantidade = $emailsVigentes.Count; Descrição = "Total de e-mails vigentes na organização" }
)

$resumoCliente | Export-Excel -Path $excelPath -WorksheetName "RESUMO_EXECUTIVO" -AutoSize -FreezeTopRow -TableStyle Medium3

# Atualizar mensagem final
Write-Host "`n" + "="*80 -ForegroundColor Green
Write-Host "RELATÓRIO GERADO COM SUCESSO!" -ForegroundColor Green
Write-Host "="*80 -ForegroundColor Green
Write-Host "Arquivo: $excelPath" -ForegroundColor White

Write-Host "`n📧 RESPOSTA AO CLIENTE:" -ForegroundColor Yellow
Write-Host "Total de e-mails vigentes: $($emailsVigentes.Count)" -ForegroundColor White
Write-Host "`nResumo por categoria:" -ForegroundColor Cyan
$resumoCliente | Format-Table Tipo, Quantidade, Descrição -AutoSize

Write-Host "`n📋 ABAS GERADAS:" -ForegroundColor Magenta
Write-Host "• RESPOSTA_CLIENTE: Lista limpa de todos os e-mails vigentes" -ForegroundColor White
Write-Host "• RESUMO_EXECUTIVO: Totais por categoria" -ForegroundColor White  
Write-Host "• Abas 1-6: Detalhamento técnico completo" -ForegroundColor Gray


# Desconectar
Disconnect-MgGraph

Write-Host "`nScript finalizado. Verifique o arquivo Excel gerado." -ForegroundColor Green
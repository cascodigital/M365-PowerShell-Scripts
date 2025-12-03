# M365 PowerShell Scripts

Colecao de scripts PowerShell para administracao Microsoft 365, Exchange Online, OneDrive e Active Directory. Projetado para uso profissional com foco em praticidade e seguranca.

## 🎯 Categorias

### Seguranca e Auditoria
- **Ver_MfaComplianceReport.ps1** - Relatorio de conformidade MFA filtrando usuarios reais
- **Relacao_Confianca.ps1** - Audita relacoes de confianca AD e identifica problemas de trust
- **Buscar_Logon.ps1** - Analise forense de eventos 4624 (logon) em multiplos hosts

### Gestao de Usuarios
- **Alterar_Senhas_365.ps1** - Gera e aplica senhas seguras em massa com exportacao CSV
- **UsarAlias.ps1** - Gerencia aliases de usuarios e habilita SendFromAliasEnabled
- **Ver_Emails.ps1** - Compila todos enderecos do tenant (usuarios, grupos, aliases)

### OneDrive e Busca
- **Procura_Arquivos.ps1** - Busca arquivos em OneDrive por nome, extensao ou conteudo
- **Remover_Email.ps1** - Search & Purge para remover mensagens especificas

### Exchange Online
- **Configura-CatchAll.ps1** - Cria grupo dinamico e regra catch-all para emails nao entregues
- **Manage-SecurityDefaults-SMTP.ps1** - Ativa/desativa Security Defaults (MFA) + SMTP AUTH individual

### Monitoramento Local
- **Procura_Eventos.ps1** - Analisa logs Windows por Event IDs com exportacao Excel
- **monitor-ping.ps1** - Monitor de latencia ICMP com gravacao CSV em tempo real

### Sistema Windows
- **office_removal.ps1** - Remocao agressiva de instalacoes Office
- **Fix-KeyboardLayout.ps1** - Resolve troca automatica de layout PT-BR para EN-US permanentemente

## 📋 Pre-requisitos

- PowerShell 5.1+ (recomendado: PowerShell 7+)
- Modulos necessarios (instalados sob demanda):
  - `Microsoft.Graph`
  - `ExchangeOnlineManagement`
  - `ImportExcel`
  - `ActiveDirectory`
- Contas com permissoes adequadas (Global Admin/Exchange Admin/Password Admin)
- Executar como Administrador para tarefas locais

## 🚀 Uso Rapido

1. Clone o repositorio:
```powershell
git clone https://github.com/cascodigital/M365-PowerShell-Scripts.git
cd M365-PowerShell-Scripts/scripts
```

2. Execute o script desejado:
```powershell
PowerShell -ExecutionPolicy Bypass -File .\NomeDoScript.ps1
```

3. Siga os prompts interativos

## 📊 Exemplos de Uso

### Gerar relatorio MFA
```powershell
.\Ver_MfaComplianceReport.ps1
# Exporta CSV com estado MFA, metodos registrados e recomendacoes
```

### Buscar arquivos em OneDrive
```powershell
.\Procura_Arquivos.ps1
# Busca por nome/extensao e exporta resultados com caminho e proprietario
```

### Auditar logons do dominio
```powershell
.\Buscar_Logon.ps1
# Coleta eventos 4624 e agrega por usuario, origem e horario
```

### Corrigir layout de teclado
```powershell
.\Fix-KeyboardLayout.ps1
# Forca ABNT2 e desativa hotkeys de troca automatica
# Requer logoff/login para aplicar mudancas de registro
```

## ⚠️ Avisos de Seguranca

### Scripts Destrutivos
- **Alterar_Senhas_365.ps1**: Gera CSV com senhas em texto claro - armazene com seguranca
- **Remover_Email.ps1**: Operacao irreversivel - testar em ambiente controlado
- **office_removal.ps1**: Remocao agressiva - avisar usuarios antes de executar

### Permissoes
Scripts solicitam consentimento e escopos Graph/Exchange. **Confirme permissoes antes de executar em producao.**

## 📂 Estrutura

```
/scripts              Scripts PowerShell principais
gpo_logons.rar        GPO para auditoria de logons
LICENSE               MIT License
```

## 🛠️ Troubleshooting

### Modulos nao encontrados
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

### Erro de execucao
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Falha de conexao Exchange
Verifique MFA habilitado e use Modern Authentication

## 📄 Licenca

Este projeto esta licenciado sob a MIT License - veja o arquivo [LICENSE](LICENSE) para detalhes.

## 👤 Autor

Andre Kittler

## 🔗 Links Uteis

- [Microsoft Graph PowerShell](https://learn.microsoft.com/en-us/powershell/microsoftgraph/)
- [Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
- [Active Directory Module](https://learn.microsoft.com/en-us/powershell/module/activedirectory/)

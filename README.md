# M365-PowerShell-Scripts

Coleção de scripts PowerShell para administração Microsoft 365, Exchange Online, OneDrive, Active Directory e análise local de Windows. Projetado para uso profissional — scripts são interativos e exigem permissões apropriadas.

## Conteúdo (scripts)

- Ver_MfaComplianceReport.ps1 — Relatório de conformidade MFA (foco em usuários "reais").
- Ver_Emails.ps1 — Levantamento completo de endereços (usuários, grupos, compartilhadas, aliases).
- Alterar_Senhas_365.ps1 — Geração e aplicação em massa de senhas para um domínio M365.
- Procura_Arquivos.ps1 — Busca por arquivos em OneDrive (usuário específico ou todos do domínio).
- Procura_Sharepoint.ps1 — Busca recursiva de arquivos em sites SharePoint via Microsoft Graph.
- Procura_Eventos.ps1 — Analisador de logs Windows (múltiplos Event IDs, exporta para Excel).
- Buscar_Logon.ps1 — Coleta forense de logons (EventID 4624) em computadores do domínio.
- Remover_Email.ps1 — Remoção global (Search & Purge) de mensagens por remetente/assunto.
- Configura-CatchAll.ps1 — Cria grupo dinâmico e implementa regra catch‑all no Exchange Online.
- UsarAlias.ps1 — Gerencia aliases e habilita SendFromAliasEnabled no tenant.
- Relacao_Confianca.ps1 — Auditoria de relação de confiança entre máquinas e o AD.
- monitor-ping.ps1 — Monitor de latência (ICMP) com CSV em tempo real.
- office_removal.ps1 — Remoção agressiva de todas as instalações do Microsoft Office (destrutivo).

Arquivo adicional:
- gpo_logons.rar — backup/objeto GPO exigido por Buscar_Logon.ps1.

## Requisitos mínimos

- PowerShell 5.1+ (preferível PowerShell 7+ para alguns módulos).
- Módulos (instaláveis quando solicitados pelos scripts): Microsoft.Graph, ExchangeOnlineManagement, ImportExcel, ActiveDirectory.
- Contas com permissões adequadas (Global/Admin/Exchange/Password admin, dependendo do script).
- Executar como Administrador para tarefas locais/AD/Office removal.

## Uso rápido

1. Navegar até a pasta dos scripts:
   PowerShell: `Set-Location '<caminho>/m365-powershell-scripts/scripts'`
2. Executar o script desejado:
   `PowerShell -ExecutionPolicy Bypass -File .\NomeDoScript.ps1`
3. Seguir prompts interativos. Leia avisos antes de confirmar ações.

## Avisos de segurança e uso

- Alterar senhas em massa gera CSV com senhas em texto claro — armazene com segurança e apague após uso.
- Remoção de emails (Remover_Email.ps1) e remoção do Office (office_removal.ps1) são operações com impacto amplo e irreversível; use com extremo cuidado.
- Scripts que usam Graph/Exchange pedem consentimento/escopos. Confirme permissões antes de executar em produção.

## Estrutura do repositório

- /scripts — scripts PowerShell (principal).
- gpo_logons.rar — GPO para habilitar auditoria de logons (utilizado por Buscar_Logon.ps1).
- LICENSE — MIT.

## Contribuições e contato

Pull requests e issues são bem-vindos. Autor: Andre Kittler.

## Licença

MIT License.

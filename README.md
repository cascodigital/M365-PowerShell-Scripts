# M365-PowerShell-Scripts

Coleção de scripts PowerShell para administração Microsoft 365, Exchange Online, OneDrive, Active Directory e análise local de Windows. Projetado para uso profissional — scripts são interativos e exigem permissões apropriadas.

## Conteúdo (scripts) — descrições detalhadas

- Ver_MfaComplianceReport.ps1  
  Gera relatório de conformidade MFA focado em "usuários reais" (filtra contas de serviço, aplicações e contas excluídas). Entrada: período/opções de filtragem. Saída: CSV/Excel com estado MFA, métodos registrados e recomendações. Requer permissões de leitura em Azure AD/Graph.

- Ver_Emails.ps1  
  Varre o tenant para compilar todos os endereços associados (usuários, grupos, caixas compartilhadas e aliases). Produz lista deduplicada para auditoria ou migração. Útil para identificar destinatários ativos e aliases ocultos. Requer permissões de Exchange/Graph.

- Alterar_Senhas_365.ps1  
  Gera senhas seguras em massa e aplica a usuários de um domínio/OU selecionado. Gera CSV com usuário → nova senha (texto claro) para distribuição controlada. Atenção: operação disruptiva; usar conta com permissão de alteração de senha em massa.

- Procura_Arquivos.ps1  
  Busca arquivos em OneDrive para usuário específico ou para todos usuários do tenant, por nome, extensão ou conteúdo (quando disponível). Exporta resultados para CSV/Excel com caminho, proprietário e informações de compartilhamento. Requer Microsoft Graph/OneDrive scopes.

- Procura_Sharepoint.ps1  
  Busca recursiva em sites SharePoint (sites, bibliotecas, pastas) via Microsoft Graph. Filtragem por padrão, tipo de arquivo, data e proprietário. Gera relatório com URLs diretas e metadados. Necessita permissões de leitura em SharePoint/Graph.

- Procura_Eventos.ps1  
  Analisa logs locais do Windows buscando múltiplos Event IDs (ex.: autenticação, falhas, remoções). Agrega e exporta para Excel com contagens, gravidade e eventos relevantes. Executar como Administrador nas máquinas alvo ou via remoting.

- Buscar_Logon.ps1  
  Coleta eventos 4624 (logon) em computadores do domínio para análise forense de acesso. Suporta varredura em múltiplos hosts, agrega por usuário, origem e hora. Usa o arquivo gpo_logons.rar para instruções/objeto GPO que habilita auditoria, se necessário. Requer privilégios de leitura de eventos remotos e permissão AD.

- Remover_Email.ps1  
  Implementa busca e purge (Search & Purge) para remover mensagens específicas por remetente, assunto ou conteúdo. Suporte a escopos amplos (caixas, grupos). Operação destrutiva — testar em ambiente controlado e ter backups/auditoria. Requer permissões de compliance/exchange.

- Configura-CatchAll.ps1  
  Cria um grupo dinâmico e adiciona regra de transporte no Exchange Online para capturar emails não entregues (catch‑all). Inclui validações e instruções de rollback. Necessita permissões de administrador Exchange.

- UsarAlias.ps1  
  Gerencia aliases de usuários e habilita SendFromAliasEnabled no tenant (quando suportado). Permite mapear aliases existentes, adicionar/remover aliases em lote e validar envio a partir de aliases. Requer permissões de gestão de usuários/Exchange.

- Relacao_Confianca.ps1  
  Audita relações de confiança entre estações/servidores e o Active Directory, identifica máquinas com trust problems (timing, senha de máquina, SID). Gera relatório com recomendações de correção. Executar com conta de domínio com privilégios de leitura de AD.

- monitor-ping.ps1  
  Monitor simples de latência (ICMP) para uma lista de hosts; grava resultados em CSV em tempo real e plota resumo. Útil para monitoramento rápido de disponibilidade de recursos de rede. Executar com privilégios suficientes para ICMP e em rede que permita ping.

- office_removal.ps1  
  Remoção agressiva de instalações do Microsoft Office de máquinas locais (scripts de desinstalação e limpeza). Operação destrutiva — testar em máquinas de laboratório e avisar usuários. Requer execução como Administrador local e, para remoção em escala, privilégios de domínio.

Arquivo adicional:
- gpo_logons.rar — backup/objeto GPO com configuração de auditoria de logons. Usado por Buscar_Logon.ps1 quando for necessário habilitar auditoria em endpoints.

## Requisitos mínimos

- PowerShell 5.1+ (recomenda-se PowerShell 7+ para melhor compatibilidade com módulos).
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
- Remoção de emails (Remover_Email.ps1) e remoção do Office (office_removal.ps1) são operações com impacto amplo e irreversível; use com extrema cautela.
- Scripts que usam Graph/Exchange pedem consentimento/escopos. Confirme permissões antes de executar em produção.

## Estrutura do repositório

- /scripts — scripts PowerShell (principal).
- gpo_logons.rar — GPO para habilitar auditoria de logons (utilizado por Buscar_Logon.ps1).
- LICENSE — MIT.

## Contribuições e contato

Pull requests e issues são bem-vindos. Autor: Andre Kittler.

## Licença

MIT License.

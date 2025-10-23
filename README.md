# M365-PowerShell-Scripts

Colecao de scripts PowerShell para administracao Microsoft 365, Exchange Online, OneDrive, Active Directory e analise local de Windows. Projetado para uso profissional — scripts sao interativos e exigem permissoes apropriadas.

---

## 📋 Conteudo (scripts) — descricoes detalhadas

### **Ver_MfaComplianceReport.ps1**
Gera relatorio de conformidade MFA focado em "usuarios reais" (filtra contas de servico, aplicacoes e contas excluidas). Entrada: periodo/opcoes de filtragem. Saida: CSV/Excel com estado MFA, metodos registrados e recomendacoes.  
🔑 **Requer** permissoes de leitura em Azure AD/Graph.

### **Ver_Emails.ps1**
Varre o tenant para compilar todos os enderecos associados (usuarios, grupos, caixas compartilhadas e aliases). Produz lista deduplicada para auditoria ou migracao. Util para identificar destinatarios ativos e aliases ocultos.  
🔑 **Requer** permissoes de Exchange/Graph.

### **Alterar_Senhas_365.ps1**
Gera senhas seguras em massa e aplica a usuarios de um dominio/OU selecionado. Gera CSV com usuario → nova senha (texto claro) para distribuicao controlada.  
⚠️ **Atencao:** operacao disruptiva; usar conta com permissao de alteracao de senha em massa.

### **Procura_Arquivos.ps1**
Busca arquivos em OneDrive para usuario especifico ou para todos usuarios do tenant, por nome, extensao ou conteudo (quando disponivel). Exporta resultados para CSV/Excel com caminho, proprietario e informacoes de compartilhamento.  
🔑 **Requer** Microsoft Graph/OneDrive scopes.

### **Procura_Sharepoint.ps1**
Busca recursiva em sites SharePoint (sites, bibliotecas, pastas) via Microsoft Graph. Filtragem por padrao, tipo de arquivo, data e proprietario. Gera relatorio com URLs diretas e metadados.  
🔑 **Necessita** permissoes de leitura em SharePoint/Graph.

### **Procura_Eventos.ps1**
Analisa logs locais do Windows buscando multiplos Event IDs (ex.: autenticacao, falhas, remocoes). Agrega e exporta para Excel com contagens, gravidade e eventos relevantes.  
💻 **Executar** como Administrador nas maquinas alvo ou via remoting.

### **Buscar_Logon.ps1**
Coleta eventos 4624 (logon) em computadores do dominio para analise forense de acesso. Suporta varredura em multiplos hosts, agrega por usuario, origem e hora. Usa o arquivo `gpo_logons.rar` para instrucoes/objeto GPO que habilita auditoria, se necessario.  
🔑 **Requer** privilegios de leitura de eventos remotos e permissao AD.

### **Remover_Email.ps1**
Implementa busca e purge (Search & Purge) para remover mensagens especificas por remetente, assunto ou conteudo. Suporte a escopos amplos (caixas, grupos).  
🔥 **Operacao destrutiva** — testar em ambiente controlado e ter backups/auditoria.  
🔑 **Requer** permissoes de compliance/exchange.

### **Configura-CatchAll.ps1**
Cria um grupo dinamico e adiciona regra de transporte no Exchange Online para capturar emails nao entregues (catch‑all). Inclui validacoes e instrucoes de rollback.  
🔑 **Necessita** permissoes de administrador Exchange.

### **UsarAlias.ps1**
Gerencia aliases de usuarios e habilita SendFromAliasEnabled no tenant (quando suportado). Permite mapear aliases existentes, adicionar/remover aliases em lote e validar envio a partir de aliases.  
🔑 **Requer** permissoes de gestao de usuarios/Exchange.

### **Relacao_Confianca.ps1**
Audita relacoes de confianca entre estacoes/servidores e o Active Directory, identifica maquinas com trust problems (timing, senha de maquina, SID). Gera relatorio com recomendacoes de correcao.  
🔑 **Executar** com conta de dominio com privilegios de leitura de AD.

### **monitor-ping.ps1**
Monitor simples de latencia (ICMP) para uma lista de hosts; grava resultados em CSV em tempo real e plota resumo. Util para monitoramento rapido de disponibilidade de recursos de rede.  
💻 **Executar** com privilegios suficientes para ICMP e em rede que permita ping.

### **office_removal.ps1**
Remocao agressiva de instalacoes do Microsoft Office de maquinas locais (scripts de desinstalacao e limpeza).  
🔥 **Operacao destrutiva** — testar em maquinas de laboratorio e avisar usuarios.  
💻 **Requer** execucao como Administrador local e, para remocao em escala, privilegios de dominio.

### 📦 Arquivo adicional
- **gpo_logons.rar** — backup/objeto GPO com configuracao de auditoria de logons. Usado por `Buscar_Logon.ps1` quando for necessario habilitar auditoria em endpoints.

---

## 🛠️ Requisitos minimos

- **PowerShell 5.1+** (recomenda-se PowerShell 7+ para melhor compatibilidade com modulos)
- **Modulos** (instalaveis quando solicitados pelos scripts):
  - `Microsoft.Graph`
  - `ExchangeOnlineManagement`
  - `ImportExcel`
  - `ActiveDirectory`
- **Contas com permissoes adequadas** (Global/Admin/Exchange/Password admin, dependendo do script)
- **Executar como Administrador** para tarefas locais/AD/Office removal

---

## 🚀 Uso rapido

1. Navegar ate a pasta dos scripts:
   ```powershell
   Set-Location '<caminho>/m365-powershell-scripts/scripts'
   ```

2. Executar o script desejado:
   ```powershell
   PowerShell -ExecutionPolicy Bypass -File .\NomeDoScript.ps1
   ```

3. Seguir prompts interativos. **Leia avisos antes de confirmar acoes.**

---

## ⚠️ Avisos de seguranca e uso

- **Alterar senhas em massa** gera CSV com senhas em texto claro — armazene com seguranca e apague apos uso
- **Remocao de emails** (`Remover_Email.ps1`) e **remocao do Office** (`office_removal.ps1`) sao operacoes com impacto amplo e irreversivel; use com extrema cautela
- Scripts que usam Graph/Exchange pedem consentimento/escopos. **Confirme permissoes antes de executar em producao**

---

## 📂 Estrutura do repositorio

```
/scripts          — scripts PowerShell (principal)
gpo_logons.rar    — GPO para habilitar auditoria de logons
LICENSE           — MIT
```

---

## 🤝 Contribuicoes e contato

Pull requests e issues sao bem-vindos.  
**Autor:** Andre Kittler

---

## 📄 Licenca

MIT License

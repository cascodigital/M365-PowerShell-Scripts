# Repositório de Scripts PowerShell para Microsoft 365

Este repositório contém uma coleção de scripts PowerShell projetados para facilitar a administração e a análise em ambientes Microsoft 365 e Active Directory. Cada script é autossuficiente e possui funcionalidades específicas que atendem a diferentes necessidades administrativas.

## Scripts Disponíveis

### 1. **Procura_Eventos.ps1**
   - **Descrição**: Analisador abrangente de logs de eventos do Windows.
   - **Funcionalidades**:
     - Busca múltiplos Event IDs em arquivos de log (.evtx).
     - Exporta resultados para Excel com formatação profissional.
     - Filtragem por intervalo de datas e tratamento de arquivos corrompidos.
   - **Uso**: Ideal para investigações de segurança e auditorias.

### 2. **Procura_Arquivos.ps1**
   - **Descrição**: Localizador avançado de arquivos no OneDrive.
   - **Funcionalidades**:
     - Busca em OneDrive de um usuário específico ou em todos os OneDrives de um domínio.
     - Permite filtros de busca personalizados.
   - **Uso**: Útil para encontrar arquivos específicos em ambientes corporativos.

### 3. **office_removal.ps1**
   - **Descrição**: Remoção completa de todas as versões do Microsoft Office.
   - **Funcionalidades**:
     - Desinstalação silenciosa e remoção de registros e perfis do Outlook.
     - Limpeza de pastas de instalação e arquivos temporários.
   - **Uso**: Preparação para reinstalação limpa do Office.

### 4. **monitor-ping.ps1**
   - **Descrição**: Monitoramento contínuo de latência via ICMP (ping).
   - **Funcionalidades**:
     - Monitora até 5 endereços IP com relatórios em tempo real.
     - Classifica latências por níveis (verde, amarelo, vermelho).
   - **Uso**: Ideal para monitorar a saúde da rede.

### 5. **Configura-CatchAll.ps1**
   - **Descrição**: Configuração automatizada de regra catch-all para domínios Microsoft 365.
   - **Funcionalidades**:
     - Cria grupos dinâmicos e regras de transporte para redirecionar emails.
     - Exceções automáticas para usuários válidos.
   - **Uso**: Gerenciamento de emails em domínios corporativos.

### 6. **Buscar_Logon.ps1**
   - **Descrição**: Coleta e análise de eventos de logon de usuários em Active Directory.
   - **Funcionalidades**:
     - Busca eventos de logon humano (EventID 4624) em computadores do domínio.
     - Exporta resultados para Excel com relatórios detalhados.
   - **Uso**: Auditoria de logons e segurança.

### 7. **Alterar_Senhas_365.ps1**
   - **Descrição**: Gerador automatizado de senhas aleatórias para usuários Microsoft 365.
   - **Funcionalidades**:
     - Gera senhas seguras e aplica em massa a usuários de um domínio.
     - Relatórios detalhados de sucessos e falhas.
   - **Uso**: Gerenciamento de senhas em ambientes corporativos.

## Requisitos

- **PowerShell**: Versão 5.1 ou superior.
- **Módulos**: Dependências específicas para cada script (ex: `ImportExcel`, `Microsoft.Graph`, `ActiveDirectory`).
- **Permissões**: Privilegios administrativos necessários para execução de scripts que interagem com o Active Directory e Microsoft 365.

## Como Usar

1. Clone o repositório:
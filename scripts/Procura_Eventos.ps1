# Repositório de Scripts PowerShell para Microsoft 365

Este repositório contém uma coleção de scripts PowerShell projetados para facilitar a administração e a análise em ambientes Microsoft 365 e Active Directory. Cada script é interativo e oferece funcionalidades específicas para atender a diversas necessidades administrativas.

## Scripts Disponíveis

### 1. **Procura_Eventos.ps1**
   - **Descrição**: Analisador abrangente de logs de eventos do Windows.
   - **Funcionalidades**:
     - Busca múltiplos Event IDs em arquivos de log (.evtx).
     - Filtragem por intervalo de datas.
     - Exportação de resultados para Excel.
   - **Uso**: Ideal para investigações de segurança e auditorias.

### 2. **Procura_Arquivos.ps1**
   - **Descrição**: Localizador avançado de arquivos no OneDrive.
   - **Funcionalidades**:
     - Busca em OneDrive de um usuário específico ou em todos os OneDrives de um domínio.
   - **Uso**: Útil para encontrar arquivos em ambientes corporativos.

### 3. **office_removal.ps1**
   - **Descrição**: Remoção completa de todas as versões do Microsoft Office e Outlook.
   - **Funcionalidades**:
     - Desinstalação silenciosa e remoção de registros.
     - Eliminação de perfis e arquivos temporários.
   - **Uso**: Preparação para reinstalação limpa do Office.

### 4. **monitor-ping.ps1**
   - **Descrição**: Monitoramento contínuo de latência via ICMP (ping).
   - **Funcionalidades**:
     - Monitora até 5 endereços IP.
     - Gera relatórios em tempo real em formato CSV.
   - **Uso**: Ideal para monitoramento de rede.

### 5. **Configura-CatchAll.ps1**
   - **Descrição**: Configuração automatizada de regra catch-all para domínios Microsoft 365.
   - **Funcionalidades**:
     - Criação de grupo dinâmico de exceção.
     - Redirecionamento de emails enviados para endereços inexistentes.
   - **Uso**: Gerenciamento de emails em ambientes corporativos.

### 6. **Buscar_Logon.ps1**
   - **Descrição**: Coleta e análise de eventos de logon de usuários em Active Directory.
   - **Funcionalidades**:
     - Busca eventos de logon humano (EventID 4624).
     - Exportação de resultados para Excel.
   - **Uso**: Auditoria de logons e segurança.

### 7. **Alterar_Senhas_365.ps1**
   - **Descrição**: Gerador automatizado de senhas aleatórias para usuários Microsoft 365.
   - **Funcionalidades**:
     - Geração de senhas seguras e aplicáveis em massa.
     - Relatório detalhado de sucessos e falhas.
   - **Uso**: Gerenciamento de senhas em ambientes corporativos.

## Requisitos

- **PowerShell**: Versão 5.1 ou superior.
- **Módulos**: Dependências específicas para cada script (ex: `ImportExcel`, `Microsoft.Graph`, `ActiveDirectory`).
- **Permissões**: Privilegios administrativos necessários para execução de alguns scripts.

## Como Usar

1. Clone o repositório:
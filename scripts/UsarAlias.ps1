# Repositório de Scripts PowerShell para Microsoft 365

Este repositório contém uma coleção de scripts PowerShell projetados para facilitar a administração e a análise em ambientes Microsoft 365 e Active Directory. Cada script é autossuficiente e possui funcionalidades específicas, permitindo a automação de tarefas comuns e a coleta de dados relevantes.

## Scripts Disponíveis

### 1. **Procura_Eventos.ps1**
   - **Descrição**: Analisador abrangente de logs de eventos do Windows.
   - **Funcionalidades**:
     - Busca múltiplos Event IDs em logs do sistema.
     - Exporta resultados para Excel com formatação profissional.
     - Filtragem por intervalo de datas e tratamento de logs arquivados.
   - **Uso**: Ideal para investigações de segurança e auditorias.

### 2. **Procura_Arquivos.ps1**
   - **Descrição**: Localizador avançado de arquivos no OneDrive.
   - **Funcionalidades**:
     - Busca em OneDrive de usuários específicos ou em todos os OneDrives de um domínio.
     - Permite filtros de busca personalizados.
   - **Uso**: Útil para administradores que precisam localizar arquivos em ambientes corporativos.

### 3. **office_removal.ps1**
   - **Descrição**: Remoção completa de todas as versões do Microsoft Office.
   - **Funcionalidades**:
     - Desinstalação silenciosa e remoção de registros e perfis do Outlook.
     - Limpeza de arquivos temporários e pastas de instalação.
   - **Uso**: Preparação para reinstalação limpa do Office.

### 4. **monitor-ping.ps1**
   - **Descrição**: Monitoramento contínuo de latência via ICMP (ping).
   - **Funcionalidades**:
     - Monitora até 5 endereços IP com relatórios em tempo real.
     - Classifica latências por níveis (verde, amarelo, vermelho).
   - **Uso**: Ideal para administradores de rede que precisam monitorar a conectividade.

### 5. **Configura-CatchAll.ps1**
   - **Descrição**: Configuração automatizada de regra catch-all para domínios Microsoft 365.
   - **Funcionalidades**:
     - Cria grupo dinâmico de exceção e regra de transporte catch-all.
     - Permite configuração de domínio como InternalRelay.
   - **Uso**: Útil para gerenciar e redirecionar emails enviados para endereços inexistentes.

### 6. **Buscar_Logon.ps1**
   - **Descrição**: Coleta e análise de eventos de logon de usuários em Active Directory.
   - **Funcionalidades**:
     - Busca eventos de logon humano (EventID 4624) em computadores do domínio.
     - Exporta resultados para Excel com relatórios detalhados.
   - **Uso**: Ideal para auditorias de segurança e análise de logons.

### 7. **Alterar_Senhas_365.ps1**
   - **Descrição**: Gerador automatizado de senhas aleatórias para usuários Microsoft 365.
   - **Funcionalidades**:
     - Gera senhas seguras e aplica em massa a usuários de um domínio específico.
     - Exporta relatórios detalhados de sucessos e falhas.
   - **Uso**: Útil para administradores que precisam redefinir senhas em massa.

## Como Usar

1. **Pré-requisitos**: Certifique-se de ter o PowerShell instalado e os módulos necessários (como `Microsoft.Graph` e `ImportExcel`).
2. **Execução**: Cada script pode ser executado diretamente no PowerShell. Siga as instruções interativas que aparecem durante a execução.
3. **Permissões**: Alguns scripts requerem permissões administrativas ou específicas do Microsoft 365.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## Licença

Este projeto está licenciado sob a [Licença MIT](LICENSE).

---

Sinta-se à vontade para personalizar ainda mais este README conforme necessário. Um README claro e conciso pode ajudar outros a entender rapidamente o propósito e a funcionalidade do seu repositório. Boa sorte com seu LinkedIn!
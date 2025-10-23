# Repositório de Scripts PowerShell para Microsoft 365

Este repositório contém uma coleção de scripts PowerShell projetados para facilitar a administração e a análise em ambientes Microsoft 365 e Active Directory. Cada script é autossuficiente e possui funcionalidades específicas que atendem a diferentes necessidades administrativas.

## Scripts Disponíveis

### 1. **Procura_Eventos.ps1**
   - **Descrição**: Analisador abrangente de logs de eventos do Windows.
   - **Funcionalidades**:
     - Busca múltiplos Event IDs em logs do sistema.
     - Exporta resultados para Excel com formatação profissional.
     - Filtragem por intervalo de datas e tratamento de arquivos corrompidos.

### 2. **Procura_Arquivos.ps1**
   - **Descrição**: Localizador avançado de arquivos no OneDrive.
   - **Funcionalidades**:
     - Busca em OneDrive de um usuário específico ou em todos os OneDrives de um domínio.
     - Permite filtros de busca personalizados.

### 3. **office_removal.ps1**
   - **Descrição**: Remoção completa de todas as versões do Microsoft Office e Outlook.
   - **Funcionalidades**:
     - Desinstalação silenciosa e remoção de registros e arquivos temporários.
     - Limpeza total do sistema para reinstalação.

### 4. **monitor-ping.ps1**
   - **Descrição**: Monitoramento contínuo de latência via ICMP (ping).
   - **Funcionalidades**:
     - Monitora até 5 endereços IP.
     - Gera relatórios em tempo real em formato CSV.

### 5. **Configura-CatchAll.ps1**
   - **Descrição**: Configuração automatizada de regra catch-all para domínios Microsoft 365.
   - **Funcionalidades**:
     - Cria grupos dinâmicos e regras de transporte para redirecionar emails.
     - Exceções para usuários válidos.

### 6. **Buscar_Logon.ps1**
   - **Descrição**: Coleta e analisa eventos de logon de usuários em computadores do domínio Active Directory.
   - **Funcionalidades**:
     - Busca eventos de logon humano (EventID 4624).
     - Exporta resultados para Excel.

### 7. **Alterar_Senhas_365.ps1**
   - **Descrição**: Gerador automatizado de senhas aleatórias para usuários Microsoft 365.
   - **Funcionalidades**:
     - Gera senhas seguras e aplica em massa.
     - Exporta relatórios detalhados em CSV.

## Como Usar

1. **Pré-requisitos**: Certifique-se de ter o PowerShell e os módulos necessários instalados.
2. **Execução**: Execute os scripts no PowerShell com privilégios de administrador.
3. **Documentação**: Cada script contém comentários e exemplos de uso para facilitar a compreensão.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

---

Sinta-se à vontade para explorar e utilizar os scripts conforme necessário. Para mais informações, consulte a documentação oficial do Microsoft 365 e PowerShell.
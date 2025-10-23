# Repositório de Scripts PowerShell para Microsoft 365

Este repositório contém uma coleção de scripts PowerShell projetados para facilitar a administração e a análise em ambientes Microsoft 365 e Active Directory. Cada script é documentado com suas funcionalidades, requisitos e exemplos de uso.

## Scripts Disponíveis

### 1. **Procura_Eventos.ps1**
   - **Descrição**: Analisador abrangente de logs de eventos do Windows. Permite busca por múltiplos Event IDs e exportação dos resultados para Excel.
   - **Funcionalidades**:
     - Busca simultânea de múltiplos Event IDs.
     - Filtragem por intervalo de datas.
     - Exportação estruturada para Excel.
   - **Uso**: Execute o script e siga as instruções interativas.

### 2. **Procura_Arquivos.ps1**
   - **Descrição**: Localizador avançado de arquivos no OneDrive, com opções para buscar em usuários específicos ou em todos os OneDrives de um domínio.
   - **Funcionalidades**:
     - Busca em OneDrive de usuário específico ou em todos os usuários do domínio.
     - Filtros de busca personalizáveis.
   - **Uso**: Execute o script e escolha o modo de busca desejado.

### 3. **office_removal.ps1**
   - **Descrição**: Script para remoção completa de todas as versões do Microsoft Office e Outlook do sistema Windows.
   - **Funcionalidades**:
     - Desinstalação silenciosa de todas as versões do Office.
     - Remoção de chaves de registro e perfis do Outlook.
   - **Uso**: Execute o script como administrador para iniciar a remoção.

### 4. **monitor-ping.ps1**
   - **Descrição**: Monitoramento contínuo de latência via ICMP (ping) para múltiplos destinos.
   - **Funcionalidades**:
     - Relatório em tempo real com resultados coloridos.
     - Exportação de resultados para CSV.
   - **Uso**: Execute o script e forneça os IPs a serem monitorados.

### 5. **Configura-CatchAll.ps1**
   - **Descrição**: Configuração automatizada de regra catch-all para domínios Microsoft 365.
   - **Funcionalidades**:
     - Criação de grupo dinâmico de exceção.
     - Implementação de regra de transporte catch-all.
   - **Uso**: Execute o script e siga as instruções interativas.

### 6. **Buscar_Logon.ps1**
   - **Descrição**: Coleta e análise de eventos de logon de usuários em computadores do domínio Active Directory.
   - **Funcionalidades**:
     - Busca por eventos de logon humano (EventID 4624).
     - Exportação para Excel com formatação profissional.
   - **Uso**: Execute o script e forneça o nome do computador ou domínio.

### 7. **Alterar_Senhas_365.ps1**
   - **Descrição**: Gerador automatizado de senhas aleatórias para usuários Microsoft 365 por domínio.
   - **Funcionalidades**:
     - Geração de senhas seguras e memoráveis.
     - Aplicação em massa de novas senhas.
   - **Uso**: Execute o script e forneça o domínio alvo.

## Requisitos

- **PowerShell**: Versão 5.1 ou superior.
- **Módulos Necessários**: 
  - `ImportExcel`
  - `Microsoft.Graph`
  - `ActiveDirectory` (para scripts relacionados ao AD)

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

---

Sinta-se à vontade para explorar os scripts e adaptá-los conforme suas necessidades!
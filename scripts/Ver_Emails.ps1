# Repositório de Scripts PowerShell para Microsoft 365

Este repositório contém uma coleção de scripts PowerShell projetados para facilitar a administração e a análise em ambientes Microsoft 365 e Active Directory. Cada script é documentado com suas funcionalidades, requisitos e exemplos de uso.

## Scripts Disponíveis

### 1. **Procura_Eventos.ps1**
   - **Descrição**: Analisador abrangente de logs de eventos do Windows. Permite busca por múltiplos Event IDs e exporta resultados para Excel.
   - **Funcionalidades**:
     - Busca simultânea de múltiplos Event IDs.
     - Filtragem por intervalo de datas.
     - Exportação estruturada para Excel.
   - **Uso**: Execute o script e siga as instruções interativas.

### 2. **Procura_Arquivos.ps1**
   - **Descrição**: Localizador avançado de arquivos no OneDrive, com opções para buscar em usuários específicos ou em todos os OneDrives de um domínio.
   - **Funcionalidades**:
     - Busca em OneDrive de um usuário específico ou de todos os usuários de um domínio.
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
     - Monitora até 5 endereços IP.
     - Gera relatórios em tempo real em formato CSV.
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
     - Exportação de resultados para Excel.
   - **Uso**: Execute o script e forneça as informações solicitadas.

### 7. **Alterar_Senhas_365.ps1**
   - **Descrição**: Gerador automatizado de senhas aleatórias para usuários Microsoft 365 por domínio.
   - **Funcionalidades**:
     - Geração de senhas seguras e memoráveis.
     - Aplicação em massa de senhas para usuários habilitados.
   - **Uso**: Execute o script e forneça o domínio alvo.

## Requisitos

- **PowerShell**: Versão 5.1 ou superior.
- **Módulos**: Dependendo do script, pode ser necessário instalar módulos como `ImportExcel`, `Microsoft.Graph`, e `ActiveDirectory`.
- **Permissões**: Alguns scripts requerem privilégios administrativos.

## Contribuições

Sinta-se à vontade para contribuir com melhorias ou novas funcionalidades. Para isso, basta abrir uma issue ou enviar um pull request.

## Licença

Este projeto está licenciado sob a [Licença MIT](LICENSE).

---

Sinta-se à vontade para explorar os scripts e adaptá-los conforme suas necessidades!
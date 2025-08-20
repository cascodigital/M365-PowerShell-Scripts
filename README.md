# M365-PowerShell-Scripts

Uma coleção de scripts PowerShell para automação e administração de ambientes Microsoft 365 e infraestrutura local (Active Directory), utilizando os módulos Microsoft.Graph e outros.

## Tabela de Conteúdos

1. [Ver_MfaComplianceReport.ps1](#ver_mfacompliancereportps1)
2. [Relacao_Confianca.ps1](#relacao_confiancaps1)
3. [Procura_Arquivos.ps1](#procura_arquivosps1)
4. [Procura_Eventos.ps1](#procura_eventosps1)
5. [BuscaLogon.ps1](#buscalogonps1)
6. [GPO - Auditoria de Logon (gpo_logons.rar)](#gpo---auditoria-de-logon-gpo_logonsrar)
7. [AlterarPerfilDeRede.ps1](#alterarperfilderedps1)
8. [Alterar_Senhas_365.ps1](#alterar_senhas_365ps1)
---

## Ver_MfaComplianceReport.ps1

Este script gera um relatório de conformidade sobre o status do MFA no Microsoft 365, focando em **contas de usuários reais** e, ao final, entra em um **modo de consulta interativo** para análise detalhada de contas específicas.

### Funcionalidades Principais

- **Filtragem Inteligente de Usuários**: Exclui contas de serviço, sincronização e sistemas para focar o relatório em pessoas, tornando a análise de conformidade mais precisa.

- **Geração de Relatório Duplo**: Cria dois arquivos de saída:
  - Um `.csv` com todos os dados técnicos para análise.
  - Um `.txt` formatado como um relatório executivo, ideal para apresentar a clientes ou gestores.

- **Análise Detalhada dos Métodos**: Identifica uma gama completa de métodos de MFA, incluindo **Microsoft Authenticator**, **Telefone/SMS**, **Apps OATH**, **Windows Hello for Business** e **chaves FIDO2**.

- **Consulta Interativa Pós-relatório**: Após gerar o relatório, o script permite ao administrador consultar os detalhes de qualquer usuário em tempo real, exibindo informações específicas como o número de telefone (mascarado), nome do dispositivo vinculado ao Authenticator, data de criação do método, etc.

- **Sumário Visual no Console**: Apresenta um resumo colorido e direto no terminal, mostrando o percentual de conformidade e listando usuários com e sem MFA.

### Pré-requisitos

- Módulo PowerShell **Microsoft.Graph.PowerShell** instalado (`Install-Module -Name Microsoft.Graph`).
- **Permissões de API do Microsoft Graph**: O administrador que executa o script precisa consentir com as permissões `User.Read.All`, `UserAuthenticationMethod.Read.All`, `Directory.Read.All` e `Policy.Read.All`.

### Como Usar

1. Baixe o script e renomeie-o para `Ver-MfaComplianceReport.ps1`.
2. **Importante**: Abra o script e personalize os filtros de exclusão na seção `# 2. Obter usuários` para corresponder aos padrões de nomes de contas de serviço da sua organização.
3. Abra o PowerShell e execute o script: `.\Ver-MfaComplianceReport.ps1`
4. Após a geração dos relatórios, o script entrará no modo de consulta. Digite o email de um usuário para ver seus métodos de MFA em detalhe ou pressione ENTER para sair.

---

## Relacao_Confianca.ps1

Este script verifica o status da **relação de confiança (trust relationship)** de todos os computadores ativos em um domínio Active Directory local. Ele gera um relatório detalhado em formato Excel (`.xlsx`), identificando máquinas offline e aquelas com a relação de confiança quebrada.

### Funcionalidades Principais

- **Verificação Abrangente**: Testa todos os computadores com status "Enabled" no Active Directory.
- **Diagnóstico Preciso**: Diferencia computadores offline (não respondem ao ping) de computadores com falha na relação de confiança.
- **Uso de nltest.exe**: Utiliza o comando nativo `nltest` para verificar a confiança, eliminando a necessidade de WinRM (PowerShell Remoting).
- **Relatório em Excel**: Exporta os resultados para um arquivo `.xlsx` formatado, com data e hora no nome.
- **Cálculo de Inatividade**: Para máquinas offline, estima há quantos dias estão sem comunicação com base na propriedade `LastLogonDate`.
- **Resumo no Console**: Exibe um resumo rápido dos resultados (OK, Falha, Offline) diretamente no terminal após a conclusão.

### Pré-requisitos

- Execução em um computador ingressado no domínio ou com as ferramentas RSAT do Active Directory instaladas.
- Módulo PowerShell **ActiveDirectory**.
- Módulo PowerShell **ImportExcel** instalado (`Install-Module -Name ImportExcel`).

### Como Usar

1. Baixe o script `Relacao-Confianca.ps1`.
2. (Opcional) Edite a variável `$reportFolder` no início do script para alterar o local onde o relatório será salvo.
3. Abra o PowerShell como Administrador em um controlador de domínio ou máquina com as ferramentas do AD.
4. Execute o script: `.\Relacao-Confianca.ps1`
5. O relatório será gerado na pasta `C:\Temp` (ou no local configurado).

---

## Procura_Arquivos.ps1

Este script interativo localiza arquivos por nome no **OneDrive for Business** de um usuário específico, utilizando o Microsoft Graph para realizar a busca de forma eficiente.

### Funcionalidades Principais

- **Busca Interativa**: Solicita o e-mail do usuário alvo e o nome/filtro do arquivo diretamente no console.
- **Suporte a Wildcards**: Permite buscas flexíveis utilizando asteriscos (*).
- **Múltiplos Drives**: O script identifica e pesquisa em todos os drives associados ao usuário.
- **Saída Detalhada**: Exibe o caminho completo do arquivo, drive de origem, tamanho, data da última modificação e o ID do item.
- **Detecção de Duplicatas**: Alerta o administrador caso arquivos com o mesmo nome sejam encontrados.

### Pré-requisitos

- Módulo PowerShell **Microsoft.Graph.PowerShell** instalado (`Install-Module -Name Microsoft.Graph`).
- **Permissões de API do Microsoft Graph**: `User.Read.All` e `Files.Read.All`.

### Como Usar

1. Baixe o script `procura_arquivos.ps1`.
2. Abra o PowerShell e execute o script: `.\procura_arquivos.ps1`
3. Siga as instruções no console.

---

## Procura_Eventos.ps1

Este script PowerShell foi criado para simplificar e acelerar a análise de **Logs de Eventos do Windows**, permitindo a busca por múltiplos Event IDs em um intervalo de datas e exportando os resultados para um arquivo Excel (`.xlsx`).

### Funcionalidades Principais

- **Busca por Múltiplos Event IDs**: Permite ao usuário inserir um ou mais IDs de evento de uma só vez.
- **Filtro por Intervalo de Datas**: A busca é restrita a um período de início e fim definido pelo usuário.
- **Busca Abrangente**: O script varre TODOS os logs de eventos do sistema (`.evtx`).
- **Exportação Detalhada para Excel**: Organiza os resultados em um arquivo Excel, com abas separadas para cada Event ID.

### Pré-requisitos

- Execução como **Administrador**.
- Módulo PowerShell **ImportExcel** instalado (`Install-Module -Name ImportExcel`).

### Como Usar

1. Baixe o script `Procura_Eventos.ps1`.
2. Abra o PowerShell como Administrador.
3. Execute o script e siga as instruções no console.

---

## BuscaLogon.ps1

Este script realiza uma busca forense por **eventos de logon (ID 4624)** em um computador específico ou em todos os computadores ativos do domínio, focando em tipos de logon que indicam atividade humana direta (incluindo logons offline/cacheados). O objetivo é consolidar, a partir de uma máquina central, os eventos ocorridos em múltiplos endpoints.

### Funcionalidades Principais

- **Escopo Flexível**: A busca pode ser feita em um único computador ou em todo o domínio.
- **Foco em Logons Relevantes**: Filtra por tipos de logon **Interativo(2)**, **Remote(10)**, **Desbloqueio(7)** e **Offline/Cache(11)**.
- **Gerenciamento de Timeout**: Usa jobs em paralelo com tempo limite para não travar em máquinas que não respondem.
- **Relatório em Excel**: Exporta os resultados para um arquivo `.xlsx` detalhado.
- **Resumo no Console**: Exibe um sumário da coleta (sucesso, offline, falhas) no terminal.

### Pré-requisitos

- **GPO de Auditoria (Essencial)**: As máquinas alvo devem ter a política de auditoria habilitada para registrar os eventos de logon. Uma GPO pronta (`gpo_logons.rar`) está incluída neste repositório.
- Módulo PowerShell **ActiveDirectory**.
- Módulo PowerShell **ImportExcel** (`Install-Module -Name ImportExcel`).
- Execução como Administrador em uma máquina com as ferramentas RSAT (AD).

### Como Usar

1. **Primeiro, importe a GPO**: Descompacte `gpo_logons.rar` e importe a política de grupo usando o GPMC no seu Domain Controller. Vincule a GPO à OU que contém os computadores a serem auditados.
2. Aguarde a replicação da GPO nos clientes.
3. Execute o script `BuscaLogon.ps1`.
4. Siga as instruções: informe o alvo ('T' para todos), a data inicial e a data final.
5. O relatório será salvo em `C:\Temp\EventLog_Exports`.

---

## GPO - Auditoria de Logon (gpo_logons.rar)

Este é um backup de uma **Group Policy Object (GPO)** pré-configurada para habilitar as políticas de auditoria necessárias para o funcionamento do script `BuscaLogon.ps1`.

### Funcionalidades Principais

- **Ativação da Auditoria Avançada**: Habilita as subcategorias de auditoria essenciais para registrar o Evento ID 4624, como "Logon", "Logon Especial", "Criação de Processos", etc.
- **Habilitação do WinRM**: Configura o serviço de Gerenciamento Remoto do Windows e as regras de firewall necessárias para que o script possa consultar os logs remotamente.

### Pré-requisitos

- Um ambiente de **Active Directory Domain Services**.
- Permissões de **Administrador de Domínio** para manipular GPOs.

### Como Usar

1. Descompacte o arquivo `gpo_logons.rar`.
2. No seu Domain Controller, abra o "Gerenciamento de Política de Grupo" (GPMC).
3. Clique com o botão direito em "Objetos de Política de Grupo" e selecione "Novo" para criar uma GPO vazia (ex: "Auditoria de Logon de Estações").
4. Clique com o botão direito na GPO recém-criada e selecione "Importar Configurações...".
5. Siga o assistente, apontando para a pasta que foi descompactada.
6. Após a importação, vincule a GPO à(s) OU(s) que contêm os computadores que serão auditados.

---

## AlterarPerfilDeRede.ps1

Este script interativo permite visualizar e alterar a categoria de **perfis de conexão de rede** (Pública, Privada ou Domínio) em uma máquina Windows local. É útil para corrigir rapidamente problemas de firewall quando uma rede é identificada incorretamente como "Pública".

### Funcionalidades Principais

- **Listagem Clara**: Exibe todos os perfis de rede ativos com seus nomes, índices e categorias atuais.
- **Seleção Interativa**: Permite ao administrador selecionar a interface de rede a ser alterada pelo seu índice.
- **Menu de Alteração Simples**: Oferece um menu claro para escolher a nova categoria de rede.
- **Verificação de Privilégios**: Garante que o script seja executado com permissões de administrador, o que é necessário para realizar as alterações.

### Pré-requisitos

- Execução como **Administrador** na máquina local onde a alteração é necessária.

### Como Usar

1. Baixe o script `AlterarPerfilDeRede.ps1`.
2. Clique com o botão direito no arquivo e selecione "Executar com o PowerShell" ou abra uma janela do PowerShell como Administrador.
3. Execute o script: `.\AlterarPerfilDeRede.ps1`
4. Siga as instruções no console para selecionar a interface e a nova categoria de rede.

---

## Alterar_Senhas_365.ps1

Este script automatiza a geração e aplicação de **senhas aleatórias** para usuários específicos do Microsoft 365, filtrando por domínio. Ideal para operações de reset em massa ou padronização de senhas temporárias.

### Funcionalidades Principais

- **Geração de Senhas Seguras**: Cria senhas no formato **2 maiúsculas + 4 números + 2 minúsculas** (ex: AB1234cd).
- **Filtro por Domínio**: Permite selecionar usuários de um domínio específico para a operação.
- **Filtro de Usuários Ativos**: Processa apenas contas habilitadas, ignorando usuários desativados.
- **Confirmação Prévia**: Exibe a lista de usuários que serão afetados antes de executar as alterações.
- **Relatório Detalhado**: Gera um arquivo CSV com todas as senhas alteradas e status das operações.
- **Senhas Definitivas**: Define as senhas como permanentes - usuários não precisam alterá-las no próximo logon.
- **Controle de Throttling**: Implementa pausas entre requisições para evitar limitação de API.

### Pré-requisitos

- Módulo PowerShell **Microsoft.Graph** instalado (`Install-Module -Name Microsoft.Graph`).
- **Permissões Administrativas Elevadas**: Executar com conta que possua um dos seguintes roles:
  - **Global Administrator** (recomendado)
  - **User Administrator**
  - **Password Administrator**
- **Permissões de API do Microsoft Graph**: `User.ReadWrite.All`, `Directory.ReadWrite.All`, `UserAuthenticationMethod.ReadWrite.All`, `Directory.AccessAsUser.All`.

### Como Usar

1. Baixe o script `Alterar_Senhas_365.ps1`.
2. **IMPORTANTE**: Execute com uma conta administrativa adequada no tenant M365.
3. Abra o PowerShell como Administrador.
4. Execute o script: `.\Alterar_Senhas_365.ps1`
5. Siga as instruções:
   - Autentique-se com suas credenciais administrativas do M365
   - Informe o domínio alvo (ex: `empresa.com.br`)
   - Confirme a lista de usuários que serão processados
6. O script irá processar todos os usuários e gerar um arquivo CSV com os resultados.

### ⚠️ Avisos Importantes

- **Teste primeiro**: Execute em um ambiente de teste ou com um usuário específico antes de aplicar em larga escala.
- **Backup**: Considere documentar as senhas atuais se necessário para rollback.
- **Comunicação**: Notifique os usuários sobre a alteração das senhas antes de executar.
- **Segurança**: O arquivo CSV contém senhas em texto plano - trate com máxima segurança.

---

## Contribuições

Contribuições são bem-vindas! Sinta-se livre para abrir issues ou enviar pull requests.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

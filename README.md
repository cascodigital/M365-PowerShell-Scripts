# M365-PowerShell-Scripts

Uma coleção de scripts PowerShell para automação e administração de ambientes Microsoft 365 e infraestrutura local (Active Directory), utilizando os módulos `Microsoft.Graph` e outros.

## Tabela de Conteúdos

1.  [Ver_MfaComplianceReport.ps1](#ver_mfacompliancereportps1)
2.  [Alterar_Senhas_365.ps1](#alterar_senhas_365ps1)
3.  [Procura_Arquivos.ps1](#procura_arquivosps1)
4.  [Remover_Email.ps1](#remover_emailps1)
5.  [Relacao_Confianca.ps1](#relacao_confiancaps1)
6.  [Procura_Eventos.ps1](#procura_eventosps1)
7.  [Buscar_Logon.ps1](#buscar_logonps1)
8.  [GPO - Auditoria de Logon (gpo_logons.rar)](#gpo---auditoria-de-logon-gpo_logonsrar)
9.  [AlterarPerfilDeRede.ps1](#alterarperfilderedeps1)
10. [Ver_Emails.ps1](#ver_emailsps1)

---

## Ver_MfaComplianceReport.ps1

Gera um relatório de conformidade sobre o status do MFA no Microsoft 365, focando em **contas de usuários reais** e, ao final, entra em um **modo de consulta interativo** para análise detalhada de contas específicas.

### Funcionalidades Principais

* **Filtragem Inteligente**: Exclui contas de serviço, sincronização e sistemas para focar o relatório em usuários reais.
* **Relatório Duplo**: Cria um `.csv` com dados completos e um `.txt` formatado como relatório executivo.
* **Análise de Métodos**: Identifica todos os métodos de MFA, como **Authenticator**, **Telefone/SMS**, **OATH**, **Windows Hello** e **FIDO2**.
* **Consulta Interativa**: Após gerar o relatório, permite consultar detalhes de qualquer usuário em tempo real.
* **Sumário Visual**: Exibe um resumo colorido no console com o percentual de conformidade.

### Pré-requisitos

* Módulo PowerShell `Microsoft.Graph`.
* Permissões de API do Microsoft Graph: `User.Read.All`, `UserAuthenticationMethod.Read.All`, `Directory.Read.All`, `Policy.Read.All`.

### Como Usar

1.  Abra o script e personalize os filtros de exclusão de contas de serviço.
2.  Execute no PowerShell: `.\Ver-MfaComplianceReport.ps1`.
3.  Após a geração dos relatórios, digite o email de um usuário para ver detalhes ou pressione ENTER para sair.

---

## Ver_Emails.ps1

Gera um **relatório completo de todos os e-mails vigentes** na organização Microsoft 365, categorizando usuários, grupos, caixas compartilhadas e aliases. Ideal para atender solicitações de levantamento de endereços de e-mail ativos.

### Funcionalidades Principais

* **Categorização Inteligente**: Separa automaticamente usuários ativos, inativos, externos, grupos e caixas compartilhadas.
* **Identificação de Tipos de Grupo**: Distingue entre Microsoft 365, Teams, Listas de Distribuição e grupos de Segurança.
* **Detecção de Aliases**: Identifica apelidos de e-mail tanto de usuários quanto de grupos através dos ProxyAddresses.
* **Relatório Executivo**: Cria aba específica com **apenas e-mails vigentes** para apresentação ao cliente.
* **Excel Organizado**: Gera arquivo Excel com múltiplas abas categorizadas e resumo executivo.
* **Filtragem de Usuários Externos**: Remove automaticamente convidados e contas `#EXT#` do relatório principal.

### Pré-requisitos

* Módulos PowerShell: `Microsoft.Graph.Users`, `Microsoft.Graph.Groups`, `ImportExcel`.
* Permissões de API do Microsoft Graph: `User.Read.All`, `Group.Read.All`, `Mail.Read`.

### Como Usar

1.  Execute no PowerShell: `.\Ver_Emails.ps1`.
2.  Aguarde a coleta de dados (pode levar alguns minutos em organizações grandes).
3.  O script gerará um arquivo Excel com timestamp no nome.
4.  **Para o cliente**: Use as abas `RESPOSTA_CLIENTE` e `RESUMO_EXECUTIVO`.
5.  **Para análise técnica**: Consulte as abas numeradas com detalhamentos.

### Estrutura do Relatório

* **RESPOSTA_CLIENTE**: Lista limpa apenas dos e-mails funcionais
* **RESUMO_EXECUTIVO**: Totais por categoria para apresentação
* **1-Usuários_Ativos**: Funcionários com licença ativa
* **2-Usuários_Sem_Licença**: Possíveis ex-funcionários
* **3-Usuários_Externos**: Convidados e contas externas
* **4-Grupos**: Todos os tipos (Microsoft 365, Teams, Distribuição)
* **5-Caixas_Compartilhadas**: E-mails genéricos compartilhados
* **6-Aliases**: Apelidos e endereços alternativos

---

## Alterar_Senhas_365.ps1

Automatiza a geração e aplicação de **senhas aleatórias** para usuários de um domínio específico no Microsoft 365.

### Funcionalidades Principais

* **Geração Segura**: Cria senhas no formato `AA1234aa`.
* **Filtro por Domínio**: Aplica a alteração apenas a usuários do domínio especificado.
* **Confirmação Prévia**: Exibe a lista de usuários que serão afetados antes de prosseguir.
* **Relatório CSV**: Gera um arquivo com todas as senhas alteradas e o status da operação.
* **Senha Permanente**: Define as senhas para não expirarem no próximo logon.
* **Controle de Throttling**: Implementa pausas para evitar bloqueios da API.

### Pré-requisitos

* Módulo PowerShell `Microsoft.Graph`.
* Permissões de **Administrador Global**, **Administrador de Usuários** ou **Administrador de Senhas**.
* Permissões de API do Graph: `User.ReadWrite.All`, `Directory.ReadWrite.All`.

### Como Usar

1.  Execute o script `.\Alterar_Senhas_365.ps1` em uma janela do PowerShell como Administrador.
2.  Autentique-se com uma conta administrativa.
3.  Informe o domínio alvo (ex: `empresa.com`).
4.  Confirme a lista de usuários para iniciar o processo.

#### ⚠️ Avisos Importantes
* **Segurança**: O arquivo CSV gerado contém senhas em texto plano. Armazene-o em local seguro e remova-o quando não for mais necessário.
* **Comunicação**: Avise os usuários sobre a alteração antes de executar o script.

---

## Procura_Arquivos.ps1

Localiza arquivos por nome no **OneDrive for Business** de um usuário específico, de forma interativa.

### Funcionalidades Principais

* **Busca Interativa**: Solicita o e-mail do usuário e o nome do arquivo no console.
* **Suporte a Wildcards**: Permite o uso de `*` para buscas flexíveis.
* **Saída Detalhada**: Exibe caminho completo, tamanho, data de modificação e ID do arquivo.

### Pré-requisitos

* Módulo PowerShell `Microsoft.Graph`.
* Permissões de API do Graph: `User.Read.All`, `Files.Read.All`.

### Como Usar

1.  Execute o script: `.\Procura_Arquivos.ps1`.
2.  Siga as instruções no console.

---

## Remover_Email.ps1

Realiza a remoção em massa de e-mails específicos de **todas as caixas de correio** do Microsoft 365, com base no remetente e assunto.

### Funcionalidades Principais

* **Remoção Global**: Busca e remove e-mails de todo o ambiente M365.
* **Operação Segura**: Utiliza o `Security & Compliance Center` e realiza um **Soft Delete** (move os e-mails para a pasta de Itens Recuperáveis).
* **Processo Automatizado**: Conecta, solicita os critérios, cria a busca, aguarda a conclusão e executa a remoção.
* **Confirmação Crítica**: Exige uma confirmação final antes de iniciar a remoção.
* **Status em Tempo Real**: Exibe o progresso da busca diretamente no console.

### Pré-requisitos

* Módulo PowerShell `ExchangeOnlineManagement`.
* Permissões administrativas que incluam o role **Search And Purge** (disponível em grupos como *Compliance Management* ou *Organization Management*).

### Como Usar

1.  Execute o script `.\Remover_Email.ps1` em uma janela do PowerShell como Administrador.
2.  O script se conectará ao *Security & Compliance Center*.
3.  Informe o **remetente** e o **assunto** do e-mail a ser removido.
4.  Confirme a operação digitando `S`.

#### ⚠️ Avisos Importantes
* **Impacto Elevado**: Este script afeta TODAS as caixas de correio. Use com extrema cautela e tenha certeza absoluta dos critérios de busca.
* **Irreversibilidade**: Embora seja um *Soft Delete*, a recuperação dos e-mails é um processo manual. Verifique os parâmetros duas vezes antes de confirmar.

---

## Relacao_Confianca.ps1

Verifica o status da **relação de confiança (trust relationship)** de todos os computadores ativos no Active Directory local e gera um relatório em Excel.

### Funcionalidades Principais

* **Diagnóstico Preciso**: Diferencia máquinas offline daquelas com a relação de confiança quebrada.
* **Relatório em Excel**: Exporta os resultados para um arquivo `.xlsx` formatado.
* **Cálculo de Inatividade**: Estima há quantos dias as máquinas offline não se comunicam.
* **Não Requer WinRM**: Usa o comando `nltest.exe` para a verificação.

### Pré-requisitos

* Execução em um computador ingressado no domínio com as ferramentas RSAT.
* Módulo PowerShell `ActiveDirectory`.
* Módulo PowerShell `ImportExcel`.

### Como Usar

1.  Execute o script `.\Relacao_Confianca.ps1` em uma janela do PowerShell como Administrador.
2.  O relatório será gerado na pasta `C:\Temp` por padrão.

---

## Procura_Eventos.ps1

Busca múltiplos **Event IDs** nos logs de eventos do Windows em um intervalo de datas e exporta os resultados para Excel.

### Funcionalidades Principais

* **Busca por Múltiplos IDs**: Permite inserir vários Event IDs de uma só vez.
* **Filtro por Data**: Restringe a busca a um período específico.
* **Busca Completa**: Varre todos os logs de eventos do sistema (`.evtx`).
* **Exportação para Excel**: Organiza os resultados em abas separadas para cada Event ID.

### Pré-requisitos

* Execução como **Administrador**.
* Módulo PowerShell `ImportExcel`.

### Como Usar

1.  Execute o script `.\Procura_Eventos.ps1` como Administrador.
2.  Siga as instruções no console.

---

## Buscar_Logon.ps1

Realiza uma busca forense por **eventos de logon (ID 4624)** em um ou todos os computadores do domínio, focando em atividades humanas diretas (logon interativo, remoto e offline).

### Funcionalidades Principais

* **Escopo Flexível**: Busca em um único PC ou em todo o domínio.
* **Foco em Logons Relevantes**: Filtra por tipos de logon 2, 7, 10 e 11.
* **Execução em Paralelo**: Usa jobs com timeout para não travar em máquinas offline.
* **Relatório em Excel**: Exporta os resultados para um arquivo `.xlsx`.

### Pré-requisitos

* **GPO de Auditoria (Essencial)**: As máquinas alvo devem ter a política de auditoria de logon habilitada. Uma GPO pronta está incluída neste repositório.
* Módulo PowerShell `ActiveDirectory` e `ImportExcel`.
* Execução como Administrador.

### Como Usar

1.  **Primeiro**: Importe e vincule a GPO `gpo_logons.rar` na OU dos computadores.
2.  Aguarde a replicação da política.
3.  Execute `.\Buscar_Logon.ps1` e siga as instruções.

---

## GPO - Auditoria de Logon (gpo_logons.rar)

Backup de uma **Group Policy Object (GPO)** pré-configurada para habilitar as políticas de auditoria necessárias para o funcionamento do script `Buscar_Logon.ps1`.

### Funcionalidades Principais

* **Ativação da Auditoria Avançada**: Habilita as subcategorias de auditoria para registrar o Evento ID 4624.
* **Habilitação do WinRM**: Configura o serviço de Gerenciamento Remoto do Windows e as regras de firewall.

### Pré-requisitos

* Ambiente Active Directory Domain Services.
* Permissões de Administrador de Domínio.

### Como Usar

1.  Descompacte o arquivo.
2.  No GPMC, crie uma nova GPO vazia.
3.  Clique com o botão direito na GPO criada e selecione **"Importar Configurações..."**.
4.  Siga o assistente, apontando para a pasta descompactada.
5.  Vincule a GPO à OU que contém os computadores a serem auditados.

---

## AlterarPerfilDeRede.ps1

Permite visualizar e alterar a categoria de perfis de rede (Pública, Privada) em uma máquina Windows local.

### Funcionalidades Principais

* **Listagem Clara**: Exibe os perfis de rede ativos com seus nomes e categorias.
* **Alteração Interativa**: Permite selecionar a interface e a nova categoria por meio de um menu.
* **Verificação de Privilégios**: Garante a execução com permissões de administrador.

### Pré-requisitos

* Execução como **Administrador** na máquina local.

### Como Usar

1.  Clique com o botão direito no arquivo `AlterarPerfilDeRede.ps1` e selecione "Executar com o PowerShell".
2.  Siga as instruções no console.

---

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

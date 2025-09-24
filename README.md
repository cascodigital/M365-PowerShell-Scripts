# M365-PowerShell-Scripts

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![PowerShell: 7.5+](https://img.shields.io/badge/PowerShell-7.5%2B-blue.svg)

Uma coleção de scripts PowerShell para automação e administração de ambientes Microsoft 365 e infraestrutura local (Active Directory), utilizando os módulos `Microsoft.Graph` e outros.

---

## 🚀 Tabela de Scripts

### Categoria: Microsoft 365

| Script | Descrição |
| :--- | :--- |
| **[Ver_MfaComplianceReport.ps1](#ver_mfacompliancereportps1)** | Gera um relatório de conformidade MFA e entra em modo de consulta interativo. |
| **[Ver_Emails.ps1](#ver_emailsps1)** | Gera um relatório completo de todos os e-mails vigentes na organização. |
| **[Alterar_Senhas_365.ps1](#alterar_senhas_365ps1)** | Automatiza a geração e aplicação de senhas aleatórias para usuários de um domínio. |
| **[Procura_Arquivos.ps1](#procura_arquivosps1)** | Localiza arquivos no OneDrive for Business de um usuário de forma interativa. |
| **[Remover_Email.ps1](#remover_emailps1)** | Remove e-mails específicos de todas as caixas de correio da organização. |
| **[Configura-CatchAll.ps1](#configura-catchallps1)** | Automatiza a configuração de um e-mail "catch-all" (coletor geral) para um domínio. |
| **[UsarAlias.ps1](#usaraliasps1)** | Habilita a funcionalidade 'Enviar como Alias' e gerencia os aliases de um usuário. |

### Categoria: Active Directory & Windows Local

| Script | Descrição |
| :--- | :--- |
| **[Relacao_Confianca.ps1](#relacao_confiancaps1)** | Verifica a relação de confiança de todos os computadores no Active Directory. |
| **[Procura_Eventos.ps1](#procura_eventosps1)** | Busca múltiplos Event IDs nos logs de eventos do Windows em um período. |
| **[Buscar_Logon.ps1](#buscar_logonps1)** | Realiza busca forense por eventos de logon (ID 4624) em computadores do domínio. |
| **[GPO - Auditoria de Logon](#gpo---auditoria-de-logon-gpo_logonsrar)** | Backup de GPO para habilitar a auditoria necessária para o script `Buscar_Logon.ps1`. |
| **[AlterarPerfilDeRede.ps1](#alterarperfilderedeps1)** | Visualiza e altera a categoria de perfis de rede (Pública/Privada) em uma máquina local. |

---

## 📜 Detalhes dos Scripts

### Microsoft 365

#### Ver_MfaComplianceReport.ps1

Gera um relatório de conformidade sobre o status do MFA no Microsoft 365, focando em **contas de usuários reais** e, ao final, entra em um **modo de consulta interativo** para análise detalhada de contas específicas.

* **Funcionalidades**: Filtragem inteligente, relatório duplo (CSV e TXT), análise de métodos, consulta interativa, sumário visual.
* **Pré-requisitos**: Módulo `Microsoft.Graph`, permissões de API (`User.Read.All`, `UserAuthenticationMethod.Read.All`, etc.).
* **Como Usar**: Execute `.\Ver-MfaComplianceReport.ps1` e siga as instruções.

#### Ver_Emails.ps1

Gera um **relatório completo de todos os e-mails vigentes** na organização Microsoft 365, categorizando usuários, grupos, caixas compartilhadas e aliases. Ideal para atender solicitações de levantamento de endereços de e-mail ativos.

* **Funcionalidades**: Categorização inteligente, identificação de tipos de grupo, detecção de aliases, relatório executivo em Excel com abas organizadas.
* **Pré-requisitos**: Módulos `Microsoft.Graph.Users`, `Microsoft.Graph.Groups`, `ImportExcel`. Permissões de API (`User.Read.All`, `Group.Read.All`).
* **Como Usar**: Execute `.\Ver_Emails.ps1` e aguarde a geração do arquivo Excel.

#### Alterar_Senhas_365.ps1

Automatiza a geração e aplicação de **senhas aleatórias** para usuários de um domínio específico no Microsoft 365.

* **Funcionalidades**: Geração segura, filtro por domínio, confirmação prévia, relatório CSV, senhas permanentes, controle de throttling.
* **Pré-requisitos**: Módulo `Microsoft.Graph`, permissões de Admin (Global, Usuário ou Senha), permissões de API (`User.ReadWrite.All`).
* **Como Usar**: Execute `.\Alterar_Senhas_365.ps1`, autentique-se e informe o domínio alvo.
    > ⚠️ **Aviso Importante**: O CSV gerado contém senhas em texto plano. Armazene-o em local seguro e remova-o após o uso.

#### Procura_Arquivos.ps1

Localiza arquivos por nome no OneDrive for Business de usuário específico ou todos os usuários de um domínio.

* **Funcionalidades**: Busca interativa com dois modos (usuário único ou domínio completo), suporte a wildcards (*, ?), detecção de duplicatas, relatório CSV automático, verificação de provisionamento.Busca interativa com dois modos (usuário único ou domínio completo), suporte a wildcards (*, ?), detecção de duplicatas, relatório CSV automático, verificação de provisionamento.
* **Pré-requisitos**: Módulo Microsoft.Graph (Authentication, Users, Files), permissões de API (User.Read.All, Files.Read.All, Directory.Read.All), privilégios administrativos.
* **Como Usar**: Execute .\Procura_Arquivos.ps1, escolha modo 1 (usuário específico) ou modo 2 (todos do domínio), digite domínio/email + filtro de busca. Modo 2 gera CSV automaticamente.


#### Remover_Email.ps1

Realiza a remoção em massa de e-mails específicos de **todas as caixas de correio** do Microsoft 365, com base no remetente e assunto.

* **Funcionalidades**: Remoção global (Soft Delete), processo automatizado via Security & Compliance Center, confirmação crítica, status em tempo real.
* **Pré-requisitos**: Módulo `ExchangeOnlineManagement`, role **Search And Purge**.
* **Como Usar**: Execute `.\Remover_Email.ps1`, informe o remetente e o assunto, e confirme a operação.
    > ⚠️ **Aviso Importante**: Este script afeta TODAS as caixas de correio. Use com extrema cautela.

#### Configura-CatchAll.ps1

Automatiza a configuração de um e-mail **"catch-all"** (coletor geral) para um domínio específico, redirecionando e-mails enviados para endereços inexistentes para uma única caixa de correio.

* **Funcionalidades**: Instalação automática do módulo `ExchangeOnlineManagement`, validação de domínio, altera o tipo do domínio para `InternalRelay`, cria regra de transporte com prioridade dinâmica para evitar conflitos.
* **Pré-requisitos**: Módulo `ExchangeOnlineManagement`, permissões de Administrador do Exchange.
* **Como Usar**: Execute `.\Configura-CatchAll.ps1` e forneça o e-mail do administrador, o domínio alvo e o e-mail coletor.
    > ⚠️ **Aviso Importante**: A propagação da regra de transporte pode levar até uma hora para ser concluída em todo o ambiente.

#### UsarAlias.ps1

Habilita a funcionalidade **"Enviar como Alias"** para toda a organização e entra em um menu interativo para visualizar e adicionar novos aliases a um usuário específico.

* **Funcionalidades**: Ativação automática do recurso `SendFromAliasEnabled` no tenant, menu interativo para listar e adicionar múltiplos aliases a um usuário, instruções de uso no final.
* **Pré-requisitos**: Módulo `ExchangeOnlineManagement`, permissões de Administrador do Exchange.
* **Como Usar**: Execute `.\UsarAlias.ps1`, autentique-se e siga as instruções para selecionar o usuário e gerenciar seus aliases.
    > ⚠️ **Aviso Importante**: Se a funcionalidade for ativada pelo script, pode levar algumas horas para propagar.

---

### Active Directory & Windows Local

#### Relacao_Confianca.ps1

Verifica o status da **relação de confiança (trust relationship)** de todos os computadores ativos no Active Directory local e gera um relatório em Excel.

* **Funcionalidades**: Diagnóstico preciso, relatório em Excel, cálculo de inatividade, não requer WinRM.
* **Pré-requisitos**: Ferramentas RSAT, módulos `ActiveDirectory` e `ImportExcel`.
* **Como Usar**: Execute `.\Relacao_Confianca.ps1` como Administrador em um computador do domínio.

#### Procura_Eventos.ps1

Busca múltiplos **Event IDs** nos logs de eventos do Windows em um intervalo de datas e exporta os resultados para Excel.

* **Funcionalidades**: Busca por múltiplos IDs, filtro por data, varredura completa dos logs, exportação para Excel com abas separadas.
* **Pré-requisitos**: Execução como Administrador, módulo `ImportExcel`.
* **Como Usar**: Execute `.\Procura_Eventos.ps1` como Administrador e siga as instruções.

#### Buscar_Logon.ps1

Realiza uma busca forense por **eventos de logon (ID 4624)** em um ou todos os computadores do domínio, focando em atividades humanas diretas.

* **Funcionalidades**: Escopo flexível, foco em logons relevantes (tipos 2, 7, 10, 11), execução em paralelo, relatório em Excel.
* **Pré-requisitos**: GPO de Auditoria de Logon habilitada (inclusa no repositório), módulos `ActiveDirectory` e `ImportExcel`.
* **Como Usar**: Importe e vincule a GPO `gpo_logons.rar`, aguarde a replicação e execute `.\Buscar_Logon.ps1`.

#### GPO - Auditoria de Logon (gpo_logons.rar)

Backup de uma **Group Policy Object (GPO)** pré-configurada para habilitar as políticas de auditoria e o WinRM, necessários para o funcionamento do script `Buscar_Logon.ps1`.

* **Como Usar**: No GPMC, crie uma GPO vazia, clique com o botão direito, selecione **"Importar Configurações..."** e aponte para a pasta descompactada. Vincule a GPO na OU desejada.

#### AlterarPerfilDeRede.ps1

Permite visualizar e alterar a categoria de perfis de rede (Pública, Privada) em uma máquina Windows local.

* **Funcionalidades**: Listagem clara, alteração interativa, verificação de privilégios.
* **Pré-requisitos**: Execução como Administrador na máquina local.
* **Como Usar**: Clique com o botão direito no arquivo e selecione "Executar com o PowerShell".

### 👨‍💻 Autor

**Andre Kittler** - *Administrador Microsoft 365*

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

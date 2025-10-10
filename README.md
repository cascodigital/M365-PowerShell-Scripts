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

### Categoria: Limpeza & Recuperação

| Script | Descrição |
| :--- | :--- |
| **[office_removal.ps1](#office_removalps1)** | Remove todas as versões do Office e Outlook: desinstala, apaga registros, perfis, AppData e temporários, tornando o sistema "zerado" de Office (ação destrutiva e irreversível). |

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

* **Funcionalidades**: Busca interativa, dois modos (usuário único ou domínio completo), suporte a wildcards, detecção de duplicatas, relatório CSV automático.
* **Pré-requisitos**: Módulo Microsoft.Graph (Authentication, Users, Files), permissões de API (User.Read.All, Files.Read.All, Directory.Read.All), privilégios administrativos.
* **Como Usar**: Execute .\Procura_Arquivos.ps1, escolha modo desejado e siga instruções.

#### Remover_Email.ps1

Realiza a remoção em massa de e-mails específicos de **todas as caixas de correio** do Microsoft 365, com base no remetente e assunto.

* **Funcionalidades**: Remoção global (Soft Delete), processo automatizado via Security & Compliance Center, confirmação crítica, status em tempo real.
* **Pré-requisitos**: Módulo `ExchangeOnlineManagement`, role **Search And Purge**.
* **Como Usar**: Execute `.\Remover_Email.ps1`, informe o remetente e o assunto, e confirme a operação.
    > ⚠️ **Aviso Importante**: Este script afeta TODAS as caixas de correio. Use com extrema cautela.

#### Configura-CatchAll.ps1

Automatiza a configuração de um e-mail **"catch-all"** (coletor geral) para um domínio específico, redirecionando e-mails enviados para endereços inexistentes para uma única caixa de correio.

* **Funcionalidades**: Instalação automática do módulo `ExchangeOnlineManagement`, validação de domínio, altera o tipo do domínio para `InternalRelay`, cria regra de transporte com prioridade dinâmica.
* **Pré-requisitos**: Módulo `ExchangeOnlineManagement`, permissões de Administrador do Exchange.
* **Como Usar**: Execute `.\Configura-CatchAll.ps1` e forneça os dados solicitados.
    > ⚠️ **Aviso Importante**: Propagação da regra pode levar até uma hora.

#### UsarAlias.ps1

Habilita a funcionalidade **"Enviar como Alias"** na organização e entra em menu interativo para adicionar/gerenciar aliases de usuários.

* **Funcionalidades**: Ativação automática de SendFromAliasEnabled no tenant, lista e adição via menu, instruções ao final.
* **Pré-requisitos**: Módulo `ExchangeOnlineManagement`, privilégios administrativos Exchange.
* **Como Usar**: Execute, autentique-se e siga o menu interativo.
    > ⚠️ **Aviso**: A propagação do recurso pode levar algumas horas.

---

### Limpeza & Recuperação

#### office_removal.ps1

Remove **todas as versões do Microsoft Office e Outlook, perfis, registros, cache e arquivos temporários** do Windows, deixando o sistema pronto para instalação limpa ou repasse. A ação é radical e irreversível.

* **Funcionalidades**: 
  - Encerra processos do Office e Outlook.
  - Desinstala qualquer versão detectada.
  - Remove registros em HKLM e HKCU.
  - Exclui pastas em Program Files, AppData e cache.
  - Apaga todos os perfis e arquivos locais do Outlook/Office.
* **Pré-requisitos**: Execução como Administrador, PowerShell 5.1+.
* **Como Usar**:  
  1. Faça backup dos dados importantes (OST/PST/OneNote).
  2. Execute como Administrador:  
     `PowerShell -ExecutionPolicy Bypass -File .\office_removal.ps1`
  3. Aguarde e reinicie o computador.
* **Avisos**: Todos os dados locais e perfis do Office serão perdidos. Não há backup por padrão.
* **Indicações**: Limpeza total antes de migração, troubleshooting crítico, devolução de máquina, reinstalação clean.

---

### 👨‍💻 Autor

**Andre Kittler** - *Administrador Microsoft 365*

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

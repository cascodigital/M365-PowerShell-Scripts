# M365-PowerShell-Scripts

Uma coleção de scripts PowerShell para automação e administração de ambientes Microsoft 365 e infraestrutura local (Active Directory), utilizando os módulos Microsoft.Graph e outros.

---

## Get_MfaComplianceReport.ps1

Este script gera um relatório de conformidade sobre o status do MFA no Microsoft 365, focando em **contas de usuários reais** e, ao final, entra em um **modo de consulta interativo** para análise detalhada de contas específicas.

### Funcionalidades Principais
- **Filtragem Inteligente de Usuários:** Exclui contas de serviço, sincronização e sistemas para focar o relatório em pessoas, tornando a análise de conformidade mais precisa.
- **Geração de Relatório Duplo:** Cria dois arquivos de saída:
    -   Um **`.csv`** com todos os dados técnicos para análise.
    -   Um **`.txt`** formatado como um relatório executivo, ideal para apresentar a clientes ou gestores.
- **Análise Detalhada dos Métodos:** Identifica uma gama completa de métodos de MFA, incluindo **Microsoft Authenticator, Telefone/SMS, Apps OATH, Windows Hello for Business e chaves FIDO2**.
- **Consulta Interativa Pós-relatório:** Após gerar o relatório, o script permite ao administrador **consultar os detalhes de qualquer usuário em tempo real**, exibindo informações específicas como o número de telefone (mascarado), nome do dispositivo vinculado ao Authenticator, data de criação do método, etc.
- **Sumário Visual no Console:** Apresenta um resumo colorido e direto no terminal, mostrando o percentual de conformidade e listando usuários com e sem MFA.

### Pré-requisitos
- Módulo PowerShell `Microsoft.Graph.PowerShell` instalado (`Install-Module -Name Microsoft.Graph`).
- Permissões de API do Microsoft Graph: O administrador que executa o script precisa consentir com as permissões **User.Read.All**, **UserAuthenticationMethod.Read.All**, **Directory.Read.All** e **Policy.Read.All**.

### Como Usar
1.  Baixe o script e renomeie-o para `Get-MfaComplianceReport.ps1`.
2.  **Importante:** Abra o script e personalize os filtros de exclusão na seção `# 2. Obter usuários` para corresponder aos padrões de nomes de contas de serviço da sua organização.
3.  Abra o PowerShell e execute o script: `.\Get-MfaComplianceReport.ps1`
4.  Após a geração dos relatórios, o script entrará no modo de consulta. Digite o email de um usuário para ver seus métodos de MFA em detalhe ou pressione ENTER para sair.

---

## Relacao_Confianca.ps1

Este script verifica o status da relação de confiança (trust relationship) de **todos** os computadores ativos em um domínio Active Directory local. Ele gera um relatório detalhado em formato Excel (.xlsx), identificando máquinas offline e aquelas com a relação de confiança quebrada.

### Funcionalidades Principais
- **Verificação Abrangente:** Testa todos os computadores com status "Enabled" no Active Directory.
- **Diagnóstico Preciso:** Diferencia computadores offline (não respondem ao ping) de computadores com falha na relação de confiança.
- **Uso de `nltest.exe`:** Utiliza o comando nativo `nltest` para verificar a confiança, eliminando a necessidade de WinRM (PowerShell Remoting).
- **Relatório em Excel:** Exporta os resultados para um arquivo `.xlsx` formatado, com data e hora no nome.
- **Cálculo de Inatividade:** Para máquinas offline, estima há quantos dias estão sem comunicação com base na propriedade `LastLogonDate`.
- **Resumo no Console:** Exibe um resumo rápido dos resultados (OK, Falha, Offline) diretamente no terminal após a conclusão.

### Pré-requisitos
- Execução em um computador ingressado no domínio ou com as ferramentas RSAT do Active Directory instaladas.
- Módulo PowerShell `ActiveDirectory`.
- Módulo PowerShell `ImportExcel` instalado (`Install-Module -Name ImportExcel`).

### Como Usar
1.  Baixe o script `Relacao-Confianca.ps1`.
2.  (Opcional) Edite a variável `$reportFolder` no início do script para alterar o local onde o relatório será salvo.
3.  Abra o PowerShell como Administrador em um controlador de domínio ou máquina com as ferramentas do AD.
4.  Execute o script: `.\Relacao-Confianca.ps1`
5.  O relatório será gerado na pasta `C:\Temp` (ou no local configurado).

---

## Procura_Arquivos.ps1

Este script interativo localiza arquivos por nome no OneDrive for Business de um usuário específico, utilizando o Microsoft Graph para realizar a busca de forma eficiente.

### Funcionalidades Principais
- **Busca Interativa:** Solicita o e-mail do usuário alvo e o nome/filtro do arquivo diretamente no console.
- **Suporte a Wildcards:** Permite buscas flexíveis utilizando asteriscos (`*`).
- **Múltiplos Drives:** O script identifica e pesquisa em todos os drives associados ao usuário.
- **Saída Detalhada:** Exibe o caminho completo do arquivo, drive de origem, tamanho, data da última modificação e o ID do item.
- **Detecção de Duplicatas:** Alerta o administrador caso arquivos com o mesmo nome sejam encontrados.

### Pré-requisitos
- Módulo PowerShell `Microsoft.Graph.PowerShell` instalado (`Install-Module -Name Microsoft.Graph`).
- Permissões de API do Microsoft Graph: **User.Read.All** e **Files.Read.All**.

### Como Usar
1.  Baixe o script `procura_arquivos.ps1`.
2.  Abra o PowerShell e execute o script: `.\procura_arquivos.ps1`
3.  Siga as instruções no console.

---

## Search-Events.ps1

Este script PowerShell foi criado para simplificar e acelerar a análise de Logs de Eventos do Windows, permitindo a busca por múltiplos Event IDs em um intervalo de datas e exportando os resultados para um arquivo Excel (.xlsx).

### Funcionalidades Principais
- **Busca por Múltiplos Event IDs:** Permite ao usuário inserir um ou mais IDs de evento de uma só vez.
- **Filtro por Intervalo de Datas:** A busca é restrita a um período de início e fim definido pelo usuário.
- **Busca Abrangente:** O script varre TODOS os logs de eventos do sistema (`.evtx`).
- **Exportação Detalhada para Excel:** Organiza os resultados em um arquivo Excel, com abas separadas para cada Event ID.

### Pré-requisitos
- Execução como Administrador.
- Módulo PowerShell `ImportExcel` instalado (`Install-Module -Name ImportExcel`).

### Como Usar
1.  Baixe o script `Search-DetailedWinEvent.ps1`.
2.  Abra o PowerShell como Administrador.
3.  Execute o script e siga as instruções no console.

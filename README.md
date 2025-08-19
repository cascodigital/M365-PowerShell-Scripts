# M365-PowerShell-Scripts

Uma coleção de scripts PowerShell para automação e administração de ambientes Microsoft 365 e infraestrutura local (Active Directory), utilizando os módulos Microsoft.Graph e outros.

---

## Relacao_Confianca.ps1

Este script verifica o status da relação de confiança (trust relationship) de **todos** os computadores ativos em um domínio Active Directory local. Ele gera um relatório detalhado em formato Excel (.xlsx), identificando máquinas offline e aquelas com a relação de confiança quebrada.

### Funcionalidades Principais
- **Verificação Abrangente:** Testa todos os computadores com status "Enabled" no Active Directory.
- **Diagnóstico Preciso:** Diferencia computadores offline (não respondem ao ping) de computadores com falha na relação de confiança.
- **Uso de `nltest.exe`:** Utiliza o comando nativo `nltest` para verificar a confiança, eliminando a necessidade de WinRM (PowerShell Remoting).
- **Relatório em Excel:** Exporta os resultados para um arquivo `.xlsx` formatado, com data e hora no nome, incluindo colunas para status, detalhes e data do último logon.
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
- **Suporte a Wildcards:** Permite buscas flexíveis utilizando asteriscos (`*`). Se nenhum wildcard for usado, o script o adiciona automaticamente para uma busca mais ampla.
- **Múltiplos Drives:** O script identifica e pesquisa em todos os drives associados ao usuário (ex: OneDrive pessoal, drives de sites do SharePoint vinculados).
- **Saída Detalhada:** Exibe o caminho completo do arquivo, drive de origem, tamanho, data da última modificação e o ID do item.
- **Detecção de Duplicatas:** Alerta o administrador caso arquivos com o mesmo nome sejam encontrados em locais diferentes.
- **Conexão Inteligente:** Verifica se já existe uma conexão ativa com o Microsoft Graph antes de solicitar uma nova.

### Pré-requisitos
- Módulo PowerShell `Microsoft.Graph.PowerShell` instalado (`Install-Module -Name Microsoft.Graph`).
- Permissões de API do Microsoft Graph: O administrador que executa o script precisa consentir com as permissões **User.Read.All** e **Files.Read.All**. O script solicitará isso na primeira conexão.

### Como Usar
1.  Baixe o script `procura_arquivos.ps1`.
2.  Abra o PowerShell.
3.  Execute o script: `.\procura_arquivos.ps1`
4.  Siga as instruções no console para inserir o e-mail do usuário e o filtro de busca para o nome do arquivo.

---

## Search_Events.ps1

Este script PowerShell foi criado para simplificar e acelerar a análise de Logs de Eventos do Windows, permitindo a busca por múltiplos Event IDs em um intervalo de datas e exportando os resultados para um arquivo Excel (.xlsx) de fácil leitura.

### Funcionalidades Principais
- **Busca por Múltiplos Event IDs:** Permite ao usuário inserir um ou mais IDs de evento (ex: 4624, 4625) de uma só vez.
- **Filtro por Intervalo de Datas:** A busca é restrita a um período de início e fim definido pelo usuário.
- **Busca Abrangente:** O script varre TODOS os logs de eventos do sistema (`.evtx`), não apenas os mais comuns.
- **Exportação Detalhada para Excel:** Extrai todas as propriedades de cada evento encontrado e as organiza em um arquivo Excel, com abas separadas para cada Event ID e uma aba consolidada.
- **Validação e Segurança:** O script verifica se está sendo executado com privilégios de administrador e se o módulo `ImportExcel` está instalado.

### Pré-requisitos
- Execução como Administrador.
- Módulo PowerShell `ImportExcel` instalado (`Install-Module -Name ImportExcel`).

### Como Usar
1.  Baixe o script `Search-DetailedWinEvent.ps1`.
2.  Abra o PowerShell como Administrador.
3.  Execute o script e siga as instruções no console para inserir os Event IDs e as datas.

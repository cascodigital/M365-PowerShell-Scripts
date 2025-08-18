# M365-PowerShell-Scripts

Uma coleção de scripts PowerShell para automação e administração de ambientes Microsoft 365, utilizando o módulo Microsoft.Graph.

---

## Search-EVENTS.ps1

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

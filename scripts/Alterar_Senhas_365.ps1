# Repositório de Scripts PowerShell para Microsoft 365

Este repositório contém uma coleção de scripts PowerShell projetados para facilitar a administração e a análise em ambientes Microsoft 365 e Active Directory. Cada script é documentado com suas funcionalidades, requisitos e exemplos de uso.

## Índice

- [1. Coleta de Eventos](#1-coleta-de-eventos)
- [2. Monitoramento de Rede](#2-monitoramento-de-rede)
- [3. Gerenciamento de Senhas](#3-gerenciamento-de-senhas)
- [4. Configuração de Regras de Email](#4-configuração-de-regras-de-email)
- [5. Remoção de Aplicativos](#5-remoção-de-aplicativos)

---

## 1. Coleta de Eventos

### `Buscar_Logon.ps1`
Coleta e analisa eventos de logon de usuários em computadores do domínio Active Directory. Este script busca eventos de logon humano (EventID 4624) e permite filtrar por período de datas.

- **Funcionalidades:**
  - Busca em computadores específicos ou em todo o domínio.
  - Exportação automática para Excel com formatação profissional.
  - Relatório detalhado de coleta com estatísticas.

---

## 2. Monitoramento de Rede

### `monitor-ping.ps1`
Realiza monitoramento contínuo de latência via ICMP (ping) para múltiplos destinos. O script exibe resultados coloridos no console e grava um relatório CSV em tempo real.

- **Funcionalidades:**
  - Monitoramento de até 5 endereços IP.
  - Classificação de latências por níveis (VERDE, AMARELO, VERMELHO).
  - Relatório em CSV com timestamp e status.

---

## 3. Gerenciamento de Senhas

### `Alterar_Senhas_365.ps1`
Gera e aplica senhas aleatórias para usuários Microsoft 365 por domínio. O script utiliza o módulo Microsoft.Graph para conectividade moderna.

- **Funcionalidades:**
  - Geração de senhas no formato memorável (AA1234qq).
  - Filtragem automática por domínio específico.
  - Relatório detalhado de sucessos e falhas.

---

## 4. Configuração de Regras de Email

### `Configura-CatchAll.ps1`
Configura uma regra catch-all para domínios Microsoft 365, permitindo redirecionar emails enviados para endereços inexistentes.

- **Funcionalidades:**
  - Criação de grupo de distribuição dinâmico.
  - Criação de regra de transporte catch-all com exceção para membros do grupo.
  - Logs coloridos no terminal para acompanhamento.

---

## 5. Remoção de Aplicativos

### `office_removal.ps1`
Remove completamente todas as versões do Microsoft Office e Outlook do sistema Windows.

- **Funcionalidades:**
  - Desinstalação silenciosa de todas as versões do Office.
  - Remoção de chaves de registro e perfis do Outlook.
  - Limpeza total de pastas de programa e arquivos temporários.

---

## Contribuições

Sinta-se à vontade para contribuir com melhorias ou sugestões. Para isso, basta abrir uma issue ou enviar um pull request.

## Licença

Este projeto está licenciado sob a [Licença MIT](LICENSE).

---

## Contato

Para mais informações, entre em contato com [Seu Nome](mailto:seuemail@dominio.com).
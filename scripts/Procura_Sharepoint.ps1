# Repositório de Scripts PowerShell para Microsoft 365

Este repositório contém uma coleção de scripts PowerShell projetados para facilitar a administração e a análise em ambientes Microsoft 365 e Active Directory. Cada script é autossuficiente e possui funcionalidades específicas que atendem a diferentes necessidades administrativas.

## Índice

- [1. Coleta de Eventos de Logon](#1-coleta-de-eventos-de-logon)
- [2. Alteração de Senhas em Massa](#2-alteração-de-senhas-em-massa)
- [3. Monitoramento de Latência de Rede](#3-monitoramento-de-latência-de-rede)
- [4. Configuração de Regra Catch-All](#4-configuração-de-regra-catch-all)
- [5. Remoção Completa do Microsoft Office](#5-remoção-completa-do-microsoft-office)
- [6. Localizador Avançado de Arquivos no OneDrive](#6-localizador-avançado-de-arquivos-no-onedrive)
- [7. Analisador de Logs de Eventos Windows](#7-analisador-de-logs-de-eventos-windows)

## 1. Coleta de Eventos de Logon

**Arquivo:** `Buscar_Logon.ps1`

Este script coleta e analisa eventos de logon de usuários em computadores do domínio Active Directory. Ele busca eventos de logon humano (EventID 4624) e permite filtrar por períodos específicos.

### Funcionalidades:
- Busca em computadores específicos ou em todo o domínio.
- Exportação automática para Excel com formatação profissional.
- Relatório detalhado com estatísticas de coleta.

---

## 2. Alteração de Senhas em Massa

**Arquivo:** `Alterar_Senhas_365.ps1`

Este script gera e aplica senhas aleatórias para usuários do Microsoft 365/Azure AD. As senhas são geradas em um formato seguro e memorável.

### Funcionalidades:
- Geração de senhas no formato `AA1234qq`.
- Filtragem automática por domínio específico.
- Relatório detalhado de sucessos e falhas.

---

## 3. Monitoramento de Latência de Rede

**Arquivo:** `monitor-ping.ps1`

Este script realiza pings em múltiplos endereços IP, exibindo resultados coloridos no console e gravando um relatório CSV em tempo real.

### Funcionalidades:
- Monitoramento contínuo de latência via ICMP.
- Classificação de latências por níveis (verde, amarelo, vermelho).
- Relatório em CSV com timestamp e status.

---

## 4. Configuração de Regra Catch-All

**Arquivo:** `Configura-CatchAll.ps1`

Este script automatiza a configuração de uma regra catch-all para domínios Microsoft 365, criando um grupo dinâmico de exceção.

### Funcionalidades:
- Criação de grupo de distribuição dinâmico.
- Regra de transporte catch-all para redirecionar emails enviados para endereços inexistentes.
- Logs detalhados no terminal.

---

## 5. Remoção Completa do Microsoft Office

**Arquivo:** `office_removal.ps1`

Este script remove completamente todas as versões do Microsoft Office e Outlook do sistema Windows.

### Funcionalidades:
- Desinstalação silenciosa de todas as versões do Office.
- Remoção de chaves de registro e perfis do Outlook.
- Limpeza total de pastas de programa e arquivos temporários.

---

## 6. Localizador Avançado de Arquivos no OneDrive

**Arquivo:** `Procura_Arquivos.ps1`

Este script permite buscar arquivos no OneDrive de um usuário específico ou em todos os OneDrives de um domínio.

### Funcionalidades:
- Busca em OneDrive de usuário específico ou em todos os OneDrives do domínio.
- Relatório detalhado dos arquivos encontrados.

---

## 7. Analisador de Logs de Eventos Windows

**Arquivo:** `Procura_Eventos.ps1`

Este script realiza uma análise forense de logs de eventos do Windows, permitindo busca por múltiplos Event IDs.

### Funcionalidades:
- Busca simultânea de múltiplos Event IDs.
- Exportação para Excel com abas separadas por tipo de evento.
- Filtragem por intervalo de datas.

---

## Contribuições

Sinta-se à vontade para contribuir com melhorias ou sugestões. Para isso, basta abrir uma issue ou enviar um pull request.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).
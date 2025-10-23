# Repositório de Scripts PowerShell para Microsoft 365

Este repositório contém uma coleção de scripts PowerShell projetados para facilitar a administração e a análise em ambientes Microsoft 365 e Active Directory. Cada script é autossuficiente e possui funcionalidades específicas que atendem a diferentes necessidades administrativas.

## Índice

- [1. Coleta de Eventos de Logon](#1-coleta-de-eventos-de-logon)
- [2. Alteração de Senhas em Massa](#2-alteração-de-senhas-em-massa)
- [3. Monitoramento de Latência de Rede](#3-monitoramento-de-latência-de-rede)
- [4. Configuração de Regras Catch-All](#4-configuração-de-regras-catch-all)
- [5. Remoção Completa do Microsoft Office](#5-remoção-completa-do-microsoft-office)
- [6. Localizador de Arquivos no OneDrive](#6-localizador-de-arquivos-no-onedrive)
- [7. Analisador de Logs de Eventos do Windows](#7-analisador-de-logs-de-eventos-do-windows)

## 1. Coleta de Eventos de Logon

**Arquivo:** `Buscar_Logon.ps1`

Este script coleta e analisa eventos de logon de usuários em computadores do domínio Active Directory. Ele busca eventos de logon humano (EventID 4624) e permite filtrar por períodos específicos.

### Funcionalidades:
- Busca em computadores específicos ou em todo o domínio.
- Exportação dos resultados para um arquivo Excel formatado.

---

## 2. Alteração de Senhas em Massa

**Arquivo:** `Alterar_Senhas_365.ps1`

Este script gera e aplica senhas aleatórias para usuários do Microsoft 365 em um domínio específico. As senhas são geradas em um formato seguro e memorável.

### Funcionalidades:
- Geração de senhas no formato `AA1234qq`.
- Processamento em massa com controle de throttling.
- Relatório detalhado de sucessos e falhas.

---

## 3. Monitoramento de Latência de Rede

**Arquivo:** `monitor-ping.ps1`

Este script realiza pings em múltiplos destinos e monitora a latência de rede, exibindo resultados coloridos no console e gravando um relatório em tempo real.

### Funcionalidades:
- Monitoramento contínuo de até 5 endereços IP.
- Relatório em CSV com estatísticas de latência.

---

## 4. Configuração de Regras Catch-All

**Arquivo:** `Configura-CatchAll.ps1`

Este script automatiza a configuração de uma regra catch-all para domínios Microsoft 365, criando um grupo dinâmico de exceção.

### Funcionalidades:
- Criação de grupo de distribuição dinâmico.
- Implementação de regra de transporte para redirecionar e-mails enviados para endereços inexistentes.

---

## 5. Remoção Completa do Microsoft Office

**Arquivo:** `office_removal.ps1`

Este script remove completamente todas as versões do Microsoft Office e Outlook do sistema Windows, incluindo aplicativos, registros e perfis.

### Funcionalidades:
- Desinstalação silenciosa de todas as versões do Office.
- Limpeza de registros e arquivos temporários.

---

## 6. Localizador de Arquivos no OneDrive

**Arquivo:** `Procura_Arquivos.ps1`

Este script permite buscar arquivos no OneDrive de um usuário específico ou em todos os OneDrives de um domínio.

### Funcionalidades:
- Busca em OneDrive de usuário específico ou em todos os usuários do domínio.
- Relatório detalhado dos arquivos encontrados.

---

## 7. Analisador de Logs de Eventos do Windows

**Arquivo:** `Procura_Eventos.ps1`

Este script analisa logs de eventos do Windows, permitindo busca por múltiplos Event IDs e exportação dos resultados para Excel.

### Funcionalidades:
- Busca simultânea de múltiplos Event IDs.
- Exportação estruturada para Excel com abas separadas por tipo de evento.

---

## Contribuições

Sinta-se à vontade para contribuir com melhorias ou sugestões. Para isso, basta abrir uma issue ou enviar um pull request.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).
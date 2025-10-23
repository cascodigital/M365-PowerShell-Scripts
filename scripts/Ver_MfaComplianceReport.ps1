# Repositório de Scripts PowerShell para Microsoft 365

Este repositório contém uma coleção de scripts PowerShell projetados para facilitar a administração e a análise em ambientes Microsoft 365 e Active Directory. Cada script é documentado com suas funcionalidades, exemplos de uso e requisitos.

## Índice

- [1. Coleta de Eventos](#1-coleta-de-eventos)
- [2. Gerenciamento de Senhas](#2-gerenciamento-de-senhas)
- [3. Monitoramento de Rede](#3-monitoramento-de-rede)
- [4. Configuração de Regras de Email](#4-configuração-de-regras-de-email)
- [5. Localização de Arquivos no OneDrive](#5-localização-de-arquivos-no-onedrive)
- [6. Remoção do Microsoft Office](#6-removal-do-microsoft-office)

## 1. Coleta de Eventos

### `Buscar_Logon.ps1`
Coleta e analisa eventos de logon de usuários em computadores do domínio Active Directory. Filtra eventos de logon humano (EventID 4624) e exporta os resultados para um arquivo Excel.

**Funcionalidades:**
- Busca em computadores específicos ou em todo o domínio.
- Filtragem por período de datas.
- Tratamento de timeouts e máquinas offline.

---

## 2. Gerenciamento de Senhas

### `Alterar_Senhas_365.ps1`
Gera e aplica senhas aleatórias para usuários do Microsoft 365. As senhas são criadas em um formato seguro e memorável.

**Funcionalidades:**
- Geração de senhas no formato AA1234qq.
- Filtragem por domínio específico.
- Relatório detalhado de sucessos e falhas.

---

## 3. Monitoramento de Rede

### `monitor-ping.ps1`
Realiza monitoramento contínuo de latência via ICMP (ping) para múltiplos destinos, exibindo resultados coloridos no console e gravando um relatório CSV em tempo real.

**Funcionalidades:**
- Monitoramento de até 5 endereços IP.
- Classificação de latências por níveis (verde, amarelo, vermelho).
- Relatório em tempo real.

---

## 4. Configuração de Regras de Email

### `Configura-CatchAll.ps1`
Configura uma regra catch-all para domínios Microsoft 365, redirecionando emails enviados para endereços inexistentes.

**Funcionalidades:**
- Criação de grupo de distribuição dinâmico.
- Exceção para membros do grupo dinâmico.
- Logs detalhados no terminal.

---

## 5. Localização de Arquivos no OneDrive

### `Procura_Arquivos.ps1`
Localiza arquivos no OneDrive de usuários específicos ou em todos os OneDrives de um domínio.

**Funcionalidades:**
- Busca em OneDrive de usuário específico ou em todos os OneDrives.
- Relatório detalhado dos arquivos encontrados.

---

## 6. Remoção do Microsoft Office

### `office_removal.ps1`
Remove completamente todas as versões do Microsoft Office e Outlook do sistema Windows.

**Funcionalidades:**
- Desinstalação silenciosa de todas as versões do Office.
- Remoção de chaves de registro e perfis do Outlook.
- Limpeza total de pastas e arquivos temporários.

---

## Como Usar

1. **Clone o repositório:**
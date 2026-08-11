# [UL] Dashboard_Taxonomias

Dashboard operacional de taxonomias Unilever BR — Grasp x StormX.

## Acesso
- **GitHub Pages:** https://falssp.github.io/ul-dashboard-taxonomias
- **Web App (API):** https://script.google.com/macros/s/AKfycbzz35AaDNYnhJ9HlYR_GUuk_UXMxIyiu0L3joPXW0lrA2P3l6tS2ReGVFWyOM8L6Hzhmg/exec

## Arquivos
| Arquivo | Onde vai | Descrição |
|---|---|---|
| `index.html` | GitHub Pages (raiz do repo) | Painel web — carrega dados via fetch |
| `01_Config.gs` | Apps Script → config | Menu, credenciais, popups, avisos de token |
| `02_Sync.gs` | Apps Script → sync | Busca paginada, buffer, gravação RAW |
| `03_Abas.gs` | Apps Script → abas | Gravação e formatação das abas |
| `04_WebApp.gs` | Apps Script → webapp | doGet (JSON + HTML), doPost (exportar) |

## Setup
1. Planilha Google Sheets → Extensões → Apps Script
2. Cola os 4 `.gs` nos arquivos correspondentes
3. Cria arquivo HTML `Painel` e cola o conteúdo de `index.html` — **apenas para o modo GAS**
4. Implantar → Nova implantação → Web App → Executar como: Eu → Qualquer pessoa
5. Menu 📋 Dashboard UL → 1. Configurar credenciais

## GitHub Pages
Push do `index.html` na raiz → Settings → Pages → Source: main / root → URL ativa em ~1min

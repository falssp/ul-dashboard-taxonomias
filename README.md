# [UL] Dashboard_Taxonomias

Dashboard operacional de taxonomias Unilever BR — Grasp x StormX.  
Roda inteiramente no Google Apps Script, sem servidor externo.

---

## Visão geral

Sincroniza issues do Jira (projeto `UL`) para uma planilha Google Sheets e expõe um painel web interativo com filtros, KPIs e exportação.

```
Jira (4967 issues)
  └── sincronizarCompleto / sincronizarJira (a cada 4h)
        └── _SYNC_BUFFER (aba oculta)
              └── RAW_PAI + RAW_FILHO + INCORRETOS
                    └── PAINEL → VISUAL → TABELA
                          └── Painel.html (Web App)
```

---

## Contagem de issues

| Categoria | Quantidade | O que são |
|---|---|---|
| Tickets de trabalho | 4.470 | Subtarefas + tickets solo (sem filhos) |
| Campanhas | 486 | Tarefas pai que têm subtarefas (agrupadores) |
| Incorretos | 11 | Tipo inválido (Epic, História etc) |
| **Total Jira** | **4.967** | `tickets + campanhas + incorretos` |

---

## Arquivos

| Arquivo | Responsabilidade |
|---|---|
| `01_Config.gs` | Menu, credenciais, popups, acionadores, avisos de token |
| `02_Sync.gs` | Busca paginada encadeada, buffer, gravação RAW, etapa 2 |
| `03_Abas.gs` | Gravação e formatação de todas as abas da planilha |
| `04_WebApp.gs` | `doGet`, `getDados`, `exportarDados`, `formatarPlanilha` |
| `Painel.html` | Painel web — filtros, KPIs, árvore BU→Agência→Marca→Campanha |
| `appsscript.json` | Manifest do Apps Script (timezone, scopes OAuth) |

---

## Abas da planilha

| Aba | Descrição |
|---|---|
| 🏠 INICIO | Instruções de uso |
| PAINEL | Visão consolidada pai+filho com veículo detectado |
| 📊 VISUAL | Árvore hierárquica colorida por BU |
| 📋 TABELA | Tabela filtrável igual ao PAINEL |
| RAW_PAI | Issues tipo Tarefa (pais) brutas do Jira |
| RAW_FILHO | Issues tipo Subtarefa brutas do Jira |
| ⚠️ INCORRETOS | Issues com tipo inválido (Epic, História etc) |
| DE_PARA | Dicionário de normalização de marca e veículo |
| _SYNC_BUFFER | Buffer temporário da busca paginada — sempre oculta |

---

## Arquitetura do sync

O sync usa **triggers one-shot encadeados** para contornar o limite de 6 minutos do Apps Script:

```
sincronizarCompleto()
  → _executarPaginaBusca()  [até 4.5min por execução]
      → _buscarProximaPagina()  [repete quantas vezes precisar]
          → _gravarRAWDoBusfer()  [lê buffer, grava RAW]
              → _sincronizarEtapa2()  [PAINEL + VISUAL + TABELA]
```

**Dupla garantia do nextPageToken:** `PropertiesService` (primário) + `CacheService` (secundário, 6h).  
Se ambos perderem o token, reinicia o sync do zero automaticamente.

---

## Avisos de expiração do token

O token de API do Jira tem validade de 1 ano. O sistema avisa em 3 camadas:

| Quando | Como |
|---|---|
| ≤ 10 dias | Banner amarelo no painel |
| ≤ 5 dias | E-mail automático (1 por dia) + banner laranja |
| ≤ 3 dias | Alert ao abrir a planilha + banner vermelho |
| Expirado | Alert bloqueante + banner vermelho + e-mail |

Para renovar: `id.atlassian.com → Segurança → Tokens de API`  
Depois: **Menu 📋 Dashboard UL → 1. Configurar credenciais**

---

## Instalação / configuração inicial

1. Abra a planilha Google Sheets
2. **Extensões → Apps Script**
3. Cole cada arquivo `.gs` no arquivo correspondente
4. Cole `Painel.html` como arquivo HTML
5. Substitua o conteúdo de `appsscript.json` pelo arquivo deste repo
6. **Implantar → Nova implantação → Web App**
   - Executar como: **Eu**
   - Quem tem acesso: **Qualquer pessoa na organização** (ou Qualquer pessoa)
7. Na planilha: **Menu 📋 Dashboard UL → 1. Configurar credenciais**
8. Insira e-mail Atlassian + API Token + data de validade
9. Aguarde o sync completo (~10 minutos para 5.000 issues)

---

## Uso do clasp (opcional)

Para sincronizar o código diretamente do terminal:

```bash
npm install -g @google/clasp
clasp login
clasp clone <SCRIPT_ID>   # ID do Apps Script (URL do editor)
```

Criar `.clasp.json` na raiz:
```json
{
  "scriptId": "<SEU_SCRIPT_ID>",
  "rootDir": "."
}
```

Publicar alterações:
```bash
clasp push
```

> **Nota:** `.clasp.json` está no `.gitignore` — nunca commitar o Script ID publicamente.

---

## Histórico de versões

| Versão | Data | Descrição |
|---|---|---|
| v1.0 | jul/2026 | Versão inicial — sync simples, painel básico |
| v2.0 | jul/2026 | Sync encadeado sem timeout, buffer oculto |
| v2.1 | jul/2026 | Contagem correta (tickets vs campanhas vs incorretos) |
| v2.2 | jul/2026 | Popups assíncronos em todos os itens do menu |
| v2.3 | jul/2026 | Avisos de expiração de token (3 camadas) |

---

## Equipe

**Grasp x StormX** — AdOps Unilever BR  
Contato: `felipe.lima@stormx.com.br`

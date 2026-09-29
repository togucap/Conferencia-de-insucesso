# Controle de Insucessos — K-OUT Auditoria Express

Web App em **Google Apps Script** (runtime V8) para conferência física de cargas de insucesso
num CD: o operador bipa etiquetas, o sistema cruza com a carga faturada, com a *Lista de Chegada*
(entrada sistémica RMA) e com o relatório *Monte o Seu (MSPC)*, e produz o veredito por item
(recebido, falta confirmada, furo de processo, bloqueio, divergência).

A planilha Google (a "Planilha Base") funciona como base de dados; cada carga é uma aba.

---

## Estrutura do repositório

```
.
├── .clasp.json            # scriptId + rootDir=src (para clasp push/pull)
├── README.md
└── src/                   # Ficheiros do projeto Apps Script, com os nomes originais
    ├── appsscript.json    # Manifesto (timezone, scopes, Sheets API v4, webapp)
    ├── Código.gs          # Backend (servidor)
    ├── Index.html         # Shell da página + todos os modais + includes
    ├── Estilos.html       # CSS (tema dark/light por variáveis)
    ├── View_PainelGerencial.html  # Markup do Painel Gerencial (filtros, KPIs, gráficos)
    ├── Script_Estado.html     # Referências DOM, estado global, arranque, dropdown de cargas, loader
    ├── Script_UI.html         # Tema, Tabela⇄Bipar, navegação SPA, gráficos Chart.js
    ├── Script_Dados.html      # Carregar carga, renderizar tabela, progresso, filtro da tabela
    ├── Script.html            # Motor de bipagem (evento Enter no input da etiqueta)
    ├── Script_Metricas.html   # Relatório "Detalhes", resumo final, finalizar/reabrir
    ├── Script_Exportacao.html # Exportação .xlsx (Consolidado / Unitários / Múltiplos)
    ├── Script_Upload.html     # Upload "Nova Carga" e "Lista Chegada" (lidos no cliente com SheetJS)
    ├── Script_Bloqueio.html   # Bloqueio em massa (marcar "Não Receber" por Pedido ou NF)
    └── Script_Filtros.html    # Filtros do Painel Gerencial (Carga/Mês/Ano)
```

> Os ficheiros `.html` do tipo `Script_*` são `<script>` puros incluídos em `Index.html` via
> `<?!= include('Nome'); ?>`. **A ordem de inclusão importa** (partilham o escopo global):
> `Script_Estado → Script_UI → Script_Dados → Script_Metricas → Script_Exportacao →
> Script_Upload → Script_Bloqueio → Script_Filtros → Script`.

Deploy com [clasp](https://github.com/google/clasp): `clasp push` a partir da raiz (usa `rootDir: src`).

---

## Modelo de dados (abas da planilha)

| Aba | Origem | Conteúdo |
|---|---|---|
| `DD-MM-AAAA` (e `DD-MM-AAAA Vn`) | Upload "Nova Carga" | Uma carga. Cor da aba `#10b981` = **finalizada** |
| `Base_Lista_Chegada` | Upload manual / importador do Drive | Etiquetas que deram entrada sistémica (RMA) |
| `Base_Historico_Eventos` | Gerada na finalização | `ID Único (Pedido\|Etiqueta)` → JSON `{ultimoStatus, ultimaCarga, eventos[]}` |

**Colunas obrigatórias numa carga:** `Mercadoria Código`, `Número do Pedido Faturado`
(com ou sem acento; o match é por cabeçalho em minúsculas).
**Colunas opcionais usadas:** `Nota Fiscal`/`NF`, `Avaliação`, coluna de nome/descrição do produto.
**Colunas criadas pelo sistema:** `Status Conferência`, `Data/Hora Conferência`,
`Conferencia MSPC` (no upload), `Status Lista Chegada`, `Avaliação` (se não existir).

**Chave composta** usada em todos os cruzamentos: `Pedido|Etiqueta`.

**Valores de status:** `Conferido`, `Bloqueado`, `Divergência (Fora da Carga)`;
Lista: `Recebido Lista`; Avaliação: `Não Receber`;
Histórico: `Bloqueado`, `Falta Confirmada`, `Falta Resolvida (Sistémica)`.

---

## Backend — `Código.gs` (funções chamadas via `google.script.run`)

| Secção | Função | Chamada por | O que faz |
|---|---|---|---|
| 1 | `doGet`, `include` | Web App | Serve `Index` |
| 2 | `getAbasPlanilha()` | Estado, Upload | Lista abas `DD-MM-AAAA`, abertas primeiro, mais recentes primeiro |
| 2 | `getDadosDaAba(aba)` | Dados | Devolve matriz; `bloqueio` se faltar coluna obrigatória ou houver linha com NF sem produto/pedido |
| 2 | `getErrosDaCarga(aba)` | Dados | Divergências já gravadas (para repor o painel "Fora da Carga") |
| 3 | `registrarConferencia` / `registrarBloqueio` | Script | Grava status + data/hora na linha (`_atualizarCelulaStatus`) |
| 3 | `registrarErroNaMesmaAba` | Script | Acrescenta linha vermelha "Divergência (Fora da Carga)" |
| 3 | `finalizarCargaStatus(aba)` | Métricas | Pinta aba de verde, sincroniza com Lista Chegada, alimenta histórico; devolve chaves recebidas |
| 3 | `reabrirCargaStatus(aba)` | Métricas | Remove cor da aba |
| 4 | `uploadNovaCarga(nome, json)` | Upload | Cria aba (sufixo `Vn` se já existir) e injeta coluna `Conferencia MSPC` via `_obterMapaMonteOSeu` |
| 4 | `aplicarRecusasEmMassa(aba, tipo, valores)` | Bloqueio | Marca `Avaliação = Não Receber` por Pedido ou NF |
| 5 | `getDashboardMetrics()` | UI (painel) | Por carga: pedidos (total/100% bipados), volumes (total/conferidos), recusas (total/realizadas) |
| 6 | `importarRelatorioListaChegada()` | *Nenhuma no front* (execução manual/acionador) | Lê CSV "Lista MSPC" do Drive em blocos de 5 MB, incremental |
| 7 | `processarUploadListaChegada(json)` | Upload | Mesmo fluxo, a partir do ficheiro enviado pelo utilizador |
| 8 | `_obterSetGlobalListaChegada`, `_sincronizarAbaComSetListaChegada`, `sincronizarStatusListaChegadaGlobal`, `getTagsRecebidasPorAba` | interno / Métricas | Marca `Status Lista Chegada = Recebido Lista` nas cargas |
| 9 | `_obterAbaHistorico`, `_atualizarHistoricoNaFinalizacao`, `_autoCurarHistorico`, `obterHistoricoEventosFrontend` | interno / Estado | "Cérebro histórico" JSON (reincidência / falta recuperada) |
| 10 | `_obterMapaMonteOSeu()` | `uploadNovaCarga` | Lê CSV "Monte o Seu - Consolidado" do Drive → `{pedido: tipo/kit}` |

Recursos externos fixos no código: pasta Drive da Lista MSPC (`12d0yusk…`), pasta Drive do
Monte o Seu (`1_Q8WflH…`) e o link da Planilha Base no rodapé do `Index.html`.

---

## Fluxos principais (frontend)

1. **Arranque** (`Script_Estado`): carrega lista de cargas e o histórico JSON em memória
   (`mapaHistoricoGlobal`, indexado só pela etiqueta).
2. **Iniciar Auditoria** (`Script_Dados`): lê a aba, identifica índices de colunas, conta itens por
   pedido, ordena (múltiplos no topo, agrupados com cores alternadas), marca conferidos/recusas,
   esconde divergências e mostra KPIs + progresso. Se a carga estiver finalizada abre o resumo.
3. **Bipagem** (`Script`): valida 8 dígitos → ignora etiqueta já divergente → consulta histórico
   (reincidente bloqueado / falta recuperada) → procura a 1.ª linha não conferida com essa etiqueta:
   - avaliação "Não Receber" ou reincidente → **NÃO RECEBER** + `registrarBloqueio`
   - encontrada → conferida + alerta de pedido parcial/múltiplo completo + tag MSPC + `registrarConferencia`
   - todas já conferidas → "volumetria esgotada"
   - não existe → **Divergência** + `registrarErroNaMesmaAba`
4. **Detalhes / Finalizar** (`Script_Metricas`): sincroniza com a Lista de Chegada e classifica:
   bipado & na lista · bipado & falta entrada · **furo** (na lista sem bipe) ·
   **falta confirmada** (sem bipe e fora da lista). Cartões clicáveis com drill-down.
5. **Exportar** (`Script_Exportacao`): `.xlsx` com abas Consolidado, Unitários e Múltiplos e
   coluna "Veredito Final".
6. **Uploads** (`Script_Upload`): ficheiro tem de conter `Carga_insucesso` ou `Lista_chegada`
   no nome; leitura com SheetJS (`raw: false`) e envio como string JSON.
7. **Painel Gerencial** (`Script_UI` + `Script_Filtros`): cache das métricas, filtros multi-seleção
   por Carga/Mês/Ano, gráfico de linha (volume) e bullet chart (prometido × auditado).

---

## Pontos de atenção (observados na análise — não alterados)

Registados para considerar quando juntarmos com o próximo código:

- **Globais implícitos por `id`**: `inputEtiqueta`, `resultadoBipe`, `containerTabela`,
  `containerNaoEncontrados`, `listaNaoEncontrados`, `visaoBipar`, `switchVisao`, `textoVisao`,
  `containerSwitchVisao`, `checkboxTheme`, `themeLabel`, `detalhesMicroTitulo`,
  `tabelaDetalhesMicro`, `tabelaDetalhesMicroCabecalho` nunca são declarados — funcionam pelo
  acesso nomeado do browser (`window[id]`). Renomear um `id` ou criar uma variável homónima quebra.
- **Filtro da tabela com dois listeners**: `Script_UI.aplicarFiltroTabela` e o listener de
  `Script_Dados` tratam o mesmo `change`; o de `Script_Dados` (mais completo, com modo `mspc`) é
  o que prevalece.
- **Escritas sem `withFailureHandler`** (`registrarConferencia`, `registrarBloqueio`,
  `registrarErroNaMesmaAba`, `reabrirCargaStatus`): se falharem, a UI mostra sucesso mas a
  planilha não é atualizada.
- **Sem `LockService`**: vários operadores na mesma carga podem criar colunas de status duplicadas
  ou corridas no `appendRow`.
- **Índice de linha** vem de `data-linha-original`; ordenar/inserir linhas na aba manualmente
  durante a conferência desalinha as gravações.
- **Histórico no frontend indexado só pela etiqueta** (o pedido da chave composta é descartado).
- `importarRelatorioListaChegada` e `processarUploadListaChegada` duplicam a mesma lógica
  (colunas alvo, dedupe, escrita em blocos); o parser CSV também está duplicado em `_obterMapaMonteOSeu`.
- Dados da planilha são inseridos com `innerHTML` em vários pontos (sem escaping).

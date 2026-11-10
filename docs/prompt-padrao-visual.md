# Prompt — Padrão visual "Controle de Insucessos"

Copie tudo o que está abaixo da linha e cole no início do pedido do próximo projeto.
No fim, acrescente o que o novo sistema tem de fazer (telas, dados, funcionalidades).

---

Você vai construir a interface seguindo **exatamente** o padrão visual descrito abaixo. As funcionalidades deste projeto são outras, mas a aparência tem de ser idêntica: mesmas cores, mesma tipografia, mesmos tamanhos, mesmas regras de uso das cores. Não invente cores, não use a paleta padrão de nenhuma biblioteca (Tailwind, Bootstrap, Material) e não troque as fontes. Quando um componente pedido não estiver descrito aqui, construa-o com os mesmos tokens e as mesmas proporções dos componentes descritos.

## 1. Identidade

- Estilo: painel operacional sóbrio, denso de informação, inspirado no GitHub (tons de grafite) com **laranja** como cor de marca.
- **Tema escuro é o padrão**; existe tema claro completo, trocado por um botão no topo (atributo `data-theme="dark" | "light"` no `<html>`).
- Cor de marca: laranja (`#f97316` no escuro, `#c2410c` no claro). É usada **só** para a ação principal, o item ativo da navegação, o foco dos campos, a barra de progresso e destaques de marca. Nunca como cor de erro.
- Superfícies planas: fundo, cartões com borda de 1 px e raio de 12 px. Sem gradientes nas superfícies, sem sombras nos cartões (sombra só em elementos que flutuam: modais, menus suspensos).

## 2. Tipografia

```html
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
```

- Texto: `'IBM Plex Sans', 'Segoe UI', system-ui, sans-serif`, 14 px base.
- **Códigos e identificadores** (etiquetas, números de pedido, NF, nomes de lote, horas em listas): `'IBM Plex Mono', monospace`, peso 600.
- Números de indicadores: `font-variant-numeric: tabular-nums`.
- Escala: título de página 22 px/600 (`letter-spacing: -0.01em`) · título de cartão 15 px/600 · valor de KPI 28 px/600 · texto 14 px · rótulos e legendas 12–13 px · rótulo de secção 11 px/600, MAIÚSCULAS, `letter-spacing: 0.08em`.
- Rótulos de campo: 12 px, 600, cor de texto de apoio.

## 3. Tokens de cor (usar estes nomes e valores)

```css
/* --- TEMA ESCURO (padrão) --- */
:root, [data-theme="dark"] {
  --bg-global: #0d1117;      /* fundo da página */
  --bg-sidebar: #0a0d12;     /* barra lateral */
  --bg-card: #151b23;        /* cartões, campos, tabelas */
  --bg-elevated: #1b2230;    /* hover, item ativo, trilhos de progresso */
  --bg-subtle: #11161d;      /* cabeçalho de tabela, linhas alternadas */
  --border-color: #232c39;   /* bordas normais e divisórias */
  --border-strong: #2f3a4b;  /* bordas de campos e botões secundários */
  --text-main: #e6edf3;
  --text-support: #9aa6b2;   /* legendas, rótulos, texto secundário */

  --brand-accent: #f97316;   /* laranja da marca */
  --brand-hover: #fb923c;
  --brand-soft: rgba(249, 115, 22, 0.14);
  --on-brand: #1a0d02;       /* texto sobre o laranja */

  --table-hover: #1a212b;

  --emerald: #3fb950;        --emerald-soft: rgba(63, 185, 80, 0.14);   /* sucesso / OK */
  --red-alert: #ff7b72;      --red-soft: rgba(255, 123, 114, 0.16);     /* erro / bloqueio */
  --red-action: #da3633;     --on-red: #ffffff;                         /* botão destrutivo */
  --violet: #c4a7ff;         --violet-soft: rgba(196, 167, 255, 0.14);  /* fora do padrão / divergência */
  --blue: #6cb6ff;           --blue-soft: rgba(108, 182, 255, 0.14);    /* informação / etiqueta especial */
  --neutral-soft: rgba(154, 166, 178, 0.14);                            /* pendente / neutro */

  /* Cartões de resultado (fundo / borda / título / subtítulo) */
  --ok-bg: #0f2418;   --ok-border: #1f6f35;   --ok-text: #7ee787;   --ok-sub: #a7c7b0;
  --stop-bg: #2a1215; --stop-border: #8e1519; --stop-text: #ffa198; --stop-sub: #e6b3ae;
  --div-bg: #1d1631;  --div-border: #5b3fa0;  --div-text: #d2bfff;  --div-sub: #c1b3de;
  --rep-bg: #1b2230;  --rep-border: #2f3a4b;  --rep-text: #e6edf3;  --rep-sub: #9aa6b2;

  /* Modo de ação crítica (ex.: "modo recusa"): rosa-carmim */
  --recusa: #ff6b8b;  --recusa-soft: rgba(255, 107, 139, 0.14);  --recusa-forte: #be123c;
  --rec-bg: #2b0f18;  --rec-border: #be123c;  --rec-text: #ff9bb0;  --rec-sub: #f0b6c3;

  --shadow-pop: 0 12px 32px -8px rgba(0, 0, 0, 0.55);
  --scrollbar-bg: #0d1117;
  --scrollbar-thumb: #2f3a4b;
}

/* --- TEMA CLARO --- */
[data-theme="light"] {
  --bg-global: #f6f7f9;
  --bg-sidebar: #ffffff;
  --bg-card: #ffffff;
  --bg-elevated: #eef1f5;
  --bg-subtle: #f8f9fb;
  --border-color: #dfe3e8;
  --border-strong: #c9d0d8;
  --text-main: #0f172a;
  --text-support: #57606a;

  --brand-accent: #c2410c;
  --brand-hover: #9a3412;
  --brand-soft: rgba(194, 65, 12, 0.10);
  --on-brand: #ffffff;

  --table-hover: #f1f4f8;

  --emerald: #1a7f37;    --emerald-soft: #e6f4ea;
  --red-alert: #cf222e;  --red-soft: #ffebe9;
  --red-action: #a40e26; --on-red: #ffffff;
  --violet: #8250df;     --violet-soft: #f3edfd;
  --blue: #0969da;       --blue-soft: #ddf4ff;
  --neutral-soft: #eef1f5;

  --ok-bg: #ecf8ef;   --ok-border: #9fd8ac;   --ok-text: #116329;   --ok-sub: #3d6b4a;
  --stop-bg: #fff0ee; --stop-border: #f3a8a0; --stop-text: #a40e26; --stop-sub: #7d3a37;
  --div-bg: #f6f1ff;  --div-border: #c8b3f3;  --div-text: #6639ba;  --div-sub: #5b4a7d;
  --rep-bg: #f3f5f8;  --rep-border: #dfe3e8;  --rep-text: #0f172a;  --rep-sub: #57606a;

  --recusa: #be123c;  --recusa-soft: #ffe4ea; --recusa-forte: #be123c;
  --rec-bg: #fff0f3;  --rec-border: #f19bb0;  --rec-text: #9f1239;  --rec-sub: #7a2a3d;

  --shadow-pop: 0 12px 32px -8px rgba(15, 23, 42, 0.18);
  --scrollbar-bg: #f6f7f9;
  --scrollbar-thumb: #c9d0d8;
}
```

Todo o CSS usa **apenas** `var(--…)`. Nenhuma cor fixa nos componentes (a única exceção são os gráficos, ver secção 6, e os e-mails, ver secção 7). O tema claro tem de funcionar sem nenhum ajuste extra nos componentes.

## 4. Significado das cores (regra mais importante)

As cores são **semânticas**: a mesma cor quer sempre dizer a mesma coisa, em qualquer tela.

| Significado | Texto/ícone | Fundo suave | Exemplos de uso |
|---|---|---|---|
| Ação principal / marca / item ativo / progresso | `--brand-accent` | `--brand-soft` | botão principal, barra de progresso, foco de campo, "Em aberto" |
| Sucesso, concluído, pode seguir | `--emerald` | `--emerald-soft` | "Conferido", "Concluída", "Receber" |
| Erro, bloqueio, não aceitar | `--red-alert` | `--red-soft` | "Não receber", KPI de bloqueados, alertas críticos |
| Fora do padrão / divergência / item estranho ao conjunto | `--violet` | `--violet-soft` | itens fora do lote, contador de divergências |
| Informação / marcação especial | `--blue` | `--blue-soft` | etiquetas informativas (ex.: "MSPC"), avisos neutros |
| Pendente / neutro / sem dados | `--text-support` ou `--text-main` | `--neutral-soft` | "Pendente", "Sem tracking", "Repetido" |
| Ação crítica em curso (modo especial ligado) | `--recusa` | `--recusa-soft` | faixa de aviso fixa, moldura do ecrã |

- Laranja **nunca** significa erro; vermelho **nunca** é usado em botões primários (só em `.btn-perigo`, ações destrutivas).
- Destaque forte só quando há ocorrência: um indicador com valor 0 fica neutro (`--text-main`/`--text-support`); com valor > 0 ganha a cor do seu significado.
- Texto secundário é sempre `--text-support`; nunca cinza "inventado".

## 5. Componentes (medidas e regras)

**Layout geral**
- `body`: `display: flex; height: 100vh; overflow: hidden;` — barra lateral fixa à esquerda + área principal com scroll próprio (padding `22px 28px 40px`); cabeçalho da página com título à esquerda e o botão de tema à direita, margem inferior 18 px.
- Barra lateral: 240 px, fundo `--bg-sidebar`, borda direita 1 px `--border-color`, padding `20px 14px`, gap 24 px entre blocos. Em telas estreitas recolhe para só ícones.
- Marca: quadrado 30×30, raio 8, fundo `--brand-accent`, texto `--on-brand` (iniciais, 13 px/700) + nome 14 px/700 e subtítulo 12 px `--text-support`.
- Item de navegação: altura mín. 40 px, padding `8px 10px`, raio 8, 14 px/500, cor `--text-support`; hover e ativo com fundo `--bg-elevated` e texto `--text-main`; **ativo com o ícone em `--brand-accent`** e peso 600.
- Títulos de secção da barra: 11 px, MAIÚSCULAS, `letter-spacing: 0.08em`, `--text-support`.
- Ícones: traço (stroke) 1.8–2.4, sem preenchimento, 18–20 px, `stroke: currentColor` (estilo Lucide/Feather).

**Cartão** (`.card-container`): fundo `--bg-card`, borda 1 px `--border-color`, raio 12, padding 20, margem inferior 16. Título 15 px/600.

**Faixa de indicadores (KPI)**: um único cartão dividido em células por bordas verticais de 1 px `--border-color`; cada célula com padding `16px 20px`, título 13 px/500 `--text-support`, valor 28 px/600 tabular, legenda 12 px. Valor em `--red-alert` para bloqueios e `--violet` para divergências. A última célula pode conter os botões de ação empilhados.

**Botões** (altura 40, padding `0 16px`, raio 8, 14 px/600, transição 0.15 s):
- `.btn-brand`: fundo e borda `--brand-accent`, texto `--on-brand`, peso 700; hover `--brand-hover`.
- `.btn-secundario`: fundo `--bg-card`, borda `--border-strong`, texto `--text-main`; hover fundo `--bg-elevated`.
- `.btn-fantasma`: transparente, borda `--border-color`.
- `.btn-perigo`: fundo `--red-action`, texto `--on-red`.
- Desativado: `opacity: 0.45; cursor: not-allowed`.

**Campos** (input, select, textarea): padding `10px 12px`, raio 8, fundo `--bg-card`, borda 1 px `--border-strong`, 14 px. Foco: borda `--brand-accent` + `box-shadow: 0 0 0 3px var(--brand-soft)`. Select sem aparência nativa, com seta SVG à direita. Campos de códigos em IBM Plex Mono.

**Chips e badges** (pílula, `border-radius: 999px`):
- Chip de estado: altura 22, padding `0 9px`, 12 px/600, texto na cor semântica sobre o fundo `-soft` correspondente.
- Chip de filtro: altura 32, padding `0 12px`, borda `--border-strong`, transparente; **ativo** = fundo `--text-main` e texto `--bg-global` (inversão).
- Contador: padding `1px 8px`, 12 px/600, fundo `-soft` + texto da cor.

**Tabelas**: dentro de cartão com raio 12 e scroll próprio; células padding `10px 16px`, borda inferior 1 px `--border-color`, 14 px; cabeçalho fixo (`position: sticky`), fundo `--bg-subtle`, 12 px/600 `--text-support`; hover de linha `--table-hover`; linhas concluídas ficam com texto `--text-support` (sem riscar); estados indicados por chip, não por fundo colorido da linha inteira.

**Cartão de resultado** (resposta a uma ação, ex.: leitura de código): raio 12, padding `22px 24px`, borda 1 px, fundo/borda/título/subtítulo do grupo `--ok-*`, `--stop-*`, `--div-*` ou `--rep-*`; ícone num círculo de 46 px com fundo na cor do título; título 22 px/700 (26 px para bloqueio); identificador em Plex Mono grande (até 44 px); entrada com animação curta (`opacity` + `translateY(4px)`, 0.18 s).

**Avisos em linha**: flex com ícone, padding `10px 14px`, raio 8, 14 px/600, borda 1 px; perigo = fundo `--red-soft`, borda `--stop-border`, texto `--stop-text`; informação = fundo `--blue-soft`, borda e texto `--blue`.

**Item de lista com estado** (ex.: item bloqueado): barra à esquerda de 4 px na cor do estado (`--red-alert`, `--emerald`, `--text-support`) e fundo suave da mesma cor; painel com alerta ativo ganha borda e título em `--red-alert`.

**Progresso**: trilho 6 px, raio 3, fundo `--bg-elevated`; preenchimento `--brand-accent` (ou `--emerald` para "concluído"), transição 0.4 s.

**Controlo segmentado** (alternar vistas): caixa com padding 3 px, fundo `--bg-card`, borda `--border-color`, raio 9; opção ativa com fundo `--bg-elevated` e `inset 0 0 0 1px var(--border-strong)`.

**Modais**: sobreposição `rgba(1, 4, 9, 0.62)` com `backdrop-filter: blur(3px)`; conteúdo é um cartão (raio 12) com `--shadow-pop`; diálogos de confirmação centrados, até 440 px, botões lado a lado (cancelar = fantasma, confirmar = brand ou perigo). Fecha com botão "×" no topo.

**Faixa de modo crítico**: fixa no topo do ecrã, texto branco sobre riscas diagonais de `#be123c`/`#9f1239`, com leve pulsação; usada apenas quando um modo perigoso está ligado.

**Scrollbar**: 10 px, trilho `--scrollbar-bg`, polegar `--scrollbar-thumb` com raio 5 e borda de 2 px da cor do trilho; hover do polegar `--text-support`.

## 6. Gráficos (Chart.js ou equivalente)

- Série principal / volume: linha `#f97316` com preenchimento em gradiente vertical do laranja (≈ 0.25 de opacidade no topo até transparente na base); pontos com fundo branco e borda laranja.
- Comparação "total × realizado": barra do total com fundo `rgba(249, 115, 22, 0.15)` e borda `#f97316`; barra do realizado sólida `#10b981`.
- Rótulos de dados sempre visíveis, 12 px negrito; sobre barras sólidas em branco.
- Tooltip: escuro `#1e293b` / claro `#ffffff`; título `#f8fafc` / `#0f172a`; corpo `#94a3b8` / `#64748b`; borda `#334155` / `#e2e8f0`.
- Grelha discreta (cor `--border-color`), sem bordas de eixo pesadas; textos dos eixos em `--text-support`.
- Ao trocar de tema, atualizar as cores do gráfico (ler `data-theme`).

## 7. E-mails HTML (se o projeto enviar e-mails)

E-mail não lê variáveis CSS: use as cores do tema claro em linha. Fonte Arial; largura máx. 640 px; título 20 px; cartões em tabela (`<table role="presentation">`) com padding `14px 16px`, borda `#e2e8f0`, raio 8, fundo `#f8fafc`, rótulo 12 px `#64748b`, valor 22–24 px/700 `#0f172a`. Cartões em destaque com borda esquerda de 4 px e fundo suave, **colorido só quando o valor é > 0**: vermelho `#b91c1c` / `#fef2f2`, laranja `#c2410c` / `#fff7ed`, roxo `#7e22ce` / `#faf5ff`, azul `#1d4ed8` / `#eff6ff`. Botão principal: fundo `#c2410c`, texto branco, raio 6.

## 8. Responsivo e acessibilidade

- Funciona a partir de 360 px de largura sem scroll horizontal; grelhas passam a uma coluna; barra lateral recolhe para ícones.
- Contraste mínimo AA em ambos os temas (os tokens acima já cumprem; não clarear textos de apoio).
- Foco visível em todos os controlos (anel `--brand-soft` de 3 px ou `outline: 2px solid var(--brand-accent)`).
- Estados nunca só por cor: sempre com texto ou ícone ("NÃO RECEBER", "Concluída", contador).
- Animações curtas (≤ 0.4 s) e só em transições de estado.

## 9. Não fazer

- Não usar outras fontes, outras paletas ou cores fora dos tokens.
- Não usar sombras em cartões, gradientes em fundos ou bordas arredondadas diferentes de 8 (controlos), 9 (controlo segmentado), 12 (cartões) e 999 (pílulas).
- Não usar laranja para erro nem vermelho para ação principal.
- Não pintar linhas de tabela inteiras para indicar estado (usar chip ou barra lateral de 4 px).
- Não esquecer o tema claro: tudo tem de ser testado nos dois temas.

## 10. Verificação final

Antes de entregar, confirme: (1) os dois temas lado a lado com os mesmos componentes; (2) nenhuma cor fixa no CSS dos componentes; (3) cada cor usada respeita a tabela da secção 4; (4) códigos/identificadores em IBM Plex Mono; (5) a 360 px não há scroll horizontal.

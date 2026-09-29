/**
 * Cruzamento de NFs x relatórios de faturamento — v4 (execução automática contínua, otimizada)
 *
 * Basta executar UMA vez a função "cruzarNotasFiscais" (ou o menu "Executar busca").
 * O Apps Script tem limite de tempo por execução (6 min em conta pessoal, 30 min no
 * Workspace), então o script trabalha em etapas: quando chega perto do limite, salva
 * o progresso (inclusive a posição dentro do relatório que estava lendo) e agenda
 * sozinho a próxima etapa por um gatilho de tempo. Ao terminar, monta a aba
 * "Resultado", remove o gatilho e envia um e-mail com o resumo.
 *
 * Correções da v2 mantidas: leitura por bytes, detecção de codificação, cabeçalho
 * sempre detectado, busca de coluna exata, parser tolerante a aspas, retentativas.
 */

// ==========================================
// CONFIGURAÇÕES
// ==========================================
const CONFIG = {
  FOLDER_ID: '1FyzDA0iDQ-SfSwuE9_nubMN9gRW_VxNy',

  NOME_ABA_PRINCIPAL: 'NFs',
  COLUNA_NOTAS_PRINCIPAL: 'A',
  LINHA_INICIO_PRINCIPAL: 2,
  COLUNA_RESULTADO_PRINCIPAL: 'B',

  // Aba onde o resultado será gravado (recriada ao final de cada cruzamento)
  NOME_ABA_RESULTADO: 'Resultado',

  // Filtro dos arquivos do Drive (sem diferenciar maiúsculas/minúsculas)
  PREFIXO_ARQUIVO: 'relatorio_pedidos_faturados',
  INCLUIR_SUBPASTAS: false,

  // Mapeamento por nome de cabeçalho
  CABECALHO_BUSCA: 'NF saída',
  CABECALHOS_TRAZER: [
    'Código Produto',
    'Número do Pedido Faturado',
    'Mercadoria Código',
    'Produto Nome',
    'Custo Produto'
  ],

  // Reserva, caso o cabeçalho não seja encontrado
  COLUNA_NOTAS_FATURAMENTO_LETRA: 'L',
  COLUNAS_PARA_TRAZER_LETRAS: ['A', 'M', 'C', 'D', 'E'],

  // Remove itens idênticos da mesma NF vindos de relatórios diferentes
  REMOVER_DUPLICADOS_ENTRE_ARQUIVOS: true,

  // Adiciona uma coluna com o nome do relatório de onde veio cada item
  INCLUIR_ARQUIVO_ORIGEM: true,

  TAMANHO_BLOCO_MB: 8,

  // Tempo de trabalho por etapa. O Google limita cada execução a 6 min (conta pessoal
  // e Workspace), então não passe de 5 — a margem é para terminar o bloco em andamento.
  TEMPO_LIMITE_MINUTOS: 5,

  // Trava de segurança: número máximo de etapas automáticas por cruzamento
  MAX_ETAPAS: 150,

  // E-mail com o resumo ao terminar (ou em caso de erro).
  // '' = envia para a conta que executa o script. null = não envia.
  EMAIL_AVISO: 'arthur.silva@kabum.com.br'
};

const ABA_CACHE_ARQUIVOS = '_cruz_arquivos';
const ABA_CACHE_RESULTADOS = '_cruz_resultados';
const ABA_CACHE_LISTA = '_cruz_lista';
const FUNCAO_CONTINUACAO = 'continuarCruzamento';
const PROP_ESTADO = 'CRUZ_ESTADO_ARQUIVO';
const PROP_ETAPAS = 'CRUZ_ETAPAS';

// ==========================================
// MENU E PONTOS DE ENTRADA
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Cruzamento NF')
    .addItem('1. Autorizar permissões', 'forcarPermissoes')
    .addItem('2. Executar busca (do zero)', 'cruzarNotasFiscais')
    .addItem('3. Cancelar busca em andamento', 'cancelarCruzamento')
    .addToUi();
}

/** Execute esta função UMA vez. O restante acontece sozinho. */
function cruzarNotasFiscais() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    notificar_('Em andamento', 'Já existe um cruzamento sendo executado neste momento.', false);
    return;
  }
  try {
    removerGatilhos_();
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(PROP_ESTADO);
    props.setProperty(PROP_ETAPAS, '0');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    [ABA_CACHE_ARQUIVOS, ABA_CACHE_RESULTADOS, ABA_CACHE_LISTA].forEach(nome => {
      const s = ss.getSheetByName(nome);
      if (s) ss.deleteSheet(s);
    });

    executarEtapa_();
  } finally {
    lock.releaseLock();
  }
}

/** Chamada automaticamente pelo gatilho de tempo. Também pode ser rodada à mão. */
function continuarCruzamento() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // outra etapa ainda rodando; ela cuida do agendamento
  try {
    executarEtapa_();
  } finally {
    lock.releaseLock();
  }
}

function cancelarCruzamento() {
  removerGatilhos_();
  PropertiesService.getScriptProperties().deleteProperty(PROP_ESTADO);
  notificar_('Cancelado', 'Etapas automáticas canceladas. Rode "cruzarNotasFiscais" para recomeçar.', false);
}

// ==========================================
// UMA ETAPA DE PROCESSAMENTO
// ==========================================
function executarEtapa_() {
  const inicio = Date.now();
  const prazo = inicio + CONFIG.TEMPO_LIMITE_MINUTOS * 60 * 1000;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  // Gatilho de segurança: se esta etapa morrer por qualquer motivo, a próxima
  // começa sozinha depois do limite. É trocado por um mais curto ao pausar.
  agendarContinuacao_(CONFIG.TEMPO_LIMITE_MINUTOS + 10);

  const etapa = Number(props.getProperty(PROP_ETAPAS) || 0) + 1;
  props.setProperty(PROP_ETAPAS, String(etapa));
  if (etapa > CONFIG.MAX_ETAPAS) {
    removerGatilhos_();
    notificar_('Interrompido', `O cruzamento passou de ${CONFIG.MAX_ETAPAS} etapas e foi parado por segurança. ` +
      'Aumente MAX_ETAPAS ou TEMPO_LIMITE_MINUTOS e rode "cruzarNotasFiscais" de novo.', true);
    return;
  }

  if (CONFIG.NOME_ABA_RESULTADO === CONFIG.NOME_ABA_PRINCIPAL) {
    return falhar_('A aba de resultado deve ser diferente da aba principal.');
  }
  const sheetNFs = ss.getSheetByName(CONFIG.NOME_ABA_PRINCIPAL);
  if (!sheetNFs) return falhar_(`Aba "${CONFIG.NOME_ABA_PRINCIPAL}" não encontrada.`);

  const ultimaLinha = sheetNFs.getLastRow();
  if (ultimaLinha < CONFIG.LINHA_INICIO_PRINCIPAL) return falhar_('Não há dados na planilha principal.');

  // 1. Notas a procurar
  const idxNota = letraParaIndice(CONFIG.COLUNA_NOTAS_PRINCIPAL);
  const numCols = Math.max(sheetNFs.getLastColumn(), idxNota + 1);
  const dadosNFs = sheetNFs.getRange(
    CONFIG.LINHA_INICIO_PRINCIPAL, 1,
    ultimaLinha - CONFIG.LINHA_INICIO_PRINCIPAL + 1, numCols
  ).getValues();

  const notas = new Set();
  dadosNFs.forEach(l => {
    const n = limparNota(l[idxNota]);
    if (n) notas.add(n);
  });

  // 2. Progresso salvo
  const nTrazer = CONFIG.CABECALHOS_TRAZER.length;
  const cacheArq = obterAbaCache_(ss, ABA_CACHE_ARQUIVOS,
    ['ID', 'Arquivo', 'Linhas lidas', 'Itens encontrados', 'Observação']);
  const cacheRes = obterAbaCache_(ss, ABA_CACHE_RESULTADOS,
    ['NF'].concat(CONFIG.CABECALHOS_TRAZER, ['Arquivo']));
  const jaProcessados = lerIdsProcessados_(cacheArq);
  const estadoSalvo = lerEstado_();

  // 3. Lista de arquivos
  let arquivos;
  try {
    arquivos = obterListaArquivos_(ss); // lista do Drive feita só na 1ª etapa
  } catch (e) {
    return falhar_('Não foi possível acessar a pasta do Drive: ' + e.message);
  }
  const pendentes = arquivos.filter(a => !jaProcessados.has(a.id));
  console.log(`▶️ Etapa ${etapa}: ${notas.size} NFs, ${arquivos.length} relatórios, ${pendentes.length} pendentes.`);

  // 4. Processamento
  let pausado = false;
  let feitosAgora = 0;

  for (const arq of pendentes) {
    if (Date.now() > prazo) { pausado = true; break; }

    const retomada = (estadoSalvo && estadoSalvo.fileId === arq.id) ? estadoSalvo : null;
    if (retomada) console.log(`↪️ Retomando ${arq.nome} a partir do byte ${retomada.pos}`);

    let r;
    const tArq = Date.now();
    try {
      r = processarArquivo_(arq, notas, prazo, retomada);
    } catch (e) {
      console.error(`❌ ${arq.nome}: ${e.message}`);
      anexarLinhas_(cacheArq, [[arq.id, arq.nome, 0, 0, 'ERRO: ' + e.message]]);
      if (retomada) props.deleteProperty(PROP_ESTADO);
      continue;
    }

    // Itens encontrados são gravados já, mesmo com o arquivo lido pela metade
    anexarLinhas_(cacheRes, r.linhas);

    if (r.interrompido) {
      salvarEstado_(Object.assign({ fileId: arq.id }, r.estado));
      pausado = true;
      break;
    }

    anexarLinhas_(cacheArq, [[arq.id, arq.nome, r.linhasLidas, r.totalItens, r.obs]]);
    if (retomada) props.deleteProperty(PROP_ESTADO);
    feitosAgora++;
    console.log(`📄 ${arq.nome}: ${(r.bytesLidos / 1048576).toFixed(1)} MB em ${Math.round((Date.now() - tArq) / 1000)}s, ` +
      `${r.linhasLidas} linhas lidas, ${r.totalItens} itens${r.obs ? ' — ' + r.obs : ''}`);
  }

  if (pausado) {
    SpreadsheetApp.flush();
    agendarContinuacao_(1);
    const total = jaProcessados.size + feitosAgora;
    notificar_('Em andamento',
      `${total} de ${arquivos.length} relatórios concluídos (etapa ${etapa}). ` +
      'A próxima etapa começa automaticamente em cerca de 1 minuto.', false);
    return;
  }

  // 5. Montar a aba Resultado pode demorar: se esta etapa já trabalhou bastante,
  //    deixa para a próxima, que começa com o tempo todo livre.
  if (feitosAgora > 0 && Date.now() - inicio > 90 * 1000) {
    agendarContinuacao_(1);
    notificar_('Em andamento', 'Todos os relatórios foram lidos. O resultado será montado na próxima etapa.', false);
    return;
  }

  // 6. Fim: monta o resultado e desliga o agendamento
  removerGatilhos_();
  props.deleteProperty(PROP_ESTADO);

  const resumo = gerarResultado_(ss, sheetNFs, dadosNFs, idxNota, numCols, cacheRes, nTrazer);
  const avisos = contarAvisos_(cacheArq);
  const minutos = etapa > 1 ? ` em ${etapa} etapas` : '';

  notificar_('Concluído',
    `${arquivos.length} relatórios verificados${minutos}.\n` +
    `NFs encontradas: ${resumo.encontradas} de ${notas.size}.\n` +
    `NFs não encontradas: ${notas.size - resumo.encontradas}.` +
    (avisos ? `\n${avisos} relatório(s) com aviso ou erro. Veja a aba oculta "${ABA_CACHE_ARQUIVOS}".` : ''),
    true);
}

function falhar_(msg) {
  removerGatilhos_();
  notificar_('Erro', msg, true);
}

// ==========================================
// LEITURA DE UM RELATÓRIO
// ==========================================
function processarArquivo_(arq, notas, prazo, retomada) {
  const url = `https://www.googleapis.com/drive/v3/files/${arq.id}?alt=media&supportsAllDrives=true`;
  const token = ScriptApp.getOAuthToken();
  const bloco = Math.floor(CONFIG.TAMANHO_BLOCO_MB * 1024 * 1024);

  let inicioByte = retomada ? retomada.pos : 0;
  const byteInicial = inicioByte;
  let sobra = '';            // última linha incompleta do bloco anterior (1 caractere = 1 byte)
  let utf8 = retomada ? retomada.charset === 'UTF-8' : true;
  let layout = retomada ? retomada.layout : null;
  let linhasLidas = retomada ? retomada.linhasLidas : 0;
  const itensAntes = retomada ? retomada.itens : 0;
  let fim = false;
  const linhas = [];
  const obs = retomada ? retomada.obs.slice() : [];

  while (!fim) {
    if (Date.now() > prazo) {
      // Guarda a posição exata (início da primeira linha ainda não lida)
      return {
        interrompido: true,
        linhas,
        estado: {
          pos: inicioByte - sobra.length,
          charset: utf8 ? 'UTF-8' : 'ISO-8859-1',
          layout, linhasLidas, obs,
          itens: itensAntes + linhas.length
        }
      };
    }

    const resp = baixarComRetry_(url, token, inicioByte, inicioByte + bloco - 1);
    const code = resp.getResponseCode();
    if (code === 416) break; // pediu além do fim do arquivo
    if (code !== 200 && code !== 206) {
      throw new Error(`HTTP ${code} ao baixar o arquivo`);
    }
    if (code === 200 && inicioByte > 0) {
      throw new Error('O servidor não aceitou leitura parcial; não foi possível retomar o arquivo');
    }

    // Decodificado no servidor como ISO-8859-1: cada byte vira exatamente 1 caractere.
    // Assim o corte nas quebras de linha é exato e não há conversão lenta de arrays de bytes.
    let bruto = resp.getContentText('ISO-8859-1');
    const total = tamanhoTotal_(resp);
    inicioByte += bruto.length;

    if (code === 200 || bruto.length < bloco || (total !== null && inicioByte >= total)) {
      fim = true;
    }

    if (sobra) { bruto = sobra + bruto; sobra = ''; }

    let parte;
    if (fim) {
      parte = bruto;
    } else {
      const ultimoNL = bruto.lastIndexOf('\n');
      if (ultimoNL === -1) { sobra = bruto; continue; }
      parte = bruto.slice(0, ultimoNL + 1);
      sobra = bruto.slice(ultimoNL + 1);
    }
    bruto = null;

    // Converte para o texto real: UTF-8 quando válido; senão o arquivo é ANSI (Windows-1252)
    let texto = null;
    if (utf8) {
      try { texto = decodeURIComponent(escape(parte)); }
      catch (e) { utf8 = false; }
    }
    if (!utf8) texto = corrigirWindows1252_(parte);
    parte = null;

    const linhasTexto = texto.split(/\r?\n/);
    texto = null;

    let i = 0;
    if (!layout) {
      linhasTexto[0] = linhasTexto[0].replace(/^\uFEFF/, '');
      layout = detectarLayout_(linhasTexto);
      if (layout.linhaCabecalho === -1) {
        obs.push(`Cabeçalho "${CONFIG.CABECALHO_BUSCA}" não encontrado; usando letras de reserva`);
      } else if (layout.faltando.length) {
        obs.push('Colunas não encontradas (usando reserva): ' + layout.faltando.join(', '));
      }
      i = layout.linhaCabecalho + 1;
    }

    const idxNF = layout.idxNota;
    const idxTrazer = layout.idxTrazer;
    const delim = layout.delim;

    for (; i < linhasTexto.length; i++) {
      const lt = linhasTexto[i];
      if (!lt) continue;
      linhasLidas++;

      const campos = dividirLinha_(lt, delim);
      if (campos.length <= idxNF) continue;

      const nf = limparNota(campos[idxNF]);
      if (!nf || !notas.has(nf)) continue;

      const valores = idxTrazer.map(idx =>
        (idx >= 0 && idx < campos.length) ? String(campos[idx]).trim() : '');
      linhas.push([nf].concat(valores, [arq.nome]));
    }
  }

  if (!utf8 && obs.indexOf('Codificação: windows-1252') === -1) obs.push('Codificação: windows-1252');
  return {
    interrompido: false, linhas, linhasLidas,
    totalItens: itensAntes + linhas.length,
    bytesLidos: inicioByte - byteInicial,
    obs: obs.join(' | ')
  };
}

// Ajusta os caracteres em que Windows-1252 difere de ISO-8859-1 (€, aspas curvas, travessões...)
const MAPA_1252_ = {
  0x80: '€', 0x82: '‚', 0x83: 'ƒ', 0x84: '„', 0x85: '…', 0x86: '†', 0x87: '‡', 0x88: 'ˆ',
  0x89: '‰', 0x8A: 'Š', 0x8B: '‹', 0x8C: 'Œ', 0x8E: 'Ž', 0x91: '‘', 0x92: '’', 0x93: '“',
  0x94: '”', 0x95: '•', 0x96: '–', 0x97: '—', 0x98: '˜', 0x99: '™', 0x9A: 'š', 0x9B: '›',
  0x9C: 'œ', 0x9E: 'ž', 0x9F: 'Ÿ'
};
function corrigirWindows1252_(s) {
  return s.replace(/[\x80-\x9F]/g, c => MAPA_1252_[c.charCodeAt(0)] || c);
}

function detectarLayout_(linhas) {
  const alvoNF = normalizarCabecalho_(CONFIG.CABECALHO_BUSCA);
  const layout = {
    delim: ';',
    idxNota: letraParaIndice(CONFIG.COLUNA_NOTAS_FATURAMENTO_LETRA),
    idxTrazer: CONFIG.COLUNAS_PARA_TRAZER_LETRAS.map(letraParaIndice),
    linhaCabecalho: -1,
    faltando: []
  };

  const limite = Math.min(30, linhas.length);
  let linhaDelim = '';

  for (let l = 0; l < limite; l++) {
    if (!linhaDelim && linhas[l]) linhaDelim = linhas[l];
    if (normalizarCabecalho_(linhas[l]).indexOf(alvoNF) === -1) continue;

    layout.delim = detectarDelimitador_(linhas[l]);
    const cabs = dividirLinha_(linhas[l], layout.delim).map(normalizarCabecalho_);

    const iNF = acharColuna_(cabs, alvoNF);
    if (iNF === -1) continue; // o texto apareceu, mas não como cabeçalho de coluna
    layout.idxNota = iNF;

    layout.idxTrazer = CONFIG.CABECALHOS_TRAZER.map((c, k) => {
      const idx = acharColuna_(cabs, normalizarCabecalho_(c));
      if (idx === -1) { layout.faltando.push(c); return layout.idxTrazer[k]; }
      return idx;
    });
    layout.linhaCabecalho = l;
    return layout;
  }

  if (linhaDelim) layout.delim = detectarDelimitador_(linhaDelim);
  return layout;
}

function detectarDelimitador_(linha) {
  let melhor = ';', max = 0;
  [';', ',', '\t', '|'].forEach(d => {
    const n = linha.split(d).length - 1;
    if (n > max) { max = n; melhor = d; }
  });
  return melhor;
}

// Exata primeiro; só depois "contém"
function acharColuna_(cabs, alvo) {
  const exata = cabs.indexOf(alvo);
  if (exata !== -1) return exata;
  return cabs.findIndex(c => c.indexOf(alvo) !== -1);
}

/**
 * Divide uma linha de CSV respeitando aspas. Tolerante a aspas soltas no meio
 * de um campo (ex.: Parafuso 1/2" aço): elas são tratadas como texto comum.
 * Trabalha com fatias do texto (indexOf/slice), bem mais rápido que caractere a caractere.
 */
function dividirLinha_(linha, d) {
  if (linha.indexOf('"') === -1) return linha.split(d);

  const campos = [];
  const n = linha.length;
  let i = 0;

  while (i <= n) {
    if (linha.charCodeAt(i) === 34) {           // campo entre aspas
      let j = i + 1;
      let valor = '';
      for (;;) {
        const q = linha.indexOf('"', j);
        if (q === -1) { valor += linha.slice(j); i = n + 1; break; }          // aspas sem fechar
        const prox = linha[q + 1];
        if (prox === '"') { valor += linha.slice(j, q + 1); j = q + 2; continue; } // "" → "
        if (prox === undefined || prox === d) { valor += linha.slice(j, q); i = q + 2; break; }
        valor += linha.slice(j, q + 1); j = q + 1;                            // aspa solta
      }
      campos.push(valor);
    } else {
      const p = linha.indexOf(d, i);
      if (p === -1) { campos.push(linha.slice(i)); break; }
      campos.push(linha.slice(i, p));
      i = p + 1;
    }
  }
  return campos;
}

function baixarComRetry_(url, token, ini, fim) {
  const opts = {
    headers: { Authorization: 'Bearer ' + token, Range: `bytes=${ini}-${fim}` },
    muteHttpExceptions: true
  };
  for (let t = 0; t < 5; t++) {
    const r = UrlFetchApp.fetch(url, opts);
    const c = r.getResponseCode();
    const limiteTaxa = c === 429 || c >= 500 ||
      (c === 403 && /rate|quota/i.test(r.getContentText()));
    if (!limiteTaxa) return r;
    Utilities.sleep(1000 * Math.pow(2, t));
  }
  throw new Error('Falha ao baixar após 5 tentativas');
}

function tamanhoTotal_(resp) {
  const h = resp.getHeaders();
  for (const k in h) {
    if (k.toLowerCase() === 'content-range') {
      const m = String(h[k]).match(/\/(\d+)\s*$/);
      return m ? Number(m[1]) : null;
    }
  }
  return null;
}

function listarArquivos_(folderId) {
  const prefixo = CONFIG.PREFIXO_ARQUIVO.toLowerCase();
  const lista = [];
  const pilha = [DriveApp.getFolderById(folderId)];

  while (pilha.length) {
    const pasta = pilha.pop();
    const it = pasta.searchFiles(`title contains '${CONFIG.PREFIXO_ARQUIVO.replace(/'/g, "\\'")}' and trashed = false`);
    while (it.hasNext()) {
      const f = it.next();
      const nome = f.getName();
      const nomeMin = nome.toLowerCase();
      if (!nomeMin.startsWith(prefixo)) continue;
      if (f.getMimeType() !== MimeType.CSV && !nomeMin.endsWith('.csv')) continue;
      lista.push({ id: f.getId(), nome: nome });
    }
    if (CONFIG.INCLUIR_SUBPASTAS) {
      const subs = pasta.getFolders();
      while (subs.hasNext()) pilha.push(subs.next());
    }
  }
  lista.sort((a, b) => a.nome.localeCompare(b.nome));
  return lista;
}

function obterListaArquivos_(ss) {
  const s = ss.getSheetByName(ABA_CACHE_LISTA);
  if (s && s.getLastRow() > 1) {
    return s.getRange(2, 1, s.getLastRow() - 1, 2).getValues()
      .map(r => ({ id: String(r[0]), nome: String(r[1]) }));
  }
  const lista = listarArquivos_(CONFIG.FOLDER_ID);
  const aba = obterAbaCache_(ss, ABA_CACHE_LISTA, ['ID', 'Arquivo']);
  anexarLinhas_(aba, lista.map(a => [a.id, a.nome]));
  return lista;
}

// ==========================================
// RESULTADO FINAL
// ==========================================
function gerarResultado_(ss, sheetNFs, dadosNFs, idxNota, numCols, cacheRes, nTrazer) {
  // Agrupa os itens encontrados por NF
  const mapa = new Map();
  const nLinhasCache = cacheRes.getLastRow() - 1;
  if (nLinhasCache > 0) {
    const dados = cacheRes.getRange(2, 1, nLinhasCache, nTrazer + 2).getValues();
    dados.forEach(l => {
      const nf = String(l[0]);
      const valores = l.slice(1, 1 + nTrazer).map(v => String(v));
      const origem = String(l[1 + nTrazer]);

      if (!mapa.has(nf)) mapa.set(nf, { chaves: new Map(), itens: [] });
      const g = mapa.get(nf);

      if (CONFIG.REMOVER_DUPLICADOS_ENTRE_ARQUIVOS) {
        const chave = valores.join('\u0001');
        const origemAnterior = g.chaves.get(chave);
        if (origemAnterior !== undefined && origemAnterior !== origem) return;
        g.chaves.set(chave, origem);
      }
      g.itens.push({ valores, origem });
    });
  }

  const idxRes = letraParaIndice(CONFIG.COLUNA_RESULTADO_PRINCIPAL);
  let largura = Math.max(numCols, idxRes + nTrazer);
  const idxOrigem = CONFIG.INCLUIR_ARQUIVO_ORIGEM ? largura : -1;
  if (CONFIG.INCLUIR_ARQUIVO_ORIGEM) largura++;

  const completar = linha => {
    const nova = linha.slice(0, largura);
    while (nova.length < largura) nova.push('');
    return nova;
  };

  // Cabeçalho: copia as linhas de cabeçalho da aba NFs e nomeia as colunas novas
  const nCab = CONFIG.LINHA_INICIO_PRINCIPAL - 1;
  const cabecalho = nCab > 0
    ? sheetNFs.getRange(1, 1, nCab, numCols).getValues().map(completar)
    : [];
  if (cabecalho.length) {
    const u = cabecalho[cabecalho.length - 1];
    CONFIG.CABECALHOS_TRAZER.forEach((c, k) => { u[idxRes + k] = c; });
    if (idxOrigem !== -1) u[idxOrigem] = 'Arquivo de origem';
  }

  const saida = [];
  const encontradas = new Set();
  const naoEncontradas = new Set();

  dadosNFs.forEach(linha => {
    const base = completar(linha);
    const nf = limparNota(linha[idxNota]);
    const g = nf ? mapa.get(nf) : null;

    if (g && g.itens.length) {
      encontradas.add(nf);
      g.itens.forEach(it => {
        const nova = base.slice();
        it.valores.forEach((v, c) => { nova[idxRes + c] = v; });
        if (idxOrigem !== -1) nova[idxOrigem] = it.origem;
        saida.push(nova);
      });
    } else {
      if (nf) naoEncontradas.add(nf);
      const nova = base.slice();
      for (let c = 0; c < nTrazer; c++) nova[idxRes + c] = '';
      if (idxOrigem !== -1) nova[idxOrigem] = '';
      saida.push(nova);
    }
  });

  if (naoEncontradas.size) {
    console.log('🔎 Exemplos de NFs não encontradas: ' + Array.from(naoEncontradas).slice(0, 30).join(', '));
  }

  let aba = ss.getSheetByName(CONFIG.NOME_ABA_RESULTADO);
  if (aba) aba.clear(); else aba = ss.insertSheet(CONFIG.NOME_ABA_RESULTADO);

  const todas = cabecalho.concat(saida);
  garantirTamanho_(aba, todas.length, largura);
  aba.getRange(1, 1, todas.length, largura).setValues(todas);
  if (cabecalho.length) aba.setFrozenRows(cabecalho.length);
  try { aba.activate(); } catch (e) { /* execução por gatilho: sem tela aberta */ }

  return { encontradas: encontradas.size };
}

// ==========================================
// AUXILIARES DE PLANILHA
// ==========================================
function obterAbaCache_(ss, nome, cabecalho) {
  let s = ss.getSheetByName(nome);
  if (!s) {
    s = ss.insertSheet(nome);
    s.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);
    s.hideSheet();
  }
  return s;
}

function lerIdsProcessados_(sheet) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return new Set();
  return new Set(sheet.getRange(2, 1, n, 1).getValues().map(r => String(r[0])));
}

function contarAvisos_(sheet) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return 0;
  return sheet.getRange(2, 5, n, 1).getValues().filter(r => {
    const t = String(r[0]);
    return t && !/^Codificação: [\w-]+$/.test(t); // aviso de codificação sozinho não conta
  }).length;
}

function garantirTamanho_(sheet, linhas, colunas) {
  if (linhas > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), linhas - sheet.getMaxRows());
  if (colunas > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), colunas - sheet.getMaxColumns());
}

// Grava como texto puro para não perder zeros à esquerda nos códigos
function anexarLinhas_(sheet, linhas) {
  if (!linhas.length) return;
  const ini = sheet.getLastRow() + 1;
  const cols = linhas[0].length;
  garantirTamanho_(sheet, ini + linhas.length - 1, cols);
  sheet.getRange(ini, 1, linhas.length, cols).setNumberFormat('@').setValues(linhas);
}

// ==========================================
// AUXILIARES GERAIS
// ==========================================
function limparNota(nota) {
  if (nota === null || nota === undefined || nota === '') return '';
  return String(nota).replace(/[^0-9]/g, '').replace(/^0+/, '');
}

function normalizarCabecalho_(s) {
  return String(s || '')
    .replace(/^\uFEFF/, '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/["']/g, '')
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .trim();
}

function letraParaIndice(letra) {
  letra = String(letra).toUpperCase();
  let indice = 0;
  for (let i = 0; i < letra.length; i++) {
    indice = indice * 26 + (letra.charCodeAt(i) - 64);
  }
  return indice - 1;
}

// ==========================================
// AGENDAMENTO, ESTADO E AVISOS
// ==========================================
function agendarContinuacao_(minutos) {
  removerGatilhos_();
  ScriptApp.newTrigger(FUNCAO_CONTINUACAO).timeBased().after(minutos * 60 * 1000).create();
}

function removerGatilhos_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === FUNCAO_CONTINUACAO) ScriptApp.deleteTrigger(t);
  });
}

function lerEstado_() {
  const txt = PropertiesService.getScriptProperties().getProperty(PROP_ESTADO);
  if (!txt) return null;
  try { return JSON.parse(txt); } catch (e) { return null; }
}

function salvarEstado_(estado) {
  PropertiesService.getScriptProperties().setProperty(PROP_ESTADO, JSON.stringify(estado));
}

/**
 * Mostra um aviso rápido na planilha (se estiver aberta) e, quando "final",
 * envia e-mail — execuções por gatilho não conseguem abrir janelas.
 */
function notificar_(titulo, mensagem, final) {
  console.log(`[${titulo}] ${mensagem}`);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(mensagem, 'Cruzamento NF — ' + titulo, final ? 60 : 15);
  } catch (e) { /* sem planilha aberta */ }

  if (!final || CONFIG.EMAIL_AVISO === null) return;
  try {
    const destino = CONFIG.EMAIL_AVISO || Session.getEffectiveUser().getEmail();
    const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    MailApp.sendEmail(destino, `[Cruzamento NF] ${titulo}`, `${mensagem}\n\nPlanilha: ${url}`);
  } catch (e) {
    console.error('Não foi possível enviar o e-mail: ' + e.message);
  }
}

function forcarPermissoes() {
  try {
    DriveApp.getRootFolder();
    SpreadsheetApp.getActiveSpreadsheet();
    UrlFetchApp.fetch('https://www.google.com');
    ScriptApp.getProjectTriggers();
    MailApp.getRemainingDailyQuota();
    Session.getEffectiveUser().getEmail();
    SpreadsheetApp.getUi().alert('Sucesso!', 'Todas as permissões concedidas.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Aviso', 'Conceda as permissões na tela do Google.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

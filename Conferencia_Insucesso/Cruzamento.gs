/**
 * ============================================================================
 * SISTEMA: K-OUT Auditoria Express
 * FICHEIRO: Cruzamento.gs — Nova carga a partir de uma lista de NFs
 * ============================================================================
 *
 * Integra o "Cruzamento de NFs x relatórios de faturamento" (projeto Cruzamento_NF, v4)
 * na Conferência de Insucessos:
 *
 *  1. No modal "Nova carga › Cruzar NFs" o utilizador cola as NFs e escolhe a data.
 *     O Web App chama iniciarCruzamentoCarga(), que grava as NFs na aba "NFs_Base".
 *  2. O cruzamento lê os CSV "relatorio_pedidos_faturados*" da pasta do Drive em
 *     etapas (limite de tempo do Apps Script). O Web App chama processarEtapaCruzamento()
 *     em ciclo e mostra o progresso; se a página for fechada, um gatilho de tempo
 *     (continuarCruzamento) continua sozinho.
 *  3. No fim, a aba da carga (DD-MM-AAAA) é criada pelo mesmo caminho do upload de
 *     ficheiro (_criarAbaCarga em Código.gs: sufixo Vn e coluna MSPC), a "NFs_Base"
 *     é limpa e os caches temporários são apagados.
 *
 * Mantidos da v4: leitura por bytes, deteção de codificação (UTF-8 / Windows-1252),
 * deteção do cabeçalho, busca de coluna exata, parser tolerante a aspas, retentativas,
 * retoma a meio de um relatório e remoção de duplicados entre relatórios.
 */

// ==========================================
// CONFIGURAÇÕES
// ==========================================
const CRUZ_CONFIG = {
  // Pasta do Drive com os relatórios de faturamento
  FOLDER_ID: '1FyzDA0iDQ-SfSwuE9_nubMN9gRW_VxNy',

  // Aba que recebe as NFs coladas (limpa ao terminar)
  NOME_ABA_NFS: 'NFs_Base',

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

  // Nome da coluna de NF na aba da carga (a Conferência reconhece "Nota Fiscal" ou "NF")
  CABECALHO_NF_CARGA: 'Nota Fiscal',

  // Reserva, caso o cabeçalho não seja encontrado
  COLUNA_NOTAS_FATURAMENTO_LETRA: 'L',
  COLUNAS_PARA_TRAZER_LETRAS: ['A', 'M', 'C', 'D', 'E'],

  // Remove itens idênticos da mesma NF vindos de relatórios diferentes
  REMOVER_DUPLICADOS_ENTRE_ARQUIVOS: true,

  TAMANHO_BLOCO_MB: 8,

  // Tempo de trabalho por etapa. Cada execução do Apps Script (inclusive as chamadas
  // do Web App) tem limite de 6 min: a margem é para terminar o bloco em andamento.
  TEMPO_LIMITE_MINUTOS: 4.5,

  // Trava de segurança: número máximo de etapas por cruzamento
  MAX_ETAPAS: 150,

  // E-mail com o resumo ao terminar (ou em caso de erro).
  // '' = envia para a conta que executa o script. null = não envia.
  EMAIL_AVISO: 'arthur.silva@kabum.com.br'
};

const CRUZ_ABA_ARQUIVOS = '_cruz_arquivos';
const CRUZ_ABA_RESULTADOS = '_cruz_resultados';
const CRUZ_ABA_LISTA = '_cruz_lista';
const CRUZ_FUNCAO_CONTINUACAO = 'continuarCruzamento';
const CRUZ_PROP_ESTADO = 'CRUZ_ESTADO_ARQUIVO';
const CRUZ_PROP_ETAPAS = 'CRUZ_ETAPAS';
const CRUZ_PROP_JOB = 'CRUZ_JOB';

// ==========================================
// MENU DA PLANILHA
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Cruzamento NF')
    .addItem('Autorizar permissões', 'forcarPermissoes')
    .addItem('Continuar cruzamento em andamento', 'continuarCruzamento')
    .addItem('Cancelar cruzamento em andamento', 'cancelarCruzamento')
    .addToUi();
}

// ==========================================
// API DO WEB APP
// ==========================================

/**
 * Grava as NFs na aba "NFs_Base" e prepara um novo cruzamento.
 * @param {string} dataISO  data da carga no formato AAAA-MM-DD (input type="date")
 * @param {string|string[]} nfsEntrada  texto colado do Excel ou lista de NFs
 */
function iniciarCruzamentoCarga(dataISO, nfsEntrada) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { erro: 'Existe uma etapa de cruzamento a ser processada neste momento. Aguarde alguns segundos e tente novamente.' };
  }
  try {
    const jobAtual = lerJobCruzamento_();
    if (jobAtual && jobAtual.status === 'em_andamento') {
      return {
        erro: `Já existe um cruzamento em andamento para a carga ${jobAtual.nomeCarga}. Aguarde terminar ou cancele-o.`,
        job: jobAtual
      };
    }

    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dataISO || '').trim());
    if (!m) return { erro: 'Selecione a data da carga.' };
    const nomeCarga = `${m[3]}-${m[2]}-${m[1]}`;

    const nfs = extrairNotas_(nfsEntrada);
    if (!nfs.length) return { erro: 'Nenhum número de NF válido foi encontrado no texto colado.' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    gravarNFsBase_(ss, nfs);
    apagarCachesCruzamento_(ss);
    removerGatilhos_();

    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(CRUZ_PROP_ESTADO);
    props.setProperty(CRUZ_PROP_ETAPAS, '0');

    let usuario = '';
    try { usuario = Session.getActiveUser().getEmail(); } catch (e) { /* sem acesso ao e-mail */ }

    const job = {
      status: 'em_andamento',
      fase: 'preparando',
      nomeCarga: nomeCarga,
      totalNFs: nfs.length,
      relatoriosTotal: 0,
      relatoriosFeitos: 0,
      etapa: 0,
      inicio: new Date().toISOString(),
      usuario: usuario,
      mensagem: `${nfs.length} NFs gravadas na aba "${CRUZ_CONFIG.NOME_ABA_NFS}". A iniciar a leitura dos relatórios...`
    };
    salvarJobCruzamento_(job);

    // Rede de segurança: se a página for fechada, o gatilho continua o trabalho
    agendarContinuacao_(CRUZ_CONFIG.TEMPO_LIMITE_MINUTOS + 2);
    return { sucesso: true, job: job };
  } catch (e) {
    return { erro: 'Não foi possível iniciar o cruzamento: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** Executa uma etapa (chamada em ciclo pelo Web App) e devolve o estado atual. */
function processarEtapaCruzamento() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    const ocupado = lerJobCruzamento_() || {};
    ocupado.ocupado = true; // outra etapa (gatilho) está a correr; o Web App tenta de novo
    return ocupado;
  }
  try {
    const job = lerJobCruzamento_();
    if (job && job.status === 'em_andamento') executarEtapa_();
  } finally {
    lock.releaseLock();
  }
  return lerJobCruzamento_();
}

/** Estado do último cruzamento (ou null). */
function obterStatusCruzamento() {
  return lerJobCruzamento_();
}

/** Chamada automaticamente pelo gatilho de tempo. Também pode ser rodada à mão. */
function continuarCruzamento() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // outra etapa ainda rodando; ela cuida do agendamento
  try {
    const job = lerJobCruzamento_();
    if (!job || job.status !== 'em_andamento') {
      removerGatilhos_();
      if (job && job.status === 'cancelado') finalizarLimpezaCruzamento_(SpreadsheetApp.getActiveSpreadsheet());
      return;
    }
    executarEtapa_();
  } finally {
    lock.releaseLock();
  }
}

/** Cancela o cruzamento em andamento e prepara a "NFs_Base" para a próxima leva. */
function cancelarCruzamento() {
  removerGatilhos_();
  const job = atualizarJobCruzamento_({
    status: 'cancelado', fase: 'cancelado', fim: new Date().toISOString(),
    mensagem: 'Cruzamento cancelado.'
  });

  // Se nenhuma etapa estiver a correr, limpa já; senão, a etapa deteta o cancelamento e para.
  const lock = LockService.getScriptLock();
  if (lock.tryLock(1000)) {
    try { finalizarLimpezaCruzamento_(SpreadsheetApp.getActiveSpreadsheet()); } finally { lock.releaseLock(); }
  }
  notificar_('Cancelado', 'Cruzamento de NFs cancelado.', false);
  return job;
}

// ==========================================
// UMA ETAPA DE PROCESSAMENTO
// ==========================================
function executarEtapa_() {
  try {
    executarEtapaInterna_();
  } catch (e) {
    console.error('Erro no cruzamento: ' + (e && e.stack || e));
    falhar_('Erro inesperado no cruzamento: ' + (e && e.message || e));
  }
}

function executarEtapaInterna_() {
  const inicio = Date.now();
  const prazo = inicio + CRUZ_CONFIG.TEMPO_LIMITE_MINUTOS * 60 * 1000;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  // Gatilho de segurança: se esta etapa morrer por qualquer motivo, a próxima
  // começa sozinha depois do limite.
  agendarContinuacao_(CRUZ_CONFIG.TEMPO_LIMITE_MINUTOS + 2);

  const etapa = Number(props.getProperty(CRUZ_PROP_ETAPAS) || 0) + 1;
  props.setProperty(CRUZ_PROP_ETAPAS, String(etapa));
  if (etapa > CRUZ_CONFIG.MAX_ETAPAS) {
    return falhar_(`O cruzamento passou de ${CRUZ_CONFIG.MAX_ETAPAS} etapas e foi parado por segurança. ` +
      'Aumente MAX_ETAPAS ou TEMPO_LIMITE_MINUTOS em Cruzamento.gs e envie as NFs de novo.');
  }

  // 1. Notas a procurar (aba NFs_Base, coluna A a partir da linha 2)
  const abaNFs = ss.getSheetByName(CRUZ_CONFIG.NOME_ABA_NFS);
  if (!abaNFs || abaNFs.getLastRow() < 2) {
    return falhar_(`A aba "${CRUZ_CONFIG.NOME_ABA_NFS}" está vazia. Envie as NFs novamente.`);
  }
  const nfsOrdem = [];
  const notas = new Set();
  abaNFs.getRange(2, 1, abaNFs.getLastRow() - 1, 1).getDisplayValues().forEach(l => {
    const n = limparNota(l[0]);
    if (n && !notas.has(n)) { notas.add(n); nfsOrdem.push(n); }
  });
  if (!notas.size) return falhar_(`A aba "${CRUZ_CONFIG.NOME_ABA_NFS}" não tem NFs válidas.`);

  // 2. Progresso salvo
  const nTrazer = CRUZ_CONFIG.CABECALHOS_TRAZER.length;
  const cacheArq = obterAbaCache_(ss, CRUZ_ABA_ARQUIVOS,
    ['ID', 'Arquivo', 'Linhas lidas', 'Itens encontrados', 'Observação']);
  const cacheRes = obterAbaCache_(ss, CRUZ_ABA_RESULTADOS,
    ['NF'].concat(CRUZ_CONFIG.CABECALHOS_TRAZER, ['Arquivo']));
  const jaProcessados = lerIdsProcessados_(cacheArq);
  const estadoSalvo = lerEstado_();

  // 3. Lista de arquivos
  let arquivos;
  try {
    arquivos = obterListaArquivos_(ss); // lista do Drive feita só na 1ª etapa
  } catch (e) {
    return falhar_('Não foi possível aceder à pasta dos relatórios no Drive: ' + e.message);
  }
  if (!arquivos.length) {
    return falhar_(`Nenhum relatório "${CRUZ_CONFIG.PREFIXO_ARQUIVO}*.csv" foi encontrado na pasta do Drive.`);
  }
  const pendentes = arquivos.filter(a => !jaProcessados.has(a.id));
  let concluidos = arquivos.length - pendentes.length;
  console.log(`▶️ Etapa ${etapa}: ${notas.size} NFs, ${arquivos.length} relatórios, ${pendentes.length} pendentes.`);

  atualizarJobCruzamento_({
    fase: 'lendo', etapa: etapa, relatoriosTotal: arquivos.length, relatoriosFeitos: concluidos,
    mensagem: `A ler os relatórios de faturamento (${concluidos} de ${arquivos.length}).`
  });

  // 4. Processamento
  let pausado = false;
  let feitosAgora = 0;

  for (const arq of pendentes) {
    if (Date.now() > prazo) { pausado = true; break; }
    if (cruzamentoCancelado_()) return finalizarLimpezaCruzamento_(ss);

    const retomada = (estadoSalvo && estadoSalvo.fileId === arq.id) ? estadoSalvo : null;
    if (retomada) console.log(`↪️ Retomando ${arq.nome} a partir do byte ${retomada.pos}`);

    let r;
    const tArq = Date.now();
    try {
      r = processarArquivo_(arq, notas, prazo, retomada);
    } catch (e) {
      console.error(`❌ ${arq.nome}: ${e.message}`);
      anexarLinhas_(cacheArq, [[arq.id, arq.nome, 0, 0, 'ERRO: ' + e.message]]);
      if (retomada) props.deleteProperty(CRUZ_PROP_ESTADO);
      concluidos++;
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
    if (retomada) props.deleteProperty(CRUZ_PROP_ESTADO);
    feitosAgora++;
    concluidos++;
    console.log(`📄 ${arq.nome}: ${(r.bytesLidos / 1048576).toFixed(1)} MB em ${Math.round((Date.now() - tArq) / 1000)}s, ` +
      `${r.linhasLidas} linhas lidas, ${r.totalItens} itens${r.obs ? ' — ' + r.obs : ''}`);

    atualizarJobCruzamento_({
      relatoriosFeitos: concluidos,
      mensagem: `A ler os relatórios de faturamento (${concluidos} de ${arquivos.length}).`
    });
  }

  if (cruzamentoCancelado_()) return finalizarLimpezaCruzamento_(ss);

  if (pausado) {
    SpreadsheetApp.flush();
    agendarContinuacao_(2);
    atualizarJobCruzamento_({
      relatoriosFeitos: concluidos,
      mensagem: `${concluidos} de ${arquivos.length} relatórios lidos. A continuar na próxima etapa...`
    });
    return;
  }

  // 5. Montar a carga pode demorar: se esta etapa já trabalhou bastante,
  //    deixa para a próxima, que começa com o tempo todo livre.
  if (feitosAgora > 0 && Date.now() - inicio > 90 * 1000) {
    agendarContinuacao_(2);
    atualizarJobCruzamento_({
      fase: 'montando', relatoriosFeitos: concluidos,
      mensagem: 'Todos os relatórios foram lidos. A carga será montada na próxima etapa...'
    });
    return;
  }

  // 6. Fim: monta a aba da carga, limpa a NFs_Base e desliga o agendamento
  atualizarJobCruzamento_({ fase: 'montando', relatoriosFeitos: concluidos, mensagem: 'A criar a aba da carga...' });

  const job = lerJobCruzamento_();
  const montagem = montarMatrizCarga_(cacheRes, nTrazer, nfsOrdem);
  const avisos = contarAvisos_(cacheArq);

  if (montagem.matriz.length <= 1) {
    finalizarLimpezaCruzamento_(ss);
    return falhar_('Nenhuma das NFs enviadas foi encontrada nos relatórios de faturamento. A carga não foi criada.', {
      resultado: {
        aba: null, itens: 0, encontradas: 0, totalNFs: notas.size,
        naoEncontradas: montagem.naoEncontradas, linhasSemProduto: montagem.ignoradas,
        avisosRelatorios: avisos, relatorios: arquivos.length, etapas: etapa
      }
    });
  }

  const criada = _criarAbaCarga(job.nomeCarga, montagem.matriz, { comoTexto: true });
  finalizarLimpezaCruzamento_(ss);

  const resultado = {
    aba: criada.aba,
    itens: montagem.matriz.length - 1,
    encontradas: montagem.encontradas,
    totalNFs: notas.size,
    naoEncontradas: montagem.naoEncontradas,
    linhasSemProduto: montagem.ignoradas,
    avisosRelatorios: avisos,
    relatorios: arquivos.length,
    etapas: etapa
  };
  atualizarJobCruzamento_({
    status: 'concluido', fase: 'concluido', fim: new Date().toISOString(), resultado: resultado,
    mensagem: `Carga ${criada.aba} criada com ${resultado.itens} itens.`
  });

  notificar_('Concluído',
    `Carga "${criada.aba}" criada com ${resultado.itens} itens.\n` +
    `NFs encontradas: ${resultado.encontradas} de ${notas.size} (${arquivos.length} relatórios verificados` +
    (etapa > 1 ? ` em ${etapa} etapas` : '') + ').' +
    (resultado.naoEncontradas.length ? `\nNFs não encontradas (${resultado.naoEncontradas.length}): ${resultado.naoEncontradas.join(', ')}` : '') +
    (avisos ? `\n${avisos} relatório(s) com aviso ou erro.` : ''),
    true);
}

function falhar_(msg, extra) {
  removerGatilhos_();
  atualizarJobCruzamento_(Object.assign({
    status: 'erro', fase: 'erro', erro: msg, mensagem: msg, fim: new Date().toISOString()
  }, extra || {}));
  notificar_('Erro', msg, true);
}

// ==========================================
// MONTAGEM DA ABA DA CARGA
// ==========================================
/**
 * Agrupa os itens encontrados por NF (na ordem da NFs_Base) e monta a matriz da carga:
 * ["Nota Fiscal", ...CABECALHOS_TRAZER]. Linhas sem Mercadoria e sem Pedido são
 * ignoradas, porque a Conferência bloqueia cargas com "NF sem produto".
 */
function montarMatrizCarga_(cacheRes, nTrazer, nfsOrdem) {
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

      if (CRUZ_CONFIG.REMOVER_DUPLICADOS_ENTRE_ARQUIVOS) {
        const chave = valores.join('\u0001');
        const origemAnterior = g.chaves.get(chave);
        if (origemAnterior !== undefined && origemAnterior !== origem) return;
        g.chaves.set(chave, origem);
      }
      g.itens.push(valores);
    });
  }

  const idxPed = CRUZ_CONFIG.CABECALHOS_TRAZER.indexOf('Número do Pedido Faturado');
  const idxMerc = CRUZ_CONFIG.CABECALHOS_TRAZER.indexOf('Mercadoria Código');

  const matriz = [[CRUZ_CONFIG.CABECALHO_NF_CARGA].concat(CRUZ_CONFIG.CABECALHOS_TRAZER)];
  const naoEncontradas = [];
  let encontradas = 0;
  let ignoradas = 0;

  nfsOrdem.forEach(nf => {
    const g = mapa.get(nf);
    let usou = false;
    if (g) {
      g.itens.forEach(valores => {
        const merc = idxMerc !== -1 ? String(valores[idxMerc]).trim() : '';
        const ped = idxPed !== -1 ? String(valores[idxPed]).trim() : '';
        if (!merc && !ped) { ignoradas++; return; }
        matriz.push([nf].concat(valores));
        usou = true;
      });
    }
    if (usou) encontradas++; else naoEncontradas.push(nf);
  });

  if (naoEncontradas.length) {
    console.log('🔎 NFs não encontradas: ' + naoEncontradas.slice(0, 50).join(', '));
  }
  return { matriz, encontradas, naoEncontradas, ignoradas };
}

// ==========================================
// NFs_BASE, CACHES E LIMPEZA
// ==========================================
/** Aceita o texto colado do Excel (uma NF por célula/linha) ou uma lista. */
function extrairNotas_(entrada) {
  const lista = Array.isArray(entrada) ? entrada : String(entrada || '').split(/[\s;,|]+/);
  const vistos = new Set();
  const nfs = [];
  lista.forEach(v => {
    // "123456.0" / "123456,00" (número com casas decimais) conta como 123456
    const n = limparNota(String(v == null ? '' : v).trim().replace(/[.,]0+$/, ''));
    if (n && !vistos.has(n)) { vistos.add(n); nfs.push(n); }
  });
  return nfs;
}

function gravarNFsBase_(ss, nfs) {
  let aba = ss.getSheetByName(CRUZ_CONFIG.NOME_ABA_NFS);
  if (!aba) aba = ss.insertSheet(CRUZ_CONFIG.NOME_ABA_NFS);
  aba.clearContents();
  const linhas = [['NF']].concat(nfs.map(n => [n]));
  garantirTamanho_(aba, linhas.length, 1);
  aba.getRange(1, 1, linhas.length, 1).setNumberFormat('@').setValues(linhas);
  aba.getRange(1, 1).setFontWeight('bold');
  aba.setFrozenRows(1);
}

function limparNFsBase_(ss) {
  const aba = ss.getSheetByName(CRUZ_CONFIG.NOME_ABA_NFS);
  if (!aba) return;
  const ultima = aba.getLastRow();
  if (ultima > 1) aba.getRange(2, 1, ultima - 1, Math.max(1, aba.getLastColumn())).clearContent();
  aba.getRange(1, 1).setValue('NF');
}

function apagarCachesCruzamento_(ss) {
  [CRUZ_ABA_ARQUIVOS, CRUZ_ABA_RESULTADOS, CRUZ_ABA_LISTA].forEach(nome => {
    const s = ss.getSheetByName(nome);
    if (s) ss.deleteSheet(s);
  });
}

/** Fim do cruzamento: gatilhos, estado, caches e a NFs_Base ficam prontos para a próxima leva. */
function finalizarLimpezaCruzamento_(ss) {
  removerGatilhos_();
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(CRUZ_PROP_ESTADO);
  props.deleteProperty(CRUZ_PROP_ETAPAS);
  apagarCachesCruzamento_(ss);
  limparNFsBase_(ss);
}

// ==========================================
// ESTADO DO CRUZAMENTO (para o Web App)
// ==========================================
function lerJobCruzamento_() {
  const txt = PropertiesService.getScriptProperties().getProperty(CRUZ_PROP_JOB);
  if (!txt) return null;
  try { return JSON.parse(txt); } catch (e) { return null; }
}

function salvarJobCruzamento_(job) {
  PropertiesService.getScriptProperties().setProperty(CRUZ_PROP_JOB, JSON.stringify(job));
}

/** Atualiza o estado, sem sobrescrever um cruzamento já cancelado/terminado. */
function atualizarJobCruzamento_(alteracoes) {
  const job = lerJobCruzamento_() || {};
  if (job.status && job.status !== 'em_andamento') return job;
  Object.assign(job, alteracoes);
  salvarJobCruzamento_(job);
  return job;
}

function cruzamentoCancelado_() {
  const job = lerJobCruzamento_();
  return !job || job.status !== 'em_andamento';
}

// ==========================================
// LEITURA DE UM RELATÓRIO
// ==========================================
function processarArquivo_(arq, notas, prazo, retomada) {
  const url = `https://www.googleapis.com/drive/v3/files/${arq.id}?alt=media&supportsAllDrives=true`;
  const token = ScriptApp.getOAuthToken();
  const bloco = Math.floor(CRUZ_CONFIG.TAMANHO_BLOCO_MB * 1024 * 1024);

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
      linhasTexto[0] = linhasTexto[0].replace(/^﻿/, '');
      layout = detectarLayout_(linhasTexto);
      if (layout.linhaCabecalho === -1) {
        obs.push(`Cabeçalho "${CRUZ_CONFIG.CABECALHO_BUSCA}" não encontrado; usando letras de reserva`);
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
  const alvoNF = normalizarCabecalho_(CRUZ_CONFIG.CABECALHO_BUSCA);
  const layout = {
    delim: ';',
    idxNota: letraParaIndice(CRUZ_CONFIG.COLUNA_NOTAS_FATURAMENTO_LETRA),
    idxTrazer: CRUZ_CONFIG.COLUNAS_PARA_TRAZER_LETRAS.map(letraParaIndice),
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

    layout.idxTrazer = CRUZ_CONFIG.CABECALHOS_TRAZER.map((c, k) => {
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
  const prefixo = CRUZ_CONFIG.PREFIXO_ARQUIVO.toLowerCase();
  const lista = [];
  const pilha = [DriveApp.getFolderById(folderId)];

  while (pilha.length) {
    const pasta = pilha.pop();
    const it = pasta.searchFiles(`title contains '${CRUZ_CONFIG.PREFIXO_ARQUIVO.replace(/'/g, "\\'")}' and trashed = false`);
    while (it.hasNext()) {
      const f = it.next();
      const nome = f.getName();
      const nomeMin = nome.toLowerCase();
      if (!nomeMin.startsWith(prefixo)) continue;
      if (f.getMimeType() !== MimeType.CSV && !nomeMin.endsWith('.csv')) continue;
      lista.push({ id: f.getId(), nome: nome });
    }
    if (CRUZ_CONFIG.INCLUIR_SUBPASTAS) {
      const subs = pasta.getFolders();
      while (subs.hasNext()) pilha.push(subs.next());
    }
  }
  lista.sort((a, b) => a.nome.localeCompare(b.nome));
  return lista;
}

function obterListaArquivos_(ss) {
  const s = ss.getSheetByName(CRUZ_ABA_LISTA);
  if (s && s.getLastRow() > 1) {
    return s.getRange(2, 1, s.getLastRow() - 1, 2).getValues()
      .map(r => ({ id: String(r[0]), nome: String(r[1]) }));
  }
  const lista = listarArquivos_(CRUZ_CONFIG.FOLDER_ID);
  const aba = obterAbaCache_(ss, CRUZ_ABA_LISTA, ['ID', 'Arquivo']);
  anexarLinhas_(aba, lista.map(a => [a.id, a.nome]));
  return lista;
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
    .replace(/^﻿/, '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
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
  ScriptApp.newTrigger(CRUZ_FUNCAO_CONTINUACAO).timeBased().after(Math.round(minutos * 60 * 1000)).create();
}

function removerGatilhos_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === CRUZ_FUNCAO_CONTINUACAO) ScriptApp.deleteTrigger(t);
  });
}

function lerEstado_() {
  const txt = PropertiesService.getScriptProperties().getProperty(CRUZ_PROP_ESTADO);
  if (!txt) return null;
  try { return JSON.parse(txt); } catch (e) { return null; }
}

function salvarEstado_(estado) {
  PropertiesService.getScriptProperties().setProperty(CRUZ_PROP_ESTADO, JSON.stringify(estado));
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

  if (!final || CRUZ_CONFIG.EMAIL_AVISO === null) return;
  try {
    const destino = CRUZ_CONFIG.EMAIL_AVISO || Session.getEffectiveUser().getEmail();
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

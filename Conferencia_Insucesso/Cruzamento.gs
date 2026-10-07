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

  // Aba com os destinatários do e-mail de resumo (uma linha por e-mail; coluna "Ativo" Sim/Não).
  // É criada automaticamente com os EMAILS_INICIAIS se ainda não existir.
  NOME_ABA_EMAILS: 'Config_Emails',
  EMAILS_INICIAIS: ['arthur.silva@kabum.com.br'],

  // Aba permanente com as NFs que não estavam nos relatórios (para preenchimento manual)
  NOME_ABA_PENDENTES: 'NFs_Nao_Encontradas',

  // Colunas da carga usadas no resumo do e-mail
  COLUNA_PEDIDO: 'Número do Pedido Faturado',
  COLUNA_ETIQUETA: 'Mercadoria Código',
  COLUNA_CUSTO: 'Custo Produto'
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
    .addSeparator()
    .addItem('Romaneio: recriar modelo no Docs', 'recriarModeloRomaneio')
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
    try { garantirAbaEmails_(ss); } catch (e) { console.warn('Aba de e-mails: ' + e.message); }
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
    return { erro: explicarErroPermissao_('Não foi possível iniciar o cruzamento: ' + e.message) };
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
    falhar_(explicarErroPermissao_('Erro inesperado no cruzamento: ' + (e && e.message || e)));
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
    return falhar_(explicarErroPermissao_('Não foi possível aceder à pasta dos relatórios no Drive: ' + e.message));
  }
  if (!arquivos.length) {
    return falhar_(`Nenhum relatório "${CRUZ_CONFIG.PREFIXO_ARQUIVO}*.csv" foi encontrado na pasta do Drive.`);
  }
  const pendentes = arquivos.filter(a => !jaProcessados.has(a.id));
  let concluidos = arquivos.length - pendentes.length;
  console.log(`▶️ Etapa ${etapa}: ${notas.size} NFs, ${arquivos.length} relatórios, ${pendentes.length} pendentes.`);

  if (pendentes.length) {
    atualizarJobCruzamento_({
      fase: 'lendo', etapa: etapa, relatoriosTotal: arquivos.length, relatoriosFeitos: concluidos,
      mensagem: `A ler os relatórios de faturamento (${concluidos} de ${arquivos.length}).`
    });
  } else {
    // Relatórios já lidos: esta etapa continua o tracking ou monta a carga
    atualizarJobCruzamento_({ etapa: etapa, relatoriosTotal: arquivos.length, relatoriosFeitos: concluidos });
  }

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

  // 4b. Tracking Intelipost: para os pedidos encontrados no faturamento, lê as ocorrências
  //     nas planilhas do drive de tracking (também por etapas)
  if (typeof TRACK_CONFIG !== 'undefined' && TRACK_CONFIG.ATIVO) {
    const jobT = lerJobCruzamento_() || {};
    if (!jobT.trackingConcluido) {
      const pedidos = pedidosDoCacheCruzamento_(cacheRes, nTrazer);
      if (pedidos.size) {
        atualizarJobCruzamento_({ fase: 'tracking', relatoriosFeitos: concluidos,
          mensagem: 'A ler o tracking Intelipost dos ' + pedidos.size + ' pedidos encontrados...' });
        const t = etapaTracking_(ss, pedidos, prazo);
        if (t.cancelado || cruzamentoCancelado_()) return finalizarLimpezaCruzamento_(ss);
        if (t.pausado) {
          SpreadsheetApp.flush();
          agendarContinuacao_(2);
          atualizarJobCruzamento_({ fase: 'tracking', trackingTotal: t.total, trackingFeitos: t.feitos,
            mensagem: `Tracking: ${t.feitos} de ${t.total} planilhas lidas. A continuar na próxima etapa...` });
          return;
        }
        atualizarJobCruzamento_({ trackingConcluido: true, trackingTotal: t.total, trackingFeitos: t.feitos,
          trackingErro: t.erro || null });
        if (t.trabalhou) feitosAgora++;
      } else {
        atualizarJobCruzamento_({ trackingConcluido: true, trackingTotal: 0, trackingFeitos: 0 });
      }
    }
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

  // Tracking: "Avaliação" (Não Receber, prazo de devolução) e "Avaliação transportes" (Recusa).
  // A referência do prazo é o momento da criação da base.
  let tracking = null;
  if (typeof TRACK_CONFIG !== 'undefined' && TRACK_CONFIG.ATIVO) {
    try {
      tracking = resumirTrackingParaJob_(aplicarTrackingNaMatriz_(ss, montagem.matriz, Date.now()), job);
    } catch (e) {
      console.error('Tracking não aplicado: ' + e.message);
      tracking = { erro: 'O tracking não foi aplicado: ' + e.message };
    }
  }

  const resumo = calcularResumoCarga_(montagem.matriz);
  const criada = _criarAbaCarga(job.nomeCarga, montagem.matriz, { comoTexto: true });
  finalizarLimpezaCruzamento_(ss);
  try {
    registrarNFsNaoEncontradas_(ss, criada.aba, montagem.naoEncontradas, job.usuario);
  } catch (e) {
    console.error('Não foi possível registar as NFs não encontradas: ' + e.message);
  }

  const resultado = {
    aba: criada.aba,
    resumo: resumo,
    itens: montagem.matriz.length - 1,
    encontradas: montagem.encontradas,
    totalNFs: notas.size,
    naoEncontradas: montagem.naoEncontradas,
    linhasSemProduto: montagem.ignoradas,
    avisosRelatorios: avisos,
    relatorios: arquivos.length,
    etapas: etapa,
    tracking: tracking
  };
  atualizarJobCruzamento_({
    status: 'concluido', fase: 'concluido', fim: new Date().toISOString(), resultado: resultado,
    mensagem: `Carga ${criada.aba} criada com ${resultado.itens} itens.`
  });

  notificar_('Concluído', `Carga "${criada.aba}" criada com ${resultado.itens} itens.`, true);
  enviarEmailConclusao_(ss, resultado, lerJobCruzamento_());
}

/** Pedidos (Número do Pedido Faturado) dos itens encontrados no faturamento. */
function pedidosDoCacheCruzamento_(cacheRes, nTrazer) {
  const pedidos = new Set();
  const idx = CRUZ_CONFIG.CABECALHOS_TRAZER.indexOf(CRUZ_CONFIG.COLUNA_PEDIDO);
  const n = cacheRes.getLastRow() - 1;
  if (idx < 0 || n < 1) return pedidos;
  cacheRes.getRange(2, 2 + idx, n, 1).getValues().forEach(l => {
    const p = limparPedidoTracking_(l[0]);
    if (p) pedidos.add(p);
  });
  return pedidos;
}

/** Resumo do tracking com listas limitadas (o estado do job vive numa propriedade de ~9 KB). */
function resumirTrackingParaJob_(t, job) {
  const LIMITE = 120;
  const cortar = lista => ({ total: lista.length, itens: lista.slice(0, LIMITE) });
  return {
    planilhas: job.trackingTotal || 0,
    erro: job.trackingErro || null,
    pedidos: t.pedidos,
    encontrados: t.encontrados,
    naoReceber: cortar(t.naoReceber),
    recusas: cortar(t.recusas),
    semTracking: cortar(t.semTracking),
    semDevolucao: cortar(t.semDevolucao)
  };
}

/** Linhas do tracking para o e-mail de conclusão (texto simples). */
function textoTrackingEmail_(t) {
  if (!t) return '';
  if (t.erro && !t.pedidos) return 'TRACKING INTELIPOST: ' + t.erro;
  const lista = g => g.itens.map(x => x.nfs || x.pedido).join(', ') + (g.total > g.itens.length ? ` (+${g.total - g.itens.length})` : '');
  return [
    'TRACKING INTELIPOST',
    `Pedidos consultados: ${t.pedidos} · com tracking: ${t.encontrados} · planilhas lidas: ${t.planilhas}`,
    `Fora do prazo de devolução (${TRACK_CONFIG.PRAZO_DEVOLUCAO_DIAS} dias), marcados "Não Receber": ${t.naoReceber.total}` + (t.naoReceber.total ? ' — NFs ' + lista(t.naoReceber) : ''),
    `Recusas dentro da promessa de entrega ("Avaliação transportes" = Recusa): ${t.recusas.total}` + (t.recusas.total ? ' — NFs ' + lista(t.recusas) : ''),
    t.semTracking.total ? `Sem tracking: ${t.semTracking.total} — NFs ${lista(t.semTracking)}` : '',
    t.semDevolucao.total ? `Sem ocorrência de devolução: ${t.semDevolucao.total} — NFs ${lista(t.semDevolucao)}` : '',
    t.erro ? 'Aviso: ' + t.erro : ''
  ].filter(Boolean).join('\n');
}

/** Bloco HTML do tracking para o e-mail de conclusão. */
function htmlTrackingEmail_(t) {
  if (!t) return '';
  const linhas = textoTrackingEmail_(t).split('\n').slice(1);
  return `<div style="margin:12px 0;padding:12px 14px;border:1px solid #cbd5e1;background:#f8fafc;border-radius:8px;font-size:13px">` +
    `<b>Tracking Intelipost</b><br>` + linhas.map(escaparHtmlEmail_).join('<br>') + `</div>`;
}

function falhar_(msg, extra) {
  removerGatilhos_();
  atualizarJobCruzamento_(Object.assign({
    status: 'erro', fase: 'erro', erro: msg, mensagem: msg, fim: new Date().toISOString()
  }, extra || {}));
  notificar_('Erro', msg, true);
  enviarEmailErro_(SpreadsheetApp.getActiveSpreadsheet(), msg, lerJobCruzamento_());
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
  const caches = [CRUZ_ABA_ARQUIVOS, CRUZ_ABA_RESULTADOS, CRUZ_ABA_LISTA]
    .concat(typeof nomesCacheTracking_ === 'function' ? nomesCacheTracking_() : []);
  caches.forEach(nome => {
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
/**
 * Os gatilhos de tempo são só a rede de segurança para continuar com a página fechada.
 * Se a permissão "script.scriptapp" ainda não foi concedida por quem publicou o Web App,
 * o cruzamento NÃO falha: continua conduzido pela página aberta e mostra um aviso.
 */
function agendarContinuacao_(minutos) {
  if (gatilhosIndisponiveisNoJob_()) return false;
  if (!removerGatilhos_()) return false;
  try {
    ScriptApp.newTrigger(CRUZ_FUNCAO_CONTINUACAO).timeBased().after(Math.round(minutos * 60 * 1000)).create();
    return true;
  } catch (e) {
    registrarFalhaGatilho_(e);
    return false;
  }
}

function removerGatilhos_() {
  if (gatilhosIndisponiveisNoJob_()) return false;
  try {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === CRUZ_FUNCAO_CONTINUACAO) ScriptApp.deleteTrigger(t);
    });
    return true;
  } catch (e) {
    registrarFalhaGatilho_(e);
    return false;
  }
}

/** Neste cruzamento já se sabe que os gatilhos não estão autorizados: não volta a tentar. */
function gatilhosIndisponiveisNoJob_() {
  const job = lerJobCruzamento_();
  return !!(job && job.status === 'em_andamento' && job.semGatilhos);
}

function registrarFalhaGatilho_(e) {
  const job = lerJobCruzamento_();
  if (!job || job.status !== 'em_andamento' || job.semGatilhos) return;
  console.warn('Gatilhos indisponíveis (o cruzamento continua pela página aberta): ' + (e && e.message || e));
  atualizarJobCruzamento_({
    semGatilhos: true,
    aviso: 'A continuação automática em segundo plano não está autorizada (permissão de gatilhos). ' +
      'Mantenha esta página aberta até o cruzamento terminar.',
    avisoUrl: obterUrlAutorizacao_()
  });
}

/**
 * Link de autorização das permissões em falta para a conta que executa o script
 * (no Web App: a conta que o publicou). null se já está tudo autorizado ou se indisponível.
 */
function obterUrlAutorizacao_() {
  try {
    const info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    if (info.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) return info.getAuthorizationUrl();
  } catch (e) { /* indisponível neste contexto */ }
  return null;
}

/** Acrescenta a orientação de autorização quando o erro é de permissão. */
function explicarErroPermissao_(msg) {
  msg = String(msg || '');
  if (!/permiss|authoriz|autoriza|scope/i.test(msg)) return msg;
  const url = obterUrlAutorizacao_();
  return msg + ' — Quem publicou o Web App deve autorizar todas as permissões' +
    (url ? ` (${url})` : ' executando "forcarPermissoes" no editor do Apps Script') +
    ' e marcar TODAS as caixas na janela do Google.';
}

function lerEstado_() {
  const txt = PropertiesService.getScriptProperties().getProperty(CRUZ_PROP_ESTADO);
  if (!txt) return null;
  try { return JSON.parse(txt); } catch (e) { return null; }
}

function salvarEstado_(estado) {
  PropertiesService.getScriptProperties().setProperty(CRUZ_PROP_ESTADO, JSON.stringify(estado));
}

/** Mostra um aviso rápido na planilha, se estiver aberta. O e-mail tem motor próprio (abaixo). */
function notificar_(titulo, mensagem, final) {
  console.log(`[${titulo}] ${mensagem}`);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(mensagem, 'Cruzamento NF — ' + titulo, final ? 60 : 15);
  } catch (e) { /* sem planilha aberta */ }
}

// ==========================================
// RESUMO DA CARGA (pedidos, etiquetas, valor)
// ==========================================
/**
 * Resumo da matriz da carga (linha 0 = cabeçalho):
 *  - pedidos: contagem ÚNICA de "Número do Pedido Faturado" (coluna C)
 *  - etiquetas: contagem de "Mercadoria Código" preenchida (coluna D)
 *  - valorTotal: soma de "Custo Produto" (coluna F), que vem no formato americano (1,234.56)
 */
function calcularResumoCarga_(matriz) {
  const cab = matriz[0].map(c => normalizarCabecalho_(c));
  const idx = nome => cab.indexOf(normalizarCabecalho_(nome));
  const iPed = idx(CRUZ_CONFIG.COLUNA_PEDIDO);
  const iEtq = idx(CRUZ_CONFIG.COLUNA_ETIQUETA);
  const iCusto = idx(CRUZ_CONFIG.COLUNA_CUSTO);

  const pedidos = new Set();
  let etiquetas = 0;
  let centavos = 0;
  let valoresInvalidos = 0;

  for (let i = 1; i < matriz.length; i++) {
    const l = matriz[i];
    const ped = iPed !== -1 ? String(l[iPed] == null ? '' : l[iPed]).trim() : '';
    if (ped) pedidos.add(ped);
    if (iEtq !== -1 && String(l[iEtq] == null ? '' : l[iEtq]).trim()) etiquetas++;
    if (iCusto !== -1) {
      const bruto = String(l[iCusto] == null ? '' : l[iCusto]).trim();
      if (!bruto) continue;
      const v = converterValor_(bruto);
      if (isNaN(v)) valoresInvalidos++; else centavos += Math.round(v * 100);
    }
  }
  return { pedidos: pedidos.size, etiquetas: etiquetas, valorTotal: centavos / 100, valoresInvalidos: valoresInvalidos };
}

/**
 * Converte um valor monetário em número. O relatório usa o formato americano
 * (ponto decimal, vírgula de milhar: "1,234.56"); também aceita "1.234,56", "R$ 12,50" e "(10.00)".
 */
function converterValor_(v) {
  if (typeof v === 'number') return v;
  let s = String(v == null ? '' : v).trim();
  if (!s) return 0;
  const negativo = /^\(.*\)$/.test(s) || /^-/.test(s.replace(/^[^\d(-]+/, ''));
  s = s.replace(/[^\d.,]/g, '');
  if (!s) return NaN;

  const ponto = s.lastIndexOf('.');
  const virgula = s.lastIndexOf(',');
  if (ponto !== -1 && virgula !== -1) {
    // O último separador é o decimal
    s = ponto > virgula ? s.replace(/,/g, '') : s.replace(/\./g, '').replace(',', '.');
  } else if (virgula !== -1) {
    // Só vírgulas: "1,234" / "1,234,567" = milhar (americano); "12,5" / "12,50" = decimal
    s = /^\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
  } else if ((s.match(/\./g) || []).length > 1) {
    s = s.replace(/\./g, ''); // "1.234.567" = milhar
  }
  const n = Number(s);
  return isNaN(n) ? NaN : (negativo ? -n : n);
}

function formatarMoeda_(valor) {
  const negativo = valor < 0;
  const partes = Math.abs(valor).toFixed(2).split('.');
  const inteiro = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (negativo ? '-' : '') + 'R$ ' + inteiro + ',' + partes[1];
}

function formatarNumero_(n) {
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

// ==========================================
// NFs NÃO ENCONTRADAS (registo permanente)
// ==========================================
const CRUZ_CAB_PENDENTES = ['Carga', 'NF', 'Registada em', 'Registada por', 'Situação', 'Resolvida em', 'Resolvida por'];

function obterAbaPendentes_(ss) {
  let aba = ss.getSheetByName(CRUZ_CONFIG.NOME_ABA_PENDENTES);
  if (!aba) {
    aba = ss.insertSheet(CRUZ_CONFIG.NOME_ABA_PENDENTES);
    aba.getRange(1, 1, 1, CRUZ_CAB_PENDENTES.length).setValues([CRUZ_CAB_PENDENTES]).setFontWeight('bold').setBackground('#f1f5f9');
    aba.setFrozenRows(1);
  }
  return aba;
}

function dataHoraAgora_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
}

/** Acrescenta as NFs sem dados no faturamento à aba permanente, como "Pendente". */
function registrarNFsNaoEncontradas_(ss, nomeCarga, nfs, usuario) {
  if (!nfs || !nfs.length) return 0;
  const aba = obterAbaPendentes_(ss);
  const quando = dataHoraAgora_();
  const linhas = nfs.map(nf => [nomeCarga, String(nf), quando, usuario || '', 'Pendente', '', '']);
  const ini = aba.getLastRow() + 1;
  garantirTamanho_(aba, ini + linhas.length - 1, CRUZ_CAB_PENDENTES.length);
  aba.getRange(ini, 1, linhas.length, CRUZ_CAB_PENDENTES.length).setNumberFormat('@').setValues(linhas);
  return linhas.length;
}

function lerPendentes_(ss) {
  const aba = ss.getSheetByName(CRUZ_CONFIG.NOME_ABA_PENDENTES);
  if (!aba || aba.getLastRow() < 2) return { aba: aba, linhas: [] };
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, CRUZ_CAB_PENDENTES.length).getDisplayValues();
  return { aba: aba, linhas: dados.map((l, i) => ({ linha: i + 2, carga: l[0], nf: l[1], registadaEm: l[2], situacao: l[4] })) };
}

/** Web App: NFs ainda pendentes de uma carga. */
function obterNFsPendentes(nomeCarga) {
  try {
    const alvo = String(nomeCarga || '').trim();
    return lerPendentes_(SpreadsheetApp.getActiveSpreadsheet()).linhas
      .filter(p => p.carga === alvo && p.situacao === 'Pendente')
      .map(p => ({ nf: p.nf, registadaEm: p.registadaEm }));
  } catch (e) {
    return [];
  }
}

/** Web App: marca uma NF como resolvida (dados preenchidos manualmente) e devolve as restantes. */
function marcarNFResolvida(nomeCarga, nf) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { aba, linhas } = lerPendentes_(ss);
  const alvo = linhas.find(p => p.carga === String(nomeCarga) && p.nf === String(nf) && p.situacao === 'Pendente');
  if (aba && alvo) {
    let usuario = '';
    try { usuario = Session.getActiveUser().getEmail(); } catch (e) { /* sem acesso */ }
    aba.getRange(alvo.linha, 5, 1, 3).setValues([['Resolvida', dataHoraAgora_(), usuario]]);
  }
  return obterNFsPendentes(nomeCarga);
}

// ==========================================
// E-MAIL DE RESUMO (destinatários na aba Config_Emails)
// ==========================================
function garantirAbaEmails_(ss) {
  let aba = ss.getSheetByName(CRUZ_CONFIG.NOME_ABA_EMAILS);
  if (aba) return aba;
  aba = ss.insertSheet(CRUZ_CONFIG.NOME_ABA_EMAILS);
  const linhas = [['E-mail', 'Ativo (Sim/Não)', 'Nome', 'Observação']]
    .concat(CRUZ_CONFIG.EMAILS_INICIAIS.map(e => [e, 'Sim', '', '']));
  aba.getRange(1, 1, linhas.length, 4).setValues(linhas);
  aba.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f1f5f9');
  aba.getRange(1, 1).setNote('Um e-mail por linha. Estes endereços recebem o resumo de cada carga criada pelo ' +
    'Cruzamento de NFs. Escreva "Não" na coluna Ativo para suspender um endereço sem o apagar.');
  aba.setFrozenRows(1);
  return aba;
}

/** Destinatários ativos e válidos da aba Config_Emails (sem repetidos). */
function lerDestinatarios_(ss) {
  const aba = ss.getSheetByName(CRUZ_CONFIG.NOME_ABA_EMAILS);
  if (!aba || aba.getLastRow() < 2) return [];
  const vistos = new Set();
  return aba.getRange(2, 1, aba.getLastRow() - 1, 2).getDisplayValues()
    .filter(l => {
      const email = String(l[0]).trim().toLowerCase();
      const ativo = String(l[1]).trim().toLowerCase();
      if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(email)) return false;
      if (/^(n|nao|não|false|falso|0|inativo)$/.test(ativo)) return false;
      if (vistos.has(email)) return false;
      vistos.add(email);
      return true;
    })
    .map(l => String(l[0]).trim());
}

function enviarEmail_(ss, assunto, texto, html) {
  let destinatarios = [];
  try {
    garantirAbaEmails_(ss);
    destinatarios = lerDestinatarios_(ss);
  } catch (e) {
    console.error('Não foi possível ler a aba de e-mails: ' + e.message);
  }
  if (!destinatarios.length) {
    console.log(`Nenhum destinatário ativo em "${CRUZ_CONFIG.NOME_ABA_EMAILS}": e-mail não enviado.`);
    return 0;
  }
  try {
    MailApp.sendEmail({ to: destinatarios.join(','), subject: assunto, body: texto, htmlBody: html, name: 'Controle de Insucessos' });
    return destinatarios.length;
  } catch (e) {
    console.error('Não foi possível enviar o e-mail: ' + e.message);
    return 0;
  }
}

function escaparHtmlEmail_(t) {
  return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function linkAba_(ss, nomeAba) {
  let url = '';
  try { url = ss.getUrl(); } catch (e) { return ''; }
  try {
    const aba = nomeAba ? ss.getSheetByName(nomeAba) : null;
    if (aba) url += '#gid=' + aba.getSheetId();
  } catch (e) { /* sem id */ }
  return url;
}

function enviarEmailConclusao_(ss, r, job) {
  const rs = r.resumo || { pedidos: 0, etiquetas: 0, valorTotal: 0, valoresInvalidos: 0 };
  const quando = dataHoraAgora_();
  const url = linkAba_(ss, r.aba);
  const falta = r.naoEncontradas || [];
  const assunto = `[Controle de Insucessos] Carga ${r.aba} criada — ${formatarNumero_(rs.pedidos)} pedidos · ` +
    `${formatarNumero_(rs.etiquetas)} etiquetas · ${formatarMoeda_(rs.valorTotal)}`;

  const texto = [
    `Carga ${r.aba} criada pelo Cruzamento de NFs em ${quando}` + (job && job.usuario ? ` (iniciado por ${job.usuario})` : '') + '.',
    '',
    'RESUMO DA CARGA',
    `Pedidos (Número do Pedido Faturado, únicos): ${formatarNumero_(rs.pedidos)}`,
    `Etiquetas (Mercadoria Código): ${formatarNumero_(rs.etiquetas)}`,
    `Valor total de devolução (Custo Produto): ${formatarMoeda_(rs.valorTotal)}`,
    rs.valoresInvalidos ? `Atenção: ${rs.valoresInvalidos} valor(es) de custo não reconhecido(s) ficaram fora da soma.` : '',
    '',
    `NFs enviadas: ${r.totalNFs} · encontradas: ${r.encontradas} · não encontradas: ${falta.length}`,
    falta.length ? `NFs não encontradas (preencher manualmente; registadas na aba ${CRUZ_CONFIG.NOME_ABA_PENDENTES}): ${falta.join(', ')}` : '',
    `Relatórios de faturamento verificados: ${r.relatorios}` + (r.avisosRelatorios ? ` (${r.avisosRelatorios} com aviso)` : ''),
    '',
    textoTrackingEmail_(r.tracking),
    '',
    url ? `Planilha: ${url}` : ''
  ].filter((l, i, a) => l !== '' || (a[i - 1] !== '' && i > 0)).join('\n');

  const cartao = (rotulo, valor) =>
    `<td style="padding:14px 16px;border:1px solid #e2e8f0;border-radius:8px;background:#f8fafc;width:33%">` +
    `<div style="font-size:12px;color:#64748b">${rotulo}</div>` +
    `<div style="font-size:22px;font-weight:700;color:#0f172a;margin-top:4px">${valor}</div></td>`;

  const html =
    `<div style="font-family:Arial,Helvetica,sans-serif;color:#0f172a;max-width:640px">` +
    `<h2 style="margin:0 0 4px;font-size:20px">Carga ${escaparHtmlEmail_(r.aba)} criada</h2>` +
    `<p style="margin:0 0 16px;color:#64748b;font-size:13px">Cruzamento de NFs concluído em ${quando}` +
    (job && job.usuario ? ` · iniciado por ${escaparHtmlEmail_(job.usuario)}` : '') + `</p>` +
    `<table role="presentation" cellspacing="8" cellpadding="0" style="width:100%;margin:0 -8px 8px"><tr>` +
    cartao('Pedidos (únicos)', formatarNumero_(rs.pedidos)) +
    cartao('Etiquetas', formatarNumero_(rs.etiquetas)) +
    cartao('Valor total de devolução', formatarMoeda_(rs.valorTotal)) +
    `</tr></table>` +
    (rs.valoresInvalidos ? `<p style="color:#b45309;font-size:13px">Atenção: ${rs.valoresInvalidos} valor(es) de custo não reconhecido(s) ficaram fora da soma.</p>` : '') +
    `<p style="font-size:14px;margin:12px 0 4px">NFs enviadas: <b>${r.totalNFs}</b> · encontradas: <b>${r.encontradas}</b> · não encontradas: <b>${falta.length}</b></p>` +
    (falta.length
      ? `<div style="margin:8px 0;padding:12px 14px;border:1px solid #fcd34d;background:#fffbeb;border-radius:8px;font-size:13px">` +
        `<b>NFs não encontradas nos relatórios de faturamento</b> — preencher manualmente na aba da carga. ` +
        `Ficam registadas na aba <b>${CRUZ_CONFIG.NOME_ABA_PENDENTES}</b>.<br><span style="font-family:monospace">${falta.map(escaparHtmlEmail_).join(', ')}</span></div>`
      : '') +
    `<p style="font-size:13px;color:#64748b">Relatórios de faturamento verificados: ${r.relatorios}` +
    (r.avisosRelatorios ? ` (${r.avisosRelatorios} com aviso)` : '') + `</p>` +
    htmlTrackingEmail_(r.tracking) +
    (url ? `<p><a href="${escaparHtmlEmail_(url)}" style="display:inline-block;padding:10px 16px;background:#c2410c;color:#ffffff;text-decoration:none;border-radius:6px;font-weight:bold">Abrir a carga na planilha</a></p>` : '') +
    `</div>`;

  return enviarEmail_(ss, assunto, texto, html);
}

function enviarEmailErro_(ss, msg, job) {
  const carga = job && job.nomeCarga ? job.nomeCarga : '';
  const falta = job && job.resultado && job.resultado.naoEncontradas ? job.resultado.naoEncontradas : [];
  const url = linkAba_(ss, null);
  const assunto = `[Controle de Insucessos] Erro no cruzamento${carga ? ' — carga ' + carga : ''}`;
  const texto = `${msg}` + (falta.length ? `\n\nNFs não encontradas: ${falta.join(', ')}` : '') + (url ? `\n\nPlanilha: ${url}` : '');
  const html = `<div style="font-family:Arial,Helvetica,sans-serif;color:#0f172a;max-width:640px">` +
    `<h2 style="margin:0 0 8px;font-size:20px;color:#b91c1c">Erro no cruzamento${carga ? ' — carga ' + escaparHtmlEmail_(carga) : ''}</h2>` +
    `<p style="font-size:14px">${escaparHtmlEmail_(msg)}</p>` +
    (falta.length ? `<p style="font-size:13px"><b>NFs não encontradas:</b> <span style="font-family:monospace">${falta.map(escaparHtmlEmail_).join(', ')}</span></p>` : '') +
    (url ? `<p><a href="${escaparHtmlEmail_(url)}">Abrir a planilha</a></p>` : '') + `</div>`;
  return enviarEmail_(ss, assunto, texto, html);
}

/**
 * Pede/valida todas as permissões usadas pela Conferência e pelo Cruzamento.
 *
 * COMO USAR: no editor do Apps Script, com a conta que PUBLICA o Web App,
 * selecione "forcarPermissoes" na barra de funções e clique em Executar.
 * O Google mostra a janela de autorização; depois publique uma nova versão
 * (Implantar › Gerenciar implantações › editar › Nova versão).
 * Também está no menu "Cruzamento NF" da planilha.
 *
 * Mostra o resultado de cada serviço (alerta na planilha ou registo de execução no editor).
 */
function forcarPermissoes() {
  // Consentimento granular do Google: o script pode ter só parte das permissões.
  // requireAllScopes interrompe esta execução e abre a janela de autorização com as que faltam
  // (no editor ou pelo menu da planilha). Depois de aceitar, execute de novo para ver o resultado.
  // Sem try/catch de propósito: a interrupção é o que faz o Google mostrar a janela.
  if (typeof ScriptApp.requireAllScopes === 'function') ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);

  const testes = [
    ['Planilha', () => SpreadsheetApp.getActiveSpreadsheet().getName()],
    ['Drive: pasta dos relatórios de faturamento', () => DriveApp.getFolderById(CRUZ_CONFIG.FOLDER_ID).getName()],
    ['Pedidos externos (download dos relatórios)', () => {
      const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about?fields=user', {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
      });
      if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
      return 'OK';
    }],
    ['Gatilhos de tempo (continuação automática)', () => ScriptApp.getProjectTriggers().length + ' gatilho(s)'],
    ['Envio de e-mail', () => MailApp.getRemainingDailyQuota() + ' e-mails restantes hoje'],
    ['E-mail da conta', () => Session.getEffectiveUser().getEmail()],
    ['Drive partilhado dos romaneios', () => {
      if (typeof ROMANEIO_CONFIG === 'undefined' || !ROMANEIO_CONFIG.DRIVE_ID) return 'não configurado';
      return DriveApp.getFolderById(ROMANEIO_CONFIG.DRIVE_ID).getName();
    }],
    ['Drive do tracking Intelipost', () => {
      if (typeof TRACK_CONFIG === 'undefined' || !TRACK_CONFIG.ATIVO) return 'desativado';
      const pasta = DriveApp.getFolderById(TRACK_CONFIG.PASTA_ID);
      const it = pasta.getFilesByType(MimeType.GOOGLE_SHEETS);
      let n = 0; while (it.hasNext() && n < 500) { it.next(); n++; }
      return pasta.getName() + ' (' + n + ' planilha(s) na raiz)';
    }],
    ['Google Docs (romaneio em PDF)', () => {
      // Abrir um documento exige a permissão "documents"; usa o modelo se já existir
      const id = PropertiesService.getScriptProperties().getProperty(
        typeof propModeloRomaneio_ === 'function' ? propModeloRomaneio_() : 'ROMANEIO_MODELO_ID');
      if (id) {
        try { return DocumentApp.openById(id).getName(); } catch (e) { if (!/not found|não encontrad|missing/i.test(e.message)) throw e; }
      }
      const teste = DocumentApp.create('teste_permissao_romaneio');
      DriveApp.getFileById(teste.getId()).setTrashed(true);
      return 'OK';
    }]
  ];

  const linhas = [];
  let falhas = 0;
  testes.forEach(([nome, fn]) => {
    try {
      const r = fn();
      linhas.push('✔ ' + nome + (r !== undefined && r !== '' ? ' — ' + r : ''));
    } catch (e) {
      falhas++;
      linhas.push('✖ ' + nome + ' — ' + (e && e.message || e));
    }
  });

  const titulo = falhas ? `Faltam ${falhas} permissão(ões)` : 'Todas as permissões estão concedidas';
  const url = falhas ? obterUrlAutorizacao_() : null;
  const texto = linhas.join('\n') + (falhas
    ? '\n\nO Google mostra uma caixa de seleção por permissão: é preciso marcar TODAS (ou "Selecionar tudo").' +
      (url ? '\n\nAbra este link com a conta que publica o Web App, marque todas as caixas e confirme:\n' + url : '') +
      '\n\nSe a janela não voltar a aparecer: em https://myaccount.google.com/connections remova o acesso ' +
      'deste projeto e execute "forcarPermissoes" de novo no editor.'
    : '\n\nSe o Web App ainda mostrar erro de permissão, publique uma nova versão da implantação.');

  console.log(titulo + '\n' + texto);
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(titulo, texto, ui.ButtonSet.OK);
  } catch (e) {
    // Executado a partir do editor: o resultado fica no "Registo de execução".
  }
  return titulo + '\n' + texto;
}

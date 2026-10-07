// ==========================================
// TRACKING INTELIPOST (etapa do Cruzamento de NFs)
// ==========================================
// Para cada pedido da carga (vindo do faturamento), procura as ocorrências de tracking nas planilhas
// do drive de tracking. Cada planilha tem as colunas "Número do Pedido | Coordenadas Maps |
// Payload JSON 1 | Payload JSON 2 | ...": o JSON da Intelipost é o texto das colunas de payload
// juntas pela ordem do número. Os ficheiros .json do drive são ignorados.
//
// Regras (avaliarTrackingIntelipost_):
//  - Devolução: ocorrências cujo nome (micro-estado) ou mensagem da transportadora contém
//    "devolu"/"devolv" (sem acentos) ou um dos PADROES_DEVOLUCAO. Conta a data mais antiga.
//    Se, até ao dia da criação da base, passaram mais de PRAZO_DEVOLUCAO_DIAS dias, o pedido
//    entra com a Avaliação "Não Receber".
//  - Recusa: ocorrências com um dos PADROES_RECUSA. Se a primeira recusa for até à data de
//    promessa de entrega ao cliente (inclusive), "Avaliação transportes" = "Recusa".
//  - Pedido sem tracking ou sem devolução: nada na planilha, só no resumo do cruzamento.

const TRACK_CONFIG = {
  ATIVO: true,
  // Drive partilhado (ou pasta) com as planilhas de tracking
  PASTA_ID: '0AP8TY_KkFeD0Uk9PVA',
  INCLUIR_SUBPASTAS: true,
  CABECALHO_PEDIDO: 'Número do Pedido',
  PREFIXO_PAYLOAD: 'Payload JSON',
  PRAZO_DEVOLUCAO_DIAS: 35,
  // Comparação sem acentos e sem diferenciar maiúsculas
  PADROES_DEVOLUCAO: [
    'Carga devolvida ao remetente',
    'DEVOLUCAO - ENTRADA FILIAL',
    'DEVOLUÇÃO AUTORIZADA',
    'Devolver - Destinatário ausente',
    'Devolver - Destinatário não localizado',
    'DEVOLVIDO',
    'EM DEVOLUCAO',
    'EM DEVOLUÇÃO'
  ],
  // Além da lista, qualquer ocorrência com "devolu" ou "devolv" conta como devolução
  RADICAIS_DEVOLUCAO: ['DEVOLU', 'DEVOLV'],
  PADROES_RECUSA: [
    'RECUSA POR DESACORDO COM PRECO DO FRETE',
    'CARGA RECUSADA PELO DESTINATARIO'
  ],
  COLUNA_AVALIACAO: 'Avaliação',
  VALOR_NAO_RECEBER: 'Não Receber',
  COLUNA_AVALIACAO_TRANSPORTES: 'Avaliação transportes',
  VALOR_RECUSA: 'Recusa',
  // Brasil sem horário de verão: dia civil de São Paulo = UTC-3
  FUSO_MS: -3 * 3600 * 1000
};

const TRACK_ABA_ARQUIVOS = '_cruz_trk_arquivos';
const TRACK_ABA_RESULTADOS = '_cruz_trk_resultados';
const TRACK_CAB_RESULTADOS = ['Pedido', 'Modificado', 'Primeira devolução', 'Primeira recusa', 'Promessa', 'Planilha'];

// ==========================================
// AVALIAÇÃO DE UM JSON (função pura)
// ==========================================

function normalizarTextoTracking_(t) {
  return String(t == null ? '' : t)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

/**
 * @param {Object|string} json payload da Intelipost
 * @return {{pedido, modificado, primeiraDevolucao, primeiraRecusa, promessa}} datas em ms (ou null)
 */
function avaliarTrackingIntelipost_(json) {
  const obj = typeof json === 'string' ? JSON.parse(json) : json;
  const c = (obj && obj.content) || obj || {};
  const ms = (num, iso) => {
    const n = Number(num);
    if (isFinite(n) && n > 0) return n;
    const d = iso ? Date.parse(iso) : NaN;
    return isFinite(d) ? d : null;
  };

  const devolucao = TRACK_CONFIG.PADROES_DEVOLUCAO.map(normalizarTextoTracking_);
  const radicais = TRACK_CONFIG.RADICAIS_DEVOLUCAO.map(normalizarTextoTracking_);
  const recusa = TRACK_CONFIG.PADROES_RECUSA.map(normalizarTextoTracking_);

  let primeiraDevolucao = null;
  let primeiraRecusa = null;
  (c.shipment_order_volume_array || []).forEach(vol => {
    (vol.shipment_order_volume_state_history_array || []).forEach(ev => {
      const data = ms(ev.event_date, ev.event_date_iso) || ms(ev.created, ev.created_iso);
      if (!data) return;
      const micro = ev.shipment_volume_micro_state || {};
      // Só o nome da ocorrência e a mensagem da transportadora: as descrições longas falam de
      // devolução como hipótese ("... será feita uma re-tentativa ou devolução") e enganariam.
      const textos = [ev.provider_message, micro.name, micro.default_name].map(normalizarTextoTracking_).filter(Boolean);
      const ehDevolucao = textos.some(t => devolucao.some(p => t.indexOf(p) >= 0) || radicais.some(r => t.indexOf(r) >= 0));
      const ehRecusa = textos.some(t => recusa.some(p => t.indexOf(p) >= 0));
      if (ehDevolucao && (primeiraDevolucao === null || data < primeiraDevolucao)) primeiraDevolucao = data;
      if (ehRecusa && (primeiraRecusa === null || data < primeiraRecusa)) primeiraRecusa = data;
    });
  });

  return {
    pedido: String(c.order_number || c.sales_order_number || (c.external_order_numbers && c.external_order_numbers.sales) || '').trim(),
    modificado: ms(c.modified, c.modified_iso) || 0,
    primeiraDevolucao,
    primeiraRecusa,
    promessa: ms(c.estimated_delivery_date, c.estimated_delivery_date_iso)
  };
}

/** Dia civil em São Paulo (número de dias desde 1970). */
function diaTracking_(ms) {
  return Math.floor((ms + TRACK_CONFIG.FUSO_MS) / 86400000);
}

/**
 * Decisão para a carga a partir da avaliação.
 * @return {{naoReceber: boolean, recusa: boolean, dias: number|null, semDevolucao: boolean}}
 */
function decidirTracking_(av, hojeMs) {
  const r = { naoReceber: false, recusa: false, dias: null, semDevolucao: !av.primeiraDevolucao };
  if (av.primeiraDevolucao) {
    r.dias = diaTracking_(hojeMs) - diaTracking_(av.primeiraDevolucao);
    r.naoReceber = r.dias > TRACK_CONFIG.PRAZO_DEVOLUCAO_DIAS;
  }
  if (av.primeiraRecusa && av.promessa && av.primeiraRecusa <= av.promessa) r.recusa = true;
  return r;
}

/** Junta as colunas "Payload JSON 1, 2, ..." pela ordem do número. */
function colunasPayloadTracking_(cabecalho) {
  const prefixo = normalizarTextoTracking_(TRACK_CONFIG.PREFIXO_PAYLOAD);
  return cabecalho.map((c, i) => {
    const n = normalizarTextoTracking_(c);
    if (n.indexOf(prefixo) !== 0) return null;
    const num = parseInt(n.slice(prefixo.length).trim(), 10);
    return { i, num: isFinite(num) ? num : 0 };
  }).filter(Boolean).sort((a, b) => a.num - b.num).map(x => x.i);
}

// ==========================================
// LEITURA DAS PLANILHAS DE TRACKING (por etapas)
// ==========================================

function limparPedidoTracking_(v) {
  return String(v == null ? '' : v).trim().replace(/\.0+$/, '');
}

/** Planilhas Google do drive de tracking (os .json e outros ficheiros são ignorados). */
function listarPlanilhasTracking_() {
  const lista = [];
  const pilha = [{ pasta: DriveApp.getFolderById(TRACK_CONFIG.PASTA_ID), nivel: 0 }];
  while (pilha.length) {
    const { pasta, nivel } = pilha.pop();
    const it = pasta.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (it.hasNext()) {
      const f = it.next();
      if (!f.isTrashed()) lista.push({ id: f.getId(), nome: f.getName() });
    }
    if (TRACK_CONFIG.INCLUIR_SUBPASTAS && nivel < 5) {
      const subs = pasta.getFolders();
      while (subs.hasNext()) pilha.push({ pasta: subs.next(), nivel: nivel + 1 });
    }
  }
  lista.sort((a, b) => a.nome.localeCompare(b.nome));
  return lista;
}

/**
 * Procura os pedidos numa planilha de tracking (todas as abas). Lê só a coluna "Número do
 * Pedido" e, nas linhas encontradas, as colunas de payload.
 * @return {{linhas: number, resultados: Array[], errosJson: number, obs: string}}
 */
function lerPlanilhaTracking_(arq, pedidos) {
  const planilha = SpreadsheetApp.openById(arq.id);
  const alvoPedido = normalizarTextoTracking_(TRACK_CONFIG.CABECALHO_PEDIDO);
  let linhas = 0, errosJson = 0, abasLidas = 0;
  const resultados = [];

  planilha.getSheets().forEach(aba => {
    const ultimaLinha = aba.getLastRow();
    const ultimaColuna = aba.getLastColumn();
    if (ultimaLinha < 2 || ultimaColuna < 1) return;
    const cab = aba.getRange(1, 1, 1, ultimaColuna).getDisplayValues()[0];
    const iPed = cab.findIndex(c => normalizarTextoTracking_(c) === alvoPedido);
    const payload = colunasPayloadTracking_(cab);
    if (iPed < 0 || !payload.length) return;
    abasLidas++;

    const colPed = aba.getRange(2, iPed + 1, ultimaLinha - 1, 1).getDisplayValues();
    linhas += colPed.length;
    const minC = Math.min.apply(null, payload);
    const maxC = Math.max.apply(null, payload);

    colPed.forEach((l, k) => {
      const pedido = limparPedidoTracking_(l[0]);
      if (!pedido || !pedidos.has(pedido)) return;
      const valores = aba.getRange(k + 2, minC + 1, 1, maxC - minC + 1).getValues()[0];
      const texto = payload.map(i => { const v = valores[i - minC]; return v == null ? '' : String(v); }).join('');
      if (!texto.trim()) return;
      try {
        const av = avaliarTrackingIntelipost_(texto);
        resultados.push([pedido, av.modificado || 0, av.primeiraDevolucao || '', av.primeiraRecusa || '', av.promessa || '', arq.nome]);
      } catch (e) {
        errosJson++;
      }
    });
  });

  const obs = [];
  if (!abasLidas) obs.push('sem colunas "' + TRACK_CONFIG.CABECALHO_PEDIDO + '" e "' + TRACK_CONFIG.PREFIXO_PAYLOAD + ' 1"');
  if (errosJson) obs.push(errosJson + ' JSON inválido(s)');
  return { linhas, resultados, errosJson, obs: obs.join('; ') };
}

/**
 * Uma etapa da leitura do tracking (dentro do cruzamento). Lê as planilhas pendentes até ao prazo.
 * @param {Set<string>} pedidos pedidos da carga
 * @return {{pausado: boolean, total: number, feitos: number, trabalhou: boolean, erro?: string}}
 */
function etapaTracking_(ss, pedidos, prazo) {
  const abaArq = obterAbaCache_(ss, TRACK_ABA_ARQUIVOS, ['ID', 'Planilha', 'Situação', 'Linhas', 'Encontrados', 'Observação']);
  const abaRes = obterAbaCache_(ss, TRACK_ABA_RESULTADOS, TRACK_CAB_RESULTADOS);

  // Lista das planilhas: feita uma vez por cruzamento (fica na aba de cache)
  if (abaArq.getLastRow() < 2) {
    let lista;
    try {
      lista = listarPlanilhasTracking_();
    } catch (e) {
      return { pausado: false, total: 0, feitos: 0, trabalhou: false,
        erro: 'Não foi possível aceder ao drive de tracking (' + TRACK_CONFIG.PASTA_ID + '): ' + e.message };
    }
    if (!lista.length) return { pausado: false, total: 0, feitos: 0, trabalhou: false, erro: 'O drive de tracking não tem planilhas.' };
    anexarLinhas_(abaArq, lista.map(a => [a.id, a.nome, 'pendente', '', '', '']));
  }

  const registos = abaArq.getRange(2, 1, abaArq.getLastRow() - 1, 6).getValues();
  let feitos = registos.filter(r => r[2] === 'feito').length;
  let trabalhou = false;

  for (let i = 0; i < registos.length; i++) {
    if (registos[i][2] === 'feito') continue;
    if (Date.now() > prazo) return { pausado: true, total: registos.length, feitos, trabalhou };
    if (typeof cruzamentoCancelado_ === 'function' && cruzamentoCancelado_()) return { pausado: false, cancelado: true, total: registos.length, feitos, trabalhou };

    const arq = { id: String(registos[i][0]), nome: String(registos[i][1]) };
    let r;
    try {
      r = lerPlanilhaTracking_(arq, pedidos);
    } catch (e) {
      r = { linhas: 0, resultados: [], obs: 'ERRO: ' + e.message };
    }
    anexarLinhas_(abaRes, r.resultados);
    abaArq.getRange(i + 2, 3, 1, 4).setValues([['feito', r.linhas, r.resultados.length, r.obs]]);
    feitos++;
    trabalhou = true;
    if (typeof atualizarJobCruzamento_ === 'function') {
      atualizarJobCruzamento_({ trackingTotal: registos.length, trackingFeitos: feitos,
        mensagem: `A ler o tracking Intelipost (${feitos} de ${registos.length} planilhas).` });
    }
  }
  return { pausado: false, total: registos.length, feitos, trabalhou };
}

/** Melhor avaliação de cada pedido (o JSON mais recente, se aparecer em mais de uma linha). */
function lerResultadosTracking_(ss) {
  const aba = ss.getSheetByName(TRACK_ABA_RESULTADOS);
  const mapa = {};
  if (!aba || aba.getLastRow() < 2) return mapa;
  const num = v => { const n = Number(v); return isFinite(n) && n > 0 ? n : null; };
  aba.getRange(2, 1, aba.getLastRow() - 1, TRACK_CAB_RESULTADOS.length).getValues().forEach(l => {
    const pedido = limparPedidoTracking_(l[0]);
    const av = { modificado: num(l[1]) || 0, primeiraDevolucao: num(l[2]), primeiraRecusa: num(l[3]), promessa: num(l[4]), planilha: String(l[5]) };
    if (!mapa[pedido] || av.modificado > mapa[pedido].modificado) mapa[pedido] = av;
  });
  return mapa;
}

/**
 * Acrescenta as colunas "Avaliação" e "Avaliação transportes" à matriz da carga e aplica as regras.
 * A data de referência do prazo é o momento da criação da base (agora).
 * @return {Object} resumo para o Web App e o e-mail
 */
function aplicarTrackingNaMatriz_(ss, matriz, hojeMs) {
  const C = TRACK_CONFIG;
  const cab = matriz[0];
  const iPed = cab.indexOf(CRUZ_CONFIG.COLUNA_PEDIDO);
  const iNF = cab.indexOf(CRUZ_CONFIG.CABECALHO_NF_CARGA);
  let iAval = cab.indexOf(C.COLUNA_AVALIACAO);
  if (iAval < 0) { cab.push(C.COLUNA_AVALIACAO); iAval = cab.length - 1; }
  cab.push(C.COLUNA_AVALIACAO_TRANSPORTES);
  const iTransp = cab.length - 1;

  const mapa = lerResultadosTracking_(ss);
  // Map: mantém a ordem da carga (num objeto, números de pedido seriam reordenados)
  const porPedido = new Map();
  for (let k = 1; k < matriz.length; k++) {
    const l = matriz[k];
    while (l.length < cab.length) l.push('');
    const pedido = limparPedidoTracking_(l[iPed]);
    if (!pedido) continue;
    if (!porPedido.has(pedido)) porPedido.set(pedido, { nfs: [], linhas: [] });
    const g = porPedido.get(pedido);
    const nf = iNF >= 0 ? String(l[iNF]) : '';
    if (nf && g.nfs.indexOf(nf) < 0) g.nfs.push(nf);
    g.linhas.push(l);
  }

  const resumo = { pedidos: 0, encontrados: 0, naoReceber: [], recusas: [], semTracking: [], semDevolucao: [] };
  const ref = (pedido, extra) => Object.assign({ pedido, nfs: porPedido.get(pedido).nfs.join(', ') }, extra || {});
  porPedido.forEach((grupo, pedido) => {
    resumo.pedidos++;
    const av = mapa[pedido];
    if (!av) { resumo.semTracking.push(ref(pedido)); return; }
    resumo.encontrados++;
    const d = decidirTracking_(av, hojeMs);
    if (d.semDevolucao) resumo.semDevolucao.push(ref(pedido));
    grupo.linhas.forEach(l => {
      if (d.naoReceber) l[iAval] = C.VALOR_NAO_RECEBER;
      if (d.recusa) l[iTransp] = C.VALOR_RECUSA;
    });
    if (d.naoReceber) resumo.naoReceber.push(ref(pedido, { dias: d.dias }));
    if (d.recusa) resumo.recusas.push(ref(pedido));
  });
  return resumo;
}

/** Abas de cache do tracking (apagadas com as do cruzamento). */
function nomesCacheTracking_() {
  return [TRACK_ABA_ARQUIVOS, TRACK_ABA_RESULTADOS];
}

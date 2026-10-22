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
//    promessa de entrega ao cliente (inclusive), "Avaliação transportes" = "Recusa de Entrega".
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
  // Motor de prazo das etiquetas fora da malha de outra carga: o resultado fica sempre gravado
  VALOR_RECEBER: 'Receber',
  VALOR_RECEBER_SEM_TRACKING: 'Receber (sem tracking)',
  COLUNA_AVALIACAO_TRANSPORTES: 'Avaliação transportes',
  VALOR_RECUSA: 'Recusa de Entrega',
  // Brasil sem horário de verão: dia civil de São Paulo = UTC-3
  FUSO_MS: -3 * 3600 * 1000
};

const TRACK_ABA_ARQUIVOS = '_cruz_trk_arquivos';
const TRACK_ABA_RESULTADOS = '_cruz_trk_resultados';
const TRACK_CAB_RESULTADOS = ['Pedido', 'Modificado', 'Primeira devolução', 'Primeira recusa', 'Promessa', 'Planilha', 'NFs'];

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
  const nfs = [];
  (c.shipment_order_volume_array || []).forEach(vol => {
    const nota = vol.shipment_order_volume_invoice || {};
    const numero = String(nota.invoice_number == null ? '' : nota.invoice_number).trim();
    if (numero && nfs.indexOf(numero) < 0) nfs.push(numero);
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
    promessa: ms(c.estimated_delivery_date, c.estimated_delivery_date_iso),
    nfs   // números das notas fiscais dos volumes (invoice_number)
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
        resultados.push([pedido, av.modificado || 0, av.primeiraDevolucao || '', av.primeiraRecusa || '', av.promessa || '', arq.nome, (av.nfs || []).join(',')]);
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
    const av = { modificado: num(l[1]) || 0, primeiraDevolucao: num(l[2]), primeiraRecusa: num(l[3]), promessa: num(l[4]), planilha: String(l[5]),
      nfs: String(l[6] == null ? '' : l[6]).split(',').filter(Boolean) };
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

/** Nomes antigos do valor de recusa na "Avaliação transportes" (comparação sem acentos/maiúsculas). */
const NOMES_ANTIGOS_RECUSA = ['RECUSA', 'RECUSA NA ENTREGA', 'RECUSA DA ENTREGA', 'RECUSA DE ENTREGA'];

/**
 * Padroniza, em todas as cargas, a "Avaliação transportes" gravada com um nome antigo
 * ("Recusa", "Recusa na Entrega"...) para TRACK_CONFIG.VALOR_RECUSA ("Recusa de Entrega").
 * Menu "Cruzamento NF" da planilha (ou executar no editor). Pode ser repetida sem efeito.
 * @return {{abas:number, celulas:number}}
 */
function padronizarRecusaDeEntrega() {
  const C = TRACK_CONFIG;
  const alvo = normalizarTextoTracking_(C.COLUNA_AVALIACAO_TRANSPORTES);
  const resumo = { abas: 0, celulas: 0 };
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(aba => {
    if (!/^\d{2}-\d{2}-\d{4}/.test(aba.getName())) return;
    const ultimaLinha = aba.getLastRow(), ultimaColuna = aba.getLastColumn();
    if (ultimaLinha < 2 || ultimaColuna < 1) return;
    const cab = aba.getRange(1, 1, 1, ultimaColuna).getDisplayValues()[0];
    const col = cab.findIndex(c => normalizarTextoTracking_(c).trim() === alvo);
    if (col < 0) return;
    const faixa = aba.getRange(2, col + 1, ultimaLinha - 1, 1);
    const valores = faixa.getValues();
    let mudou = 0;
    valores.forEach(l => {
      const atual = String(l[0] == null ? '' : l[0]);
      if (atual !== C.VALOR_RECUSA && NOMES_ANTIGOS_RECUSA.indexOf(normalizarTextoTracking_(atual).trim()) >= 0) {
        l[0] = C.VALOR_RECUSA;
        mudou++;
      }
    });
    if (mudou) { faixa.setValues(valores); resumo.abas++; resumo.celulas += mudou; }
  });
  try {
    SpreadsheetApp.getUi().alert(`"${C.VALOR_RECUSA}": ${resumo.celulas} célula(s) corrigida(s) em ${resumo.abas} carga(s).`);
  } catch (e) { /* executado fora da planilha */ }
  return resumo;
}

/** Abas de cache do tracking (apagadas com as do cruzamento). */
function nomesCacheTracking_() {
  return [TRACK_ABA_ARQUIVOS, TRACK_ABA_RESULTADOS];
}

// ==========================================
// PRAZO DAS ETIQUETAS FORA DA MALHA QUE PERTENCEM A OUTRA CARGA
// ==========================================
// Uma mercadoria prometida numa carga passada (com o prazo OK) pode não ter sido entregue e
// aparecer numa carga futura, quando o prazo já não está OK. Quando o bipe é "fora da malha" e a
// etiqueta existe noutra carga, o navegador chama verificarPrazoForaDaMalha em segundo plano:
//  1. pega o pedido/NF da etiqueta na outra carga (a mais recente primeiro);
//  2. procura o tracking desse pedido nas planilhas do drive de tracking (uma passagem para todas
//     as etiquetas pedidas; resultado guardado em cache 20 min por pedido);
//  3. aplica a mesma regra do cruzamento (primeira devolução há mais de PRAZO_DEVOLUCAO_DIAS dias),
//     com a data de CRIAÇÃO DA CARGA ATUAL como referência.
// O resultado fica SEMPRE gravado na "Avaliação" da linha da divergência (com o motivo numa nota
// da célula), para não voltar a consultar o tracking cada vez que a carga é aberta:
//   fora do prazo           -> "Não Receber" (e alerta vermelho no ecrã)
//   dentro do prazo / pedido sem devolução -> "Receber"
//   pedido sem tracking     -> "Receber (sem tracking)"
// Só não grava quando o drive de tracking não pôde ser lido (volta a tentar ao abrir a carga).

const PRAZO_FORA_MALHA_CACHE_S = 20 * 60;

/**
 * @param {string} nomeAba carga atual
 * @param {string|string[]} etiquetas etiquetas fora da malha (bipadas nesta carga)
 * @return {{carga, referencia, origemReferencia, erro, resultados: {etiqueta, situacao, pedido, nf,
 *          cargaOrigem, dias, primeiraDevolucao, planilha}[]}}
 *   situacao: NAO_RECEBER | PRAZO_OK | SEM_DEVOLUCAO | SEM_TRACKING | SEM_OUTRA_CARGA
 */
function verificarPrazoForaDaMalha(nomeAba, etiquetas) {
  const C = TRACK_CONFIG;
  const lista = [].concat(etiquetas || []).map(e => String(e).trim()).filter((e, i, a) => e && a.indexOf(e) === i);
  const saida = { carga: nomeAba, referencia: '', origemReferencia: '', resultados: [], erro: '' };
  if (!C.ATIVO || !lista.length) return saida;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) { saida.erro = 'A carga ' + nomeAba + ' não existe.'; return saida; }

  const ref = dataCriacaoCarga_(aba);
  saida.referencia = dataCurtaTracking_(ref.ms);
  saida.origemReferencia = ref.origem;

  // 1. Pedido/NF da etiqueta nas outras cargas (mais recente primeiro)
  const origem = {};
  const pedidos = new Set();
  lista.forEach(et => {
    origem[et] = _localizarEtiquetaEmOutrasCargas_(ss, et, nomeAba)
      .filter(o => o.pedido && o.pedido !== 'N/A')
      .sort((a, b) => _parseDataDoNome(b.carga) - _parseDataDoNome(a.carga));
    origem[et].forEach(o => pedidos.add(limparPedidoTracking_(o.pedido)));
  });

  // 2. Tracking dos pedidos
  let tracking = {};
  try { tracking = buscarTrackingPedidos_(pedidos); }
  catch (e) { saida.erro = 'Sem acesso ao drive de tracking: ' + e.message; }

  // 3. Decisão por etiqueta
  const gravar = {};
  lista.forEach(et => {
    const occ = origem[et];
    const r = { etiqueta: et, situacao: 'SEM_OUTRA_CARGA', pedido: '', nf: '', cargaOrigem: '', dias: null, primeiraDevolucao: '', planilha: '' };
    if (occ.length) {
      Object.assign(r, { situacao: 'SEM_TRACKING', pedido: occ[0].pedido, nf: occ[0].nf, cargaOrigem: occ[0].carga });
      for (let k = 0; k < occ.length; k++) {
        const av = tracking[limparPedidoTracking_(occ[k].pedido)];
        if (!av) continue;
        const d = decidirTracking_(av, ref.ms);
        const atual = { pedido: occ[k].pedido, nf: occ[k].nf, cargaOrigem: occ[k].carga, dias: d.dias,
          primeiraDevolucao: dataCurtaTracking_(av.primeiraDevolucao), planilha: av.planilha || '' };
        if (d.naoReceber) { Object.assign(r, atual, { situacao: 'NAO_RECEBER' }); break; }
        if (r.situacao === 'SEM_TRACKING') Object.assign(r, atual, { situacao: d.semDevolucao ? 'SEM_DEVOLUCAO' : 'PRAZO_OK' });
      }
    }
    r.avaliacao = avaliacaoPrazo_(r, !!saida.erro);
    if (r.avaliacao) {
      r.nota = notaPrazo_(r, saida.referencia, saida.origemReferencia);
      gravar[et] = { valor: r.avaliacao, nota: r.nota };
    }
    saida.resultados.push(r);
  });

  if (Object.keys(gravar).length) marcarAvaliacaoDivergencias_(aba, gravar);
  return saida;
}

/** Valor da "Avaliação" para o resultado ('' = não gravar: sem outra carga ou tracking ilegível). */
function avaliacaoPrazo_(r, erroTracking) {
  const C = TRACK_CONFIG;
  switch (r.situacao) {
    case 'NAO_RECEBER': return C.VALOR_NAO_RECEBER;
    case 'PRAZO_OK':
    case 'SEM_DEVOLUCAO': return C.VALOR_RECEBER;
    case 'SEM_TRACKING': return erroTracking ? '' : C.VALOR_RECEBER_SEM_TRACKING;
    default: return '';
  }
}

/** Motivo da avaliação (nota da célula e texto no ecrã). */
function notaPrazo_(r, referencia, origemReferencia) {
  const C = TRACK_CONFIG;
  const origem = r.informadoPor
    ? `NF ${r.nf || '-'} (pedido ${r.pedido || '-'}) informada por ${r.informadoPor} em ${r.informadoEm}`
    : `NF ${r.nf || '-'} (pedido ${r.pedido || '-'}) da carga ${r.cargaOrigem || '-'}`;
  const ref = `${referencia} (${origemReferencia})`;
  switch (r.situacao) {
    case 'NAO_RECEBER':
      return `Fora do prazo: ${origem}; 1.ª devolução em ${r.primeiraDevolucao}, ${r.dias} dias até ${ref}. Limite: ${C.PRAZO_DEVOLUCAO_DIAS} dias.`;
    case 'PRAZO_OK':
      return `Dentro do prazo: ${origem}; 1.ª devolução em ${r.primeiraDevolucao}, ${r.dias} dias até ${ref}. Limite: ${C.PRAZO_DEVOLUCAO_DIAS} dias.`;
    case 'SEM_DEVOLUCAO':
      return `Dentro do prazo: ${origem}; o tracking não tem ocorrência de devolução.`;
    case 'SEM_TRACKING':
      return `${origem}: pedido sem tracking nas planilhas Intelipost; não há prazo a contar.`;
    default:
      return '';
  }
}

/**
 * Grava a "Avaliação" (e o motivo como nota) nas linhas de divergência destas etiquetas.
 * @param {Object<string, {valor, nota}>} porEtiqueta
 */
function marcarAvaliacaoDivergencias_(aba, porEtiqueta) {
  const C = TRACK_CONFIG;
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return 0;
  const dados = aba.getRange(1, 1, ultimaLinha, aba.getLastColumn()).getDisplayValues();
  const cab = dados[0].map(c => String(c).trim().toLowerCase());
  const col = function () {
    for (let i = 0; i < arguments.length; i++) if (cab.indexOf(arguments[i]) >= 0) return cab.indexOf(arguments[i]);
    return -1;
  };
  const iMerc = col('mercadoria código', 'mercadoria codigo');
  const iStatus = col('status conferência', 'status conferencia');
  if (iMerc < 0 || iStatus < 0) return 0;
  let iAval = col('avaliação', 'avaliacao');
  if (iAval < 0) {
    iAval = dados[0].length;
    aba.getRange(1, iAval + 1).setValue(C.COLUNA_AVALIACAO).setFontWeight('bold');
  }
  let n = 0;
  for (let i = 1; i < dados.length; i++) {
    const status = String(dados[i][iStatus]);
    const alvo = porEtiqueta[String(dados[i][iMerc]).trim()];
    if (!alvo) continue;
    if (status.indexOf('Divergência') < 0 && status.indexOf('Não Encontrado') < 0) continue;
    const celula = aba.getRange(i + 1, iAval + 1);
    celula.setValue(alvo.valor).setNote(alvo.nota || '');
    if (alvo.valor === C.VALOR_NAO_RECEBER) celula.setFontWeight('bold');
    n++;
  }
  return n;
}

/**
 * Tracking mais recente de cada pedido, lido das planilhas do drive de tracking.
 * @param {Set<string>} pedidos
 * @return {Object<string, {modificado, primeiraDevolucao, primeiraRecusa, promessa, planilha}>}
 */
function buscarTrackingPedidos_(pedidos, opcoes) {
  const mapa = {};
  const faltam = new Set();
  const cache = cacheTracking_();
  const semCache = !!(opcoes && opcoes.semCache);   // consulta pedida pelo operador: sempre o tracking atual
  pedidos.forEach(p => {
    const guardado = cache && !semCache ? cache.get('trkped:' + p) : null;
    if (guardado === null || guardado === undefined) { faltam.add(p); return; }
    const v = JSON.parse(guardado);
    if (v) mapa[p] = v;
  });
  if (!faltam.size) return mapa;

  const num = v => { const n = Number(v); return isFinite(n) && n > 0 ? n : null; };
  listarPlanilhasTracking_().forEach(arq => {
    let lida;
    try { lida = lerPlanilhaTracking_(arq, faltam); }
    catch (e) { console.warn('Tracking: planilha ' + arq.nome + ' não lida: ' + e.message); return; }
    lida.resultados.forEach(l => {
      const av = { modificado: num(l[1]) || 0, primeiraDevolucao: num(l[2]), primeiraRecusa: num(l[3]), promessa: num(l[4]), planilha: String(l[5]),
        nfs: String(l[6] == null ? '' : l[6]).split(',').filter(Boolean) };
      if (!mapa[l[0]] || av.modificado > mapa[l[0]].modificado) mapa[l[0]] = av;
    });
  });
  if (cache) faltam.forEach(p => {
    try { cache.put('trkped:' + p, JSON.stringify(mapa[p] || null), PRAZO_FORA_MALHA_CACHE_S); } catch (e) { /* cache cheia */ }
  });
  return mapa;
}

function cacheTracking_() {
  try { return typeof CacheService !== 'undefined' ? CacheService.getScriptCache() : null; } catch (e) { return null; }
}

/** dd/mm/aaaa no fuso de São Paulo. */
function dataCurtaTracking_(ms) {
  if (!ms) return '';
  const d = new Date(ms + TRACK_CONFIG.FUSO_MS);
  const p = n => String(n).padStart(2, '0');
  return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
}

// ==========================================
// TELA "PRODUTOS FORA DA MALHA"
// ==========================================
// Lista as etiquetas bipadas fora da malha de uma carga. O operador informa o Pedido e a NF de cada
// uma e confirma: o Pedido/NF ficam gravados na linha da divergência, o tracking Intelipost do pedido
// é consultado (a NF é conferida com as notas do JSON) e o prazo segue a mesma regra das etiquetas
// de outra carga (PRAZO_DEVOLUCAO_DIAS a contar da criação da carga). O resultado fica na "Avaliação"
// (com o motivo na nota) e aparece no painel "Divergências (fora da carga)" da conferência.

/** Colunas usadas pela tela (índices base 0; -1 se não existir). */
function colunasForaDaMalha_(cabecalho) {
  const cab = cabecalho.map(c => String(c).trim().toLowerCase());
  const col = function () {
    for (let i = 0; i < arguments.length; i++) if (cab.indexOf(arguments[i]) >= 0) return cab.indexOf(arguments[i]);
    return -1;
  };
  return {
    merc: col('mercadoria código', 'mercadoria codigo'),
    status: col('status conferência', 'status conferencia'),
    pedido: col('número do pedido faturado', 'numero do pedido faturado'),
    nf: col('nota fiscal', 'nf'),
    hora: col('data/hora conferência', 'data/hora conferencia'),
    outras: col('outras cargas'),
    aval: col('avaliação', 'avaliacao')
  };
}

const ehLinhaDivergencia_ = status => {
  status = String(status || '');
  return status.indexOf('Divergência') >= 0 || status.indexOf('Não Encontrado') >= 0;
};
const valorInformado_ = v => { v = String(v == null ? '' : v).trim(); return v === 'N/A' ? '' : v; };

/**
 * Etiquetas fora da malha da carga, uma por etiqueta.
 * @return {{carga, finalizada, referencia, origemReferencia, itens: {etiqueta, hora, pedido, nf,
 *   pedidoSugerido, nfSugerida, outrasCargas, avaliacao, nota}[]}}
 */
function listarForaDaMalhaCarga(nomeAba) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) throw new Error('A carga "' + nomeAba + '" não existe.');
  const ref = dataCriacaoCarga_(aba);
  const saida = { carga: nomeAba, finalizada: _cargaFinalizada_(aba), referencia: dataCurtaTracking_(ref.ms),
    origemReferencia: ref.origem, itens: [] };
  const ultimaLinha = aba.getLastRow(), ultimaColuna = aba.getLastColumn();
  if (ultimaLinha < 2 || ultimaColuna < 1) return saida;

  const dados = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getDisplayValues();
  const c = colunasForaDaMalha_(dados[0]);
  if (c.merc < 0 || c.status < 0) return saida;
  const notas = c.aval >= 0 ? aba.getRange(2, c.aval + 1, ultimaLinha - 1, 1).getNotes() : [];
  const vistas = {};
  for (let i = 1; i < dados.length; i++) {
    const l = dados[i];
    if (!ehLinhaDivergencia_(l[c.status])) continue;
    const etiqueta = String(l[c.merc]).trim();
    if (!etiqueta || vistas[etiqueta]) continue;
    vistas[etiqueta] = true;
    const item = {
      etiqueta,
      hora: c.hora >= 0 ? String(l[c.hora]).trim() : '',
      pedido: c.pedido >= 0 ? valorInformado_(l[c.pedido]) : '',
      nf: c.nf >= 0 ? valorInformado_(l[c.nf]) : '',
      outrasCargas: c.outras >= 0 ? String(l[c.outras]).split(/\s*[,;]\s*/).filter(Boolean) : [],
      avaliacao: c.aval >= 0 ? String(l[c.aval]).trim() : '',
      nota: notas[i - 1] ? String(notas[i - 1][0] || '') : '',
      pedidoSugerido: '', nfSugerida: ''
    };
    // Etiqueta que existe noutra carga: sugere o Pedido/NF de lá (o operador confirma ou corrige)
    if (!item.pedido && item.outrasCargas.length) {
      try {
        const o = _localizarEtiquetaEmOutrasCargas_(ss, etiqueta, nomeAba)
          .filter(x => x.pedido && x.pedido !== 'N/A')
          .sort((a, b) => _parseDataDoNome(b.carga) - _parseDataDoNome(a.carga))[0];
        if (o) { item.pedidoSugerido = o.pedido; item.nfSugerida = o.nf; }
      } catch (e) { /* sem sugestão */ }
    }
    saida.itens.push(item);
  }
  return saida;
}

/**
 * Confirmação do operador: grava Pedido/NF na divergência, consulta o tracking e avalia o prazo.
 * @return {{etiqueta, situacao, avaliacao, nota, pedido, nf, dias, primeiraDevolucao, planilha,
 *   nfConfere: boolean|null, nfsTracking: string[], referencia, erro}}
 */
function verificarPrazoInformado(nomeAba, etiqueta, pedido, nf) {
  const C = TRACK_CONFIG;
  etiqueta = String(etiqueta || '').trim();
  pedido = limparPedidoTracking_(pedido).replace(/\s+/g, '');
  nf = String(nf == null ? '' : nf).replace(/\s+/g, '');
  if (!etiqueta) throw new Error('Etiqueta em falta.');
  if (!pedido) throw new Error('Informe o número do pedido.');
  if (!nf) throw new Error('Informe o número da NF.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) throw new Error('A carga "' + nomeAba + '" não existe.');
  const ultimaLinha = aba.getLastRow();
  const dados = aba.getRange(1, 1, Math.max(1, ultimaLinha), Math.max(1, aba.getLastColumn())).getDisplayValues();
  const c = colunasForaDaMalha_(dados[0]);
  const linhas = [];
  for (let i = 1; i < dados.length; i++) {
    if (c.merc >= 0 && c.status >= 0 && String(dados[i][c.merc]).trim() === etiqueta && ehLinhaDivergencia_(dados[i][c.status])) linhas.push(i + 1);
  }
  if (!linhas.length) throw new Error('A etiqueta ' + etiqueta + ' não consta como fora da malha na carga ' + nomeAba + '.');

  // 1. Pedido e NF informados ficam na linha da divergência (como texto: sem zeros perdidos)
  let largura = dados[0].length;
  const coluna = (idx, titulo) => {
    if (idx >= 0) return idx;
    aba.getRange(1, largura + 1).setValue(titulo).setFontWeight('bold');
    return largura++;
  };
  const colPed = coluna(c.pedido, 'Número do Pedido Faturado');
  const colNF = coluna(c.nf, 'Nota Fiscal');
  linhas.forEach(n => {
    aba.getRange(n, colPed + 1).setNumberFormat('@').setValue(pedido);
    aba.getRange(n, colNF + 1).setNumberFormat('@').setValue(nf);
  });

  // 2. Tracking atual do pedido
  const ref = dataCriacaoCarga_(aba);
  let usuario = '';
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) { /* sem e-mail */ }
  const r = { etiqueta, situacao: 'SEM_TRACKING', pedido, nf, cargaOrigem: '', dias: null, primeiraDevolucao: '', planilha: '',
    nfConfere: null, nfsTracking: [], referencia: dataCurtaTracking_(ref.ms), erro: '',
    informadoPor: usuario || 'operador', informadoEm: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') };
  let tracking = {};
  try { tracking = buscarTrackingPedidos_(new Set([pedido]), { semCache: true }); }
  catch (e) { r.erro = 'Sem acesso ao drive de tracking: ' + e.message; }

  const av = tracking[pedido];
  if (av) {
    const d = decidirTracking_(av, ref.ms);
    r.dias = d.dias;
    r.primeiraDevolucao = dataCurtaTracking_(av.primeiraDevolucao);
    r.planilha = av.planilha || '';
    r.situacao = d.naoReceber ? 'NAO_RECEBER' : (d.semDevolucao ? 'SEM_DEVOLUCAO' : 'PRAZO_OK');
    // A NF informada tem de ser uma das notas do tracking deste pedido (sem zeros à esquerda)
    const semZeros = v => String(v).replace(/^0+/, '');
    r.nfsTracking = (av.nfs || []).slice();
    if (r.nfsTracking.length) r.nfConfere = r.nfsTracking.some(x => semZeros(x) === semZeros(nf));
  }

  // 3. Avaliação + motivo na linha da divergência
  r.avaliacao = avaliacaoPrazo_(r, !!r.erro);
  if (r.avaliacao) {
    r.nota = notaPrazo_(r, r.referencia, ref.origem);
    if (r.nfConfere === false) r.nota += ` ATENÇÃO: a NF ${nf} não consta no tracking deste pedido (NFs: ${r.nfsTracking.join(', ')}). Confira o pedido e a NF.`;
    const porEtiqueta = {};
    porEtiqueta[etiqueta] = { valor: r.avaliacao, nota: r.nota };
    marcarAvaliacaoDivergencias_(aba, porEtiqueta);
  } else {
    r.nota = '';
  }
  delete r.informadoEm;
  return r;
}

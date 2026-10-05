// ==========================================
// MOTOR DE JULGAMENTO (único para toda a aplicação)
// ==========================================
// julgarCargaMotor_ é uma função pura: não usa SpreadsheetApp nem variáveis globais.
// O mesmo código corre:
//  - no servidor (obterJulgamentoCarga, finalização, romaneio);
//  - no navegador (pop-up "Detalhes", resumo da finalização, Excel): o Index.html injeta o
//    código-fonte desta função com <?!= fonteMotorJulgamento(); ?>, por isso as regras nunca divergem.
//
// Regras por etiqueta (primeira que se aplica):
//  1. Divergência (bipada, mas não pertence à carga) -> Fora da malha / pertence a outra carga
//  2. Pedido com bloqueio ("Não Receber" em qualquer linha do pedido, ou status Bloqueado):
//       bipada -> Bloqueado: recusar (prazo excedido) | não bipada -> Bloqueado: não recebido (ignorado)
//  3. Status "Recusado Manualmente" (modo recusa) -> Recusado
//  4. Bipada (Conferido) -> OK
//  5. Não bipada, mas na Lista de Chegada -> OK (pela Lista)
//  6. Senão -> Falta
// Regras por pedido (o pedido tem uma única NF):
//  bloqueado -> Bloqueado (recusar se algum volume foi bipado; senão não recebido)
//  todas OK -> OK | todas em falta -> Não encontrado na carga (falta total)
//  todas recusadas -> Recusado | mistura -> Aceito parcialmente
// Alertas: bloqueado na Lista de Chegada, recusado na Lista de Chegada, reincidente de bloqueio
// noutra carga e etiqueta fora da malha encontrada noutra carga.

/**
 * @param {Object} entrada
 *   carga: nome da aba
 *   cabecalho: linha de cabeçalho da aba
 *   linhas: linhas de dados (sem o cabeçalho)
 *   numerosLinha: (opcional) número da linha na planilha de cada linha; por omissão 2, 3, ...
 *   recebidosLista: (opcional) chaves "pedido|etiqueta" na Lista de Chegada; sem ele usa a
 *                   coluna "Status Lista Chegada" = "Recebido Lista"
 *   historico: (opcional) {etiqueta: {ultimoStatus, ultimaCarga, eventos:[{carga, status}]}}
 *   divergenciasExtras: (opcional) [{etiqueta, outrasCargas:[...]}] ainda não gravadas na aba
 */
function julgarCargaMotor_(entrada) {
  const SIT = {
    OK_BIPADO: 'OK_BIPADO', OK_LISTA: 'OK_LISTA', RECUSADO: 'RECUSADO',
    BLOQ_RECUSAR: 'BLOQ_RECUSAR', BLOQ_IGNORADO: 'BLOQ_IGNORADO', FALTA: 'FALTA',
    FORA_MALHA: 'FORA_MALHA', OUTRA_CARGA: 'OUTRA_CARGA'
  };
  const PED = {
    OK: 'OK', PARCIAL: 'PARCIAL', FALTA_TOTAL: 'FALTA_TOTAL', RECUSADO: 'RECUSADO',
    BLOQ_RECUSAR: 'BLOQ_RECUSAR', BLOQ_NAO_RECEBIDO: 'BLOQ_NAO_RECEBIDO'
  };
  const ROTULO_PEDIDO = {
    OK: 'OK', PARCIAL: 'Aceito parcialmente', FALTA_TOTAL: 'Não encontrado na carga',
    RECUSADO: 'Recusado', BLOQ_RECUSAR: 'Bloqueado: recusar', BLOQ_NAO_RECEBIDO: 'Bloqueado: não recebido'
  };
  const STATUS_BIPADO = ['Conferido', 'Recusado Manualmente', 'Bloqueado'];

  const txt = v => String(v === null || v === undefined ? '' : v).trim();
  const cab = (entrada.cabecalho || []).map(c => txt(c).toLowerCase());
  const coluna = function () {
    for (let i = 0; i < arguments.length; i++) {
      const k = cab.indexOf(arguments[i]);
      if (k >= 0) return k;
    }
    return -1;
  };
  const iNF = coluna('nota fiscal', 'nf');
  const iPed = coluna('número do pedido faturado', 'numero do pedido faturado');
  const iMerc = coluna('mercadoria código', 'mercadoria codigo');
  const iProd = coluna('produto nome', 'produto');
  const iAval = coluna('avaliação', 'avaliacao');
  const iStatus = coluna('status conferência', 'status conferencia');
  const iLista = coluna('status lista chegada');
  const iOutras = coluna('outras cargas');
  const val = (l, i) => (i >= 0 ? txt(l[i]) : '');
  const listaOutras = v => txt(v).split(/\s*[,;]\s*/).filter(Boolean);

  const usarColunaLista = !entrada.recebidosLista;
  const lista = new Set(entrada.recebidosLista || []);
  const historico = entrada.historico || {};
  const carga = txt(entrada.carga);

  const itens = [];
  const divergencias = [];
  const vistasDiv = new Set();

  (entrada.linhas || []).forEach((l, k) => {
    const linha = entrada.numerosLinha ? entrada.numerosLinha[k] : k + 2;
    const etiqueta = val(l, iMerc);
    const pedido = val(l, iPed);
    const status = val(l, iStatus);
    if (!etiqueta && !pedido) return;
    if (etiqueta && (etiqueta.toLowerCase() === 'mercadoria código' || etiqueta.toLowerCase() === 'mercadoria codigo')) return;

    if (status.indexOf('Divergência') >= 0 || status.indexOf('Não Encontrado') >= 0) {
      if (vistasDiv.has(etiqueta)) return;
      vistasDiv.add(etiqueta);
      divergencias.push({ etiqueta, linha, outrasCargas: listaOutras(val(l, iOutras)) });
      return;
    }

    const aval = val(l, iAval).toLowerCase();
    itens.push({
      linha, nf: val(l, iNF), pedido, etiqueta, produto: val(l, iProd), status,
      marcadoBloqueio: aval === 'não receber' || aval === 'nao receber' || status === 'Bloqueado',
      bipado: STATUS_BIPADO.indexOf(status) >= 0,
      naLista: usarColunaLista ? val(l, iLista) === 'Recebido Lista' : lista.has(pedido + '|' + etiqueta)
    });
  });

  (entrada.divergenciasExtras || []).forEach(d => {
    const etiqueta = txt(d.etiqueta);
    if (!etiqueta || vistasDiv.has(etiqueta)) return;
    vistasDiv.add(etiqueta);
    divergencias.push({ etiqueta, linha: null, outrasCargas: (d.outrasCargas || []).slice() });
  });

  // ---------- pedidos ----------
  const pedidos = [];
  const porChave = {};
  itens.forEach(it => {
    const chave = it.pedido || (it.nf ? 'NF ' + it.nf : 'linha ' + it.linha);
    if (!porChave[chave]) {
      porChave[chave] = { chave, pedido: it.pedido, nf: it.nf, itens: [] };
      pedidos.push(porChave[chave]);
    }
    const p = porChave[chave];
    if (!p.nf && it.nf) p.nf = it.nf;
    p.itens.push(it);
  });

  pedidos.forEach(p => {
    p.total = p.itens.length;
    p.multiplo = p.total > 1;
    p.bloqueado = p.itens.some(it => it.marcadoBloqueio);
    p.contagem = { ok: 0, recusados: 0, faltas: 0, bipados: 0 };

    p.itens.forEach(it => {
      if (p.bloqueado) it.situacao = it.bipado ? SIT.BLOQ_RECUSAR : SIT.BLOQ_IGNORADO;
      else if (it.status === 'Recusado Manualmente') it.situacao = SIT.RECUSADO;
      else if (it.status === 'Conferido') it.situacao = SIT.OK_BIPADO;
      else if (it.naLista) it.situacao = SIT.OK_LISTA;
      else it.situacao = SIT.FALTA;

      if (it.bipado) p.contagem.bipados++;
      if (it.situacao === SIT.OK_BIPADO || it.situacao === SIT.OK_LISTA) p.contagem.ok++;
      if (it.situacao === SIT.RECUSADO) p.contagem.recusados++;
      if (it.situacao === SIT.FALTA) p.contagem.faltas++;
    });

    if (p.bloqueado) p.situacao = p.contagem.bipados > 0 ? PED.BLOQ_RECUSAR : PED.BLOQ_NAO_RECEBIDO;
    else if (p.contagem.ok === p.total) p.situacao = PED.OK;
    else if (p.contagem.faltas === p.total) p.situacao = PED.FALTA_TOTAL;
    else if (p.contagem.recusados === p.total) p.situacao = PED.RECUSADO;
    else p.situacao = PED.PARCIAL;
    p.rotulo = ROTULO_PEDIDO[p.situacao];

    p.itens.forEach(it => {
      it.situacaoPedido = p.situacao;
      it.multiplo = p.multiplo;
      it.totalPedido = p.total;
    });
  });

  // ---------- vereditos ----------
  const veredito = it => {
    switch (it.situacao) {
      case SIT.OK_BIPADO: return 'OK';
      case SIT.OK_LISTA: return 'OK (na Lista de Chegada, sem bipe)';
      case SIT.RECUSADO: return 'Recusado';
      case SIT.BLOQ_RECUSAR: return 'Bloqueado: recusar (prazo excedido)';
      case SIT.BLOQ_IGNORADO: return 'Bloqueado: não recebido';
      default: return it.situacaoPedido === PED.FALTA_TOTAL ? 'Falta: nota não encontrada na carga' : 'Falta: pedido parcial';
    }
  };
  itens.forEach(it => { it.veredito = veredito(it); });
  divergencias.forEach(d => {
    d.situacao = d.outrasCargas.length ? SIT.OUTRA_CARGA : SIT.FORA_MALHA;
    d.veredito = d.outrasCargas.length ? 'Fora da malha: pertence à carga ' + d.outrasCargas.join(', ') : 'Fora da malha';
  });

  // ---------- alertas ----------
  const alertas = [];
  const ref = it => `Pedido ${it.pedido || '-'}${it.nf ? ' (NF ' + it.nf + ')' : ''}, etiqueta ${it.etiqueta}`;
  itens.forEach(it => {
    if ((it.situacao === SIT.BLOQ_RECUSAR || it.situacao === SIT.BLOQ_IGNORADO) && it.naLista) {
      alertas.push({ tipo: 'BLOQUEADO_NA_LISTA', nivel: 'critico', pedido: it.pedido, nf: it.nf, etiqueta: it.etiqueta,
        texto: ref(it) + ': está bloqueado e consta na Lista de Chegada. Falha de processo, notifique.' });
    }
    if (it.situacao === SIT.RECUSADO && it.naLista) {
      alertas.push({ tipo: 'RECUSADO_NA_LISTA', nivel: 'critico', pedido: it.pedido, nf: it.nf, etiqueta: it.etiqueta,
        texto: ref(it) + ': foi recusado manualmente e consta na Lista de Chegada.' });
    }
    const h = it.etiqueta ? historico[it.etiqueta] : null;
    if (h) {
      const eventos = (h.eventos && h.eventos.length) ? h.eventos : [{ carga: h.ultimaCarga, status: h.ultimoStatus }];
      const cargas = [];
      eventos.forEach(e => {
        if (e && e.status === 'Bloqueado' && e.carga && e.carga !== carga && cargas.indexOf(e.carga) < 0) cargas.push(e.carga);
      });
      if (cargas.length) {
        alertas.push({ tipo: 'REINCIDENTE_BLOQUEIO', nivel: 'atencao', pedido: it.pedido, nf: it.nf, etiqueta: it.etiqueta,
          texto: ref(it) + ': reincidente, já bloqueado na carga ' + cargas.join(', ') + '.' });
      }
    }
  });
  divergencias.forEach(d => {
    if (d.outrasCargas.length) {
      alertas.push({ tipo: 'OUTRA_CARGA', nivel: 'atencao', etiqueta: d.etiqueta,
        texto: 'Etiqueta ' + d.etiqueta + ' fora da malha: pertence à carga ' + d.outrasCargas.join(', ') + '.' });
    }
  });

  // ---------- contagens ----------
  const contar = (lista, campo) => {
    const c = {};
    lista.forEach(x => { c[x[campo]] = (c[x[campo]] || 0) + 1; });
    return c;
  };
  return {
    carga,
    itens,
    pedidos,
    divergencias,
    alertas,
    contagens: {
      itens: contar(itens, 'situacao'),
      pedidos: contar(pedidos, 'situacao'),
      divergencias: contar(divergencias, 'situacao'),
      alertas: contar(alertas, 'tipo'),
      totalItens: itens.length,
      totalPedidos: pedidos.length
    },
    SIT, PED, ROTULO_PEDIDO
  };
}

/** Código-fonte do motor, injetado no Index.html para o navegador usar as mesmas regras. */
function fonteMotorJulgamento() {
  return julgarCargaMotor_.toString();
}

// ==========================================
// LEITURA DA CARGA NO SERVIDOR
// ==========================================

/**
 * Julga a carga a partir da planilha (sincroniza antes com a Lista de Chegada).
 * Usado pela finalização e pelo romaneio; pode ser chamado pelo Web App.
 */
function obterJulgamentoCarga(nomeAba, opcoes) {
  opcoes = opcoes || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) throw new Error('A carga "' + nomeAba + '" não existe na planilha.');

  let recebidos = null;
  if (!opcoes.semSincronizar) {
    const setLista = _obterSetGlobalListaChegada();
    if (setLista.size > 0) recebidos = _sincronizarAbaComSetListaChegada(aba, setLista);
  }

  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();
  if (ultimaLinha < 1 || ultimaColuna < 1) return julgarCargaMotor_({ carga: nomeAba, cabecalho: [], linhas: [] });
  const valores = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getDisplayValues();

  let historico = {};
  try { historico = obterHistoricoEventosFrontend(); } catch (e) { /* sem histórico */ }

  // Sem recebidos explícitos, o motor lê a coluna "Status Lista Chegada" (já sincronizada)
  return julgarCargaMotor_({
    carga: nomeAba,
    cabecalho: valores[0],
    linhas: valores.slice(1),
    historico
  });
}

// ==========================================
// ETIQUETA FORA DA MALHA: PROCURA NAS OUTRAS CARGAS
// ==========================================

/**
 * Procura a etiqueta (coluna "Mercadoria Código") em todas as abas de carga, exceto a atual.
 * Usa o TextFinder do Sheets (uma pesquisa na planilha inteira), por isso não percorre aba a aba.
 * Linhas que também foram divergências noutra carga não contam.
 * @return {string[]} nomes das cargas, pela ordem das abas
 */
function _procurarEtiquetaEmOutrasCargas_(ss, etiqueta, nomeAbaAtual) {
  etiqueta = String(etiqueta || '').trim();
  if (!etiqueta) return [];
  const encontrados = ss.createTextFinder(etiqueta).matchEntireCell(true).findAll();
  const cabecalhos = {};
  const cargas = [];

  encontrados.forEach(celula => {
    const aba = celula.getSheet();
    const nome = aba.getName();
    if (nome === nomeAbaAtual || !/^\d{2}-\d{2}-\d{4}/.test(nome) || cargas.indexOf(nome) >= 0) return;

    if (!cabecalhos[nome]) {
      const cab = aba.getRange(1, 1, 1, aba.getLastColumn()).getDisplayValues()[0].map(c => String(c).trim().toLowerCase());
      cabecalhos[nome] = {
        merc: cab.indexOf('mercadoria código') >= 0 ? cab.indexOf('mercadoria código') : cab.indexOf('mercadoria codigo'),
        status: cab.indexOf('status conferência') >= 0 ? cab.indexOf('status conferência') : cab.indexOf('status conferencia')
      };
    }
    const c = cabecalhos[nome];
    if (celula.getColumn() - 1 !== c.merc || celula.getRow() === 1) return;
    if (c.status >= 0) {
      const status = String(aba.getRange(celula.getRow(), c.status + 1).getDisplayValue()).trim();
      if (status.indexOf('Divergência') >= 0 || status.indexOf('Não Encontrado') >= 0) return;
    }
    cargas.push(nome);
  });
  return cargas;
}

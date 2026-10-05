// ==========================================
// ROMANEIO DE DEVOLUÇÃO (PDF para a transportadora)
// ==========================================
// Fluxo:
//  1. Um HTML com o layout do romaneio é convertido pelo Google Drive num documento Google Docs
//     (o "modelo"). Isto acontece uma única vez; depois a equipa pode ajustar o modelo no Docs.
//  2. Para cada romaneio: copia o modelo, troca os marcadores {{...}}, escreve as linhas na
//     área central, exporta em PDF para a pasta "Romaneios" e apaga a cópia.
//  3. O logótipo é inserido em cada romaneio (marcador {{LOGO}}), por isso trocar a imagem
//     no Drive não obriga a recriar o modelo.
//
// Um romaneio por carga: montarDadosRomaneioCarga_ lê a aba da carga e decide o que entra em
// cada secção; gerarRomaneioPdf_(dados) só desenha.

const ROMANEIO_CONFIG = {
  // Drive partilhado com as pastas do romaneio (logótipo, PDFs e modelo)
  DRIVE_ID: '0AN3saKJPOCg6Uk9PVA',
  // Pasta do logótipo dentro do drive (usa a primeira imagem que encontrar)
  NOME_PASTA_LOGO: 'Logo_KBM',
  // Pasta dos PDFs dentro do drive (criada se não existir)
  NOME_PASTA: 'Romaneios',
  // Opcional: ID de um ficheiro ou pasta do logótipo, que substitui NOME_PASTA_LOGO
  LOGO_ID: '',
  // Opcional: ID de uma pasta para os PDFs, que substitui NOME_PASTA
  PASTA_ID: '',
  NOME_MODELO: 'Modelo_Romaneio_Devolucao (não apagar)',
  // Suba este número quando o layout do HTML mudar: o modelo é recriado na próxima geração
  VERSAO_MODELO: 3,
  // Transportadora fixa (o campo no Web App fica travado)
  TRANSPORTADORA: 'GFL',
  // Página: 'A4' ou 'CARTA'
  PAGINA: 'A4',
  MARGEM_PT: 28,
  ALTURA_LOGO_PT: 58,
  FONTE: 'Montserrat',
  COR_BARRA: '#2e75b6',
  COR_CINZA: '#a6a6a6',
  // Altura mínima da área central, para as assinaturas ficarem no fundo da página
  ALTURA_CONTEUDO_PT: 480,
  // Caixas de estado do volume: saem todas marcadas
  MARCAR_ESTADOS: true,
  ESTADOS: [
    ['perfeito', 'Volumes em perfeito estado'],
    ['amassado', 'Volume amassado'],
    ['rasgado', 'Volume rasgado']
  ],
  // Secções da área central, pela ordem do romaneio. mesmaLinha:false = NFs a partir da linha seguinte
  SECOES: [
    { chave: 'nao_encontradas', titulo: 'Notas não encontradas na carga' },
    { chave: 'parciais', titulo: 'Notas aceitas parcialmente', mesmaLinha: false },
    { chave: 'avaria', titulo: 'Notas recusadas por avaria' },
    { chave: 'embalagem_vazia', titulo: 'Notas recusadas com embalagem vazia' },
    { chave: 'prazo', titulo: 'Notas recusadas por prazo indenizatório' },
    { chave: 'nao_kabum', titulo: 'Notas recusadas não pertencentes ao KaBuM' },
    { chave: 'improcedentes', titulo: 'Notas recusadas por produtos improcedentes' },
    { chave: 'fora_romaneio', titulo: 'Nota recusada fora do romaneio' }
  ],
  PROP_MODELO: 'ROMANEIO_MODELO_ID',
  PROP_VERSAO: 'ROMANEIO_MODELO_VERSAO',
  PROP_PASTA: 'ROMANEIO_PASTA_ID'
};

const ROMANEIO_PAGINAS = {
  A4: { largura: 595.28, altura: 841.89 },
  CARTA: { largura: 612, altura: 792 }
};

// ==========================================
// API DO WEB APP
// ==========================================

/**
 * Gera o romaneio de uma carga e devolve o link do PDF.
 * @param {Object} opcoes {carga: nome da aba, dataISO: data do cabeçalho (yyyy-MM-dd)}
 */
function gerarRomaneioCarga(opcoes) {
  opcoes = opcoes || {};
  const inicio = Date.now();
  if (!opcoes.carga) throw new Error('Selecione a carga.');

  const dados = montarDadosRomaneioCarga_(opcoes.carga, opcoes.dataISO);
  let pdf;
  try {
    pdf = gerarRomaneioPdf_(dados);
  } catch (e) {
    throw new Error(explicarErroRomaneio_(e));
  }

  const resultado = {
    ok: true,
    pdf: { id: pdf.id, url: pdf.url, nome: pdf.nome },
    resumo: dados.resumo,
    avisos: pdf.avisos,
    segundos: Math.round((Date.now() - inicio) / 100) / 10
  };
  try {
    const modeloId = PropertiesService.getScriptProperties().getProperty(ROMANEIO_CONFIG.PROP_MODELO);
    if (modeloId) resultado.modeloUrl = 'https://docs.google.com/document/d/' + modeloId + '/edit';
  } catch (e) { /* sem modelo */ }
  return resultado;
}

/** Apaga a referência ao modelo atual; o próximo romaneio cria um novo a partir do HTML. */
function recriarModeloRomaneio() {
  const props = PropertiesService.getScriptProperties();
  const antigo = props.getProperty(ROMANEIO_CONFIG.PROP_MODELO);
  if (antigo) {
    try { DriveApp.getFileById(antigo).setName(ROMANEIO_CONFIG.NOME_MODELO + ' (substituído)'); } catch (e) { /* já não existe */ }
  }
  props.deleteProperty(ROMANEIO_CONFIG.PROP_MODELO);
  const modelo = obterModeloRomaneio_();
  const msg = 'Modelo do romaneio recriado: ' + modelo.getUrl();
  console.log(msg);
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Romaneio', 8); } catch (e) { /* editor */ }
  return msg;
}

// ==========================================
// MOTOR PRINCIPAL (modelo Google Docs)
// ==========================================

/**
 * @param {Object} dados
 *   transportadora: string
 *   data: Date | 'dd/MM/yyyy' | 'yyyy-MM-dd' (data do cabeçalho)
 *   listagem: {pedidos: number, data: ...} (opcional)
 *   secoes: [{titulo, linhas: [string | {nf, itens:[{qtd, produto}], obs}], mesmaLinha?}]
 *   volumes: number | string
 *   estados: {perfeito, amassado, rasgado} (opcional; sem ele vale MARCAR_ESTADOS)
 *   nomeArquivo: string (opcional)
 * @return {{id, url, nome, avisos: string[]}}
 */
function gerarRomaneioPdf_(dados) {
  const pasta = obterPastaRomaneios_();
  const modelo = obterModeloRomaneio_();
  const nome = nomeArquivoRomaneio_(dados);

  const copia = modelo.makeCopy(nome + ' (temp)', pasta);
  try {
    const doc = DocumentApp.openById(copia.getId());
    const corpo = doc.getBody();

    const avisos = [];
    const avisoLogo = inserirLogo_(corpo);
    if (avisoLogo) avisos.push(avisoLogo);

    const campos = camposRomaneio_(dados);
    Object.keys(campos).forEach(chave => substituirMarcador_(corpo, chave, campos[chave]));
    preencherConteudoDoc_(corpo, linhasConteudoRomaneio_(dados));

    doc.saveAndClose();
    const pdf = copia.getAs('application/pdf').setName(nome + '.pdf');
    const arquivo = pasta.createFile(pdf);
    return { id: arquivo.getId(), url: arquivo.getUrl(), nome: arquivo.getName(), avisos };
  } finally {
    try { copia.setTrashed(true); } catch (e) { console.warn('Não foi possível apagar a cópia temporária: ' + e.message); }
  }
}

/** Troca o parágrafo {{CONTEUDO}} (dentro da célula central) pelas linhas do romaneio. */
function preencherConteudoDoc_(corpo, linhas) {
  const achado = corpo.findText('\\{\\{CONTEUDO\\}\\}');
  if (!achado) throw new Error('O modelo do romaneio não tem o marcador {{CONTEUDO}}. Execute "recriarModeloRomaneio".');

  let base = achado.getElement();
  while (base && base.getType() !== DocumentApp.ElementType.PARAGRAPH) base = base.getParent();
  if (!base) throw new Error('Marcador {{CONTEUDO}} fora de um parágrafo.');

  if (!linhas.length) { base.asParagraph().setText(' '); return; }

  const recipiente = base.getParent();
  const indice = recipiente.getChildIndex(base);
  const atributos = atributosDeTexto_(base.asParagraph());
  linhas.forEach((texto, i) => {
    const p = recipiente.insertParagraph(indice + 1 + i, base.copy().asParagraph());
    p.setText(texto === '' ? ' ' : texto);
    // setText pode perder a formatação de caractere do marcador (negrito, fonte, tamanho)
    if (atributos) p.editAsText().setAttributes(0, p.getText().length - 1, atributos);
  });
  base.removeFromParent();
}

/** Formatação de caractere do primeiro caractere do parágrafo, sem os valores nulos. */
function atributosDeTexto_(paragrafo) {
  try {
    const texto = paragrafo.editAsText();
    if (!texto.getText().length) return null;
    const brutos = texto.getAttributes(0);
    const limpos = {};
    let algum = false;
    Object.keys(brutos).forEach(k => {
      if (brutos[k] !== null && brutos[k] !== undefined) { limpos[k] = brutos[k]; algum = true; }
    });
    return algum ? limpos : null;
  } catch (e) {
    return null;
  }
}

/** Troca {{CHAVE}} pelo valor mantendo a formatação do texto à volta (sem regras de regex no valor). */
function substituirMarcador_(corpo, chave, valor) {
  const padrao = '\\{\\{' + chave + '\\}\\}';
  valor = String(valor === undefined || valor === null ? '' : valor);
  for (let guarda = 0; guarda < 50; guarda++) {
    const achado = corpo.findText(padrao);
    if (!achado) return;
    const texto = achado.getElement().asText();
    const ini = achado.getStartOffset();
    const fim = achado.getEndOffsetInclusive();
    if (valor) {
      texto.insertText(ini, valor);
      texto.deleteText(ini + valor.length, fim + valor.length);
    } else {
      texto.deleteText(ini, fim);
    }
  }
}

// ==========================================
// CONTEÚDO
// ==========================================

function camposRomaneio_(dados) {
  const estados = dados.estados;
  const campos = {
    TRANSPORTADORA: String(dados.transportadora || ROMANEIO_CONFIG.TRANSPORTADORA).trim(),
    DATA: formatarDataRomaneio_(dados.data || new Date()),
    VOLUMES: dados.volumes === undefined || dados.volumes === null ? '' : String(dados.volumes)
  };
  ROMANEIO_CONFIG.ESTADOS.forEach(([chave]) => {
    const marcada = estados ? !!estados[chave] : ROMANEIO_CONFIG.MARCAR_ESTADOS;
    campos['CX_' + chave.toUpperCase()] = marcada ? '☑' : '☐';
  });
  return campos;
}

/**
 * Converte as secções em linhas de texto, no formato do romaneio atual:
 *   "Titulo: primeira NF" / "segunda NF" / ... e uma linha em branco entre secções.
 */
function linhasConteudoRomaneio_(dados) {
  const linhas = [];
  if (dados.listagem) {
    linhas.push(`Listagem ${dados.listagem.pedidos} pedidos - ${formatarDataRomaneio_(dados.listagem.data)} :`);
    linhas.push('');
  }
  (dados.secoes || []).forEach((secao, s) => {
    const itens = (secao.linhas || []).map(formatarItemRomaneio_).filter(t => t);
    const titulo = String(secao.titulo || '').trim().replace(/:?$/, ':');
    if (secao.mesmaLinha === false || !itens.length) {
      linhas.push(titulo);
      itens.forEach(t => linhas.push(t));
    } else {
      linhas.push(titulo + ' ' + itens[0]);
      itens.slice(1).forEach(t => linhas.push(t));
    }
    if (s < dados.secoes.length - 1) linhas.push('');
  });
  return linhas;
}

/** "NF 123 (1 x Produto; 2 x Outro) observação" a partir de um objeto, ou o texto tal como veio. */
function formatarItemRomaneio_(item) {
  if (item === null || item === undefined) return '';
  if (typeof item !== 'object') return String(item).trim();

  const produtos = (item.itens || []).map(i => {
    const qtd = i.qtd === undefined || i.qtd === null || i.qtd === '' ? 1 : i.qtd;
    return `${qtd} x ${String(i.produto || '').trim()}`;
  }).filter(t => t.trim());
  let texto = 'NF ' + String(item.nf || '').trim();
  if (produtos.length) texto += ' (' + produtos.join('; ') + ')';
  if (item.obs && String(item.obs).trim()) texto += ' ' + String(item.obs).trim();
  return texto;
}

function nomeArquivoRomaneio_(dados) {
  if (dados.nomeArquivo) return String(dados.nomeArquivo);
  const limpar = t => String(t).trim().replace(/[\\/:*?"<>|]+/g, '-');
  const transp = limpar(dados.transportadora || ROMANEIO_CONFIG.TRANSPORTADORA);
  const referencia = dados.carga ? 'carga ' + limpar(dados.carga)
    : formatarDataRomaneio_(dados.data || new Date()).replace(/\//g, '-');
  const hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmmss');
  return `Romaneio ${transp} ${referencia} ${hora}`.replace(/\s+/g, ' ');
}

function formatarDataRomaneio_(valor) {
  if (valor instanceof Date) return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const texto = String(valor || '').trim();
  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(texto);
  return iso ? `${iso[3]}/${iso[2]}/${iso[1]}` : texto;
}

// ==========================================
// HTML DO ROMANEIO (modelo e motor direto)
// ==========================================

/**
 * @param {Object} dados  (ignorado no modo "modelo")
 * @param {Object} opcoes {modo: 'modelo'|'direto', logo: dataUri}
 */
function montarHtmlRomaneio_(dados, opcoes) {
  const C = ROMANEIO_CONFIG;
  const modelo = opcoes.modo === 'modelo';
  const esc = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const campos = modelo ? null : camposRomaneio_(dados);
  const v = chave => modelo ? '{{' + chave + '}}' : esc(campos[chave]);

  const borda = '1px solid #000000';
  const semBorda = 'none';
  const celula = (b, extra) => {
    const [t, r, bo, l] = b;
    return `border-top:${t};border-right:${r};border-bottom:${bo};border-left:${l};` +
      `padding:3pt 4pt;vertical-align:middle;${extra || ''}`;
  };
  const tudo = [borda, borda, borda, borda];
  const p = (texto, estilo) =>
    `<p style="margin:0;padding:0;line-height:1.3;${estilo || ''}">${texto}</p>`;

  let logo;
  if (modelo) logo = p('{{LOGO}}', 'text-align:center;font-size:20pt;font-weight:bold;color:#0b5cad;');
  else if (opcoes.logo) logo = `<p style="margin:0;text-align:center;"><img src="${opcoes.logo}" style="height:${C.ALTURA_LOGO_PT}pt;"></p>`;
  else logo = p('KaBuM!', 'text-align:center;font-size:20pt;font-weight:bold;color:#0b5cad;');

  let conteudo;
  if (modelo) {
    conteudo = p('{{CONTEUDO}}', 'font-weight:bold;font-size:9pt;');
  } else {
    conteudo = linhasConteudoRomaneio_(dados)
      .map(l => p(l === '' ? '&nbsp;' : esc(l), 'font-weight:bold;font-size:9pt;'))
      .join('');
  }

  const estados = C.ESTADOS.map(([chave, rotulo]) =>
    p(`${v('CX_' + chave.toUpperCase())}&nbsp;&nbsp;&nbsp;&nbsp;${esc(rotulo)}`,
      'font-size:8pt;margin-left:90pt;')
  ).join('');

  const col = ['22%', '23%', '27%', '28%'];
  const fonte = `font-family:${C.FONTE},Arial,sans-serif;`;

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Romaneio de Devolução</title>
<style>
  @page { size: ${C.PAGINA === 'CARTA' ? 'letter' : 'A4'}; margin: ${C.MARGEM_PT}pt; }
  body { margin: 0; ${fonte} font-size: 9pt; color: #000000; }
  table { border-collapse: collapse; width: 100%; table-layout: fixed; }
</style></head>
<body>
<table style="border-collapse:collapse;width:100%;table-layout:fixed;${fonte}">
  <colgroup>${col.map(w => `<col style="width:${w}">`).join('')}</colgroup>
  <tr style="height:70pt;">
    <td style="${celula(tudo, 'text-align:center;')}">${p('Transportadora:', 'font-weight:bold;text-align:center;')}${p('Data:', 'font-weight:bold;text-align:center;')}</td>
    <td style="${celula(tudo, 'text-align:center;')}">${p(v('TRANSPORTADORA'), 'font-weight:bold;text-align:center;')}${p(v('DATA'), 'font-weight:bold;text-align:center;')}</td>
    <td colspan="2" style="${celula(tudo, 'text-align:center;')}">${logo}</td>
  </tr>
  <tr>
    <td colspan="4" style="${celula(tudo, `background-color:${C.COR_BARRA};padding:1pt 4pt;`)}">${p('ROMANEIO DE DEVOLUÇÃO', 'font-weight:bold;color:#ffffff;text-align:center;font-size:9pt;')}</td>
  </tr>
  <tr style="height:${C.ALTURA_CONTEUDO_PT}pt;">
    <td colspan="4" style="${celula(tudo, 'vertical-align:top;padding:4pt 4pt;')}">${conteudo}</td>
  </tr>
  <tr>
    <td colspan="4" style="${celula(tudo, `background-color:${C.COR_CINZA};padding:1pt 4pt;`)}">${p('Volumes: ' + v('VOLUMES'), 'font-weight:bold;text-align:center;font-size:9pt;')}</td>
  </tr>
  <tr>
    <td colspan="4" style="${celula([borda, borda, semBorda, borda], 'padding:10pt 4pt 14pt 4pt;')}">${estados}</td>
  </tr>
  <tr>
    <td style="${celula([borda, borda, borda, borda], 'text-align:center;padding:1pt 4pt;')}">${p('Carimbo/ass', 'font-weight:bold;text-align:center;')}</td>
    <td style="${celula([semBorda, borda, semBorda, borda])}">${p('&nbsp;')}</td>
    <td style="${celula([borda, borda, borda, borda], 'text-align:center;padding:1pt 4pt;')}">${p('Carimbo/ass', 'font-weight:bold;text-align:center;')}</td>
    <td style="${celula([semBorda, borda, semBorda, borda])}">${p('&nbsp;')}</td>
  </tr>
  <tr style="height:85pt;">
    <td style="${celula(tudo)}">${p('&nbsp;')}</td>
    <td style="${celula([semBorda, borda, semBorda, borda])}">${p('&nbsp;')}</td>
    <td style="${celula(tudo)}">${p('&nbsp;')}</td>
    <td style="${celula([semBorda, borda, semBorda, borda])}">${p('&nbsp;')}</td>
  </tr>
  <tr>
    <td style="${celula(tudo, `background-color:${C.COR_CINZA};padding:1pt 4pt;`)}">${p('MOTORISTA', 'font-weight:bold;text-align:center;')}</td>
    <td style="${celula([semBorda, borda, borda, borda])}">${p('&nbsp;')}</td>
    <td style="${celula(tudo, `background-color:${C.COR_CINZA};padding:1pt 4pt;`)}">${p('KABUM', 'font-weight:bold;text-align:center;')}</td>
    <td style="${celula([semBorda, borda, borda, borda])}">${p('&nbsp;')}</td>
  </tr>
</table>
</body></html>`;
}

// ==========================================
// MODELO NO GOOGLE DOCS
// ==========================================

function obterModeloRomaneio_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(ROMANEIO_CONFIG.PROP_MODELO);
  const versao = props.getProperty(ROMANEIO_CONFIG.PROP_VERSAO);

  if (id && versao === String(ROMANEIO_CONFIG.VERSAO_MODELO)) {
    try {
      const arquivo = DriveApp.getFileById(id);
      if (!arquivo.isTrashed()) return arquivo;
    } catch (e) { /* apagado ou sem acesso: recria */ }
  } else if (id) {
    try { DriveApp.getFileById(id).setName(ROMANEIO_CONFIG.NOME_MODELO + ' (versão antiga)'); } catch (e) { /* ignora */ }
  }

  const arquivo = criarModeloRomaneio_();
  props.setProperty(ROMANEIO_CONFIG.PROP_MODELO, arquivo.getId());
  props.setProperty(ROMANEIO_CONFIG.PROP_VERSAO, String(ROMANEIO_CONFIG.VERSAO_MODELO));
  return arquivo;
}

function criarModeloRomaneio_() {
  const pasta = obterPastaRomaneios_();
  const html = montarHtmlRomaneio_({}, { modo: 'modelo' });
  const id = converterHtmlEmDoc_(html, ROMANEIO_CONFIG.NOME_MODELO, pasta.getId());

  const doc = DocumentApp.openById(id);
  const corpo = doc.getBody();
  const pagina = ROMANEIO_PAGINAS[ROMANEIO_CONFIG.PAGINA] || ROMANEIO_PAGINAS.A4;
  const m = ROMANEIO_CONFIG.MARGEM_PT;
  corpo.setPageWidth(pagina.largura).setPageHeight(pagina.altura)
    .setMarginTop(m).setMarginBottom(m).setMarginLeft(m).setMarginRight(m);

  ajustarTabelaDoModelo_(corpo, pagina.largura - 2 * m);

  reduzirParagrafosVazios_(corpo);
  doc.saveAndClose();
  return DriveApp.getFileById(id);
}

/**
 * O Docs exige um parágrafo antes e depois de uma tabela. Os vazios que sobram da conversão
 * são apagados e os obrigatórios ficam com 1 pt, para não empurrar o romaneio nem criar página em branco.
 */
function reduzirParagrafosVazios_(corpo) {
  const vazio = f => f.getType() === DocumentApp.ElementType.PARAGRAPH && !f.asParagraph().getText().trim();
  const minimizar = par => {
    try {
      par.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
      par.setText(' ');
      par.editAsText().setFontSize(1);
    } catch (e) { /* fica como está */ }
  };
  const apagarOuMinimizar = (lista, manter) => lista.forEach(f => {
    if (f === manter) return minimizar(f.asParagraph());
    try { f.removeFromParent(); } catch (e) { minimizar(f.asParagraph()); }
  });

  const filhos = [];
  for (let i = 0; i < corpo.getNumChildren(); i++) filhos.push(corpo.getChild(i));
  const primeiraTabela = filhos.findIndex(f => f.getType() === DocumentApp.ElementType.TABLE);
  if (primeiraTabela < 0) return;

  const antes = filhos.slice(0, primeiraTabela);
  if (antes.every(vazio)) apagarOuMinimizar(antes, antes[antes.length - 1]);

  let k = filhos.length;
  while (k > 0 && vazio(filhos[k - 1])) k--;
  const depois = filhos.slice(k);
  if (depois.length) apagarOuMinimizar(depois, depois[depois.length - 1]);
}

// Larguras das colunas e alturas mínimas das linhas (a conversão do HTML nem sempre as respeita)
const ROMANEIO_COLUNAS = [0.22, 0.23, 0.27, 0.28];
const ROMANEIO_ALTURAS = { 0: 70, 2: ROMANEIO_CONFIG.ALTURA_CONTEUDO_PT, 6: 85 };

function ajustarTabelaDoModelo_(corpo, larguraUtil) {
  const tabelas = corpo.getTables();
  if (!tabelas.length) throw new Error('A conversão do HTML não gerou a tabela do romaneio.');
  const tabela = tabelas[0];
  try {
    ROMANEIO_COLUNAS.forEach((f, i) => tabela.setColumnWidth(i, Math.round(larguraUtil * f)));
  } catch (e) { console.warn('Larguras das colunas: ' + e.message); }
  Object.keys(ROMANEIO_ALTURAS).forEach(i => {
    try {
      if (Number(i) < tabela.getNumRows()) tabela.getRow(Number(i)).setMinimumHeight(ROMANEIO_ALTURAS[i]);
    } catch (e) { console.warn('Altura da linha ' + i + ': ' + e.message); }
  });
}

/**
 * Troca {{LOGO}} pela imagem. Sem imagem, escreve "KaBuM!" e devolve o motivo (aviso para o Web App).
 * Se o modelo já tiver uma imagem colocada à mão (sem o marcador), não mexe.
 */
function inserirLogo_(corpo) {
  const achado = corpo.findText('\\{\\{LOGO\\}\\}');
  if (!achado) return null;

  const texto = achado.getElement().asText();
  let par = texto;
  while (par.getType() !== DocumentApp.ElementType.PARAGRAPH) par = par.getParent();
  par = par.asParagraph();

  let motivo;
  try {
    const logo = obterLogoRomaneio_();
    // A imagem entra antes do texto e só depois o marcador sai: o Docs não aceita
    // deixar o parágrafo com um texto vazio (setText('') falha com "elemento de texto vazio").
    const img = par.insertInlineImage(0, logo);
    if (texto.getText().replace(/\{\{LOGO\}\}/g, '').trim()) {
      texto.deleteText(achado.getStartOffset(), achado.getEndOffsetInclusive());
    } else {
      texto.removeFromParent();
    }
    const altura = ROMANEIO_CONFIG.ALTURA_LOGO_PT;
    const w = img.getWidth(), h = img.getHeight();
    if (w && h) img.setHeight(altura).setWidth(Math.round(w * altura / h));
    par.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    return null;
  } catch (e) {
    motivo = e && e.message || String(e);
  }
  console.warn('Logótipo indisponível, fica o texto KaBuM!: ' + motivo);
  par.setText('KaBuM!');
  return 'Logótipo não inserido: ' + motivo;
}

/** Envia o HTML ao Drive pedindo conversão para Google Docs. Devolve o ID do documento. */
function converterHtmlEmDoc_(html, nome, pastaId) {
  const limite = 'romaneio_' + Date.now();
  const metadados = { name: nome, mimeType: 'application/vnd.google-apps.document', parents: [pastaId] };
  const corpo =
    '--' + limite + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadados) + '\r\n' +
    '--' + limite + '\r\n' +
    'Content-Type: text/html; charset=UTF-8\r\n\r\n' +
    html + '\r\n' +
    '--' + limite + '--';

  const resposta = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id',
    {
      method: 'post',
      contentType: 'multipart/related; boundary=' + limite,
      payload: Utilities.newBlob(corpo).getBytes(),
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }
  );
  const codigo = resposta.getResponseCode();
  if (codigo !== 200) {
    throw new Error('O Drive recusou criar o modelo do romaneio (HTTP ' + codigo + '): ' +
      String(resposta.getContentText()).slice(0, 300));
  }
  return JSON.parse(resposta.getContentText()).id;
}

// ==========================================
// DRIVE: PASTA E LOGÓTIPO
// ==========================================

function obterPastaRomaneios_() {
  if (ROMANEIO_CONFIG.PASTA_ID) return abrirPastaRomaneio_(ROMANEIO_CONFIG.PASTA_ID, 'pasta dos romaneios');

  if (ROMANEIO_CONFIG.DRIVE_ID) {
    const drive = abrirPastaRomaneio_(ROMANEIO_CONFIG.DRIVE_ID, 'drive partilhado');
    const existentes = drive.getFoldersByName(ROMANEIO_CONFIG.NOME_PASTA);
    return existentes.hasNext() ? existentes.next() : drive.createFolder(ROMANEIO_CONFIG.NOME_PASTA);
  }

  // Sem drive configurado: pasta "Romaneios" ao lado da Planilha Base
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(ROMANEIO_CONFIG.PROP_PASTA);
  if (id) {
    try {
      const pasta = DriveApp.getFolderById(id);
      if (!pasta.isTrashed()) return pasta;
    } catch (e) { /* recria */ }
  }
  let mae;
  try {
    const pais = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
    mae = pais.hasNext() ? pais.next() : DriveApp.getRootFolder();
  } catch (e) {
    mae = DriveApp.getRootFolder();
  }
  const existentes = mae.getFoldersByName(ROMANEIO_CONFIG.NOME_PASTA);
  const pasta = existentes.hasNext() ? existentes.next() : mae.createFolder(ROMANEIO_CONFIG.NOME_PASTA);
  props.setProperty(ROMANEIO_CONFIG.PROP_PASTA, pasta.getId());
  return pasta;
}

function abrirPastaRomaneio_(id, descricao) {
  try {
    return DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('Não foi possível abrir o ' + descricao + ' (' + id + '). A conta que publica o Web App ' +
      'precisa de ser membro do drive partilhado, com acesso de "Gestor de conteúdo". Detalhe: ' + e.message);
  }
}

const ROMANEIO_MIME_PASTA = 'application/vnd.google-apps.folder';
const ROMANEIO_MIME_ATALHO = 'application/vnd.google-apps.shortcut';
// Formatos que o Google Docs aceita numa imagem; os outros são convertidos para PNG
const ROMANEIO_MIMES_DOCS = ['image/png', 'image/jpeg', 'image/gif'];

/**
 * O ID pode ser do ficheiro da imagem ou da pasta onde ela está. Aceita atalhos do Drive e
 * converte para PNG o que o Docs não aceita diretamente (WebP, BMP, Google Desenhos, etc.).
 * Em caso de falha, a mensagem diz o que foi encontrado, para se perceber o motivo.
 */
function obterLogoRomaneio_() {
  const id = ROMANEIO_CONFIG.LOGO_ID;
  if (!id) {
    if (!ROMANEIO_CONFIG.DRIVE_ID) throw new Error('logótipo não configurado (DRIVE_ID ou LOGO_ID).');
    const drive = abrirPastaRomaneio_(ROMANEIO_CONFIG.DRIVE_ID, 'drive partilhado');
    const pastas = drive.getFoldersByName(ROMANEIO_CONFIG.NOME_PASTA_LOGO);
    if (!pastas.hasNext()) throw new Error('a pasta "' + ROMANEIO_CONFIG.NOME_PASTA_LOGO + '" não existe no drive partilhado.');
    return primeiraImagemDaPasta_(pastas.next());
  }

  let arquivo = null;
  try { arquivo = DriveApp.getFileById(id); } catch (e) { /* pode ser uma pasta */ }
  if (arquivo && arquivo.getMimeType() !== ROMANEIO_MIME_PASTA) {
    return imagemDoArquivo_(arquivo);
  }

  let pasta;
  try {
    pasta = DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('o ID ' + id + ' não abre como ficheiro nem como pasta para a conta que publica o Web App (' + e.message + ').');
  }
  return primeiraImagemDaPasta_(pasta);
}

function primeiraImagemDaPasta_(pasta) {
  const vistos = [];
  const erros = [];
  const arquivos = pasta.getFiles();
  while (arquivos.hasNext()) {
    const f = arquivos.next();
    vistos.push(f.getName() + ' (' + f.getMimeType() + ')');
    try {
      return imagemDoArquivo_(f);
    } catch (e) {
      erros.push(e.message);
    }
  }
  if (!vistos.length) throw new Error('a pasta "' + pasta.getName() + '" está vazia ou os ficheiros não estão partilhados com a conta que publica o Web App.');
  throw new Error('nenhum ficheiro da pasta "' + pasta.getName() + '" serviu como imagem: ' + erros.join(' | '));
}

function imagemDoArquivo_(arquivo) {
  let f = arquivo;
  let mime = f.getMimeType();
  if (mime === ROMANEIO_MIME_ATALHO) {
    f = DriveApp.getFileById(f.getTargetId());
    mime = f.getMimeType();
  }
  const nome = f.getName() + ' (' + mime + ')';
  if (ROMANEIO_MIMES_DOCS.indexOf(mime) >= 0) return f.getBlob();
  try {
    const png = f.getAs('image/png');
    if (png && png.getBytes().length) return png;
  } catch (e) {
    throw new Error(nome + ' não pode ser usado como imagem: ' + e.message);
  }
  throw new Error(nome + ' não pode ser usado como imagem.');
}

// ==========================================
// DADOS DA CARGA
// ==========================================

/**
 * Monta o romaneio a partir do motor de julgamento (Julgamento.gs), o mesmo do pop-up Detalhes:
 *  - Listagem: pedidos da carga e a data do nome da aba; Volumes: etiquetas da carga;
 *  - Notas não encontradas na carga: pedidos em falta total (todos os volumes);
 *  - Notas aceitas parcialmente: pedidos parciais, só os volumes em falta;
 *  - Notas recusadas por prazo indenizatório: pedidos bloqueados, só os volumes bipados;
 *  - Notas recusadas por avaria: PROVISÓRIO, volumes recusados manualmente (o motivo de cada
 *    recusa será escolhido num ecrã próprio ao gerar o romaneio);
 *  - Fora da malha / outras cargas: ainda não entram (passam por validação noutro fluxo).
 */
function montarDadosRomaneioCarga_(nomeAba, dataISO) {
  const j = obterJulgamentoCarga(nomeAba);
  if (!j.itens.length) throw new Error('A carga "' + nomeAba + '" está vazia.');
  const S = j.SIT, P = j.PED;

  const registos = filtro => j.itens.filter(filtro).map(it => ({ nf: it.nf, produto: it.produto || it.etiqueta }));
  const porSecao = {
    nao_encontradas: agruparPorNF_(registos(it => it.situacaoPedido === P.FALTA_TOTAL)),
    parciais: agruparPorNF_(registos(it => it.situacaoPedido === P.PARCIAL && it.situacao === S.FALTA)),
    prazo: agruparPorNF_(registos(it => it.situacao === S.BLOQ_RECUSAR)),
    avaria: agruparPorNF_(registos(it => it.situacao === S.RECUSADO))
  };
  const secoes = ROMANEIO_CONFIG.SECOES.map(s => ({
    titulo: s.titulo, mesmaLinha: s.mesmaLinha, linhas: porSecao[s.chave] || []
  }));

  const cp = j.contagens.pedidos, ci = j.contagens.itens;
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const dataCarga = /^(\d{2})-(\d{2})-(\d{4})/.exec(nomeAba);
  return {
    carga: nomeAba,
    transportadora: ROMANEIO_CONFIG.TRANSPORTADORA,
    data: dataISO || hoje,
    listagem: { pedidos: j.pedidos.length, data: dataCarga ? `${dataCarga[1]}/${dataCarga[2]}/${dataCarga[3]}` : nomeAba },
    secoes,
    volumes: j.itens.length,
    resumo: {
      pedidos: j.pedidos.length,
      volumes: j.itens.length,
      faltaTotal: cp[P.FALTA_TOTAL] || 0,
      parciais: cp[P.PARCIAL] || 0,
      bloqueados: cp[P.BLOQ_RECUSAR] || 0,
      recusados: ci[S.RECUSADO] || 0
    }
  };
}

/** [{nf, produto}] -> [{nf, itens:[{qtd, produto}]}], pela ordem em que as NFs aparecem. */
function agruparPorNF_(registos) {
  const mapa = new Map();
  registos.forEach(r => {
    const chave = r.nf || '(sem NF)';
    if (!mapa.has(chave)) mapa.set(chave, new Map());
    const produtos = mapa.get(chave);
    produtos.set(r.produto, (produtos.get(r.produto) || 0) + 1);
  });
  return Array.from(mapa.entries()).map(([nf, produtos]) => ({
    nf, itens: Array.from(produtos.entries()).map(([produto, qtd]) => ({ qtd, produto }))
  }));
}

function explicarErroRomaneio_(e) {
  const msg = String(e && e.message || e);
  if (/DocumentApp|documents|permiss/i.test(msg)) {
    return msg + ' — Falta autorizar o acesso ao Google Docs: execute "forcarPermissoes" no editor ' +
      '(ou menu Cruzamento NF › Autorizar permissões), marque todas as caixas e publique uma nova versão.';
  }
  return msg;
}

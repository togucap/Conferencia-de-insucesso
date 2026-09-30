// ==========================================
// ROMANEIO DE DEVOLUÇÃO (PDF para a transportadora)
// ==========================================
// Fluxo:
//  1. Um HTML com o layout do romaneio é convertido pelo Google Drive num documento Google Docs
//     (o "modelo"). Isto acontece uma única vez; depois a equipa pode ajustar o modelo no Docs.
//  2. Para cada romaneio: copia o modelo, troca os marcadores {{...}}, escreve as linhas na
//     área central, exporta em PDF para a pasta "Romaneios" e apaga a cópia.
//  3. Motor alternativo (comparação): o mesmo HTML convertido diretamente em PDF, sem Docs.
//
// As regras de negócio (que NFs entram em cada secção) ficam fora deste ficheiro: o motor recebe
// os dados já prontos em gerarRomaneioPdf_(dados).

const ROMANEIO_CONFIG = {
  // Ficheiro do logótipo, ou pasta onde ele é a única imagem
  LOGO_ID: '1QQjtN3RkW0IhiNYfcgy8yMA5cW5p4twH',
  // Pasta dos PDFs. Vazio = cria "Romaneios" na mesma pasta da Planilha Base
  PASTA_ID: '',
  NOME_PASTA: 'Romaneios',
  NOME_MODELO: 'Modelo_Romaneio_Devolucao (não apagar)',
  // Suba este número quando o layout do HTML mudar: o modelo é recriado na próxima geração
  VERSAO_MODELO: 1,
  // Página: 'A4' ou 'CARTA'
  PAGINA: 'A4',
  MARGEM_PT: 28,
  ALTURA_LOGO_PT: 58,
  FONTE: 'Montserrat',
  COR_BARRA: '#2e75b6',
  COR_CINZA: '#a6a6a6',
  // Altura mínima da área central, para as assinaturas ficarem no fundo da página
  ALTURA_CONTEUDO_PT: 480,
  ESTADOS: [
    ['perfeito', 'Volumes em perfeito estado'],
    ['amassado', 'Volume amassado'],
    ['rasgado', 'Volume rasgado']
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
 * Gera um romaneio com o conteúdo do exemplo, pelos dois motores, para validar o visual.
 * @param {Object} opcoes {transportadora, dataISO, textoLongo, marcarEstados}
 */
function gerarRomaneioTeste(opcoes) {
  opcoes = opcoes || {};
  const dados = montarDadosRomaneioTeste_(opcoes);
  const resultado = { ok: true, avisos: [] };
  const inicio = Date.now();

  try {
    resultado.docs = gerarRomaneioPdf_(dados);
  } catch (e) {
    resultado.docs = { erro: explicarErroRomaneio_(e) };
  }
  try {
    resultado.html = gerarRomaneioPdfHtml_(dados);
  } catch (e) {
    resultado.html = { erro: explicarErroRomaneio_(e) };
  }

  resultado.ok = !resultado.docs.erro || !resultado.html.erro;
  resultado.segundos = Math.round((Date.now() - inicio) / 100) / 10;
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
 *   estados: {perfeito, amassado, rasgado} (true = caixa marcada)
 *   nomeArquivo: string (opcional)
 * @return {{id, url, nome, paginas?}}
 */
function gerarRomaneioPdf_(dados) {
  const pasta = obterPastaRomaneios_();
  const modelo = obterModeloRomaneio_();
  const nome = nomeArquivoRomaneio_(dados);

  const copia = modelo.makeCopy(nome + ' (temp)', pasta);
  try {
    const doc = DocumentApp.openById(copia.getId());
    const corpo = doc.getBody();

    const campos = camposRomaneio_(dados);
    Object.keys(campos).forEach(chave => substituirMarcador_(corpo, chave, campos[chave]));
    preencherConteudoDoc_(corpo, linhasConteudoRomaneio_(dados));

    doc.saveAndClose();
    const pdf = copia.getAs('application/pdf').setName(nome + '.pdf');
    const arquivo = pasta.createFile(pdf);
    return { id: arquivo.getId(), url: arquivo.getUrl(), nome: arquivo.getName() };
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
// MOTOR ALTERNATIVO (HTML direto para PDF)
// ==========================================

function gerarRomaneioPdfHtml_(dados) {
  const pasta = obterPastaRomaneios_();
  const nome = nomeArquivoRomaneio_(dados) + ' (motor HTML)';
  let logo = '';
  try {
    const blob = obterLogoRomaneio_();
    logo = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) { console.warn('Logótipo indisponível: ' + e.message); }

  const html = montarHtmlRomaneio_(dados, { modo: 'direto', logo });
  const pdf = Utilities.newBlob(html, 'text/html', nome + '.html').getAs('application/pdf').setName(nome + '.pdf');
  const arquivo = pasta.createFile(pdf);
  return { id: arquivo.getId(), url: arquivo.getUrl(), nome: arquivo.getName() };
}

// ==========================================
// CONTEÚDO
// ==========================================

function camposRomaneio_(dados) {
  const estados = dados.estados || {};
  const campos = {
    TRANSPORTADORA: String(dados.transportadora || '').trim(),
    DATA: formatarDataRomaneio_(dados.data || new Date()),
    VOLUMES: dados.volumes === undefined || dados.volumes === null ? '' : String(dados.volumes)
  };
  ROMANEIO_CONFIG.ESTADOS.forEach(([chave]) => {
    campos['CX_' + chave.toUpperCase()] = estados[chave] ? '☒' : '☐';
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
  const data = formatarDataRomaneio_(dados.data || new Date()).replace(/\//g, '-');
  const transp = String(dados.transportadora || 'Transportadora').trim().replace(/[\\/:*?"<>|]+/g, '-');
  const hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmmss');
  return `Romaneio ${transp} ${data} ${hora}`.replace(/\s+/g, ' ');
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
  inserirLogoNoModelo_(corpo);
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

function inserirLogoNoModelo_(corpo) {
  const achado = corpo.findText('\\{\\{LOGO\\}\\}');
  if (!achado) return;
  let blob;
  try {
    blob = obterLogoRomaneio_();
  } catch (e) {
    console.warn('Logótipo indisponível, fica o texto KaBuM!: ' + e.message);
    substituirMarcador_(corpo, 'LOGO', 'KaBuM!');
    return;
  }
  let par = achado.getElement();
  while (par.getType() !== DocumentApp.ElementType.PARAGRAPH) par = par.getParent();
  par = par.asParagraph();
  par.setText('');
  const img = par.appendInlineImage(blob);
  const altura = ROMANEIO_CONFIG.ALTURA_LOGO_PT;
  const w = img.getWidth(), h = img.getHeight();
  if (w && h) img.setHeight(altura).setWidth(Math.round(w * altura / h));
  par.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
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
  if (ROMANEIO_CONFIG.PASTA_ID) return DriveApp.getFolderById(ROMANEIO_CONFIG.PASTA_ID);

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

/** O ID pode ser do ficheiro da imagem ou da pasta onde ela está (usa a primeira imagem). */
function obterLogoRomaneio_() {
  const id = ROMANEIO_CONFIG.LOGO_ID;
  if (!id) throw new Error('LOGO_ID não configurado.');

  let erroFicheiro;
  try {
    const arquivo = DriveApp.getFileById(id);
    if (/^image\//.test(arquivo.getMimeType())) return arquivo.getBlob();
    erroFicheiro = new Error('O ficheiro ' + arquivo.getName() + ' não é uma imagem (' + arquivo.getMimeType() + ').');
  } catch (e) {
    erroFicheiro = e;
  }

  let pasta;
  try {
    pasta = DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('Não foi possível abrir o logótipo (' + id + '): ' + (erroFicheiro && erroFicheiro.message || e.message));
  }
  const arquivos = pasta.getFiles();
  while (arquivos.hasNext()) {
    const f = arquivos.next();
    if (/^image\//.test(f.getMimeType())) return f.getBlob();
  }
  throw new Error('A pasta "' + pasta.getName() + '" não tem nenhuma imagem para o logótipo.');
}

// ==========================================
// DADOS DE TESTE E ERROS
// ==========================================

function montarDadosRomaneioTeste_(opcoes) {
  const secoes = [
    { titulo: 'Notas não encontradas na carga', linhas: [
      { nf: '28660325', itens: [{ qtd: 1, produto: 'Memória RAM Rise Mode Z, 8GB, 3200MHz, DDR4' }] }
    ] },
    { titulo: 'Notas aceitas parcialmente', mesmaLinha: false, linhas: [
      { nf: '28691026', itens: [{ qtd: 1, produto: 'Processador AMD Ryzen 7 5700X' }],
        obs: 'Produto veio fora da sua embalagem original, sendo assim será feito a recusa parcial da NF' }
    ] },
    { titulo: 'Notas recusadas por avaria', linhas: [
      { nf: '28655181', itens: [{ qtd: 1, produto: 'Gabinete Gamer Kalkan Skye' }] },
      { nf: '28609612', itens: [{ qtd: 1, produto: 'Monitor Profissional ASUS ProArt 27' }] },
      { nf: '28602432', itens: [{ qtd: 1, produto: 'Monitor Gamer Curvo Rise Mode Prime 32' }] },
      { nf: '28613464', itens: [{ qtd: 1, produto: 'Mouse Gamer Sem Fio Attack Shark X8SE' }] },
      { nf: '28606140', itens: [{ qtd: 1, produto: 'Teclado Mecânico Gamer Husky Anchorage Full Size' }] }
    ] },
    { titulo: 'Notas recusadas com embalagem vazia', linhas: [
      { nf: '28525085', itens: [{ qtd: 1, produto: 'Processador AMD Ryzen 7 5800X3D' }] }
    ] },
    { titulo: 'Notas recusadas por prazo indenizatório', linhas: [
      { nf: '28483831', itens: [{ qtd: 1, produto: 'MacBook Pro de 14' }] },
      { nf: '28569952', itens: [{ qtd: 1, produto: 'Placa de Vídeo MSI RTX 5060 Shadow 2X OC NVIDIA GeForce' }],
        obs: '(Foi enviado um Headset improcedente no local da Placa de vídeo)' }
    ] },
    { titulo: 'Notas recusadas não pertencentes ao KaBuM', linhas: [] },
    { titulo: 'Notas recusadas por produtos improcedentes', linhas: [] },
    { titulo: 'Nota recusada fora do romaneio', linhas: [] }
  ];

  if (opcoes.textoLongo) {
    const produtos = ['Placa de Vídeo Gigabyte GeForce RTX 5070 Ti Gaming OC 16GB GDDR7', 'SSD Kingston NV3 2TB M.2 NVMe',
      'Monitor Gamer LG UltraGear 27" QHD 180Hz', 'Cadeira Gamer DT3 Sports Rhino', 'Fonte Corsair RM850e 850W 80 Plus Gold'];
    const extra = [];
    for (let i = 0; i < 70; i++) {
      extra.push({
        nf: String(28700000 + i * 137),
        itens: [{ qtd: 1 + (i % 3), produto: produtos[i % produtos.length] }].concat(i % 4 === 0 ? [{ qtd: 1, produto: produtos[(i + 2) % produtos.length] }] : []),
        obs: i % 5 === 0 ? '(Caixa com sinais de violação na lateral, lacre rompido e produto sem os acessórios originais; volume separado para análise)' : ''
      });
    }
    secoes[2].linhas = secoes[2].linhas.concat(extra);
  }

  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return {
    transportadora: opcoes.transportadora || 'GFL',
    data: opcoes.dataISO || hoje,
    listagem: { pedidos: 175, data: opcoes.dataListagemISO || opcoes.dataISO || hoje },
    secoes,
    volumes: 175,
    estados: opcoes.marcarEstados ? { perfeito: true, amassado: true, rasgado: true } : {}
  };
}

function explicarErroRomaneio_(e) {
  const msg = String(e && e.message || e);
  if (/DocumentApp|documents|permiss/i.test(msg)) {
    return msg + ' — Falta autorizar o acesso ao Google Docs: execute "forcarPermissoes" no editor ' +
      '(ou menu Cruzamento NF › Autorizar permissões), marque todas as caixas e publique uma nova versão.';
  }
  return msg;
}

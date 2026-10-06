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
  VERSAO_MODELO: 4,
  // Identifica o código do motor no resultado do Web App (para confirmar a versão publicada)
  VERSAO_MOTOR: '2026-10-06 continuação',
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
  FONTE_CONTEUDO_PT: 9,
  // Opcional: tamanhos de letra a tentar para caber numa página antes de criar continuação
  // (vazio = a letra nunca muda)
  FONTES_AJUSTE_PT: [],
  // Rodapé (carimbo/assinaturas) partido entre páginas: a última categoria passa para uma página
  // de continuação, junto com o rodapé. Uma categoria maior do que isto só passa o fim
  // (cortado numa NF), para a continuação nunca ficar maior do que uma página.
  CONTINUACAO_MAX_CARACTERES: 700,
  // Linha em branco entre categorias, em proporção da letra (mais baixa que uma linha de texto)
  PROPORCAO_LINHA_VAZIA: 0.5,
  // Caixas de estado do volume: saem todas marcadas
  MARCAR_ESTADOS: true,
  ESTADOS: [
    ['perfeito', 'Volumes em perfeito estado'],
    ['amassado', 'Volume amassado'],
    ['rasgado', 'Volume rasgado']
  ],
  // Secções da área central, pela ordem do romaneio (só aparecem as que a carga tiver).
  // mesmaLinha:false = NFs a partir da linha seguinte
  SECOES: [
    { chave: 'nao_encontradas', titulo: 'Notas não encontradas na carga' },
    { chave: 'parciais', titulo: 'Notas aceitas parcialmente' },
    { chave: 'avaria', titulo: 'Notas recusadas por avaria' },
    { chave: 'embalagem_vazia', titulo: 'Notas recusadas com embalagem vazia' },
    { chave: 'prazo', titulo: 'Notas recusadas por prazo indenizatório' },
    { chave: 'nao_kabum', titulo: 'Notas recusadas não pertencentes ao KaBuM' },
    { chave: 'improcedentes', titulo: 'Notas recusadas por produtos improcedentes' },
    { chave: 'fora_romaneio', titulo: 'Nota recusada fora do romaneio' }
  ],
  // Motivos de recusa manual escolhidos por NF ao gerar o romaneio (chave = secção)
  MOTIVOS: [
    ['avaria', 'Avaria'],
    ['improcedentes', 'Improcedente'],
    ['embalagem_vazia', 'Embalagem vazia']
  ],
  // Aba oculta com os motivos por NF e as notas "não pertencentes ao KaBuM" escritas à mão
  NOME_ABA_SELECOES: 'Romaneio_Selecoes',
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
 * Dados do ecrã de seleção prévia: NFs recusadas manualmente (com o motivo já gravado, se houver),
 * notas "não pertencentes ao KaBuM" já escritas e o que entra automaticamente no romaneio.
 */
function obterSelecaoRomaneio(nomeAba) {
  if (!nomeAba) throw new Error('Selecione a carga.');
  const j = obterJulgamentoCarga(nomeAba);
  const salvas = lerSelecoesRomaneio_(nomeAba);
  const recusados = nfsRecusadasManualmente_(j).map(r => {
    const s = salvas.motivos[r.nf] || {};
    return {
      nf: r.nf, pedido: r.pedido, qtd: r.qtd, produtoBase: r.produtoBase,
      etiquetas: r.itens.map(i => i.etiqueta),
      motivo: s.motivo || '',
      produto: s.produto || r.produtoBase
    };
  });
  const automaticas = montarDadosRomaneioCarga_(nomeAba, null, { julgamento: j, semRecusados: true, selecoes: [], naoKabum: [] });
  return {
    carga: nomeAba,
    motivos: ROMANEIO_CONFIG.MOTIVOS.map(([valor, rotulo]) => ({ valor, rotulo })),
    recusados,
    naoKabum: salvas.naoKabum,
    automaticas: automaticas.secoes.filter(s => s.linhas.length).map(s => ({ titulo: s.titulo, nfs: s.linhas.length }))
  };
}

/** Texto da área central com as escolhas atuais (sem gerar o PDF nem gravar nada). */
function previaRomaneioCarga(opcoes) {
  opcoes = opcoes || {};
  if (!opcoes.carga) throw new Error('Selecione a carga.');
  const dados = montarDadosRomaneioCarga_(opcoes.carga, opcoes.dataISO, {
    selecoes: opcoes.selecoes, naoKabum: opcoes.naoKabum, permitirSemMotivo: true
  });
  return { linhas: linhasConteudoRomaneio_(dados), volumes: dados.volumes, pendentes: dados.semMotivo };
}

/**
 * Gera o romaneio de uma carga e devolve o link do PDF.
 * @param {Object} opcoes {carga, dataISO, selecoes:[{nf, motivo, produto}], naoKabum:[{nf, qtd, produto}]}
 *   Sem "selecoes"/"naoKabum" usa o que está gravado em Romaneio_Selecoes.
 */
function gerarRomaneioCarga(opcoes) {
  opcoes = opcoes || {};
  const inicio = Date.now();
  if (!opcoes.carga) throw new Error('Selecione a carga.');

  const dados = montarDadosRomaneioCarga_(opcoes.carga, opcoes.dataISO, {
    selecoes: opcoes.selecoes, naoKabum: opcoes.naoKabum
  });
  if (opcoes.selecoes || opcoes.naoKabum) {
    gravarSelecoesRomaneio_(opcoes.carga, dados.selecoesUsadas, dados.naoKabumUsadas);
  }

  let pdf;
  try {
    pdf = gerarRomaneioPdf_(dados);
  } catch (e) {
    throw new Error(explicarErroRomaneio_(e));
  }

  const resultado = {
    ok: true,
    pdf: { id: pdf.id, url: pdf.url, nome: pdf.nome },
    diagnostico: { paginas: pdf.paginas, fonte: pdf.fonte, continuacao: !!pdf.continuacao, motor: ROMANEIO_CONFIG.VERSAO_MOTOR },
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
  const avisos = [];

  // Cópias "(temp)" que ficaram para trás em gerações anteriores (ex.: sem permissão para apagar)
  const restos = limparTemporariosRomaneio_(pasta);
  if (restos) avisos.push(restos);

  // A cópia de trabalho fica no "O meu Drive" da conta que publica o Web App: lá pode sempre ser
  // apagada (num drive partilhado, apagar exige o papel "Gestor de conteúdo").
  let destinoTemp;
  try { destinoTemp = DriveApp.getRootFolder(); } catch (e) { destinoTemp = pasta; }
  const copia = modelo.makeCopy(nome + ' (temp)', destinoTemp);
  try {
    const doc = DocumentApp.openById(copia.getId());
    const corpo = doc.getBody();

    const avisoLogo = inserirLogo_(corpo);
    if (avisoLogo) avisos.push(avisoLogo);

    const campos = camposRomaneio_(dados);
    Object.keys(campos).forEach(chave => substituirMarcador_(corpo, chave, campos[chave]));
    preencherConteudoDoc_(corpo, linhasConteudoRomaneio_(dados));

    doc.saveAndClose();
    const ajuste = ajustarRomaneioAPagina_(copia.getId());
    const pdf = ajuste.pdf.setName(nome + '.pdf');
    const arquivo = pasta.createFile(pdf);
    return {
      id: arquivo.getId(), url: arquivo.getUrl(), nome: arquivo.getName(), avisos,
      paginas: ajuste.paginas, fonte: ajuste.fonte, continuacao: ajuste.continuacao
    };
  } finally {
    try {
      copia.setTrashed(true);
    } catch (e) {
      console.warn('Não foi possível apagar a cópia temporária: ' + e.message);
      avisos.push('A cópia de trabalho "' + nome + ' (temp)" não pôde ser apagada (' + e.message + '). Pode apagá-la à mão.');
    }
  }
}

/** Apaga cópias "(temp)" esquecidas na pasta dos romaneios. Devolve um aviso se alguma resistir. */
function limparTemporariosRomaneio_(pasta) {
  let falhas = 0;
  try {
    const arquivos = pasta.searchFiles("title contains '(temp)' and trashed = false");
    while (arquivos.hasNext()) {
      const f = arquivos.next();
      if (!/ \(temp\)$/.test(f.getName())) continue;
      try { f.setTrashed(true); } catch (e) { falhas++; }
    }
  } catch (e) {
    return null;
  }
  return falhas
    ? `Há ${falhas} cópia(s) "(temp)" antigas na pasta Romaneios que não puderam ser apagadas: a conta que publica o Web App precisa do papel "Gestor de conteúdo" no drive partilhado. Pode apagá-las à mão.`
    : null;
}

// ==========================================
// AJUSTE À PÁGINA (medido no PDF exportado)
// ==========================================

/**
 * Exporta o documento e garante que o rodapé (Volumes, estados, carimbo e assinaturas) nunca
 * sai partido entre duas páginas:
 *  1. uma página: fica assim;
 *  2. (opcional) letra menor da área central, se FONTES_AJUSTE_PT tiver tamanhos;
 *  3. várias páginas: testa uma quebra antes do rodapé; se o número de páginas não aumenta,
 *     o rodapé estava partido, e o fim do texto passa para uma página de continuação (com o
 *     cabeçalho repetido) junto com o rodapé.
 * @return {{pdf: Blob, paginas: number|null, fonte: number, continuacao: boolean}}
 */
function ajustarRomaneioAPagina_(docId) {
  const C = ROMANEIO_CONFIG;
  const exportar = () => {
    const pdf = DriveApp.getFileById(docId).getAs('application/pdf');
    return { pdf, paginas: contarPaginasPdf_(pdf) };
  };
  const fim = (r, fonte, continuacao) => ({ pdf: r.pdf, paginas: r.paginas, fonte, continuacao: !!continuacao });

  let r = exportar();
  let fonte = C.FONTE_CONTEUDO_PT;
  if (!r.paginas || r.paginas <= 1) return fim(r, fonte);

  for (const tamanho of (C.FONTES_AJUSTE_PT || [])) {
    if (!aplicarFonteConteudo_(docId, tamanho)) break;
    fonte = tamanho;
    r = exportar();
    if (r.paginas && r.paginas <= 1) return fim(r, fonte);
  }

  // O rodapé está partido? Com uma quebra antes dele, se estava inteiro surge mais uma página.
  if (!quebraAntesDoRodape_(docId, true)) return fim(r, fonte);
  const comQuebra = exportar();
  quebraAntesDoRodape_(docId, false);
  if (!comQuebra.paginas || comQuebra.paginas > r.paginas) return fim(r, fonte);

  // Rodapé partido: parte do texto vai para a página seguinte, junto com o carimbo
  if (!criarPaginaContinuacao_(docId)) return fim(comQuebra, fonte);
  return fim(exportar(), fonte, true);
}

/**
 * Passa o fim do texto da área central para uma tabela de continuação (cópia do cabeçalho do
 * romaneio) numa página nova, logo antes do rodapé.
 */
function criarPaginaContinuacao_(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const corpo = doc.getBody();
    const tabelas = corpo.getTables();
    if (tabelas.length < 2) { doc.saveAndClose(); return false; }
    const tabela = tabelas[0];
    const celula = tabela.getRow(2).getCell(0);

    let ultimo = null;
    for (let i = celula.getNumChildren() - 1; i >= 0; i--) {
      const f = celula.getChild(i);
      if (f.getType() === DocumentApp.ElementType.PARAGRAPH && f.asParagraph().getText().trim()) { ultimo = f.asParagraph(); break; }
    }
    if (!ultimo || celula.getChildIndex(ultimo) === 0) { doc.saveAndClose(); return false; }

    const texto = ultimo.getText();
    const atributos = atributosDeTexto_(ultimo);
    const reaplicar = p => { if (atributos && p.getText().length) p.editAsText().setAttributes(0, p.getText().length - 1, atributos); };
    const limite = ROMANEIO_CONFIG.CONTINUACAO_MAX_CARACTERES;
    let movido = null;

    if (texto.length > limite) {
      // Categoria longa: corta numa fronteira ", NF " e passa só o fim
      let pos = texto.indexOf(', NF ');
      while (pos >= 0 && texto.length - pos > limite) pos = texto.indexOf(', NF ', pos + 1);
      if (pos > 0) {
        const doisPontos = texto.indexOf(':');
        const titulo = doisPontos > 0 ? texto.slice(0, doisPontos) : '';
        movido = (titulo ? titulo + ' (continuação): ' : '') + texto.slice(pos + 2);
        ultimo.setText(texto.slice(0, pos));
        reaplicar(ultimo);
      }
    }
    if (movido === null) {
      // Categoria inteira (e a linha em branco antes dela)
      movido = texto;
      const indice = celula.getChildIndex(ultimo);
      ultimo.removeFromParent();
      const anterior = indice > 0 ? celula.getChild(indice - 1) : null;
      if (anterior && anterior.getType() === DocumentApp.ElementType.PARAGRAPH &&
          !anterior.asParagraph().getText().trim() && celula.getNumChildren() > 1) {
        anterior.removeFromParent();
      }
    }

    // Tabela de continuação: mesmo cabeçalho, só com o texto movido
    const continuacao = tabela.copy();
    const celulaC = continuacao.getRow(2).getCell(0);
    while (celulaC.getNumChildren() > 1) celulaC.getChild(celulaC.getNumChildren() - 1).removeFromParent();
    const p0 = celulaC.getChild(0).asParagraph();
    p0.setText(movido);
    reaplicar(p0);

    const indiceTabela = corpo.getChildIndex(tabela);
    const quebra = corpo.insertPageBreak(indiceTabela + 1);
    try { quebra.getParent().asParagraph().setSpacingBefore(0).setSpacingAfter(0); } catch (e) { /* sem ajuste */ }
    corpo.insertTable(indiceTabela + 2, continuacao);
    doc.saveAndClose();
    return true;
  } catch (e) {
    console.warn('Página de continuação do romaneio falhou: ' + e.message);
    return false;
  }
}

/** Número de páginas de um PDF (lê a árvore /Pages; null se não conseguir). */
function contarPaginasPdf_(blob) {
  try {
    const texto = Utilities.newBlob(blob.getBytes()).getDataAsString('ISO-8859-1');
    let maior = 0;
    const reCount = /\/Type\s*\/Pages\b[^>]*?\/Count\s+(\d+)|\/Count\s+(\d+)[^>]*?\/Type\s*\/Pages\b/g;
    let m;
    while ((m = reCount.exec(texto))) maior = Math.max(maior, Number(m[1] || m[2]));
    if (maior) return maior;
    const paginas = texto.match(/\/Type\s*\/Page(?![a-zA-Z])/g);
    return paginas ? paginas.length : null;
  } catch (e) {
    return null;
  }
}

/** Muda o tamanho da letra da área central (linhas em branco proporcionais). */
function aplicarFonteConteudo_(docId, tamanho) {
  try {
    const doc = DocumentApp.openById(docId);
    const celula = doc.getBody().getTables()[0].getRow(2).getCell(0);
    for (let i = 0; i < celula.getNumChildren(); i++) {
      const filho = celula.getChild(i);
      if (filho.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
      const texto = filho.asParagraph().editAsText();
      if (!texto.getText().length) continue;
      const vazia = !texto.getText().trim();
      texto.setFontSize(vazia ? tamanho * ROMANEIO_CONFIG.PROPORCAO_LINHA_VAZIA : tamanho);
    }
    doc.saveAndClose();
    return true;
  } catch (e) {
    console.warn('Ajuste da letra do romaneio falhou: ' + e.message);
    return false;
  }
}

/** Liga/desliga uma quebra de página entre a tabela do conteúdo e a do rodapé. */
function quebraAntesDoRodape_(docId, ligar) {
  try {
    const doc = DocumentApp.openById(docId);
    const corpo = doc.getBody();
    if (ligar) {
      const tabelas = corpo.getTables();
      if (tabelas.length < 2) { doc.saveAndClose(); return false; }
      corpo.insertPageBreak(corpo.getChildIndex(tabelas[0]) + 1);
    } else {
      const achado = corpo.findElement(DocumentApp.ElementType.PAGE_BREAK);
      if (achado) achado.getElement().getParent().removeFromParent();
    }
    doc.saveAndClose();
    return true;
  } catch (e) {
    console.warn('Quebra de página antes do rodapé falhou: ' + e.message);
    return false;
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
    if (texto === '') p.editAsText().setFontSize(ROMANEIO_CONFIG.FONTE_CONTEUDO_PT * ROMANEIO_CONFIG.PROPORCAO_LINHA_VAZIA);
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
  // Só entram as categorias que a carga tem (secções sem NFs não aparecem),
  // salvo dados.mostrarVazias = true
  const secoes = (dados.secoes || [])
    .map(secao => ({ secao, itens: (secao.linhas || []).map(formatarItemRomaneio_).filter(t => t) }))
    .filter(x => x.itens.length || dados.mostrarVazias);
  // Romaneio enxuto: cada categoria é um único parágrafo, com as NFs separadas por ", "
  // (a quebra de linha fica a cargo do documento). Uma linha em branco entre categorias.
  secoes.forEach(({ secao, itens }, s) => {
    const titulo = String(secao.titulo || '').trim().replace(/:?$/, ':');
    const nfs = itens.join(', ');
    if (!nfs) linhas.push(titulo);
    else if (secao.mesmaLinha === false) { linhas.push(titulo); linhas.push(nfs); }
    else linhas.push(titulo + ' ' + nfs);
    if (s < secoes.length - 1) linhas.push('');
  });
  // Sem nenhuma categoria, não fica uma linha em branco solta depois da Listagem
  while (linhas.length && linhas[linhas.length - 1] === '') linhas.pop();
  return linhas;
}

/**
 * Padrão de escrita do romaneio: "NF {{NF}} ({{contagem de itens}}, {{nome do produto}})".
 * Vários produtos na mesma NF: "NF 123 (2, Monitor; 1, Cabo)". Observação opcional no fim.
 * Um texto (string) entra tal como veio.
 */
function formatarItemRomaneio_(item) {
  if (item === null || item === undefined) return '';
  if (typeof item !== 'object') return String(item).trim();

  const produtos = (item.itens || []).map(i => {
    const qtd = i.qtd === undefined || i.qtd === null || i.qtd === '' ? 1 : i.qtd;
    return `${qtd}, ${String(i.produto || '').trim()}`;
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
</table>
<table style="border-collapse:collapse;width:100%;table-layout:fixed;${fonte}">
  <colgroup>${col.map(w => `<col style="width:${w}">`).join('')}</colgroup>
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

  // Entre a tabela do conteúdo e a do rodapé o Docs exige um parágrafo: fica com 1 pt
  const tabelas = filhos.map((f, i) => f.getType() === DocumentApp.ElementType.TABLE ? i : -1).filter(i => i >= 0);
  if (tabelas.length > 1) {
    const entre = filhos.slice(tabelas[0] + 1, tabelas[1]);
    if (entre.length && entre.every(vazio)) apagarOuMinimizar(entre, entre[0]);
  }

  let k = filhos.length;
  while (k > 0 && vazio(filhos[k - 1])) k--;
  const depois = filhos.slice(k);
  if (depois.length) apagarOuMinimizar(depois, depois[depois.length - 1]);
}

// Larguras das colunas e alturas mínimas das linhas (a conversão do HTML nem sempre as respeita)
const ROMANEIO_COLUNAS = [0.22, 0.23, 0.27, 0.28];
// [tabela do cabeçalho e conteúdo, tabela do rodapé] -> {linha: altura mínima em pt}
const ROMANEIO_ALTURAS = [{ 0: 70, 2: ROMANEIO_CONFIG.ALTURA_CONTEUDO_PT }, { 3: 85 }];

function ajustarTabelaDoModelo_(corpo, larguraUtil) {
  const tabelas = corpo.getTables();
  if (tabelas.length < 2) throw new Error('A conversão do HTML não gerou as tabelas do romaneio.');
  tabelas.slice(0, 2).forEach((tabela, t) => {
    try {
      ROMANEIO_COLUNAS.forEach((f, i) => tabela.setColumnWidth(i, Math.round(larguraUtil * f)));
    } catch (e) { console.warn('Larguras das colunas: ' + e.message); }
    const alturas = ROMANEIO_ALTURAS[t];
    Object.keys(alturas).forEach(i => {
      try {
        if (Number(i) < tabela.getNumRows()) tabela.getRow(Number(i)).setMinimumHeight(alturas[i]);
      } catch (e) { console.warn('Altura da linha ' + i + ': ' + e.message); }
    });
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
 * Monta o romaneio a partir do motor de julgamento (Julgamento.gs) e das escolhas do ecrã de seleção:
 *  - Notas não encontradas na carga: pedidos em falta total (todos os volumes);
 *  - Notas aceitas parcialmente: pedidos parciais, só os volumes em falta;
 *  - Notas recusadas por prazo indenizatório: volumes bipados com status "Bloqueado";
 *  - Avaria / Improcedente / Embalagem vazia: NFs recusadas manualmente, pelo motivo escolhido
 *    por NF (na embalagem vazia o nome do produto pode ser corrigido à mão);
 *  - Não pertencentes ao KaBuM: NF, quantidade e produto escritos à mão;
 *  - Fora da malha / outras cargas: ainda não entram (passam por validação noutro fluxo).
 * Volumes = soma dos volumes listados no romaneio.
 *
 * @param {Object} opcoes {julgamento, selecoes, naoKabum, permitirSemMotivo, semRecusados}
 *   sem selecoes/naoKabum usa o que está gravado na aba Romaneio_Selecoes.
 */
function montarDadosRomaneioCarga_(nomeAba, dataISO, opcoes) {
  opcoes = opcoes || {};
  const j = opcoes.julgamento || obterJulgamentoCarga(nomeAba);
  if (!j.itens.length) throw new Error('A carga "' + nomeAba + '" está vazia.');
  const S = j.SIT, P = j.PED;

  const registos = filtro => j.itens.filter(filtro).map(it => ({ nf: it.nf, produto: it.produto || it.etiqueta }));
  const porSecao = {
    nao_encontradas: agruparPorNF_(registos(it => it.situacaoPedido === P.FALTA_TOTAL)),
    parciais: agruparPorNF_(registos(it => it.situacaoPedido === P.PARCIAL && it.situacao === S.FALTA)),
    // Prazo expirado: só volumes que passaram na bipagem e ficaram com o status "Bloqueado"
    prazo: agruparPorNF_(registos(it => it.situacao === S.BLOQ_RECUSAR && it.status === 'Bloqueado'))
  };

  // ---- recusados manualmente: motivo por NF ----
  const salvas = (opcoes.selecoes && opcoes.naoKabum) ? null : lerSelecoesRomaneio_(nomeAba);
  const mapaSel = {};
  (opcoes.selecoes || []).forEach(s => { if (s && s.nf) mapaSel[String(s.nf).trim()] = s; });
  const motivosValidos = ROMANEIO_CONFIG.MOTIVOS.map(m => m[0]);
  const selecoesUsadas = [];
  const semMotivo = [];
  if (!opcoes.semRecusados) {
    nfsRecusadasManualmente_(j).forEach(r => {
      const escolha = opcoes.selecoes ? (mapaSel[r.nf] || {}) : (salvas.motivos[r.nf] || {});
      const motivo = String(escolha.motivo || '').trim();
      if (motivosValidos.indexOf(motivo) < 0) { semMotivo.push(r.nf); return; }
      const produtoManual = String(escolha.produto || '').trim();
      const linha = (motivo === 'embalagem_vazia' && produtoManual)
        ? { nf: r.nf, itens: [{ qtd: r.qtd, produto: produtoManual }] }
        : agruparPorNF_(r.itens.map(i => ({ nf: r.nf, produto: i.produto || i.etiqueta })))[0];
      (porSecao[motivo] = porSecao[motivo] || []).push(linha);
      selecoesUsadas.push({ nf: r.nf, motivo, produto: motivo === 'embalagem_vazia' ? (produtoManual || r.produtoBase) : '' });
    });
    if (semMotivo.length && !opcoes.permitirSemMotivo) {
      throw new Error('Escolha o motivo da recusa para a(s) NF(s): ' + semMotivo.join(', ') + '.');
    }
  }

  // ---- não pertencentes ao KaBuM (escritas à mão) ----
  const naoKabumUsadas = [];
  (opcoes.naoKabum || (salvas ? salvas.naoKabum : [])).forEach(n => {
    const nf = String(n && n.nf || '').trim();
    const produto = String(n && n.produto || '').trim();
    const qtd = Math.max(1, parseInt(n && n.qtd, 10) || 1);
    if (!nf && !produto) return;
    if (!nf || !produto) throw new Error('Nas notas não pertencentes ao KaBuM, preencha a NF e o produto.');
    naoKabumUsadas.push({ nf, qtd, produto });
  });
  porSecao.nao_kabum = naoKabumUsadas.map(n => ({ nf: n.nf, itens: [{ qtd: n.qtd, produto: n.produto }] }));

  const secoes = ROMANEIO_CONFIG.SECOES.map(s => ({
    titulo: s.titulo, mesmaLinha: s.mesmaLinha, linhas: porSecao[s.chave] || []
  }));
  const volumes = secoes.reduce((total, s) => total + s.linhas.reduce((t, l) =>
    t + (l.itens || []).reduce((q, i) => q + (Number(i.qtd) || 0), 0), 0), 0);

  const cp = j.contagens.pedidos, ci = j.contagens.itens;
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const dataCarga = /^(\d{2})-(\d{2})-(\d{4})/.exec(nomeAba);
  return {
    carga: nomeAba,
    transportadora: ROMANEIO_CONFIG.TRANSPORTADORA,
    data: dataISO || hoje,
    listagem: { pedidos: j.pedidos.length, data: dataCarga ? `${dataCarga[1]}/${dataCarga[2]}/${dataCarga[3]}` : nomeAba },
    secoes,
    volumes,
    semMotivo,
    selecoesUsadas,
    naoKabumUsadas,
    resumo: {
      pedidos: j.pedidos.length,
      volumes,
      faltaTotal: cp[P.FALTA_TOTAL] || 0,
      parciais: cp[P.PARCIAL] || 0,
      bloqueados: cp[P.BLOQ_RECUSAR] || 0,
      recusados: ci[S.RECUSADO] || 0
    }
  };
}

/** Volumes recusados manualmente, agrupados por NF (pela ordem da carga). */
function nfsRecusadasManualmente_(j) {
  const mapa = {};
  const lista = [];
  j.itens.filter(it => it.situacao === j.SIT.RECUSADO).forEach(it => {
    const nf = it.nf || ('Pedido ' + it.pedido);
    if (!mapa[nf]) { mapa[nf] = { nf, pedido: it.pedido, itens: [] }; lista.push(mapa[nf]); }
    mapa[nf].itens.push(it);
  });
  lista.forEach(r => {
    r.qtd = r.itens.length;
    const nomes = [];
    r.itens.forEach(i => { const n = i.produto || i.etiqueta; if (nomes.indexOf(n) < 0) nomes.push(n); });
    r.produtoBase = nomes.join('; ');
  });
  return lista;
}

// ==========================================
// ESCOLHAS GRAVADAS (aba oculta Romaneio_Selecoes)
// ==========================================
// Colunas: Carga | Tipo | NF | Motivo | Quantidade | Produto | Atualizado em | Por
// Tipo "RECUSA" = motivo escolhido para uma NF recusada manualmente; "NAO_KABUM" = nota escrita à mão.

const ROMANEIO_CAB_SELECOES = ['Carga', 'Tipo', 'NF', 'Motivo', 'Quantidade', 'Produto', 'Atualizado em', 'Por'];

function abaSelecoesRomaneio_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(ROMANEIO_CONFIG.NOME_ABA_SELECOES);
  if (!aba && criar) {
    aba = ss.insertSheet(ROMANEIO_CONFIG.NOME_ABA_SELECOES);
    aba.getRange(1, 1, 1, ROMANEIO_CAB_SELECOES.length).setValues([ROMANEIO_CAB_SELECOES]).setFontWeight('bold');
    aba.setFrozenRows(1);
    try { aba.hideSheet(); } catch (e) { /* fica visível */ }
  }
  return aba;
}

function lerSelecoesRomaneio_(nomeAba) {
  const resultado = { motivos: {}, naoKabum: [] };
  const aba = abaSelecoesRomaneio_(false);
  if (!aba || aba.getLastRow() < 2) return resultado;
  aba.getRange(2, 1, aba.getLastRow() - 1, ROMANEIO_CAB_SELECOES.length).getDisplayValues().forEach(l => {
    if (String(l[0]).trim() !== nomeAba) return;
    const tipo = String(l[1]).trim();
    if (tipo === 'RECUSA') {
      resultado.motivos[String(l[2]).trim()] = { motivo: String(l[3]).trim(), produto: String(l[5]).trim() };
    } else if (tipo === 'NAO_KABUM') {
      resultado.naoKabum.push({ nf: String(l[2]).trim(), qtd: parseInt(l[4], 10) || 1, produto: String(l[5]).trim() });
    }
  });
  return resultado;
}

/** Substitui as escolhas gravadas desta carga pelas atuais. */
function gravarSelecoesRomaneio_(nomeAba, selecoes, naoKabum) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const aba = abaSelecoesRomaneio_(true);
    const n = ROMANEIO_CAB_SELECOES.length;
    const outras = aba.getLastRow() > 1
      ? aba.getRange(2, 1, aba.getLastRow() - 1, n).getValues().filter(l => String(l[0]).trim() !== nomeAba)
      : [];
    const quando = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    let quem = '';
    try { quem = Session.getActiveUser().getEmail(); } catch (e) { /* sem e-mail */ }
    const novas = (selecoes || []).map(s => [nomeAba, 'RECUSA', s.nf, s.motivo, '', s.produto || '', quando, quem])
      .concat((naoKabum || []).map(k => [nomeAba, 'NAO_KABUM', k.nf, 'nao_kabum', k.qtd, k.produto, quando, quem]));
    const todas = outras.concat(novas);
    if (aba.getLastRow() > 1) aba.getRange(2, 1, aba.getLastRow() - 1, n).clearContent();
    if (todas.length) {
      const intervalo = aba.getRange(2, 1, todas.length, n);
      intervalo.setNumberFormat('@');
      intervalo.setValues(todas.map(l => l.map(v => String(v))));
    }
  } finally {
    lock.releaseLock();
  }
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

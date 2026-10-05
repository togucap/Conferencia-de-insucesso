/**
 * ============================================================================
 * SISTEMA: K-OUT Auditoria Express
 * FICHEIRO: Codigo.gs (Backend / Servidor Google Apps Script)
 * ============================================================================
 */

/* STREAMING_CHUNK:Configurando funcoes base e servidor web... */
// ==========================================
// 1. CONFIGURAÇÕES INICIAIS E SERVIDOR WEB
// ==========================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle("Controle de Insucessos")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function _parseDataDoNome(nomeStr) {
  const partes = nomeStr.substring(0, 10).split('-'); 
  return new Date(partes[2], partes[1] - 1, partes[0]).getTime();
}

/* STREAMING_CHUNK:Implementando funcoes de leitura e getters... */
// ==========================================
// 2. FUNÇÕES DE LEITURA (GETTERS)
// ==========================================
function getAbasPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ss.getSheets();
  const listaAba = [];
  
  abas.forEach(aba => {
    const nome = aba.getName();
    if (/^\d{2}-\d{2}-\d{4}/.test(nome)) {
      const cor = aba.getTabColor();
      const isFinalizada = (cor === "#10b981" || cor === "#00ff00"); 
      listaAba.push({ nome: nome, finalizada: isFinalizada });
    }
  });
  
  listaAba.sort((a, b) => {
      if (a.finalizada === false && b.finalizada === true) return -1;
      if (a.finalizada === true && b.finalizada === false) return 1;
      const dataA = _parseDataDoNome(a.nome);
      const dataB = _parseDataDoNome(b.nome);
      return dataB - dataA; 
  });
  
  return listaAba;
}

function getDadosDaAba(nomeAba) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) throw new Error("Aba não encontrada.");

    const ultimaLinha = aba.getLastRow();
    const ultimaColuna = aba.getLastColumn();
    
    if (ultimaLinha === 0 || ultimaColuna === 0) {
       return { dados: [], bloqueio: "" };
    }

    const dados = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getDisplayValues();
    const cabecalho = dados[0].map(c => String(c).trim().toLowerCase());
    
    let bloqueio = "";
    if (!cabecalho.includes("mercadoria código") && !cabecalho.includes("mercadoria codigo")) {
        bloqueio = "FALHA ESTRUTURAL: A coluna 'Mercadoria Código' é obrigatória e não foi encontrada.";
    } else if (!cabecalho.includes("número do pedido faturado") && !cabecalho.includes("numero do pedido faturado")) {
        bloqueio = "FALHA ESTRUTURAL: A coluna 'Número do Pedido Faturado' é obrigatória e não foi encontrada.";
    }

    // ========================================================================
    // NOVA TRAVA: VERIFICAÇÃO DE INTEGRIDADE (LINHAS CORROMPIDAS COM APENAS NF)
    // ========================================================================
    if (bloqueio === "") {
        let idxNF = cabecalho.indexOf("nota fiscal") !== -1 ? cabecalho.indexOf("nota fiscal") : cabecalho.indexOf("nf");
        let idxMerc = cabecalho.indexOf("mercadoria código") !== -1 ? cabecalho.indexOf("mercadoria código") : cabecalho.indexOf("mercadoria codigo");
        let idxPed = cabecalho.indexOf("número do pedido faturado") !== -1 ? cabecalho.indexOf("número do pedido faturado") : cabecalho.indexOf("numero do pedido faturado");

        if (idxNF !== -1 && idxMerc !== -1) {
            for (let i = 1; i < dados.length; i++) {
                let valNF = String(dados[i][idxNF]).trim();
                let valMerc = String(dados[i][idxMerc]).trim();
                let valPed = idxPed !== -1 ? String(dados[i][idxPed]).trim() : "";

                // Se a linha tem NF, mas NÃO tem Mercadoria nem Pedido, a base está corrompida!
                if (valNF !== "" && valMerc === "" && valPed === "") {
                    bloqueio = `<strong>FALHA DE INTEGRIDADE:</strong><br><br>A linha ${i + 1} possui a "NF" preenchida (${valNF}), mas faltam os dados obrigatórios do produto.<br><br>A auditoria foi bloqueada. Corrija o ficheiro base antes de iniciar.`;
                    break;
                }
            }
        }
    }
    // ========================================================================

    return { dados: dados, bloqueio: bloqueio };
  } catch (e) {
    return { erro: "Erro ao ler base: " + e.message };
  }
}

function getErrosDaCarga(nomeAba) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return [];

    const dados = aba.getDataRange().getDisplayValues();
    const cabecalho = dados[0].map(c => String(c).trim().toLowerCase());
    const idxStatus = cabecalho.indexOf("status conferência") !== -1 ? cabecalho.indexOf("status conferência") : cabecalho.indexOf("status conferencia");
    const idxMercadoria = cabecalho.indexOf("mercadoria código") !== -1 ? cabecalho.indexOf("mercadoria código") : cabecalho.indexOf("mercadoria codigo");
    const idxData = cabecalho.indexOf("data/hora conferência") !== -1 ? cabecalho.indexOf("data/hora conferência") : cabecalho.indexOf("data/hora conferencia");
    const idxOutras = cabecalho.indexOf("outras cargas");

    if (idxStatus === -1 || idxMercadoria === -1) return [];

    let erros = [];
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][idxStatus]).includes("Divergência") || String(dados[i][idxStatus]).includes("Não Encontrado")) {
        erros.push({
          codigo: dados[i][idxMercadoria],
          hora: idxData !== -1 ? dados[i][idxData] : "Sessão Anterior",
          outrasCargas: idxOutras !== -1 ? String(dados[i][idxOutras]).split(/\s*[,;]\s*/).filter(Boolean) : []
        });
      }
    }
    return erros;
  } catch (e) {
    return [];
  }
}

/* STREAMING_CHUNK:Estruturando logica de finalizacao e gravacao de dados... */
// ==========================================
// 3. FUNÇÕES DE ESCRITA E FINALIZAÇÃO
// ==========================================
function registrarConferencia(nomeAba, rowIndex, dataHora) {
  _atualizarCelulaStatus(nomeAba, rowIndex, "Conferido", dataHora);
}

function registrarBloqueio(nomeAba, rowIndex, dataHora) {
  _atualizarCelulaStatus(nomeAba, rowIndex, "Bloqueado", dataHora);
}

// Modo recusa: o volume foi bipado, mas separado e recusado pelo operador.
// O status é "cru" (sem motivo); o motivo é escolhido mais tarde, ao gerar o romaneio.
const STATUS_RECUSA_MANUAL = "Recusado Manualmente";
function registrarRecusaManual(nomeAba, rowIndex, dataHora) {
  _atualizarCelulaStatus(nomeAba, rowIndex, STATUS_RECUSA_MANUAL, dataHora);
}

function _atualizarCelulaStatus(nomeAba, rowIndex, statusTxt, dataHora) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) return;

  const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(c => String(c).trim().toLowerCase());
  let idxStatus = cabecalho.indexOf("status conferência");
  if (idxStatus === -1) idxStatus = cabecalho.indexOf("status conferencia");
  
  let idxDataHora = cabecalho.indexOf("data/hora conferência");
  if (idxDataHora === -1) idxDataHora = cabecalho.indexOf("data/hora conferencia");

  if (idxStatus === -1) {
    idxStatus = aba.getLastColumn();
    aba.getRange(1, idxStatus + 1).setValue("Status Conferência").setFontWeight("bold");
  }
  
  if (idxDataHora === -1) {
    idxDataHora = aba.getLastColumn();
    if (idxDataHora === idxStatus) idxDataHora++; 
    aba.getRange(1, idxDataHora + 1).setValue("Data/Hora Conferência").setFontWeight("bold");
  }

  aba.getRange(rowIndex, idxStatus + 1).setValue(statusTxt);
  aba.getRange(rowIndex, idxDataHora + 1).setValue(dataHora);
}

function registrarErroNaMesmaAba(nomeAba, codigoBipado, dataHora) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) return;
  
  const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  let novaLinha = new Array(cabecalho.length).fill("");
  
  const cbLower = cabecalho.map(c => String(c).trim().toLowerCase());
  
  const idxMerc = cbLower.indexOf("mercadoria código") !== -1 ? cbLower.indexOf("mercadoria código") : cbLower.indexOf("mercadoria codigo");
  const idxStatus = cbLower.indexOf("status conferência") !== -1 ? cbLower.indexOf("status conferência") : cbLower.indexOf("status conferencia");
  const idxData = cbLower.indexOf("data/hora conferência") !== -1 ? cbLower.indexOf("data/hora conferência") : cbLower.indexOf("data/hora conferencia");
  const idxPed = cbLower.indexOf("número do pedido faturado") !== -1 ? cbLower.indexOf("número do pedido faturado") : cbLower.indexOf("numero do pedido faturado");

  if (idxMerc !== -1) novaLinha[idxMerc] = codigoBipado;
  if (idxPed !== -1) novaLinha[idxPed] = "N/A"; 
  
  if (idxStatus !== -1) {
     novaLinha[idxStatus] = "Divergência (Fora da Carga)";
  } else {
     aba.getRange(1, aba.getLastColumn() + 1).setValue("Status Conferência");
     novaLinha.push("Divergência (Fora da Carga)");
  }
  
  if (idxData !== -1) {
     novaLinha[idxData] = dataHora;
  } else {
     aba.getRange(1, aba.getLastColumn() + 1).setValue("Data/Hora Conferência");
     novaLinha.push(dataHora);
  }

  // Fora da malha: procura a etiqueta nas outras cargas e regista onde ela existe
  let outrasCargas = [];
  try {
    outrasCargas = _procurarEtiquetaEmOutrasCargas_(ss, codigoBipado, nomeAba);
  } catch (e) {
    console.warn("Procura em outras cargas falhou: " + e.message);
  }
  let idxOutras = cbLower.indexOf("outras cargas");
  if (idxOutras === -1) {
    idxOutras = Math.max(aba.getLastColumn(), novaLinha.length);
    aba.getRange(1, idxOutras + 1).setValue("Outras Cargas").setFontWeight("bold");
  }
  while (novaLinha.length <= idxOutras) novaLinha.push("");
  novaLinha[idxOutras] = outrasCargas.join(", ");

  aba.appendRow(novaLinha);
  
  const lastRow = aba.getLastRow();
  aba.getRange(lastRow, 1, 1, aba.getLastColumn()).setBackground("#fee2e2").setFontColor("#b91c1c");
  return { etiqueta: String(codigoBipado), outrasCargas: outrasCargas };
}

function finalizarCargaStatus(nomeAba) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  let recebidos = [];
  if (aba) {
    aba.setTabColor("#10b981"); // Verde Esmeralda = Finalizada
    const setListaChegada = _obterSetGlobalListaChegada();
    recebidos = _sincronizarAbaComSetListaChegada(aba, setListaChegada);
    
    // GATILHO JSON: Ao finalizar, a carga alimenta o histórico
    _atualizarHistoricoNaFinalizacao(nomeAba, aba, setListaChegada);
  }
  return recebidos;
}

function reabrirCargaStatus(nomeAba) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  if (aba) aba.setTabColor(null); 
}

/* STREAMING_CHUNK:Inicializando metodos para upload e manipulacao de dados... */
// ==========================================
// 4. FERRAMENTAS AVANÇADAS (UPLOAD E BLOQUEIO)
// ==========================================
function uploadNovaCarga(nomeAba, matrizDadosInput) {
  try {
    // Tratamento anti-crash para receber Strings JSON pesadas do Javascript do cliente
    const matrizDados = typeof matrizDadosInput === 'string' ? JSON.parse(matrizDadosInput) : matrizDadosInput;
    return _criarAbaCarga(nomeAba, matrizDados);
  } catch (e) {
    return { erro: "Erro ao importar: " + e.message };
  }
}

/**
 * Cria a aba de uma carga a partir de uma matriz (linha 0 = cabeçalho).
 * Usada pelo upload de ficheiro (uploadNovaCarga) e pelo Cruzamento de NFs (Cruzamento.gs).
 * - Se a aba já existir, cria "<nome> V2", "V3"...
 * - Injeta a coluna "Conferencia MSPC" a partir do relatório "Monte o Seu".
 * opcoes.comoTexto: grava as células como texto puro (preserva zeros à esquerda).
 * Lança exceção em caso de erro.
 */
function _criarAbaCarga(nomeAba, matrizDados, opcoes) {
    opcoes = opcoes || {};
    if (!matrizDados || !matrizDados.length || !matrizDados[0].length) throw new Error("A carga está vazia.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let aba = ss.getSheetByName(nomeAba);
    
    if (aba) {
       let contador = 2;
       while (ss.getSheetByName(`${nomeAba} V${contador}`)) {
         contador++;
       }
       nomeAba = `${nomeAba} V${contador}`;
    }
    
    // ========================================================================
    // GATILHO ON-THE-FLY: Cruzamento automático com o "Monte o Seu"
    // ========================================================================
    const mapaMonteOSeu = _obterMapaMonteOSeu(); // Vai ao Drive buscar a memória
    const cabecalhoMatriz = matrizDados[0].map(c => String(c).trim().toLowerCase());
    
    let idxPed = cabecalhoMatriz.indexOf("número do pedido faturado");
    if (idxPed === -1) idxPed = cabecalhoMatriz.indexOf("numero do pedido faturado");

    // Se a carga tiver a coluna de pedido, insere a nova coluna "Conferencia MSPC"
    if (idxPed !== -1) {
        matrizDados[0].push("Conferencia MSPC"); // Rótulo da coluna na tabela final
        for (let i = 1; i < matrizDados.length; i++) {
            let ped = matrizDados[i][idxPed] ? String(matrizDados[i][idxPed]).trim() : "";
            
            // Se o pedido existir no relatório do Drive, injeta o valor. Senão, fica em branco.
            if (ped && mapaMonteOSeu[ped]) {
                matrizDados[i].push(mapaMonteOSeu[ped]);
            } else {
                matrizDados[i].push(""); 
            }
        }
    }
    // ========================================================================
    
    aba = ss.insertSheet(nomeAba);
    const intervalo = aba.getRange(1, 1, matrizDados.length, matrizDados[0].length);
    if (opcoes.comoTexto) intervalo.setNumberFormat('@');
    intervalo.setValues(matrizDados);
    aba.getRange(1, 1, 1, matrizDados[0].length).setFontWeight("bold").setBackground("#f1f5f9");
    aba.setFrozenRows(1);

    return { sucesso: true, aba: nomeAba };
}

function aplicarRecusasEmMassa(nomeAba, tipoBusca, arrayValores) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) throw new Error("Aba não encontrada.");

    const lastRow = aba.getLastRow();
    const lastCol = aba.getLastColumn();
    if (lastRow < 2) throw new Error("Aba vazia.");

    const cabecalho = aba.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(c => String(c).trim().toLowerCase());
    
    let idxBusca = -1;
    if (tipoBusca === "NF") {
        idxBusca = cabecalho.indexOf("nota fiscal") !== -1 ? cabecalho.indexOf("nota fiscal") : cabecalho.indexOf("nf");
    } else {
        idxBusca = cabecalho.indexOf("número do pedido faturado") !== -1 ? cabecalho.indexOf("número do pedido faturado") : cabecalho.indexOf("numero do pedido faturado");
    }

    if (idxBusca === -1) throw new Error(`Coluna de busca (${tipoBusca}) não encontrada na estrutura da planilha.`);

    let idxAvaliacao = cabecalho.indexOf("avaliação");
    if (idxAvaliacao === -1) idxAvaliacao = cabecalho.indexOf("avaliacao");
    
    if (idxAvaliacao === -1) {
        idxAvaliacao = lastCol;
        aba.getRange(1, idxAvaliacao + 1).setValue("Avaliação").setFontWeight("bold");
    }

    const colunaBusca = aba.getRange(2, idxBusca + 1, lastRow - 1, 1).getDisplayValues();
    const colunaAvaliacao = aba.getRange(2, idxAvaliacao + 1, lastRow - 1, 1).getValues();

    const setProcurados = new Set(arrayValores.map(v => String(v).trim()));
    let alteracoesFeitas = 0;

    for (let i = 0; i < colunaBusca.length; i++) {
        const valCelula = String(colunaBusca[i][0]).trim();
        if (setProcurados.has(valCelula)) {
            colunaAvaliacao[i][0] = "Não Receber";
            alteracoesFeitas++;
        }
    }

    if (alteracoesFeitas > 0) {
       aba.getRange(2, idxAvaliacao + 1, lastRow - 1, 1).setValues(colunaAvaliacao);
    }
    
    return { sucesso: true, modificados: alteracoesFeitas };
    
  } catch (e) {
    return { erro: "Erro ao processar o cruzamento: " + e.message };
  }
}

/* STREAMING_CHUNK:Carregando metodos para obter os dashboards e importacao de drives... */
// ==========================================
// 5. PAINEL GERENCIAL E IMPORTADORES
// ==========================================
function getDashboardMetrics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abas = ss.getSheets();
    let chartData = [];
    
    abas.forEach(aba => {
      const nome = aba.getName();
      if (/^\d{2}-\d{2}-\d{4}/.test(nome)) {
        
        const dados = aba.getDataRange().getDisplayValues();
        if (dados.length > 1) {
          const cabecalho = dados[0].map(c => String(c).trim().toLowerCase());
          const idxPed = cabecalho.indexOf("número do pedido faturado") !== -1 ? cabecalho.indexOf("número do pedido faturado") : cabecalho.indexOf("numero do pedido faturado");
          const idxStatus = cabecalho.indexOf("status conferência") !== -1 ? cabecalho.indexOf("status conferência") : cabecalho.indexOf("status conferencia");
          const idxAvaliacao = cabecalho.indexOf("avaliação") !== -1 ? cabecalho.indexOf("avaliação") : cabecalho.indexOf("avaliacao");
          
          let mapPedTotal = {};
          let mapPedConferido = {};
          
          let volTotal = 0;
          let volRealizado = 0;
          let recTotal = 0;
          let recRealizado = 0;
          let recManual = 0;
          
          for (let i = 1; i < dados.length; i++) {
             const linha = dados[i];
             const status = idxStatus !== -1 ? String(linha[idxStatus]).trim() : "";
             const isDivergencia = status.includes("Divergência") || status.includes("Não Encontrado");
             const ped = idxPed !== -1 ? String(linha[idxPed]).trim() : "";
             
             if (!isDivergencia) {
                // Conta o total de itens para cada pedido
                if (ped) {
                    mapPedTotal[ped] = (mapPedTotal[ped] || 0) + 1;
                }
                
                let isBloqueado = false;
                
                if (idxAvaliacao !== -1) {
                    const aval = String(linha[idxAvaliacao]).trim().toLowerCase();
                    if (aval === 'não receber' || aval === 'nao receber') {
                        recTotal++;
                        isBloqueado = true;
                    }
                }

                if (status === "Bloqueado") {
                    recRealizado++;
                    isBloqueado = true;
                }
                
                // Só entra para a volumetria de Bipagem se NÃO estiver bloqueado (Igual à tela de carga!)
                if (!isBloqueado) {
                    volTotal++;
                    
                    if (status === "Conferido" || status === STATUS_RECUSA_MANUAL) {
                        volRealizado++;
                        if (status === STATUS_RECUSA_MANUAL) recManual++;
                        // Conta quantos itens daquele pedido foram bipados
                        if (ped) {
                            mapPedConferido[ped] = (mapPedConferido[ped] || 0) + 1;
                        }
                    }
                }
             }
          }
          
          let pedsTotalCount = Object.keys(mapPedTotal).length;
          let pedsConcluidosCount = 0;
          
          // O pedido só entra para o Dashboard se 100% dos seus itens foram bipados (Igual à tela de carga!)
          for (const p in mapPedTotal) {
              if (mapPedConferido[p] === mapPedTotal[p]) {
                  pedsConcluidosCount++;
              }
          }
          
          chartData.push({
            carga: nome,
            pedidos: pedsTotalCount,
            realPedidos: pedsConcluidosCount,
            volume: volTotal,
            realVolumes: volRealizado,
            recusados: recTotal,
            realRecusados: recRealizado,
            recusadosManual: recManual
          });
        }
      }
    });
    
    chartData.sort((a, b) => {
        const dataA = _parseDataDoNome(a.carga);
        const dataB = _parseDataDoNome(b.carga);
        return dataA - dataB;
    });
    
    return { chartData: chartData }; 
  } catch (e) {
    return { erro: e.message };
  }
}

/**
 * ============================================================================
 * 6. FUNÇÃO DE EXTRAÇÃO AVANÇADA: Importação Automática da Lista Chegada
 * ============================================================================
 */
function importarRelatorioListaChegada() {
  try {
    Logger.log("1. A preparar o motor de carga incremental...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nomeAba = "Base_Lista_Chegada";
    let aba = ss.getSheetByName(nomeAba);
    
    let etiquetasExistentes = new Set();
    let isAppend = false; 
    let lastRow = 0;
    
    if (aba) {
      lastRow = aba.getLastRow();
      if (lastRow > 0) {
        isAppend = true;
        const cabecalhoSheet = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(c => String(c).trim().toLowerCase());
        const idxEtiquetaSheet = cabecalhoSheet.indexOf("etiqueta do produto");
        
        let idxPedidoSheet = cabecalhoSheet.indexOf("num do pedido");
        if (idxPedidoSheet === -1) idxPedidoSheet = cabecalhoSheet.indexOf("número do pedido");
        if (idxPedidoSheet === -1) idxPedidoSheet = cabecalhoSheet.indexOf("numero do pedido");
        
        if (idxEtiquetaSheet !== -1 && lastRow > 1) {
          const dadosAba = aba.getRange(2, 1, lastRow - 1, aba.getLastColumn()).getDisplayValues();
          dadosAba.forEach(r => {
             const etiq = r[idxEtiquetaSheet] ? String(r[idxEtiquetaSheet]).trim() : "";
             const ped = idxPedidoSheet !== -1 && r[idxPedidoSheet] ? String(r[idxPedidoSheet]).trim() : "";
             if (etiq) etiquetasExistentes.add(ped + "|" + etiq); 
          });
        }
      }
    }

    const folderId = '12d0yusktFNnojVKtxlLHGXGUiE-7ivNc';
    const folder = DriveApp.getFolderById(folderId);
    
    const files = folder.searchFiles("title contains 'Lista MSPC'");
    if (!files.hasNext()) throw new Error("Arquivo 'Lista MSPC' não encontrado na pasta.");
    
    const file = files.next();
    const fileId = file.getId();
    const fileSize = file.getSize(); 
    
    const chunkSize = 5 * 1024 * 1024; 
    let start = 0;
    let remainder = "";
    let cabecalhoOriginal = null;
    let indices = [];
    let cabecalhoFiltrado = [];
    let separador = ',';
    let dadosFiltrados = [];
    let idxDataEntrada = -1;
    let idxEtiquetaCsv = -1;
    let idxPedidoCsv = -1;
    
    const colunasAlvo = [
      { titulo: "Etiqueta do produto", variacoes: ["etiqueta do produto"] },
      { titulo: "Num do pedido", variacoes: ["num do pedido", "número do pedido", "numero do pedido"] },
      { titulo: "Data de entrada RMA", variacoes: ["data de entrada rma", "entrada rma"] },
      { titulo: "Valor unitário do produto", variacoes: ["valor unit", "rio do produto", "valor unitario", "valor unitário"] },
      { titulo: "Código do produto", variacoes: ["digo do produto", "código do produto", "codigo do produto", "cdigo do produto"] },
      { titulo: "CD Destino", variacoes: ["cd destino"] }
    ];
    
    function parseCsvLine(text, sep) {
       let result = [];
       let inQuotes = false;
       let value = "";
       for (let i = 0; i < text.length; i++) {
         let char = text[i];
         if (inQuotes) {
           if (char === '"') {
             if (i + 1 < text.length && text[i + 1] === '"') { value += '"'; i++; } 
             else inQuotes = false;
           } else value += char;
         } else {
           if (char === '"') inQuotes = true;
           else if (char === sep) { result.push(value); value = ""; }
           else value += char;
         }
       }
       result.push(value);
       return result;
    }
    
    while (start < fileSize) {
       let end = Math.min(start + chunkSize - 1, fileSize - 1);
       
       let response = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + fileId + "?alt=media", {
         headers: {
           "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
           "Range": "bytes=" + start + "-" + end
         },
         muteHttpExceptions: true
       });
       
       if (response.getResponseCode() !== 200 && response.getResponseCode() !== 206) {
           throw new Error("Falha HTTP no download: " + response.getResponseCode());
       }
       
       let chunkText = response.getContentText("UTF-8");
       let fullText = remainder + chunkText;
       let lastNewlineIdx = fullText.lastIndexOf("\n");
       let textToProcess = fullText;
       
       if (lastNewlineIdx !== -1 && end < fileSize - 1) {
          textToProcess = fullText.substring(0, lastNewlineIdx);
          remainder = fullText.substring(lastNewlineIdx + 1); 
       } else {
          remainder = "";
       }
       
       let lines = textToProcess.split(/\r?\n/);
       
       for (let i = 0; i < lines.length; i++) {
          let lineStr = lines[i].trim();
          if (!lineStr) continue;
          
          if (!cabecalhoOriginal) {
              separador = lineStr.indexOf(';') !== -1 ? ';' : ',';
              cabecalhoOriginal = parseCsvLine(lineStr, separador);
              
              colunasAlvo.forEach(function(alvo) {
                  const idx = cabecalhoOriginal.findIndex(c => {
                     const cLower = String(c).trim().toLowerCase();
                     return alvo.variacoes.some(variacao => cLower.includes(variacao));
                  });
                  if (idx !== -1) {
                    indices.push(idx);
                    cabecalhoFiltrado.push(alvo.titulo);
                    if (alvo.titulo === "Data de entrada RMA") idxDataEntrada = idx;
                    if (alvo.titulo === "Etiqueta do produto") idxEtiquetaCsv = idx;
                    if (alvo.titulo === "Num do pedido") idxPedidoCsv = idx;
                  }
              });
              
              if (indices.length === 0) throw new Error("Nenhuma coluna alvo encontrada no ficheiro.");
              if (!isAppend) dadosFiltrados.push(cabecalhoFiltrado); 
          } else {
              let parsedLine = parseCsvLine(lineStr, separador);
              
              if (isAppend && idxEtiquetaCsv !== -1) {
                  const valorEtiqueta = parsedLine[idxEtiquetaCsv] ? String(parsedLine[idxEtiquetaCsv]).trim() : "";
                  const valorPedido = idxPedidoCsv !== -1 && parsedLine[idxPedidoCsv] ? String(parsedLine[idxPedidoCsv]).trim() : "";
                  
                  const chaveConsulta = valorPedido + "|" + valorEtiqueta;
                  if (etiquetasExistentes.has(chaveConsulta)) continue; 
              }

              if (idxDataEntrada !== -1) {
                  const valorData = parsedLine[idxDataEntrada] ? String(parsedLine[idxDataEntrada]) : "";
                  // REMOVIDA A TRAVA ESTRITA DE ANO. Agora aceita qualquer data válida/não-vazia.
                  if (!valorData || valorData.trim() === "") continue; 
              }

              let novaLinha = indices.map(idx => parsedLine[idx] !== undefined ? parsedLine[idx] : "");
              dadosFiltrados.push(novaLinha);
          }
       }
       start = end + 1; 
    }
    
    const novosItensQtd = isAppend ? dadosFiltrados.length : dadosFiltrados.length - 1;
    
    if (dadosFiltrados.length > 0) {
        if (!aba) aba = ss.insertSheet(nomeAba);
        const sheetChunk = 20000;
        for (let k = 0; k < dadosFiltrados.length; k += sheetChunk) {
            let dChunk = dadosFiltrados.slice(k, k + sheetChunk);
            aba.getRange(lastRow + 1 + k, 1, dChunk.length, dChunk[0].length).setValues(dChunk);
        }
        
        if (!isAppend) {
            aba.getRange(1, 1, 1, cabecalhoFiltrado.length).setFontWeight("bold").setBackground("#f8fafc");
            aba.setFrozenRows(1);
            const totalColunasUsadas = cabecalhoFiltrado.length;
            const totalColunasNaAba = aba.getMaxColumns();
            if (totalColunasNaAba > totalColunasUsadas) {
                aba.deleteColumns(totalColunasUsadas + 1, totalColunasNaAba - totalColunasUsadas);
            }
        }
    }
    
    Logger.log("Sincronizando as cargas com as novas entradas...");
    sincronizarStatusListaChegadaGlobal();
    
    return { sucesso: true, linhasProcessadas: novosItensQtd, abaAtualizada: nomeAba };
    
  } catch (e) {
    return { erro: e.message };
  }
}

/* STREAMING_CHUNK:Configurando funcoes para parse de upload e unificacao com listas externas... */
/**
 * ============================================================================
 * 7. UPLOAD MANUAL DA LISTA DE CHEGADA
 * ============================================================================
 */
function processarUploadListaChegada(matrizDadosInput) {
  try {
    // Tratamento anti-crash para receber JSON pesado
    const matrizDados = typeof matrizDadosInput === 'string' ? JSON.parse(matrizDadosInput) : matrizDadosInput;
    if (!matrizDados || matrizDados.length === 0) throw new Error("O ficheiro enviado está vazio.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nomeAba = "Base_Lista_Chegada";
    let aba = ss.getSheetByName(nomeAba);
    
    let etiquetasExistentes = new Set();
    let isAppend = false; 
    let lastRow = 0;
    
    if (aba) {
      lastRow = aba.getLastRow();
      if (lastRow > 0) {
        isAppend = true;
        const cabecalhoSheet = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(c => String(c).trim().toLowerCase());
        const idxEtiquetaSheet = cabecalhoSheet.indexOf("etiqueta do produto");
        
        let idxPedidoSheet = cabecalhoSheet.indexOf("num do pedido");
        if (idxPedidoSheet === -1) idxPedidoSheet = cabecalhoSheet.indexOf("número do pedido");
        if (idxPedidoSheet === -1) idxPedidoSheet = cabecalhoSheet.indexOf("numero do pedido");
        
        if (idxEtiquetaSheet !== -1 && lastRow > 1) {
          const dadosAba = aba.getRange(2, 1, lastRow - 1, aba.getLastColumn()).getDisplayValues();
          dadosAba.forEach(r => {
             const etiq = r[idxEtiquetaSheet] ? String(r[idxEtiquetaSheet]).trim() : "";
             const ped = idxPedidoSheet !== -1 && r[idxPedidoSheet] ? String(r[idxPedidoSheet]).trim() : "";
             if (etiq) etiquetasExistentes.add(ped + "|" + etiq);
          });
        }
      }
    }

    const cabecalhoOriginal = matrizDados[0];
    const colunasAlvo = [
      { titulo: "Etiqueta do produto", variacoes: ["etiqueta do produto"] },
      { titulo: "Num do pedido", variacoes: ["num do pedido", "número do pedido", "numero do pedido"] },
      { titulo: "Data de entrada RMA", variacoes: ["data de entrada rma", "entrada rma"] },
      { titulo: "Valor unitário do produto", variacoes: ["valor unit", "rio do produto", "valor unitario", "valor unitário"] },
      { titulo: "Código do produto", variacoes: ["digo do produto", "código do produto", "codigo do produto", "cdigo do produto"] },
      { titulo: "CD Destino", variacoes: ["cd destino"] }
    ];
    
    let indices = [];
    let cabecalhoFiltrado = [];
    let idxDataEntrada = -1;
    let idxEtiquetaCsv = -1;
    let idxPedidoCsv = -1;

    colunasAlvo.forEach(function(alvo) {
        const idx = cabecalhoOriginal.findIndex(c => {
            const cLower = String(c).trim().toLowerCase();
            return alvo.variacoes.some(variacao => cLower.includes(variacao));
        });
        
        if (idx !== -1) {
          indices.push(idx);
          cabecalhoFiltrado.push(alvo.titulo);
          if (alvo.titulo === "Data de entrada RMA") idxDataEntrada = idx;
          if (alvo.titulo === "Etiqueta do produto") idxEtiquetaCsv = idx;
          if (alvo.titulo === "Num do pedido") idxPedidoCsv = idx;
        }
    });

    if (indices.length === 0) throw new Error("Nenhuma das colunas alvo foi encontrada neste ficheiro.");

    let dadosFiltrados = [];
    if (!isAppend) dadosFiltrados.push(cabecalhoFiltrado); 

    for (let i = 1; i < matrizDados.length; i++) {
        let linha = matrizDados[i];
        if (!linha || linha.join('').trim() === '') continue;

        if (isAppend && idxEtiquetaCsv !== -1) {
            const valorEtiqueta = linha[idxEtiquetaCsv] ? String(linha[idxEtiquetaCsv]).trim() : "";
            const valorPedido = idxPedidoCsv !== -1 && linha[idxPedidoCsv] ? String(linha[idxPedidoCsv]).trim() : "";
            
            const chaveConsulta = valorPedido + "|" + valorEtiqueta;
            if (etiquetasExistentes.has(chaveConsulta)) continue; 
        }

        if (idxDataEntrada !== -1) {
            const valorData = linha[idxDataEntrada] ? String(linha[idxDataEntrada]) : "";
            // REMOVIDA A TRAVA ESTRITA DE ANO. Agora aceita qualquer data válida/não-vazia.
            if (!valorData || valorData.trim() === "") continue; 
        }

        let novaLinha = indices.map(idx => linha[idx] !== undefined ? linha[idx] : "");
        dadosFiltrados.push(novaLinha);
    }

    const novosItensQtd = isAppend ? dadosFiltrados.length : dadosFiltrados.length - 1;

    if (dadosFiltrados.length > 0) {
        if (!aba) aba = ss.insertSheet(nomeAba);
        const sheetChunk = 20000;
        for (let k = 0; k < dadosFiltrados.length; k += sheetChunk) {
            let dChunk = dadosFiltrados.slice(k, k + sheetChunk);
            aba.getRange(lastRow + 1 + k, 1, dChunk.length, dChunk[0].length).setValues(dChunk);
        }

        if (!isAppend) {
            aba.getRange(1, 1, 1, cabecalhoFiltrado.length).setFontWeight("bold").setBackground("#f8fafc");
            aba.setFrozenRows(1);
            const totalColunasUsadas = cabecalhoFiltrado.length;
            const totalColunasNaAba = aba.getMaxColumns();
            if (totalColunasNaAba > totalColunasUsadas) {
                aba.deleteColumns(totalColunasUsadas + 1, totalColunasNaAba - totalColunasUsadas);
            }
        }
    }
    
    sincronizarStatusListaChegadaGlobal();

    return { sucesso: true, linhasProcessadas: novosItensQtd, abaAtualizada: nomeAba };
    
  } catch (e) {
    return { erro: e.message };
  }
}

/* STREAMING_CHUNK:Finalizando integracoes sistemicas e historico JSON... */
/**
 * ============================================================================
 * 8. MOTOR DE SINCRONIZAÇÃO SISTÉMICA (COM CHAVE COMPOSTA)
 * ============================================================================
 */

function _obterSetGlobalListaChegada() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaLista = ss.getSheetByName("Base_Lista_Chegada");
    let setListaChegada = new Set();
    if (!abaLista) return setListaChegada;

    const lastRowLista = abaLista.getLastRow();
    if (lastRowLista < 2) return setListaChegada;

    const dadosLista = abaLista.getDataRange().getValues();
    const cabecalhoLista = dadosLista[0].map(c => String(c).trim().toLowerCase());
    
    let idxEtiqLista = cabecalhoLista.indexOf("etiqueta do produto");
    if (idxEtiqLista === -1) idxEtiqLista = cabecalhoLista.indexOf("mercadoria código");
    
    let idxPedLista = cabecalhoLista.indexOf("num do pedido");
    if (idxPedLista === -1) idxPedLista = cabecalhoLista.indexOf("número do pedido");

    if (idxEtiqLista === -1) return setListaChegada;

    for (let i = 1; i < dadosLista.length; i++) {
        const etiq = String(dadosLista[i][idxEtiqLista]).trim();
        const ped = idxPedLista !== -1 && dadosLista[i][idxPedLista] ? String(dadosLista[i][idxPedLista]).trim() : "";
        
        if (etiq) {
            setListaChegada.add(ped + "|" + etiq);
        }
    }
    return setListaChegada;
}

function _sincronizarAbaComSetListaChegada(aba, setListaChegada) {
    const lastRow = aba.getLastRow();
    const lastCol = aba.getLastColumn();
    if (lastRow < 2) return [];

    const dados = aba.getDataRange().getValues();
    const cabecalho = dados[0].map(c => String(c).trim().toLowerCase());

    let idxMerc = cabecalho.indexOf("mercadoria código");
    if (idxMerc === -1) idxMerc = cabecalho.indexOf("mercadoria codigo");
    
    let idxPed = cabecalho.indexOf("número do pedido faturado");
    if (idxPed === -1) idxPed = cabecalho.indexOf("numero do pedido faturado");
    
    if (idxMerc === -1) return [];

    let idxStatusLista = cabecalho.indexOf("status lista chegada");
    if (idxStatusLista === -1) {
        idxStatusLista = aba.getLastColumn(); 
        aba.getRange(1, idxStatusLista + 1).setValue("Status Lista Chegada").setFontWeight("bold");
        for(let r=0; r<dados.length; r++) dados[r].push(r===0 ? "Status Lista Chegada" : "");
    }

    let alterou = false;
    let arrayUpdate = [];
    let recebidos = [];

    for (let i = 1; i < dados.length; i++) {
        const merc = String(dados[i][idxMerc]).trim();
        const ped = idxPed !== -1 && dados[i][idxPed] ? String(dados[i][idxPed]).trim() : "";
        const chave = ped + "|" + merc; 
        
        let statusAtual = String(dados[i][idxStatusLista]).trim();

        if (statusAtual !== "Recebido Lista" && setListaChegada.has(chave)) {
            dados[i][idxStatusLista] = "Recebido Lista";
            statusAtual = "Recebido Lista";
            alterou = true;
        }
        
        if (statusAtual === "Recebido Lista") {
            recebidos.push(chave);
        }
        
        arrayUpdate.push([dados[i][idxStatusLista]]);
    }

    if (alterou) {
        aba.getRange(2, idxStatusLista + 1, arrayUpdate.length, 1).setValues(arrayUpdate);
    }
    
    return recebidos;
}

function sincronizarStatusListaChegadaGlobal() {
    const setListaChegada = _obterSetGlobalListaChegada();
    if (setListaChegada.size === 0) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abas = ss.getSheets();
    abas.forEach(aba => {
        const nome = aba.getName();
        if (/^\d{2}-\d{2}-\d{4}/.test(nome)) {
            // TRAVA REMOVIDA: Sincroniza a aba com a Lista de Chegada, 
            // independentemente de a carga estar finalizada ou em aberto.
            _sincronizarAbaComSetListaChegada(aba, setListaChegada);
        }
    });
    
    // GATILHO AUTO-CURA
    _autoCurarHistorico(setListaChegada);
}

function getTagsRecebidasPorAba(nomeAba) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return [];
    
    // NOVO GATILHO (Botão Detalhes): Sincroniza a aba em tempo real 
    // com a "Base_Lista_Chegada" imediatamente antes de ler e devolver as tags.
    const setListaChegada = _obterSetGlobalListaChegada();
    if (setListaChegada.size > 0) {
        _sincronizarAbaComSetListaChegada(aba, setListaChegada);
    }
    
    const lastRow = aba.getLastRow();
    if (lastRow < 2) return [];
    
    const dados = aba.getDataRange().getValues();
    const cab = dados[0].map(c => String(c).trim().toLowerCase());
    
    let idxMerc = cab.indexOf("mercadoria código");
    if (idxMerc === -1) idxMerc = cab.indexOf("mercadoria codigo");
    let idxPed = cab.indexOf("número do pedido faturado");
    if (idxPed === -1) idxPed = cab.indexOf("numero do pedido faturado");
    const idxLista = cab.indexOf("status lista chegada");
    
    if (idxMerc === -1 || idxLista === -1) return [];
    
    const recebidos = [];
    for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][idxLista]).trim() === "Recebido Lista") {
            const merc = String(dados[i][idxMerc]).trim();
            const ped = idxPed !== -1 && dados[i][idxPed] ? String(dados[i][idxPed]).trim() : "";
            recebidos.push(ped + "|" + merc); 
        }
    }
    return recebidos;
}

/**
 * ============================================================================
 * 9. MOTOR DE SINCRONIZAÇÃO E CÉREBRO HISTÓRICO (JSON)
 * ============================================================================
 */

function _obterAbaHistorico(ss) {
    let aba = ss.getSheetByName("Base_Historico_Eventos");
    if (!aba) {
        aba = ss.insertSheet("Base_Historico_Eventos");
        aba.getRange("A1:B1").setValues([["ID Único (Pedido|Etiqueta)", "Currículo JSON"]]).setFontWeight("bold");
        aba.setFrozenRows(1);
    }
    return aba;
}

function _atualizarHistoricoNaFinalizacao(nomeAba, aba, setListaChegada) {
    const lastRow = aba.getLastRow();
    if (lastRow < 2) return;
    const dados = aba.getDataRange().getValues();
    const cabecalho = dados[0].map(c => String(c).trim().toLowerCase());
    
    let idxMerc = cabecalho.indexOf("mercadoria código");
    if (idxMerc === -1) idxMerc = cabecalho.indexOf("mercadoria codigo");
    let idxPed = cabecalho.indexOf("número do pedido faturado");
    if (idxPed === -1) idxPed = cabecalho.indexOf("numero do pedido faturado");
    let idxStatus = cabecalho.indexOf("status conferência");
    if (idxStatus === -1) idxStatus = cabecalho.indexOf("status conferencia");
    let idxAval = cabecalho.indexOf("avaliação");
    if (idxAval === -1) idxAval = cabecalho.indexOf("avaliacao");
    
    if (idxMerc === -1) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaHist = _obterAbaHistorico(ss);
    const lastRowHist = abaHist.getLastRow();
    
    let mapHist = {};
    if (lastRowHist > 1) {
        const dadosHist = abaHist.getRange(2, 1, lastRowHist - 1, 2).getValues();
        for (let i = 0; i < dadosHist.length; i++) {
            mapHist[String(dadosHist[i][0])] = { row: i + 2, jsonStr: dadosHist[i][1] };
        }
    }

    let updates = [];
    let appends = [];

    // Mesmo motor do Web App: o bloqueio vale para o pedido inteiro (os volumes não bipados de um
    // pedido bloqueado ficam no histórico como "Bloqueado", para o alerta de reincidente).
    const julgamento = julgarCargaMotor_({
        carga: nomeAba,
        cabecalho: dados[0],
        linhas: dados.slice(1).map(l => l.map(v => String(v))),
        recebidosLista: Array.from(setListaChegada)
    });
    const SIT = julgamento.SIT;

    for (const item of julgamento.itens) {
        const merc = item.etiqueta;
        if (!merc) continue;
        const chave = item.pedido + "|" + merc;

        const isBloqueado = item.situacao === SIT.BLOQ_RECUSAR || item.situacao === SIT.BLOQ_IGNORADO;
        const isRecusaManual = item.situacao === SIT.RECUSADO;
        const isFalta = item.situacao === SIT.FALTA;

        if (isBloqueado || isRecusaManual || isFalta) {
            const novoStatus = isBloqueado ? "Bloqueado" : (isRecusaManual ? STATUS_RECUSA_MANUAL : "Falta Confirmada");
            
            let obj;
            if (mapHist[chave]) {
                try { obj = JSON.parse(mapHist[chave].jsonStr); } catch(e) { obj = { eventos: [] }; }
                obj.ultimoStatus = novoStatus;
                obj.ultimaCarga = nomeAba;
                if (!obj.eventos) obj.eventos = [];
                obj.eventos.push({ carga: nomeAba, status: novoStatus });
                
                updates.push({ row: mapHist[chave].row, val: JSON.stringify(obj) });
            } else {
                obj = {
                    ultimoStatus: novoStatus,
                    ultimaCarga: nomeAba,
                    eventos: [{ carga: nomeAba, status: novoStatus }]
                };
                appends.push([chave, JSON.stringify(obj)]);
            }
        }
    }

    if (updates.length > 0) {
        updates.forEach(u => abaHist.getRange(u.row, 2).setValue(u.val));
    }
    if (appends.length > 0) {
        abaHist.getRange(abaHist.getLastRow() + 1, 1, appends.length, 2).setValues(appends);
    }
}

function _autoCurarHistorico(setListaChegada) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaHist = ss.getSheetByName("Base_Historico_Eventos");
    if (!abaHist) return;
    
    const lastRow = abaHist.getLastRow();
    if (lastRow < 2) return;

    const dados = abaHist.getRange(2, 1, lastRow - 1, 2).getValues();
    let alterou = false;

    for (let i = 0; i < dados.length; i++) {
        const chave = String(dados[i][0]);
        if (setListaChegada.has(chave)) {
            try {
                let obj = JSON.parse(dados[i][1]);
                if (obj.ultimoStatus === "Falta Confirmada") {
                    obj.ultimoStatus = "Falta Resolvida (Sistémica)";
                    dados[i][1] = JSON.stringify(obj);
                    alterou = true;
                }
            } catch(e) {}
        }
    }

    if (alterou) {
        abaHist.getRange(2, 1, lastRow - 1, 2).setValues(dados);
    }
}

function obterHistoricoEventosFrontend() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("Base_Historico_Eventos");
    let result = {};
    if (!aba || aba.getLastRow() < 2) return result;
    
    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues();
    
    for (let i = 0; i < dados.length; i++) {
        const id = String(dados[i][0]);
        const jsonStr = String(dados[i][1]);
        if (id && jsonStr) {
            try {
                const obj = JSON.parse(jsonStr);
                const partes = id.split("|");
                const etiq = partes.length > 1 ? partes[1] : partes[0];
                result[etiq] = obj; 
            } catch(e) {}
        }
    }
    return result;
}

/**
 * ============================================================================
 * 10. MOTOR DE EXTRAÇÃO: "MONTE O SEU - CONSOLIDADO"
 * ============================================================================
 */
function _obterMapaMonteOSeu() {
    let mapa = {};
    try {
        const folderId = '1_Q8WflHdr1i3oMrMaf079zyP2pvu0hz6';
        const folder = DriveApp.getFolderById(folderId);
        const files = folder.searchFiles("title contains 'Monte o Seu - Consolidado'");
        
        if (!files.hasNext()) return mapa; 
        
        const file = files.next();
        const fileId = file.getId();
        const fileSize = file.getSize();
        
        const chunkSize = 5 * 1024 * 1024;
        let start = 0;
        let remainder = "";
        let cabecalhoOriginal = null;
        let idxPedido = -1;
        let idxKit = -1;
        let separador = ',';

        function parseCsvLine(text, sep) {
           let result = []; let inQuotes = false; let value = "";
           for (let i = 0; i < text.length; i++) {
             let char = text[i];
             if (inQuotes) {
               if (char === '"') {
                 if (i + 1 < text.length && text[i + 1] === '"') { value += '"'; i++; } else inQuotes = false;
               } else value += char;
             } else {
               if (char === '"') inQuotes = true;
               else if (char === sep) { result.push(value); value = ""; }
               else value += char;
             }
           }
           result.push(value);
           return result;
        }

        while (start < fileSize) {
           let end = Math.min(start + chunkSize - 1, fileSize - 1);
           let response = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + fileId + "?alt=media", {
             headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken(), "Range": "bytes=" + start + "-" + end },
             muteHttpExceptions: true
           });
           
           if (response.getResponseCode() !== 200 && response.getResponseCode() !== 206) break;
           
           let chunkText = response.getContentText("UTF-8");
           let fullText = remainder + chunkText;
           let lastNewlineIdx = fullText.lastIndexOf("\n");
           let textToProcess = fullText;
           
           if (lastNewlineIdx !== -1 && end < fileSize - 1) {
              textToProcess = fullText.substring(0, lastNewlineIdx);
              remainder = fullText.substring(lastNewlineIdx + 1); 
           } else {
              remainder = "";
           }
           
           let lines = textToProcess.split(/\r?\n/);
           
           for (let i = 0; i < lines.length; i++) {
              let lineStr = lines[i].trim();
              if (!lineStr) continue;
              
              if (!cabecalhoOriginal) {
                  separador = lineStr.indexOf(';') !== -1 ? ';' : ',';
                  cabecalhoOriginal = parseCsvLine(lineStr, separador);
                  const cbLower = cabecalhoOriginal.map(c => String(c).trim().toLowerCase());
                  
                  idxPedido = cbLower.indexOf("pedido");
                  idxKit = cbLower.findIndex(c => c.includes("pertence ao kit") || c.includes("tipo"));
                  
                  if (idxPedido === -1 || idxKit === -1) return mapa; 
              } else {
                  let parsedLine = parseCsvLine(lineStr, separador);
                  let ped = parsedLine[idxPedido] ? String(parsedLine[idxPedido]).trim() : "";
                  let kit = parsedLine[idxKit] ? String(parsedLine[idxKit]).trim() : "";
                  
                  if (ped) mapa[ped] = kit;
              }
           }
           start = end + 1; 
        }
    } catch(e) {
        Logger.log("Erro no cruzamento Monte o Seu: " + e.message);
    }
    return mapa;
}

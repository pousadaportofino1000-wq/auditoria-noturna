/* Pousada Estoque & Gastos - Apps Script (V8) */

const SHEETS = {
  PRODUTOS: "Produtos",
  NOTAS: "Notas",
  ITENS: "Itens_Nota",
  MOV: "Movimentacoes",
  INVENTARIOS: "Inventarios",
  INV_ITENS: "Inventario_Itens",
  ESTOQUE: "Estoque_Atual",
  CONSUMO: "Consumo_Semanal",
  PAINEL: "Painel_Gastos",
};

const LISTS = {
  categorias: ["Padaria", "Laticinios", "Frios", "Bebidas", "Hortifruti", "Secos", "Outros"],
  unidades: ["un", "pct", "cx", "kg", "g", "l", "ml"],
  ativo: ["SIM", "NAO"],
  formasPagamento: ["Dinheiro", "Cartao", "PIX", "Boleto", "Outros"],
  tiposMov: ["ENTRADA_COMPRA", "AJUSTE_INVENTARIO_POS", "AJUSTE_INVENTARIO_NEG"],
};

const TZ = "America/Sao_Paulo";
const CURRENCY_FORMAT = "R$ #,##0.00";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Pousada Estoque & Gastos")
    .addItem("Setup inicial / Recriar estrutura", "setupAll")
    .addItem("Lancar compra (nota)", "showCompraDialog")
    .addItem("Inventario semanal (contagem)", "showInventarioDialog")
    .addItem("Atualizar relatorios", "refreshReports")
    .addItem("Abrir painel de gastos", "openGastosSidebar")
    .addToUi();
}

function setupAll() {
  const ss = SpreadsheetApp.getActive();
  setupProdutos_(ss);
  setupNotas_(ss);
  setupItensNota_(ss);
  setupMovimentacoes_(ss);
  setupInventarios_(ss);
  setupInventarioItens_(ss);
  setupEstoqueAtual_(ss);
  setupConsumoSemanal_(ss);
  setupPainelGastos_(ss);
  refreshDynamicValidations_();
  toast_("Setup concluido.");
}

function showCompraDialog() {
  const html = HtmlService.createHtmlOutputFromFile("CompraDialog")
    .setWidth(720)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, "Lancar compra (nota)");
}

function showInventarioDialog() {
  const html = HtmlService.createHtmlOutputFromFile("InventarioDialog")
    .setWidth(760)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Inventario semanal");
}

function openGastosSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("SidebarGastos")
    .setTitle("Painel de Gastos");
  SpreadsheetApp.getUi().showSidebar(html);
}

function openPainelGastosSheet() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEETS.PAINEL);
  if (sh) ss.setActiveSheet(sh);
}

function refreshReports() {
  const ss = SpreadsheetApp.getActive();
  setupEstoqueAtual_(ss);
  rebuildConsumoSemanal_(ss);
  refreshPainelGastos_();
  toast_("Relatorios atualizados.");
}

function refreshPainelGastos() {
  refreshPainelGastos_();
}

/* ==============================
 *  Setup de abas
 * ============================== */
function setupProdutos_(ss) {
  const sh = ensureSheet_(ss, SHEETS.PRODUTOS);
  const headers = ["ID/SKU", "Produto", "Categoria", "Unidade", "Estoque_Minimo", "Ativo"];
  setHeader_(sh, headers);
  sh.setFrozenRows(1);
  setColumnWidths_(sh, [140, 220, 160, 100, 140, 90]);
  applyValidationList_(sh, 3, LISTS.categorias);
  applyValidationList_(sh, 4, LISTS.unidades);
  applyValidationList_(sh, 6, LISTS.ativo);
  protectHeader_(sh, headers.length);
}

function setupNotas_(ss) {
  const sh = ensureSheet_(ss, SHEETS.NOTAS);
  const headers = ["Data_Compra", "Fornecedor", "Numero_Documento", "Forma_Pagamento", "Total_Nota", "Observacoes", "ID_Nota"];
  setHeader_(sh, headers);
  sh.setFrozenRows(1);
  setColumnWidths_(sh, [120, 200, 160, 150, 120, 260, 200]);
  applyValidationList_(sh, 4, LISTS.formasPagamento);
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat("dd/MM/yyyy");
  sh.getRange(2, 5, sh.getMaxRows() - 1, 1).setNumberFormat(CURRENCY_FORMAT);
  protectHeader_(sh, headers.length);
}

function setupItensNota_(ss) {
  const sh = ensureSheet_(ss, SHEETS.ITENS);
  const headers = ["ID_Nota", "Data_Compra", "Produto", "Categoria", "Unidade", "Quantidade", "Preco_Unitario", "Total_Linha"];
  setHeader_(sh, headers);
  sh.setFrozenRows(1);
  setColumnWidths_(sh, [180, 120, 220, 140, 90, 100, 120, 120]);
  // Validacoes dinamicas sao aplicadas em refreshDynamicValidations_

  // Formulas
  const dataFormula = `=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,{${SHEETS.NOTAS}!$G$2:$G,${SHEETS.NOTAS}!$A$2:$A},2,FALSE),"")))`;
  const catFormula = `=ARRAYFORMULA(IF(C2:C="","",IFERROR(VLOOKUP(C2:C,{${SHEETS.PRODUTOS}!$B$2:$B,${SHEETS.PRODUTOS}!$C$2:$C},2,FALSE),"")))`;
  const unFormula = `=ARRAYFORMULA(IF(C2:C="","",IFERROR(VLOOKUP(C2:C,{${SHEETS.PRODUTOS}!$B$2:$B,${SHEETS.PRODUTOS}!$D$2:$D},2,FALSE),"")))`;
  const totalFormula = `=ARRAYFORMULA(IF(F2:F="","",F2:F*G2:G))`;
  sh.getRange("B2").setFormula(dataFormula);
  sh.getRange("D2").setFormula(catFormula);
  sh.getRange("E2").setFormula(unFormula);
  sh.getRange("H2").setFormula(totalFormula);

  sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat("dd/MM/yyyy");
  sh.getRange(2, 6, sh.getMaxRows() - 1, 1).setNumberFormat("0.00");
  sh.getRange(2, 7, sh.getMaxRows() - 1, 2).setNumberFormat(CURRENCY_FORMAT);
  protectHeader_(sh, headers.length);
}

function setupMovimentacoes_(ss) {
  const sh = ensureSheet_(ss, SHEETS.MOV);
  const headers = ["Timestamp", "Data", "Tipo", "Referencia", "Produto", "Quantidade", "Custo_Unitario", "Valor", "Observacoes"];
  setHeader_(sh, headers);
  sh.setFrozenRows(1);
  setColumnWidths_(sh, [160, 110, 200, 180, 220, 110, 120, 120, 260]);
  applyValidationList_(sh, 3, LISTS.tiposMov);
  sh.getRange(2, 1, sh.getMaxRows() - 1, 2).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sh.getRange(2, 6, sh.getMaxRows() - 1, 1).setNumberFormat("0.00");
  sh.getRange(2, 7, sh.getMaxRows() - 1, 2).setNumberFormat(CURRENCY_FORMAT);
  protectHeader_(sh, headers.length);
}

function setupInventarios_(ss) {
  const sh = ensureSheet_(ss, SHEETS.INVENTARIOS);
  const headers = ["Data_Inventario", "Responsavel", "Observacoes", "ID_Inventario", "Inventario_Anterior"];
  setHeader_(sh, headers);
  sh.setFrozenRows(1);
  setColumnWidths_(sh, [140, 180, 280, 200, 200]);
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat("dd/MM/yyyy");
  protectHeader_(sh, headers.length);
}

function setupInventarioItens_(ss) {
  const sh = ensureSheet_(ss, SHEETS.INV_ITENS);
  const headers = ["ID_Inventario", "Produto", "Unidade", "Estoque_Sistema", "Estoque_Contado", "Diferenca", "Ajuste_Gerado"];
  setHeader_(sh, headers);
  sh.setFrozenRows(1);
  setColumnWidths_(sh, [200, 220, 90, 130, 130, 110, 120]);
  // Validacoes dinamicas sao aplicadas em refreshDynamicValidations_
  protectHeader_(sh, headers.length);
}

function setupEstoqueAtual_(ss) {
  const sh = ensureSheet_(ss, SHEETS.ESTOQUE);
  const headers = ["Produto", "Categoria", "Unidade", "Estoque_Atual", "Estoque_Minimo", "Status"];
  setHeader_(sh, headers);
  sh.setFrozenRows(1);
  setColumnWidths_(sh, [240, 150, 90, 140, 140, 100]);

  const produtoFormula = `=FILTER(${SHEETS.PRODUTOS}!$B$2:$B,${SHEETS.PRODUTOS}!$F$2:$F="SIM")`;
  const catFormula = `=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,{${SHEETS.PRODUTOS}!$B$2:$B,${SHEETS.PRODUTOS}!$C$2:$C},2,FALSE),"")))`;
  const unFormula = `=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,{${SHEETS.PRODUTOS}!$B$2:$B,${SHEETS.PRODUTOS}!$D$2:$D},2,FALSE),"")))`;
  const minFormula = `=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,{${SHEETS.PRODUTOS}!$B$2:$B,${SHEETS.PRODUTOS}!$E$2:$E},2,FALSE),"")))`;
  const estoqueFormula = `=ARRAYFORMULA(IF(A2:A="","",SUMIF(${SHEETS.MOV}!$E:$E,A2:A,${SHEETS.MOV}!$F:$F)))`;
  const statusFormula = `=ARRAYFORMULA(IF(A2:A="","",IF(D2:D<=E2:E,"BAIXO","OK")))`;

  sh.getRange("A2").setFormula(produtoFormula);
  sh.getRange("B2").setFormula(catFormula);
  sh.getRange("C2").setFormula(unFormula);
  sh.getRange("E2").setFormula(minFormula);
  sh.getRange("D2").setFormula(estoqueFormula);
  sh.getRange("F2").setFormula(statusFormula);

  sh.getRange(2, 4, sh.getMaxRows() - 1, 2).setNumberFormat("0.00");

  // Condicional BAIXO
  const rules = sh.getConditionalFormatRules();
  const newRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("BAIXO")
    .setBackground("#F4CCCC")
    .setRanges([sh.getRange(2, 6, sh.getMaxRows() - 1, 1)])
    .build();
  rules.push(newRule);
  sh.setConditionalFormatRules(rules);
  protectSheet_(sh, [sh.getRange("A2:F")]);
}

function setupConsumoSemanal_(ss) {
  const sh = ensureSheet_(ss, SHEETS.CONSUMO);
  const headers = ["ID_Inventario", "Data_Inventario", "Produto", "Consumo_No_Periodo", "Custo_Unit_Medio", "Valor_Consumo"];
  setHeader_(sh, headers);
  sh.setFrozenRows(1);
  setColumnWidths_(sh, [200, 120, 220, 160, 150, 150]);
  sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat("dd/MM/yyyy");
  sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setNumberFormat("0.00");
  sh.getRange(2, 5, sh.getMaxRows() - 1, 2).setNumberFormat(CURRENCY_FORMAT);
  protectSheet_(sh, []);
}

function setupPainelGastos_(ss) {
  const sh = ensureSheet_(ss, SHEETS.PAINEL);
  sh.clear();

  sh.getRange("A1").setValue("Painel de Gastos - Filtros");
  sh.getRange("A3").setValue("Data_Inicial");
  sh.getRange("C3").setValue("Data_Final");
  sh.getRange("E3").setValue("Fornecedor");
  sh.getRange("G3").setValue("Categoria");
  sh.getRange("I3").setValue("Forma_Pagamento");
  sh.getRange("K3").setValue("Numero_Documento");
  sh.getRange("A6").setValue("Gastos Filtrados");
  sh.getRange("A20").setValue("Gastos por Mes");
  sh.getRange("E20").setValue("Gastos por Fornecedor");
  sh.getRange("I20").setValue("Gastos por Categoria");

  sh.getRange("A1:L1").setFontWeight("bold");
  sh.getRange("A3:L3").setFontWeight("bold");
  sh.getRange("A6").setFontWeight("bold");

  sh.getRange("B3").setNumberFormat("dd/MM/yyyy");
  sh.getRange("D3").setNumberFormat("dd/MM/yyyy");
  // Validacoes dinamicas sao aplicadas em refreshDynamicValidations_

  sh.setFrozenRows(1);
  setColumnWidths_(sh, [160, 120, 160, 120, 160, 140, 140, 120, 140, 140, 160, 140]);

  protectSheet_(sh, [sh.getRange("B3"), sh.getRange("D3"), sh.getRange("F3"), sh.getRange("H3"), sh.getRange("J3"), sh.getRange("L3")]);
}

/* ==============================
 *  Compra (Nota)
 * ============================== */
function getCompraFormData() {
  return {
    products: getActiveProducts_(),
    formas: LISTS.formasPagamento,
  };
}

function saveCompra(payload) {
  const ss = SpreadsheetApp.getActive();
  const notas = ss.getSheetByName(SHEETS.NOTAS);
  const itens = ss.getSheetByName(SHEETS.ITENS);
  const mov = ss.getSheetByName(SHEETS.MOV);

  if (!payload || !payload.cabecalho || !payload.itens || !payload.itens.length) {
    throw new Error("Informe cabecalho e ao menos um item.");
  }

  const dataCompra = coerceDate_(payload.cabecalho.dataCompra);
  const fornecedor = String(payload.cabecalho.fornecedor || "").trim();
  const numero = String(payload.cabecalho.numeroDocumento || "").trim();
  const forma = String(payload.cabecalho.formaPagamento || "").trim();
  const totalNota = Number(payload.cabecalho.totalNota || 0);
  const obs = String(payload.cabecalho.observacoes || "").trim();

  if (!dataCompra || !fornecedor || !numero || !forma) {
    throw new Error("Campos obrigatorios: data, fornecedor, numero, forma de pagamento.");
  }

  if (notaDuplicada_(notas, dataCompra, fornecedor, numero)) {
    throw new Error("Nota duplicada (Data + Fornecedor + Numero).");
  }

  const idNota = generateNotaId_(dataCompra, numero);

  // Notas
  appendRows_(notas, [[dataCompra, fornecedor, numero, forma, totalNota, obs, idNota]]);

  // Itens_Nota
  const itemRows = payload.itens.map(it => {
    const produto = String(it.produto || "").trim();
    const qtd = Number(it.quantidade || 0);
    const preco = Number(it.precoUnitario || 0);
    if (!produto || !qtd || !preco) throw new Error("Itens precisam produto, quantidade e preco.");
    return [idNota, "", produto, "", "", qtd, preco, ""];
  });
  appendRows_(itens, itemRows);

  // Movimentacoes
  const now = new Date();
  const movRows = payload.itens.map(it => {
    const produto = String(it.produto || "").trim();
    const qtd = Number(it.quantidade || 0);
    const preco = Number(it.precoUnitario || 0);
    const valor = qtd * preco;
    return [now, dataCompra, "ENTRADA_COMPRA", idNota, produto, qtd, preco, valor, fornecedor];
  });
  appendRows_(mov, movRows);

  refreshDynamicValidations_();
  return { ok: true, idNota };
}

/* ==============================
 *  Inventario
 * ============================== */
function getInventarioFormData() {
  return {
    products: getActiveProducts_(),
  };
}

function saveInventario(payload) {
  const ss = SpreadsheetApp.getActive();
  const shInv = ss.getSheetByName(SHEETS.INVENTARIOS);
  const shItens = ss.getSheetByName(SHEETS.INV_ITENS);
  const shMov = ss.getSheetByName(SHEETS.MOV);

  if (!payload || !payload.dataInventario || !payload.responsavel || !payload.itens) {
    throw new Error("Informe data, responsavel e itens.");
  }

  const dataInv = coerceDate_(payload.dataInventario);
  if (!dataInv) throw new Error("Data de inventario invalida.");

  const responsavel = String(payload.responsavel || "").trim();
  const obs = String(payload.observacoes || "").trim();

  const prev = findInventarioAnterior_(shInv, dataInv);
  const idInv = generateInventarioId_(dataInv);

  appendRows_(shInv, [[dataInv, responsavel, obs, idInv, prev ? prev.id : ""]]);

  const movValues = getMovValues_(shMov);
  const stockMap = computeStockMapAtDate_(movValues, dataInv);
  const avgCostMap = computeAvgCostMapAtDate_(movValues, dataInv);

  const itemRows = [];
  const movRows = [];

  payload.itens.forEach(it => {
    const produto = String(it.produto || "").trim();
    if (!produto) return;
    const unidade = String(it.unidade || "").trim();
    const contado = Number(it.estoqueContado || 0);
    const sistema = Number(stockMap[produto] || 0);
    const diff = contado - sistema;
    const ajuste = diff !== 0 ? "SIM" : "NAO";

    itemRows.push([idInv, produto, unidade, sistema, contado, diff, ajuste]);

    if (diff !== 0) {
      const tipo = diff > 0 ? "AJUSTE_INVENTARIO_POS" : "AJUSTE_INVENTARIO_NEG";
      const custo = avgCostMap[produto] != null ? avgCostMap[produto] : 0;
      const valor = diff * custo;
      movRows.push([new Date(), dataInv, tipo, idInv, produto, diff, custo, valor, "Ajuste inventario"]);
    }
  });

  if (itemRows.length) appendRows_(shItens, itemRows);
  if (movRows.length) appendRows_(shMov, movRows);

  updateConsumoSemanalForInventory_(ss, idInv, dataInv, prev ? prev.id : "", prev ? prev.data : null);

  refreshDynamicValidations_();
  return { ok: true, idInventario: idInv };
}

/* ==============================
 *  Painel de Gastos
 * ============================== */
function refreshPainelGastos_() {
  const ss = SpreadsheetApp.getActive();
  const painel = ss.getSheetByName(SHEETS.PAINEL);
  if (!painel) return;

  const filtros = {
    dataIni: coerceDate_(painel.getRange("B3").getValue()),
    dataFim: coerceDate_(painel.getRange("D3").getValue()),
    fornecedor: String(painel.getRange("F3").getValue() || "").trim(),
    categoria: String(painel.getRange("H3").getValue() || "").trim(),
    forma: String(painel.getRange("J3").getValue() || "").trim(),
    numero: String(painel.getRange("L3").getValue() || "").trim(),
  };

  const notas = getNotasData_();
  const itens = getItensNotaData_();

  const rows = [];
  itens.forEach(it => {
    const nota = notas[it.idNota];
    if (!nota) return;
    if (filtros.dataIni && nota.data < filtros.dataIni) return;
    if (filtros.dataFim && nota.data > filtros.dataFim) return;
    if (filtros.fornecedor && nota.fornecedor !== filtros.fornecedor) return;
    if (filtros.categoria && it.categoria !== filtros.categoria) return;
    if (filtros.forma && nota.forma !== filtros.forma) return;
    if (filtros.numero && nota.numero !== filtros.numero) return;

    rows.push([
      nota.data, nota.fornecedor, nota.numero, nota.forma,
      it.produto, it.categoria, it.quantidade, it.totalLinha, it.idNota
    ]);
  });

  // Limpa area de resultados
  painel.getRange(8, 1, Math.max(1, painel.getMaxRows() - 7), 9).clearContent();
  painel.getRange("A8:I8").setValues([["Data", "Fornecedor", "Numero", "Forma", "Produto", "Categoria", "Quantidade", "Total", "ID_Nota"]])
    .setFontWeight("bold");
  if (rows.length) {
    painel.getRange(9, 1, rows.length, 9).setValues(rows);
    painel.getRange(9, 1, rows.length, 1).setNumberFormat("dd/MM/yyyy");
    painel.getRange(9, 8, rows.length, 1).setNumberFormat(CURRENCY_FORMAT);
  }

  // Resumos
  writeResumo_(painel, rows, 20, 1, "mes");
  writeResumo_(painel, rows, 20, 5, "fornecedor");
  writeResumo_(painel, rows, 20, 9, "categoria");
}

/* ==============================
 *  Consumo Semanal
 * ============================== */
function updateConsumoSemanalForInventory_(ss, idInv, dataInv, prevId, prevDate) {
  const shCons = ss.getSheetByName(SHEETS.CONSUMO);
  const shInvItens = ss.getSheetByName(SHEETS.INV_ITENS);
  const shMov = ss.getSheetByName(SHEETS.MOV);

  const invItems = getInventarioItens_(shInvItens, idInv);
  const prevItems = prevId ? getInventarioItens_(shInvItens, prevId) : {};

  const movValues = getMovValues_(shMov);
  const purchaseMap = buildPurchaseMap_(movValues);
  const avgCostMap = computeAvgCostMapAtDate_(movValues, dataInv);

  // Remove linhas anteriores do mesmo inventario
  removeConsumoByInventario_(shCons, idInv);

  const rows = [];
  Object.keys(invItems).forEach(prod => {
    const atual = invItems[prod];
    const anterior = prevItems[prod] || { contado: 0 };
    let entradas = 0;
    if (prevDate) {
      entradas = sumPurchasesBetween_(purchaseMap, prod, prevDate, dataInv);
    }
    const consumo = prevDate ? (anterior.contado + entradas) - atual.contado : 0;
    const custoMedio = avgCostMap[prod] != null ? avgCostMap[prod] : 0;
    const valor = consumo * custoMedio;
    rows.push([idInv, dataInv, prod, consumo, custoMedio, valor]);
  });

  if (rows.length) appendRows_(shCons, rows);
}

function rebuildConsumoSemanal_(ss) {
  const shInv = ss.getSheetByName(SHEETS.INVENTARIOS);
  const shCons = ss.getSheetByName(SHEETS.CONSUMO);
  if (!shInv || !shCons) return;

  const invValues = shInv.getDataRange().getValues();
  if (invValues.length <= 1) return;

  // Limpa corpo e reescreve
  shCons.getRange(2, 1, Math.max(1, shCons.getLastRow() - 1), shCons.getMaxColumns()).clearContent();

  const movValues = getMovValues_(ss.getSheetByName(SHEETS.MOV));
  const purchaseMap = buildPurchaseMap_(movValues);

  const invItens = ss.getSheetByName(SHEETS.INV_ITENS);

  for (let i = 1; i < invValues.length; i++) {
    const row = invValues[i];
    const dataInv = coerceDate_(row[0]);
    const idInv = String(row[3] || "").trim();
    const prevId = String(row[4] || "").trim();
    if (!dataInv || !idInv) continue;
    const prev = prevId ? findInventarioById_(shInv, prevId) : null;

    const invItems = getInventarioItens_(invItens, idInv);
    const prevItems = prevId ? getInventarioItens_(invItens, prevId) : {};
    const avgCostMap = computeAvgCostMapAtDate_(movValues, dataInv);

    const rows = [];
    Object.keys(invItems).forEach(prod => {
      const atual = invItems[prod];
      const anterior = prevItems[prod] || { contado: 0 };
      let entradas = 0;
      if (prev && prev.data) {
        entradas = sumPurchasesBetween_(purchaseMap, prod, prev.data, dataInv);
      }
      const consumo = prev ? (anterior.contado + entradas) - atual.contado : 0;
      const custoMedio = avgCostMap[prod] != null ? avgCostMap[prod] : 0;
      const valor = consumo * custoMedio;
      rows.push([idInv, dataInv, prod, consumo, custoMedio, valor]);
    });
    if (rows.length) appendRows_(shCons, rows);
  }
}

/* ==============================
 *  Utilitarios
 * ============================== */
function ensureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function setHeader_(sh, headers) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
}

function setColumnWidths_(sh, widths) {
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function applyValidationList_(sh, col, list) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).build();
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(rule);
}

function applyValidationListToCell_(range, list) {
  if (!list || !list.length) {
    range.clearDataValidations();
    return;
  }
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).build();
  range.setDataValidation(rule);
}

function applyValidationListToColumn_(sh, col, list) {
  if (!list || !list.length) {
    sh.getRange(2, col, sh.getMaxRows() - 1, 1).clearDataValidations();
    return;
  }
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).build();
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(rule);
}

function refreshDynamicValidations_() {
  const ss = SpreadsheetApp.getActive();
  const itens = ss.getSheetByName(SHEETS.ITENS);
  const invItens = ss.getSheetByName(SHEETS.INV_ITENS);
  const painel = ss.getSheetByName(SHEETS.PAINEL);

  const produtos = getActiveProductNames_();
  const notas = getNotaIds_();
  const fornecedores = getFornecedorList_();
  const categorias = getCategoriaList_();

  if (itens) {
    applyValidationListToColumn_(itens, 1, notas);
    applyValidationListToColumn_(itens, 3, produtos);
  }
  if (invItens) {
    applyValidationListToColumn_(invItens, 2, produtos);
  }
  if (painel) {
    applyValidationListToCell_(painel.getRange("F3"), fornecedores);
    applyValidationListToCell_(painel.getRange("H3"), categorias);
    applyValidationListToCell_(painel.getRange("J3"), LISTS.formasPagamento);
  }
}

function protectHeader_(sh, cols) {
  const range = sh.getRange(1, 1, 1, cols);
  const p = range.protect();
  p.setDescription("Cabecalho protegido");
  p.setWarningOnly(true);
}

function protectSheet_(sh, unprotectedRanges) {
  const p = sh.protect();
  p.setDescription("Planilha protegida");
  p.setWarningOnly(true);
  if (unprotectedRanges && unprotectedRanges.length) p.setUnprotectedRanges(unprotectedRanges);
}

function appendRows_(sh, rows) {
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

function toast_(msg) {
  SpreadsheetApp.getActive().toast(msg, "Info", 5);
}

function coerceDate_(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  const n = Number(v);
  if (isFinite(n) && n > 20000) {
    return new Date(Math.round((n - 25569) * 86400 * 1000));
  }
  const t = Date.parse(v);
  return isNaN(t) ? null : new Date(t);
}

function generateNotaId_(dataCompra, numero) {
  const stamp = Utilities.formatDate(new Date(), TZ, "yyyyMMddHHmmss");
  const num = String(numero || "").replace(/\s+/g, "");
  return `${stamp}_${num}`;
}

function generateInventarioId_(dataInv) {
  const stamp = Utilities.formatDate(new Date(), TZ, "yyyyMMddHHmmss");
  const d = Utilities.formatDate(dataInv, TZ, "yyyyMMdd");
  return `INV_${d}_${stamp}`;
}

function notaDuplicada_(sh, dataCompra, fornecedor, numero) {
  const values = sh.getDataRange().getValues();
  const target = Utilities.formatDate(dataCompra, TZ, "yyyy-MM-dd");
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const d = coerceDate_(row[0]);
    const f = String(row[1] || "").trim();
    const n = String(row[2] || "").trim();
    if (!d) continue;
    const key = Utilities.formatDate(d, TZ, "yyyy-MM-dd");
    if (key === target && f === fornecedor && n === numero) return true;
  }
  return false;
}

function getActiveProducts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PRODUTOS);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const [sku, produto, categoria, unidade, minimo, ativo] = values[i];
    if (String(ativo || "").toUpperCase() !== "SIM") continue;
    if (!produto) continue;
    out.push({ sku, produto, categoria, unidade, minimo });
  }
  return out;
}

function getActiveProductNames_() {
  return getActiveProducts_().map(p => p.produto).filter(Boolean);
}

function getNotaIds_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.NOTAS);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const id = String(values[i][6] || "").trim();
    if (id) out.push(id);
  }
  return out;
}

function getFornecedorList_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.NOTAS);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  const set = {};
  for (let i = 1; i < values.length; i++) {
    const f = String(values[i][1] || "").trim();
    if (f) set[f] = true;
  }
  return Object.keys(set).sort();
}

function getCategoriaList_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PRODUTOS);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  const set = {};
  for (let i = 1; i < values.length; i++) {
    const c = String(values[i][2] || "").trim();
    if (c) set[c] = true;
  }
  return Object.keys(set).sort();
}

function getMovValues_(shMov) {
  if (!shMov) return [];
  const values = shMov.getDataRange().getValues();
  return values.length > 1 ? values.slice(1) : [];
}

function computeStockMapAtDate_(movs, date) {
  const limit = date ? date.getTime() : Infinity;
  const map = {};
  movs.forEach(r => {
    const d = coerceDate_(r[1]);
    if (!d || d.getTime() > limit) return;
    const prod = String(r[4] || "").trim();
    const qtd = Number(r[5] || 0);
    if (!prod) return;
    map[prod] = (map[prod] || 0) + qtd;
  });
  return map;
}

function computeAvgCostMapAtDate_(movs, date) {
  const limit = date ? date.getTime() : Infinity;
  const acc = {};
  movs.forEach(r => {
    const d = coerceDate_(r[1]);
    if (!d || d.getTime() > limit) return;
    const tipo = String(r[2] || "").trim();
    if (tipo !== "ENTRADA_COMPRA") return;
    const prod = String(r[4] || "").trim();
    const qtd = Number(r[5] || 0);
    const custo = Number(r[6] || 0);
    if (!prod || !qtd) return;
    if (!acc[prod]) acc[prod] = { sumQty: 0, sumVal: 0, lastCost: 0, lastDate: 0 };
    acc[prod].sumQty += qtd;
    acc[prod].sumVal += qtd * custo;
    if (d.getTime() >= acc[prod].lastDate) {
      acc[prod].lastDate = d.getTime();
      acc[prod].lastCost = custo;
    }
  });

  const out = {};
  Object.keys(acc).forEach(prod => {
    const a = acc[prod];
    if (a.sumQty > 0) out[prod] = a.sumVal / a.sumQty;
    else out[prod] = a.lastCost || 0;
  });
  return out;
}

function findInventarioAnterior_(shInv, dataInv) {
  const values = shInv.getDataRange().getValues();
  let best = null;
  for (let i = 1; i < values.length; i++) {
    const d = coerceDate_(values[i][0]);
    const id = String(values[i][3] || "").trim();
    if (!d || !id) continue;
    if (d.getTime() < dataInv.getTime()) {
      if (!best || d.getTime() > best.data.getTime()) best = { id, data: d };
    }
  }
  return best;
}

function findInventarioById_(shInv, id) {
  const values = shInv.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][3] || "").trim() === id) {
      return { data: coerceDate_(values[i][0]), id };
    }
  }
  return null;
}

function getInventarioItens_(sh, idInv) {
  const values = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[0] || "").trim() !== idInv) continue;
    const produto = String(row[1] || "").trim();
    const contado = Number(row[4] || 0);
    if (!produto) continue;
    map[produto] = { contado };
  }
  return map;
}

function buildPurchaseMap_(movs) {
  const map = {};
  movs.forEach(r => {
    const tipo = String(r[2] || "").trim();
    if (tipo !== "ENTRADA_COMPRA") return;
    const prod = String(r[4] || "").trim();
    const d = coerceDate_(r[1]);
    const qtd = Number(r[5] || 0);
    if (!prod || !d) return;
    if (!map[prod]) map[prod] = [];
    map[prod].push({ date: d, qty: qtd });
  });
  return map;
}

function sumPurchasesBetween_(purchaseMap, produto, startDate, endDate) {
  const list = purchaseMap[produto] || [];
  const start = startDate ? startDate.getTime() : -Infinity;
  const end = endDate ? endDate.getTime() : Infinity;
  let sum = 0;
  list.forEach(p => {
    const t = p.date.getTime();
    if (t > start && t <= end) sum += p.qty;
  });
  return sum;
}

function removeConsumoByInventario_(shCons, idInv) {
  const values = shCons.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === idInv) rowsToDelete.push(i + 1);
  }
  rowsToDelete.reverse().forEach(r => shCons.deleteRow(r));
}

function getNotasData_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.NOTAS);
  const values = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const id = String(row[6] || "").trim();
    if (!id) continue;
    map[id] = {
      data: coerceDate_(row[0]),
      fornecedor: String(row[1] || "").trim(),
      numero: String(row[2] || "").trim(),
      forma: String(row[3] || "").trim(),
    };
  }
  return map;
}

function getItensNotaData_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.ITENS);
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const idNota = String(row[0] || "").trim();
    if (!idNota) continue;
    out.push({
      idNota,
      produto: String(row[2] || "").trim(),
      categoria: String(row[3] || "").trim(),
      quantidade: Number(row[5] || 0),
      totalLinha: Number(row[7] || 0),
    });
  }
  return out;
}

function writeResumo_(sh, rows, startRow, startCol, mode) {
  const map = {};
  rows.forEach(r => {
    const data = r[0];
    const fornecedor = r[1];
    const categoria = r[5];
    const total = Number(r[7] || 0);
    let key = "";
    if (mode === "mes") {
      key = data ? Utilities.formatDate(data, TZ, "yyyy-MM") : "";
    } else if (mode === "fornecedor") {
      key = fornecedor;
    } else if (mode === "categoria") {
      key = categoria;
    }
    if (!key) return;
    map[key] = (map[key] || 0) + total;
  });

  const entries = Object.keys(map).sort().map(k => [k, map[k]]);
  const clearRange = sh.getRange(startRow + 1, startCol, 200, 2);
  clearRange.clearContent();
  sh.getRange(startRow, startCol, 1, 2).setValues([[titleForMode_(mode), "Total"]]).setFontWeight("bold");
  if (entries.length) {
    sh.getRange(startRow + 1, startCol, entries.length, 2).setValues(entries);
    sh.getRange(startRow + 1, startCol + 1, entries.length, 1).setNumberFormat(CURRENCY_FORMAT);
  }
}

function titleForMode_(mode) {
  if (mode === "mes") return "Mes";
  if (mode === "fornecedor") return "Fornecedor";
  return "Categoria";
}

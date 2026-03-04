/**
 * ----------------------------
 * CONFIGURAÇÃO AUTOMÁTICA
 * ----------------------------
 * Guarda o ID da planilha em Script Properties para o Web App conseguir ler
 * os dados mesmo quando não existe "active spreadsheet".
 */
const PROP_SHEET_ID = "DASHBOARD_SHEET_ID";

/**
 * WEB APP ENTRYPOINT
 */
function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile("index")
      .setTitle("Dashboard Microbiano HRRO")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  } catch (e) {
    return HtmlService.createHtmlOutput(
      "Erro ao carregar o arquivo HTML 'index'. Verifique se existe index.html no Apps Script.<br><br>Detalhes: " +
        e.message
    );
  }
}

/**
 * Menu no Sheets
 */
function onOpen(e) {
  // ✅ garante que o ID da planilha fique salvo para o Web App
  ensureSpreadsheetId_();

  SpreadsheetApp.getUi()
    .createMenu("📊 Dashboard")
    .addItem("🖥️ Abrir Dashboard", "showDashboard")
    .addSeparator()
    .addItem("🔗 Ver Link Externo", "showWebAppUrl")
    .addSeparator()
    .addItem("⚙️ Regravar ID da Planilha (se necessário)", "forceSaveSpreadsheetId")
    .addToUi();
}

/**
 * Força salvar o ID (caso tenha duplicado script / copiado planilha etc.)
 */
function forceSaveSpreadsheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty(PROP_SHEET_ID, ss.getId());
  SpreadsheetApp.getUi().alert("ID da planilha salvo com sucesso para o Web App.");
}

/**
 * Abre modal dentro do Sheets
 */
function showDashboard() {
  try {
    // garante novamente
    ensureSpreadsheetId_();

    const html = HtmlService.createHtmlOutputFromFile("index")
      .setWidth(1980)
      .setHeight(1080);

    SpreadsheetApp.getUi().showModalDialog(html, "Relatório de Perfil Microbiano - Q2 2025");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro ao abrir o Dashboard: " + e.message);
  }
}

/**
 * Mostra URL do Web App
 */
function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();

  if (url && url.includes("/exec")) {
    const html = `
      <div style="padding: 16px; font-family: Inter, Arial, sans-serif;">
        <p style="margin: 0 0 10px;">Link para o Dashboard em ecrã inteiro:</p>
        <input type="text" value="${url}"
          style="width:100%; padding:10px; border:1px solid #cbd5e1; border-radius:8px;"
          readonly onclick="this.select();" />
        <p style="font-size: 11px; color: #64748b; margin-top: 10px;">
          Copie este link e abra numa nova aba do navegador.
        </p>
      </div>`;
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(180), "Link de Acesso");
  } else {
    ui.alert("O App ainda não foi implantado.\nVá em 'Implantar' > 'Nova implementação' e selecione 'App da Web'.");
  }
}

/**
 * (LEGADO) Lê Base_BI (se você ainda usar)
 */
function getSpreadsheetData() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName("Base_BI");

    if (!sheet) {
      console.error("Aba 'Base_BI' não encontrada.");
      return null;
    }

    const data = sheet.getDataRange().getValues();
    return data.filter(row => row.some(v => v !== "" && v !== null));
  } catch (e) {
    console.error("Erro ao ler dados (Base_BI): " + e.message);
    return null;
  }
}

/**
 * ✅ Função chamada pelo HTML: google.script.run.getDashboardData()
 */
function getDashboardData() {
  const ss = getSpreadsheet_();

  // ✅ nomes EXATOS
  const prod   = readMatrixExact_(ss, "1. Produção Laboratorial");
  const topo   = readMatrixExact_(ss, "2. Distribuição por Topografia");
  const micro  = readMatrixExact_(ss, "3. Microrganismos Isolados (Frequência)");
  const resist = readMatrixExact_(ss, "4. Perfil de Resistência (MDR/XDR)");

  const months = ["Abril", "Maio", "Junho"];

  // --- Produção ---
  const positive     = getRowByKey_(prod, "Resultado", "Positivo", months);
  const negative     = getRowByKey_(prod, "Resultado", "Negativo", months);
  const contaminated = getRowByKey_(prod, "Resultado", "Amostra Contaminada", months);

  const totalSamples  = Number(getValueByKey_(prod, "Resultado", "Total Geral", "Total Trimestre")) || 0;
  const totalPositive = Number(getValueByKey_(prod, "Resultado", "Positivo", "Total Trimestre")) || 0;

  // --- Topografia ---
  const topographyOrder = [
    "Tecido (partes moles)",
    "Fragmento ósseo",
    "Urocultura",
    "Peça cirúrgica",
    "Secreção de ferida",
  ];

  const topographyTotals = topographyOrder.map(name =>
    Number(getValueByKey_(topo, "Topografia", name, "Total")) || 0
  );

  const selected = new Set(topographyOrder);
  const otherTotal = topo.rows.reduce((acc, r) => {
    const name = String(r["Topografia"] || "");
    const total = Number(r["Total"] || 0);
    return acc + (!selected.has(name) ? total : 0);
  }, 0);

  // --- Microrganismos (TOP 6 por Total) ---
  const microTop = micro.rows
    .map(r => ({ name: String(r["Espécie"] || ""), total: Number(r["Total"] || 0) }))
    .filter(x => x.name && x.total > 0)
    .sort((a, b) => b.total - a.total)
    .slice(0, 6);

  // --- Resistência ---
  const resistOrder = [
    "Acinetobacter baumannii (MDR/CRE)",
    "Staphylococcus aureus (MRSA)",
    "Klebsiella pneumoniae (ESBL/CRE)",
    "Escherichia coli (ESBL)",
  ];

  const resistTotals = resistOrder.map(name =>
    Number(getValueByKey_(resist, "Microrganismo / Resistência", name, "Total")) || 0
  );

  const criticalEvents = resistTotals.reduce((a, b) => a + (Number(b) || 0), 0);
  const positivityRate = totalSamples > 0 ? (totalPositive / totalSamples) * 100 : 0;

  return {
    updatedAt: new Date().toISOString(),
    summary: {
      totalSamples,
      totalPositive,
      positivityRate,
      criticalEvents,
    },
    charts: {
      production: { labels: months, positive, negative, contaminated },
      topography: { labels: [...topographyOrder, "Outros"], totals: [...topographyTotals, otherTotal] },
      microbio:   { labels: microTop.map(x => x.name), totals: microTop.map(x => x.total) },
      resistance: { labels: resistOrder, totals: resistTotals },
    },
  };
}

/* ----------------- HELPERS ----------------- */

/**
 * Pega a planilha:
 * - se estiver dentro do Sheets: getActiveSpreadsheet()
 * - se for Web App/link externo: openById(propriedade salva)
 */
function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const sheetId = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
  if (!sheetId) {
    throw new Error(
      "Não encontrei o ID da planilha para o Web App. " +
      "Abra a planilha pelo Google Sheets e use o menu 📊 Dashboard > ⚙️ Regravar ID da Planilha."
    );
  }
  return SpreadsheetApp.openById(sheetId);
}

/**
 * Salva o ID da planilha automaticamente quando abrir no Sheets.
 */
function ensureSpreadsheetId_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PROP_SHEET_ID)) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    props.setProperty(PROP_SHEET_ID, ss.getId());
  }
}

function readMatrixExact_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Aba não encontrada: " + sheetName);

  const values = sh.getDataRange().getValues().filter(r => r.some(v => v !== "" && v !== null));
  if (values.length < 2) return { headers: [], rows: [] };

  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = r[i]));
    return obj;
  });

  return { headers, rows };
}

function getRowByKey_(matrix, keyCol, keyValue, cols) {
  const row = matrix.rows.find(r => String(r[keyCol]).trim() === keyValue);
  if (!row) return cols.map(() => 0);
  return cols.map(c => Number(row[c] || 0));
}

function getValueByKey_(matrix, keyCol, keyValue, targetCol) {
  const row = matrix.rows.find(r => String(r[keyCol]).trim() === keyValue);
  return row ? row[targetCol] : null;
}

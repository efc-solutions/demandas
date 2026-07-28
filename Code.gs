/**
 * PAINEL DE DEMANDAS — ABIHPEC
 * Backend Google Apps Script + Google Sheets
 *
 * Aba usada: "Demandas"
 * Colunas (linha 1 = cabeçalho, nessa ordem exata):
 * ID | Tipo | Data | Área | Solicitante | Demanda | Observações | Entrega | Status | Resultado | Nota
 *
 * Tipo aceita: ABIHPEC, EFC — alimenta o filtro "tipo" no painel.
 */

const SHEET_NAME = "Demandas";
const COLS = ["ID","Tipo","Data","Área","Solicitante","Demanda","Observações","Entrega","Status","Resultado","Nota"];

// ---------- SETUP (rode uma vez manualmente pelo editor de Apps Script) ----------
function setupPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(SHEET_NAME);
  if (!aba) aba = ss.insertSheet(SHEET_NAME);
  aba.clear();
  aba.getRange(1, 1, 1, COLS.length).setValues([COLS]);
  aba.setFrozenRows(1);
  aba.getRange(1, 1, 1, COLS.length).setFontWeight("bold");
}

// ---------- ROTEAMENTO ----------
function doGet(e) {
  const action = e.parameter.action;
  if (action === "listar") return listarDemandas();
  return responder({ status: "ok" });
}

function doPost(e) {
  const dados = JSON.parse(e.postData.contents);
  if (dados.action === "salvar") return salvarDemanda(dados);
  if (dados.action === "excluir") return excluirDemanda(dados);
  return responder({ status: "erro", mensagem: "ação desconhecida" });
}

// ---------- LISTAR ----------
function listarDemandas() {
  const aba = getAba();
  if (aba.getLastRow() < 2) return responder({ status: "ok", demandas: [] });

  const valores = aba.getRange(2, 1, aba.getLastRow() - 1, COLS.length).getValues();
  const demandas = valores
    .filter(row => row[0]) // ignora linhas vazias
    .map(row => ({
      id: row[0].toString(),
      tipo: row[1],
      data: formatarData(row[2]),
      area: row[3],
      solicitante: row[4],
      demanda: row[5],
      obs: row[6],
      entrega: formatarData(row[7]),
      status: row[8],
      resultado: row[9],
      nota: row[10]
    }));

  return responder({ status: "ok", demandas: demandas });
}

// ---------- SALVAR (cria ou atualiza) ----------
function salvarDemanda(dados) {
  const aba = getAba();
  const d = dados.demanda;

  if (d.id) {
    // Atualizar linha existente
    const linha = encontrarLinhaPorId(aba, d.id);
    if (linha > 0) {
      aba.getRange(linha, 1, 1, COLS.length).setValues([[
        d.id, d.tipo, d.data, d.area, d.solicitante, d.demanda, d.obs, d.entrega, d.status, d.resultado, d.nota
      ]]);
      return responder({ status: "ok", id: d.id });
    }
  }

  // Criar nova linha com ID novo
  const novoId = gerarProximoId(aba);
  aba.appendRow([novoId, d.tipo, d.data, d.area, d.solicitante, d.demanda, d.obs, d.entrega, d.status, d.resultado, d.nota]);
  return responder({ status: "ok", id: novoId });
}

// ---------- EXCLUIR ----------
function excluirDemanda(dados) {
  const aba = getAba();
  const linha = encontrarLinhaPorId(aba, dados.id);
  if (linha > 0) aba.deleteRow(linha);
  return responder({ status: "ok" });
}

// ---------- HELPERS ----------
function getAba() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(SHEET_NAME);
  if (!aba) throw new Error('Aba "' + SHEET_NAME + '" não encontrada. Rode setupPlanilha() primeiro.');
  return aba;
}

function encontrarLinhaPorId(aba, id) {
  if (aba.getLastRow() < 2) return -1;
  const ids = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0].toString() === id.toString()) return i + 2;
  }
  return -1;
}

function gerarProximoId(aba) {
  const ultimaLinha = aba.getLastRow();
  const numero = ultimaLinha < 2 ? 1 : ultimaLinha; // linha 2 = primeiro registro = D001
  return "D" + String(numero).padStart(3, "0");
}

function formatarData(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const d = String(val.getDate()).padStart(2, "0");
    const m = String(val.getMonth() + 1).padStart(2, "0");
    return d + "/" + m + "/" + val.getFullYear();
  }
  const str = val.toString().trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) return str;
  if (str.includes("GMT") || str.includes("UTC")) {
    const dt = new Date(str);
    if (!isNaN(dt.getTime())) {
      const d = String(dt.getUTCDate()).padStart(2, "0");
      const m = String(dt.getUTCMonth() + 1).padStart(2, "0");
      return d + "/" + m + "/" + dt.getUTCFullYear();
    }
  }
  return str;
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

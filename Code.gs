// Google Apps Script — recebe as confirmações do site, grava na planilha
// e entrega a lista para a página de admin (protegida por senha).
const SENHA_ADMIN = 'troque-esta-senha';   // <<< defina a senha do admin aqui
const ABA = 'Confirmações';
const CAB = ['Data/hora', 'Nome', 'Idade', 'Vai?', 'Nº acompanhantes', 'Acompanhantes (nome e idade)', 'Total de pessoas'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(ABA);
    if (!sh) sh = ss.insertSheet(ABA);
    if (sh.getLastRow() === 0) {
      sh.appendRow(CAB);
      sh.getRange(1, 1, 1, CAB.length).setFontWeight('bold').setBackground('#1d43a6').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    montarResumo(ss);
    const p = e.parameter;
    sh.appendRow([new Date(), p.nome, Number(p.idade), p.presenca, Number(p.qtd_acompanhantes || 0), p.acompanhantes || '', Number(p.total_pessoas || 0)]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  if (!e || !e.parameter || e.parameter.key !== SENHA_ADMIN) return out({ ok: false, erro: 'senha' });
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA);
  if (!sh || sh.getLastRow() < 2) return out({ ok: true, rows: [] });
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, CAB.length).getValues().map(r => ({
    data: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    nome: r[1], idade: r[2], vai: r[3], qtd: r[4], acompanhantes: r[5], total: r[6]
  }));
  return out({ ok: true, rows });
}

// Aba "Resumo": total de pessoas confirmadas, atualizado sozinho.
function montarResumo(ss) {
  if (ss.getSheetByName('Resumo')) return;
  const r = ss.insertSheet('Resumo', 0);
  r.getRange('A1:B1').setValues([['15 anos da Letícia Maria', '']]).setFontWeight('bold');
  r.getRange('A3:A6').setValues([['Pessoas confirmadas (total)'], ['Respostas "Sim"'], ['Respostas "Não"'], ['Acompanhantes']]);
  r.getRange('B3:B6').setFormulas([
    ["=SUMIF('Confirmações'!D:D;\"Sim\";'Confirmações'!G:G)"],
    ["=COUNTIF('Confirmações'!D:D;\"Sim\")"],
    ["=COUNTIF('Confirmações'!D:D;\"Não\")"],
    ["=SUMIF('Confirmações'!D:D;\"Sim\";'Confirmações'!E:E)"]
  ]);
  r.getRange('B3').setFontSize(18).setFontWeight('bold').setFontColor('#1d43a6');
  r.setColumnWidth(1, 240);
}

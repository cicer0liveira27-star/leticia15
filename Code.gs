// Google Apps Script — recebe as confirmações do site, grava na planilha
// e entrega a lista para a página de admin (protegida por senha).
const SENHA_ADMIN = 'Master@2026';   // senha do painel de admin
const PLANILHA_ID = '168mDcju2Ww8dVZ6AawAKXreKmaMs3XuVXCarVGCxdCg'; // sua planilha
const ABA = 'Confirmações';
const CAB = ['Data/hora', 'Nome', 'Idade', 'Vai?'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    let sh = ss.getSheetByName(ABA);
    if (!sh) sh = ss.insertSheet(ABA);
    if (sh.getLastRow() === 0) {
      sh.appendRow(CAB);
      sh.getRange(1, 1, 1, CAB.length).setFontWeight('bold').setBackground('#1d43a6').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    montarResumo(ss);
    const p = e.parameter;
    sh.appendRow([new Date(), p.nome, Number(p.idade), p.presenca]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  if (!e || !e.parameter || e.parameter.key !== SENHA_ADMIN) return out({ ok: false, erro: 'senha' });
  const sh = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA);
  if (e.parameter.acao === 'excluir') return out(excluir(sh, e.parameter.data, e.parameter.nome));
  if (!sh || sh.getLastRow() < 2) return out({ ok: true, rows: [] });
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, CAB.length).getValues().map(r => ({
    data: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    nome: r[1], idade: r[2], vai: r[3]
  }));
  return out({ ok: true, rows });
}

// Exclui uma resposta (ex.: nome repetido). Acha a linha pela data/hora + nome.
function excluir(sh, data, nome) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!sh || sh.getLastRow() < 2) return { ok: false, excluido: false, erro: 'nao_encontrado' };
    const alvo = new Date(data).getTime();
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const d = vals[i][0], n = String(vals[i][1]).trim();
      const mesmaData = d instanceof Date ? Math.abs(d.getTime() - alvo) < 1000 : String(d) === String(data);
      if (mesmaData && n === String(nome).trim()) {
        sh.deleteRow(i + 2);
        return { ok: true, excluido: true };
      }
    }
    return { ok: false, excluido: false, erro: 'nao_encontrado' };
  } finally {
    lock.releaseLock();
  }
}

// Aba "Resumo": total de pessoas confirmadas, atualizado sozinho.
function montarResumo(ss) {
  if (ss.getSheetByName('Resumo')) return;
  const r = ss.insertSheet('Resumo', 0);
  r.getRange('A1:B1').setValues([['15 anos da Letícia Maria', '']]).setFontWeight('bold');
  r.getRange('A3:A5').setValues([['Pessoas confirmadas (vão)'], ['Não vão'], ['Total de respostas']]);
  r.getRange('B3:B5').setFormulas([
    ["=COUNTIF('Confirmações'!D:D;\"Sim\")"],
    ["=COUNTIF('Confirmações'!D:D;\"Não\")"],
    ["=COUNTA('Confirmações'!B:B)-1"]
  ]);
  r.getRange('B3').setFontSize(18).setFontWeight('bold').setFontColor('#1d43a6');
  r.setColumnWidth(1, 240);
}

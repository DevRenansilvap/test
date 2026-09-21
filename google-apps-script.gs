/**
 * Rota+ — Sincronização com Google Sheets
 *
 * COMO USAR:
 * 1. Crie uma planilha nova no Google Sheets (sheets.new)
 * 2. No menu, vá em Extensões > Apps Script
 * 3. Apague o conteúdo padrão e cole todo este arquivo
 * 4. Clique em "Implantar" > "Nova implantação"
 * 5. Tipo: "App da Web"
 * 6. Executar como: "Eu"
 * 7. Quem pode acessar: "Qualquer pessoa"
 * 8. Clique em "Implantar", autorize o acesso, e copie a URL gerada
 *    (termina em /exec)
 * 9. Cole essa URL nas Configurações do app Rota+ (ícone de engrenagem)
 */

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  writeSheet(ss, 'Turnos', body.turnos, ['id','data','horaInicio','horaFim','kmInicial','kmFinal','uberAcumInicio','uberAcumFim','noventaNoveAcumInicio','noventaNoveAcumFim','pix','status']);
  writeSheet(ss, 'DespesasOperacionais', body.despesasOperacionais, ['id','data','hora','descricao','valor','turnoId']);
  writeSheet(ss, 'ContasFixas', body.contasFixas, ['id','nome','valor','dia','pago']);
  writeSheet(ss, 'GastosPessoais', body.gastosPessoais, ['id','descricao','valor','data','categoria','forma']);
  writeSheet(ss, 'Aportes', body.aportes, ['id','nome','valor','data']);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, syncedAt: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Endpoint de sincronização do Rota+ está ativo.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeSheet(ss, name, rows, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clearContents();
  sheet.appendRow(headers);
  rows = rows || [];
  if (rows.length === 0) return;
  var values = rows.map(function (r) {
    return headers.map(function (h) {
      var v = r[h];
      return (v === undefined || v === null) ? '' : v;
    });
  });
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

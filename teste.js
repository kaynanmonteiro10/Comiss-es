function setupPlanilha() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Criar aba "Contratos" se não existir
  let sheetContratos = spreadsheet.getSheetByName('Contratos');
  if (!sheetContratos) {
    sheetContratos = spreadsheet.insertSheet('Contratos');
  }
  
  // Cabeçalhos para Contratos
  const headersContratos = [
    'ID', 'TipoCliente', 'NomeCliente', 'NomeSindicato', 'CNPJ', 'SindicatoVinculado',
    'Responsavel', 'Email', 'Telefone', 'Produto', 'ValorVida', 'VidasIniciais',
    'TotalInicial', 'ComissaoInicial', 'Vendedor', 'Dia', 'Mes', 'Obs'
  ];
  sheetContratos.getRange(1, 1, 1, headersContratos.length).setValues([headersContratos]);
  
  // Criar aba "Movimentos" se não existir
  let sheetMovs = spreadsheet.getSheetByName('Movimentos');
  if (!sheetMovs) {
    sheetMovs = spreadsheet.insertSheet('Movimentos');
  }
  
  // Cabeçalhos para Movimentos
  const headersMovs = [
    'ID', 'ContratoID', 'NumMes', 'VidasAdicionadas', 'ValorVida', 'TotalMes', 'ComissaoMes'
  ];
  sheetMovs.getRange(1, 1, 1, headersMovs.length).setValues([headersMovs]);
  
  Logger.log('Planilha configurada com sucesso!');
}

function fixIDs() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetContratos = spreadsheet.getSheetByName('Contratos');
  const sheetMovs = spreadsheet.getSheetByName('Movimentos');
  
  let contratosFixados = 0;
  let movimentosFixados = 0;
  
  // Corrigir IDs em Contratos
  const contratosData = sheetContratos.getDataRange().getValues();
  for (let i = 1; i < contratosData.length; i++) {
    if (!contratosData[i][0] || contratosData[i][0].toString().trim() === '') {
      const vendedor = contratosData[i][14] || 'Sistema';
      const novoID = 'C_' + vendedor + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
      sheetContratos.getRange(i + 1, 1).setValue(novoID);
      contratosFixados++;
    }
  }
  
  // Corrigir IDs em Movimentos
  const movimentosData = sheetMovs.getDataRange().getValues();
  for (let i = 1; i < movimentosData.length; i++) {
    if (!movimentosData[i][0] || movimentosData[i][0].toString().trim() === '') {
      const contratoID = movimentosData[i][1] || 'Sistema';
      const novoID = 'M_' + contratoID + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
      sheetMovs.getRange(i + 1, 1).setValue(novoID);
      movimentosFixados++;
    }
  }
  
  Logger.log('✅ Contratos fixados: ' + contratosFixados + ' | Movimentos fixados: ' + movimentosFixados);
}

function doGet(e) {
  const action = e.parameter.action;
  const sheetContratos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Contratos');
  const sheetMovs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Movimentos');

  try {
    if (action === 'READ_ALL') {
      const contratos = sheetContratos.getDataRange().getValues().slice(1).map(row => ({
        ID: row[0], TipoCliente: row[1], NomeCliente: row[2], NomeSindicato: row[3], CNPJ: row[4],
        SindicatoVinculado: row[5], Responsavel: row[6], Email: row[7], Telefone: row[8], Produto: row[9],
        ValorVida: row[10], VidasIniciais: row[11], TotalInicial: row[12], ComissaoInicial: row[13],
        Vendedor: row[14], Dia: row[15], Mes: row[16], Obs: row[17]
      }));
      const movimentos = sheetMovs.getDataRange().getValues().slice(1).map(row => ({
        ID: row[0], ContratoID: row[1], NumMes: row[2], VidasAdicionadas: row[3], ValorVida: row[4],
        TotalMes: row[5], ComissaoMes: row[6]
      }));
      return ContentService.createTextOutput(JSON.stringify({ ok: true, contratos, movimentos })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'SAVE_CONTRACT') {
      const vendedor = e.parameter.Vendedor || 'Sistema';
      let contractID = e.parameter.ID || '';
      if (!contractID || contractID.toString().trim() === '') {
        contractID = 'C_' + vendedor + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
      }
      const data = [contractID, e.parameter.TipoCliente, e.parameter.NomeCliente, e.parameter.NomeSindicato,
                    e.parameter.CNPJ, e.parameter.SindicatoVinculado, e.parameter.Responsavel, e.parameter.Email,
                    e.parameter.Telefone, e.parameter.Produto, parseFloat(e.parameter.ValorVida), parseFloat(e.parameter.VidasIniciais),
                    parseFloat(e.parameter.ValorVida) * parseFloat(e.parameter.VidasIniciais),
                    (parseFloat(e.parameter.ValorVida) * parseFloat(e.parameter.VidasIniciais)) * 0.4,
                    vendedor, e.parameter.Dia, e.parameter.Mes, e.parameter.Obs];
      if (e.parameter.ID && e.parameter.ID.toString().trim() !== '') {
        const rows = sheetContratos.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
          if (String(rows[i][0]) == String(e.parameter.ID)) {
            sheetContratos.getRange(i+1, 1, 1, data.length).setValues([data]);
            break;
          }
        }
      } else {
        sheetContratos.appendRow(data);
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'SAVE_MOV') {
      const vendedor = e.parameter.Vendedor || 'Sistema';
      let movID = e.parameter.ID || '';
      if (!movID || movID.toString().trim() === '') {
        movID = 'M_' + vendedor + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
      }
      const totalMes = parseFloat(e.parameter.VidasAdicionadas) * parseFloat(e.parameter.ValorVida);
      const comissaoMes = totalMes * 0.4;
      const data = [movID, e.parameter.ContratoID, e.parameter.NumMes, parseFloat(e.parameter.VidasAdicionadas),
                    parseFloat(e.parameter.ValorVida), totalMes, comissaoMes];
      if (e.parameter.ID && e.parameter.ID.toString().trim() !== '') {
        const rows = sheetMovs.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
          if (String(rows[i][0]) == String(e.parameter.ID)) {
            sheetMovs.getRange(i+1, 1, 1, data.length).setValues([data]);
            break;
          }
        }
      } else {
        sheetMovs.appendRow(data);
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'DELETE_CONTRACT') {
      const rows = sheetContratos.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) == String(e.parameter.ID)) {
          sheetContratos.deleteRow(i+1);
          break;
        }
      }
      const movRows = sheetMovs.getDataRange().getValues();
      for (let i = movRows.length - 1; i >= 1; i--) {
        if (String(movRows[i][1]) == String(e.parameter.ID)) {
          sheetMovs.deleteRow(i+1);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'DELETE_MOV') {
      const rows = sheetMovs.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) == String(e.parameter.ID)) {
          sheetMovs.deleteRow(i+1);
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: 'Ação desconhecida' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

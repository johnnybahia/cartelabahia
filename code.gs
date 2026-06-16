function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Relação de Cartelas Bahia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- UTILITÁRIOS ---
function getDataHoraBR() {
  return Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
}

// --- VERIFICAR SENHA (Aba ACESSO) ---
function verificarCredenciais(usuario, senha) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACESSO');
  if (!sheet) return true; // Se não tiver aba acesso, deixa passar (modo dev)
  
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    const u = data[i][0];
    const s = data[i][1];
    if (u && s && u.toString().trim().toUpperCase() === usuario.toString().trim().toUpperCase() && s.toString().trim() === senha.toString().trim()) {
      return true;
    }
  }
  return false;
}

// --- LER DADOS (Novas Colunas Bahia) ---
function getCartelas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CARTELAS');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => h.toUpperCase().trim());
  const rows = data.slice(1);
  
  // Mapeamento atualizado para "Bahia"
  const map = {
    nome: headers.indexOf('NOME'),
    cor: headers.indexOf('COR'),
    ref: headers.indexOf('REF'),
    largura: headers.indexOf('LARGURA'),
    cliente: headers.indexOf('CLIENTE'),
    caixa: headers.indexOf('CAIXA'),
    data: headers.indexOf('DATA ENTRADA'),
    obs: headers.indexOf('OBSERVAÇÃO'),
    codigo: headers.indexOf('CODIGO'),
    enchimento: headers.indexOf('ENCHIMENTO'),
    alt: headers.indexOf('ALTERAÇÕES')
  };

  return rows.map((row, index) => ({
    id: index + 2,
    nome: map.nome > -1 ? row[map.nome] : '',
    cor: map.cor > -1 ? row[map.cor] : '',
    ref: map.ref > -1 ? row[map.ref] : '',
    largura: map.largura > -1 ? row[map.largura] : '',
    cliente: map.cliente > -1 ? row[map.cliente] : '',
    caixa: map.caixa > -1 ? row[map.caixa] : '',
    dataEntrada: map.data > -1 ? row[map.data] : '',
    observacao: map.obs > -1 ? row[map.obs] : '',
    codigo: map.codigo > -1 ? row[map.codigo] : '',
    enchimento: map.enchimento > -1 ? row[map.enchimento] : '',
    alteracoes: map.alt > -1 ? row[map.alt] : ''
  })).reverse();
}

// --- SALVAR NOVA CARTELA ---
function addCartela(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CARTELAS');
  if (!sheet) throw new Error("Aba 'CARTELAS' não encontrada.");
  
  let dataFormatada = form.dataEntrada;
  if(form.dataEntrada && form.dataEntrada.includes('-')) {
     const p = form.dataEntrada.split('-');
     dataFormatada = `${p[2]}/${p[1]}/${p[0]}`;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toUpperCase().trim());
  const newRow = new Array(headers.length).fill(''); 
  
  const setVal = (headerName, val) => {
    const idx = headers.indexOf(headerName);
    if (idx > -1) newRow[idx] = val;
  };

  setVal('NOME', form.nome.toUpperCase());
  setVal('COR', form.cor.toUpperCase());
  setVal('REF', form.ref.toUpperCase());
  setVal('LARGURA', form.largura.toUpperCase());
  setVal('CLIENTE', form.cliente.toUpperCase());
  setVal('CAIXA', form.caixa.toUpperCase());
  setVal('DATA ENTRADA', dataFormatada);
  setVal('OBSERVAÇÃO', form.observacao.toUpperCase());
  setVal('CODIGO', form.codigo);
  setVal('ENCHIMENTO', form.enchimento.toUpperCase());

  sheet.appendRow(newRow);
  return true;
}

// --- ATUALIZAR LOCALIZAÇÃO (COM SENHA) ---
function updateLocation(rowIndex, newCaixa, usuario, senha) {
  if (!verificarCredenciais(usuario, senha)) throw new Error("Usuário ou Senha incorretos.");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CARTELAS');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toUpperCase().trim());
  
  const colCaixa = headers.indexOf('CAIXA');
  const colAlt = headers.indexOf('ALTERAÇÕES');
  
  if (colCaixa === -1) throw new Error("Coluna CAIXA não encontrada.");
  
  sheet.getRange(rowIndex, colCaixa + 1).setValue(newCaixa.toUpperCase());
  
  if (colAlt > -1) {
    const log = `${usuario.toUpperCase()} em ${getDataHoraBR()}`;
    sheet.getRange(rowIndex, colAlt + 1).setValue(log);
  }
  
  return true;
}

// --- EXCLUIR ITEM (COM SENHA) ---
function deleteItem(rowIndex, usuario, senha) {
   if (!verificarCredenciais(usuario, senha)) throw new Error("Usuário ou Senha incorretos.");

   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CARTELAS');
   sheet.deleteRow(rowIndex);
   return true;
}

var SPREADSHEET_ID = '1ODIQnLUyeJ7t7943b2aeaJOoXK6ah5Up5loxIl575VI';
var _SPREADSHEET_CACHE = null;

function doGet(e) {
  var action = e.parameter.action;
  var data = e.parameter;
  var result;

  try {
    switch (action) {
      case 'buscarMaterial':
        result = buscarMaterial(data);
        break;
      case 'listarMateriais':
        result = listarMateriais();
        break;
      case 'salvarMaterial':
        result = salvarMaterial(data);
        break;
      case 'atualizarCadastroMaterial':
        result = atualizarCadastroMaterial(data);
        break;
      case 'excluirCadastroMaterial':
        result = excluirCadastroMaterial(data);
        break;
      case 'salvarCadastro':
        result = salvarCadastro(data);
        break;
      case 'listarCadastros':
        result = listarCadastros(data.tipo);
        break;
      case 'atualizarCadastro':
        result = atualizarCadastro(data);
        break;
      case 'excluirCadastro':
        result = excluirCadastro(data);
        break;
      case 'salvarPedido':
        result = salvarPedido(data);
        break;
      case 'listarPedidos':
        result = listarPedidos();
        break;
      case 'gerarProximoNumero':
        result = gerarProximoNumero();
        break;
      case 'buscarMaterialPorReferencia':
        result = buscarMaterialPorReferencia(data);
        break;
      case 'buscarReferencia':
        result = buscarReferencia(data);
        break;
      case 'atualizarItensPedido':
        result = atualizarItensPedido(data);
        break;
      case 'listarItensPedido':
        result = listarItensPedido(data);
        break;
      default:
        result = { success: false, mensagem: 'Ação não reconhecida: ' + action };
    }
  } catch (err) {
    result = { success: false, mensagem: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return doGet(e);
}

function configurarBancoDeDados() {
  var ss = _SPREADSHEET_CACHE || SpreadsheetApp.openById(SPREADSHEET_ID);
  _SPREADSHEET_CACHE = ss;
  var abas = {
    "Materiais": ["Código de Barras", "", "ID", "Descrição", "Referência", "ANVISA", "Lote", "Validade"],
    "Pedidos": ["Número", "Hospital", "Médico", "Convênio", "Paciente", "Data Cirurgia", "Status", "Data Emissão", "Responsável"],
    "Itens_Pedido": ["Número Pedido", "Código", "Descrição", "Quantidade", "Lote", "Validade", "ANVISA"],
    "Hospitais": ["Nome", "Data Cadastro"],
    "Convênios": ["Nome", "Data Cadastro"],
    "Médicos": ["Nome", "Data Cadastro"]
  };

  for (var nomeAba in abas) {
    var aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      aba = ss.insertSheet(nomeAba);
      aba.appendRow(abas[nomeAba]);
    }
  }
}

function buscarMaterial(data) {
  var sheet = getSheet('Materiais');
  if (!sheet) return { success: false, mensagem: 'Aba Materiais não encontrada' };
  var codigo = String(data.codigo || '').trim();
  var ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return { success: false, mensagem: 'Material não encontrado' };

  var valores = sheet.getRange(2, 1, ultimaLinha - 1, 8).getValues();
  for (var i = 0; i < valores.length; i++) {
    var barcode = String(valores[i][0] || '').trim();
    var codigoPlanilha = String(valores[i][2] || '').trim();
    if (codigoPlanilha === codigo || barcode === codigo) {
      return {
        success: true,
        material: {
          barcode: barcode,
          codigo: codigoPlanilha,
          descricao: valores[i][3] || '',
          referencia: valores[i][4] || '',
          anvisa: valores[i][5] || '',
          lote: valores[i][6] || '',
          validade: valores[i][7] || ''
        }
      };
    }
  }

  for (var j = 0; j < valores.length; j++) {
    var barcode2 = String(valores[j][0] || '').trim();
    var codigoPlanilha2 = String(valores[j][2] || '').trim();
    if ((codigoPlanilha2 && codigo.indexOf(codigoPlanilha2) === 0) || (barcode2 && codigo.indexOf(barcode2) === 0)) {
      return {
        success: true,
        material: {
          barcode: barcode2,
          codigo: codigoPlanilha2,
          descricao: valores[j][3] || '',
          referencia: valores[j][4] || '',
          anvisa: valores[j][5] || '',
          lote: valores[j][6] || '',
          validade: valores[j][7] || ''
        }
      };
    }
  }

  return { success: false, mensagem: 'Material não encontrado' };
}

function buscarMaterialPorReferencia(data) {
  var sheet = getSheet('Materiais');
  if (!sheet) return { success: false, mensagem: 'Aba Materiais não encontrada' };
  var referencia = data.referencia;
  if (!referencia) return { success: false, mensagem: 'Referência não informada' };
  var ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return { success: false, mensagem: 'Material não encontrado' };

  var valores = sheet.getRange(2, 3, ultimaLinha - 1, 6).getValues();
  for (var i = 0; i < valores.length; i++) {
    var refPlanilha = String(valores[i][2]).trim().toLowerCase();
    if (refPlanilha === String(referencia).trim().toLowerCase()) {
      return {
        success: true,
        material: {
          codigo: valores[i][0] || '',
          descricao: valores[i][1] || '',
          referencia: valores[i][2] || '',
          anvisa: valores[i][3] || '',
          lote: valores[i][4] || '',
          validade: valores[i][5] || ''
        }
      };
    }
  }

  return { success: false, mensagem: 'Material não encontrado para esta referência' };
}

function buscarReferencia(data) {
  var sheet = getSheet('Materiais');
  if (!sheet) return { success: true, materiais: [] };
  var busca = data.busca;
  if (!busca) return { success: true, materiais: [] };
  busca = busca.toLowerCase();
  var ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return { success: true, materiais: [] };
  var valores = sheet.getRange(2, 1, ultimaLinha - 1, 8).getValues();
  var resultados = [];
  for (var i = 0; i < valores.length; i++) {
    var codigo = String(valores[i][0] || '').toLowerCase();
    var descricao = String(valores[i][3] || '').toLowerCase();
    var referencia = String(valores[i][4] || '').toLowerCase();
    if (codigo.indexOf(busca) !== -1 || descricao.indexOf(busca) !== -1 || referencia.indexOf(busca) !== -1) {
      resultados.push({
        codigo: valores[i][0] || '',
        descricao: valores[i][3] || '',
        referencia: valores[i][4] || '',
        anvisa: valores[i][5] || ''
      });
      if (resultados.length >= 10) break;
    }
  }
  return { success: true, materiais: resultados };
}

function listarMateriais() {
  var sheet = getSheet('Materiais');
  if (!sheet) return { success: true, dados: [] };
  var ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return { success: true, dados: [] };
  var valores = sheet.getRange(2, 1, ultimaLinha - 1, 8).getValues();
  var dados = [];
  for (var i = 0; i < valores.length; i++) {
    dados.push({
      barcode: valores[i][0] || '',
      codigo: valores[i][2] || '',
      descricao: valores[i][3] || '',
      referencia: valores[i][4] || '',
      anvisa: valores[i][5] || '',
      lote: valores[i][6] || '',
      validade: valores[i][7] || ''
    });
  }
  return { success: true, dados: dados };
}

function salvarMaterial(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet('Materiais');
    if (!sheet) return { success: false, mensagem: 'Aba Materiais não encontrada' };
    var codigo = data.codigo;
    var descricao = data.descricao;
    var referencia = data.referencia || '';
    var anvisa = data.anvisa || '';
    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha >= 2) {
      var valores = sheet.getRange(2, 3, ultimaLinha - 1, 1).getValues();
      for (var i = 0; i < valores.length; i++) {
        if (String(valores[i][0]) === String(codigo)) {
          return { success: false, mensagem: 'Código já cadastrado' };
        }
      }
    }
    var dataAtual = new Date().toLocaleDateString('pt-BR');
    sheet.appendRow(['', '', codigo, descricao, referencia, anvisa, '', '', dataAtual]);
    return { success: true, mensagem: 'Material cadastrado com sucesso' };
  } finally {
    lock.releaseLock();
  }
}

function atualizarCadastroMaterial(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet('Materiais');
    if (!sheet) return { success: false, mensagem: 'Aba Materiais não encontrada' };
    var codigoOriginal = data.codigoOriginal;
    var codigo = data.codigo;
    var descricao = data.descricao;
    var referencia = data.referencia || '';
    var anvisa = data.anvisa || '';
    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) return { success: false, mensagem: 'Material não encontrado' };
    var valores = sheet.getRange(2, 3, ultimaLinha - 1, 1).getValues();
    for (var i = 0; i < valores.length; i++) {
      if (String(valores[i][0]) === String(codigoOriginal)) {
        var linha = i + 2;
        sheet.getRange(linha, 3).setValue(codigo);
        sheet.getRange(linha, 4).setValue(descricao);
        sheet.getRange(linha, 5).setValue(referencia);
        sheet.getRange(linha, 6).setValue(anvisa);
        return { success: true, mensagem: 'Material atualizado com sucesso' };
      }
    }
    return { success: false, mensagem: 'Material não encontrado' };
  } finally {
    lock.releaseLock();
  }
}

function excluirCadastroMaterial(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet('Materiais');
    if (!sheet) return { success: false, mensagem: 'Aba Materiais não encontrada' };
    var codigo = data.codigo;
    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) return { success: false, mensagem: 'Material não encontrado' };
    var valores = sheet.getRange(2, 3, ultimaLinha - 1, 1).getValues();
    for (var i = 0; i < valores.length; i++) {
      if (String(valores[i][0]) === String(codigo)) {
        sheet.deleteRow(i + 2);
        return { success: true, mensagem: 'Material excluído com sucesso' };
      }
    }
    return { success: false, mensagem: 'Material não encontrado' };
  } finally {
    lock.releaseLock();
  }
}

function salvarCadastro(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var tipo = data.tipo;
    var nome = data.nome;
    var nomeAba = tipo + 's';
    var sheet = getSheet(nomeAba);
    if (!sheet) {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      sheet = ss.insertSheet(nomeAba);
      sheet.appendRow(['Nome', 'Data Cadastro']);
    }
    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha >= 2) {
      var valores = sheet.getRange(2, 1, ultimaLinha - 1, 1).getValues();
      for (var i = 0; i < valores.length; i++) {
        if (String(valores[i][0]).toLowerCase() === String(nome).toLowerCase()) {
          return { success: false, mensagem: tipo + ' já cadastrado' };
        }
      }
    }
    var dataAtual = new Date().toLocaleDateString('pt-BR');
    sheet.appendRow([nome, dataAtual]);
    return { success: true, mensagem: tipo + ' cadastrado com sucesso' };
  } finally {
    lock.releaseLock();
  }
}

function listarCadastros(tipo) {
  var nomeAba = tipo + 's';
  var sheet;
  try {
    sheet = getSheet(nomeAba);
  } catch (e) {
    return { success: true, dados: [] };
  }
  if (!sheet) return { success: true, dados: [] };
  var ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return { success: true, dados: [] };
  var valores = sheet.getRange(2, 1, ultimaLinha - 1, 2).getValues();
  var dados = [];
  for (var i = 0; i < valores.length; i++) {
    dados.push({
      nome: valores[i][0],
      dataCadastro: valores[i][1]
    });
  }
  return { success: true, dados: dados };
}

function atualizarCadastro(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var tipo = data.tipo;
    var nomeOriginal = data.nomeOriginal;
    var nome = data.nome;
    var nomeAba = tipo + 's';
    var sheet = getSheet(nomeAba);
    if (!sheet) return { success: false, mensagem: tipo + ' não encontrado' };
    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) return { success: false, mensagem: tipo + ' não encontrado' };
    var valores = sheet.getRange(2, 1, ultimaLinha - 1, 1).getValues();
    for (var i = 0; i < valores.length; i++) {
      if (String(valores[i][0]).toLowerCase() === String(nomeOriginal).toLowerCase()) {
        var linha = i + 2;
        sheet.getRange(linha, 1).setValue(nome);
        return { success: true, mensagem: tipo + ' atualizado com sucesso' };
      }
    }
    return { success: false, mensagem: tipo + ' não encontrado' };
  } finally {
    lock.releaseLock();
  }
}

function excluirCadastro(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var tipo = data.tipo;
    var nome = data.nome;
    var nomeAba = tipo + 's';
    var sheet = getSheet(nomeAba);
    if (!sheet) return { success: false, mensagem: tipo + ' não encontrado' };
    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) return { success: false, mensagem: tipo + ' não encontrado' };
    var valores = sheet.getRange(2, 1, ultimaLinha - 1, 1).getValues();
    for (var i = 0; i < valores.length; i++) {
      if (String(valores[i][0]).toLowerCase() === String(nome).toLowerCase()) {
        sheet.deleteRow(i + 2);
        return { success: true, mensagem: tipo + ' excluído com sucesso' };
      }
    }
    return { success: false, mensagem: tipo + ' não encontrado' };
  } finally {
    lock.releaseLock();
  }
}

function salvarPedido(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheetPedidos = getSheet('Pedidos');
    var sheetItens = getSheet('Itens_Pedido');
    if (!sheetPedidos || !sheetItens) return { success: false, mensagem: 'Abas não encontradas' };

    var numero = data.numeroPedido;

    var ultimaLinhaPedidos = sheetPedidos.getLastRow();
    if (ultimaLinhaPedidos >= 2) {
      var encontrado = sheetPedidos.getRange(2, 1, ultimaLinhaPedidos - 1, 1)
        .createTextFinder(String(numero).trim())
        .matchEntireCell(true)
        .findNext();
      if (encontrado) {
        return { success: false, mensagem: 'Pedido já salvo com este número' };
      }
    }

    var hospital = data.hospital || '';
    var medico = data.medico || '';
    var convenio = data.convenio || '';
    var paciente = data.paciente || '';
    var dataCirurgia = data.dataCirurgia || '';
    var responsavel = data.responsavel || '';
    var dataAtual = new Date().toLocaleDateString('pt-BR');

    var itens;
    try {
      itens = typeof data.itens === 'string' ? JSON.parse(data.itens) : data.itens;
    } catch (e) {
      itens = [];
    }

    var pedidoLinha = [numero, hospital, medico, convenio, paciente, dataCirurgia, 'Aberto', dataAtual, responsavel];
    ensureRows(sheetPedidos, sheetPedidos.getLastRow() + 1);
    var linhaNum = sheetPedidos.getLastRow() + 1;
    sheetPedidos.getRange(linhaNum, 1, 1, 1).setNumberFormat('@');
    sheetPedidos.getRange(linhaNum, 1, 1, pedidoLinha.length).setValues([pedidoLinha]);

    if (itens && itens.length) {
      var linhasItens = [];
      for (var i = 0; i < itens.length; i++) {
        var item = itens[i];
        linhasItens.push([
          numero,
          item.codigo || '',
          item.descricao || '',
          item.qtd || 0,
          item.lote || '',
          item.validade || '',
          item.anvisa || ''
        ]);
      }
      ensureRows(sheetItens, sheetItens.getLastRow() + linhasItens.length);
      sheetItens.getRange(sheetItens.getLastRow() + 1, 1, linhasItens.length, 7).setValues(linhasItens);
    }

    return { success: true, mensagem: 'Pedido salvo com sucesso' };
  } finally {
    lock.releaseLock();
  }
}

function gerarProximoNumero() {
  var maxNum = 0;
  var sheetPedidos = getSheet('Pedidos');
  if (sheetPedidos && sheetPedidos.getLastRow() >= 2) {
    var nums = sheetPedidos.getRange(2, 1, sheetPedidos.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < nums.length; i++) {
      var n = parseInt(String(nums[i][0]).replace(/^0+/, ''), 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
  }
  var sheetItens = getSheet('Itens_Pedido');
  if (sheetItens && sheetItens.getLastRow() >= 2) {
    var nums2 = sheetItens.getRange(2, 1, sheetItens.getLastRow() - 1, 1).getValues();
    for (var j = 0; j < nums2.length; j++) {
      var n2 = parseInt(String(nums2[j][0]).replace(/^0+/, ''), 10);
      if (!isNaN(n2) && n2 > maxNum) maxNum = n2;
    }
  }
  var proximo = maxNum + 1;
  return { success: true, numero: String(proximo).padStart(6, '0') };
}

function listarPedidos() {
  var sheet = getSheet('Pedidos');
  if (!sheet) return { success: true, dados: [] };
  var ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return { success: true, dados: [] };
  var valores = sheet.getRange(2, 1, ultimaLinha - 1, 9).getValues();
  var dados = [];
  for (var i = 0; i < valores.length; i++) {
    dados.push({
      numero: String(valores[i][0]),
      hospital: valores[i][1],
      medico: valores[i][2],
      convenio: valores[i][3],
      paciente: valores[i][4],
      dataCirurgia: valores[i][5],
      status: valores[i][6],
      dataCadastro: valores[i][7],
      responsavel: valores[i][8]
    });
  }
  return { success: true, dados: dados };
}

function atualizarItensPedido(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheetItens = getSheet('Itens_Pedido');
    if (!sheetItens) return { success: false, mensagem: 'Aba Itens_Pedido não encontrada' };

    var numero = data.numeroPedido;
    if (!numero) return { success: false, mensagem: 'Número do pedido não informado' };

    var itens;
    try {
      itens = typeof data.itens === 'string' ? JSON.parse(data.itens) : data.itens;
    } catch (e) {
      itens = [];
    }

    var ultimaLinha = sheetItens.getLastRow();
    var linhasRestantes = [];
    if (ultimaLinha >= 2) {
      var valores = sheetItens.getRange(2, 1, ultimaLinha - 1, 7).getValues();
      for (var i = 0; i < valores.length; i++) {
        if (String(valores[i][0]).trim() !== String(numero).trim()) {
          linhasRestantes.push(valores[i]);
        }
      }
    }

    var novasLinhas = [];
    for (var k = 0; k < itens.length; k++) {
      var item = itens[k];
      novasLinhas.push([
        numero,
        item.codigo || '',
        item.descricao || '',
        item.qtd || 0,
        item.lote || '',
        item.validade || '',
        item.anvisa || ''
      ]);
    }

    var linhasFinais = linhasRestantes.concat(novasLinhas);
    if (ultimaLinha >= 2) {
      sheetItens.getRange(2, 1, ultimaLinha - 1, 7).clearContent();
    }
    if (linhasFinais.length) {
      ensureRows(sheetItens, linhasFinais.length + 1);
      sheetItens.getRange(2, 1, linhasFinais.length, 7).setValues(linhasFinais);
    }

    var sheetPedidos = getSheet('Pedidos');
    if (sheetPedidos && (data.hospital || data.medico || data.convenio || data.paciente || data.dataCirurgia)) {
      var ultimaLinhaPedidos = sheetPedidos.getLastRow();
      if (ultimaLinhaPedidos >= 2) {
        var valoresPedidos = sheetPedidos.getRange(2, 1, ultimaLinhaPedidos - 1, 9).getValues();
        for (var p = 0; p < valoresPedidos.length; p++) {
          if (String(valoresPedidos[p][0]).trim() === String(numero).trim()) {
            if (data.hospital !== undefined) sheetPedidos.getRange(p + 2, 2).setValue(data.hospital);
            if (data.medico !== undefined) sheetPedidos.getRange(p + 2, 3).setValue(data.medico);
            if (data.convenio !== undefined) sheetPedidos.getRange(p + 2, 4).setValue(data.convenio);
            if (data.paciente !== undefined) sheetPedidos.getRange(p + 2, 5).setValue(data.paciente);
            if (data.dataCirurgia !== undefined) sheetPedidos.getRange(p + 2, 6).setValue(data.dataCirurgia);
            break;
          }
        }
      }
    }

    return { success: true, mensagem: 'Pedido atualizado com sucesso' };
  } finally {
    lock.releaseLock();
  }
}

function listarItensPedido(data) {
  var sheet = getSheet('Itens_Pedido');
  if (!sheet) return { success: true, dados: [] };
  var numero = data.numeroPedido;
  if (!numero) return { success: true, dados: [] };
  var ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return { success: true, dados: [] };
  var valores = sheet.getRange(2, 1, ultimaLinha - 1, 7).getValues();
  var dados = [];
  for (var i = 0; i < valores.length; i++) {
    if (String(valores[i][0]).trim() === String(numero).trim()) {
      dados.push({
        numero: valores[i][0],
        codigo: valores[i][1],
        descricao: valores[i][2],
        qtd: valores[i][3],
        lote: valores[i][4],
        validade: valores[i][5],
        anvisa: valores[i][6]
      });
    }
  }
  return { success: true, dados: dados };
}

function getSheet(name) {
  var ss = _SPREADSHEET_CACHE || SpreadsheetApp.openById(SPREADSHEET_ID);
  _SPREADSHEET_CACHE = ss;
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    configurarBancoDeDados();
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

function ensureRows(sheet, requiredRowCount) {
  var maxRows = sheet.getMaxRows();
  if (requiredRowCount > maxRows) {
    sheet.insertRowsAfter(maxRows, requiredRowCount - maxRows);
  }
}

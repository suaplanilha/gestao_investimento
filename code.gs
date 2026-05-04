/**
 * SISTEMA APOLLO ENTERPRISE (SAE) - GESTÃO MINI ÍNDICE
 * Desenvolvido para: Anderson
 * Versão: 2.0 (Escalonamento de Cadastro e Segurança)
 */

const CONFIG = {
  TAXA_POR_CONTRATO: 0.25,
  VALOR_PONTO: 0.20,
  APP_ID: typeof __app_id !== 'undefined' ? __app_id : 'SaeMiniIndice'
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SAE - Gestão Mini Índice')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Setup com Atualização Incremental de Colunas
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tables = {
    'clientes': ['uuid', 'id_sequencial', 'data_cadastro', 'status', 'nome', 'idade', 'telefone', 'email', 'cidade', 'estado', 'redes_sociais', 'valor_mensalidade', 'vencimento_dia', 'capital_inicial_contrato'],
    'operacoes': ['uuid', 'cliente_id', 'data_iso', 'contratos', 'valor_por_contrato', 'pontos_pos', 'pontos_neg', 'take', 'stop', 'lucro_bruto', 'taxas', 'lucro_liquido', 'percentual_ganho'],
    'saldos': ['uuid', 'cliente_id', 'data_atualizacao', 'saldo_atual']
  };

  Object.keys(tables).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(tables[name]);
      sheet.getRange(1, 1, 1, tables[name].length).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    } else {
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      tables[name].forEach(h => {
        if (currentHeaders.indexOf(h) === -1) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
        }
      });
    }
  });
  return "Estrutura SAE sincronizada com sucesso!";
}

/**
 * CRUD Clientes com Validação de Segurança
 */
function getClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('clientes');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  return data.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function salvarCliente(cliente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('clientes');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 1. Verificação de Duplicidade (Segurança)
  const colTelefone = headers.indexOf('telefone');
  const colEmail = headers.indexOf('email');
  const colUuid = headers.indexOf('uuid');

  for (let i = 1; i < data.length; i++) {
    // Se for edição, ignora o próprio registro
    if (cliente.uuid && data[i][colUuid] === cliente.uuid) continue;
    
    if (data[i][colTelefone] == cliente.telefone) throw new Error("Telefone já cadastrado no sistema.");
    if (data[i][colEmail] == cliente.email) throw new Error("E-mail já cadastrado no sistema.");
  }

  // 2. Lógica de ID Sequencial
  let idSequencial = cliente.id_sequencial;
  if (!idSequencial) {
    const ids = data.slice(1).map(r => Number(r[headers.indexOf('id_sequencial')]) || 0);
    idSequencial = ids.length > 0 ? Math.max(...ids) + 1 : 1;
  }

  const payload = [
    cliente.uuid || Utilities.getUuid(),
    idSequencial,
    cliente.data_cadastro || new Date().toISOString(),
    cliente.status || 'Ativo',
    cliente.nome,
    cliente.idade,
    cliente.telefone,
    cliente.email,
    cliente.cidade,
    cliente.estado,
    cliente.redes_sociais,
    cliente.valor_mensalidade || 0,
    cliente.vencimento_dia || 10,
    cliente.capital_inicial_contrato || 500
  ];

  if (cliente.uuid) {
    // Modo Edição
    for (let i = 1; i < data.length; i++) {
      if (data[i][colUuid] === cliente.uuid) {
        sheet.getRange(i + 1, 1, 1, payload.length).setValues([payload]);
        break;
      }
    }
  } else {
    // Modo Novo
    sheet.appendRow(payload);
  }

  return { success: true, id: idSequencial };
}

/**
 * Operacional
 */
function registrarOperacao(op) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('operacoes');
  
  const pontosSaldo = Number(op.pontos_pos || 0) - Number(op.pontos_neg || 0);
  const vlrPorContrato = Number(op.valor_por_contrato || CONFIG.VALOR_PONTO);
  
  const lucroBruto = pontosSaldo * op.contratos * vlrPorContrato;
  const taxas = op.contratos * CONFIG.TAXA_POR_CONTRATO;
  const lucroLiquido = lucroBruto - taxas;
  
  let percentual = 0;
  if (op.capital_base > 0) {
    percentual = (lucroBruto / (op.capital_base * op.contratos)) * 100;
  }
  
  sheet.appendRow([
    Utilities.getUuid(),
    op.cliente_id,
    new Date().toISOString(),
    op.contratos,
    vlrPorContrato,
    op.pontos_pos,
    op.pontos_neg,
    op.take,
    op.stop,
    lucroBruto,
    taxas,
    lucroLiquido,
    percentual.toFixed(2) + '%'
  ]);
  
  return { success: true };
}

function getOperacoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('operacoes');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data.map(r => ({ data_iso: r[2], lucro_liquido: r[11], cliente_id: r[1] }));
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const opSheet = ss.getSheetByName('operacoes');
  if (!opSheet || opSheet.getLastRow() < 2) return { totalLiquido: 0, totalTaxas: 0, totalOperacoes: 0 };
  
  const ops = opSheet.getDataRange().getValues();
  ops.shift();
  
  let totalLiquido = 0;
  let totalTaxas = 0;
  ops.forEach(row => {
    totalTaxas += Number(row[10] || 0);
    totalLiquido += Number(row[11] || 0);
  });
  
  return { totalLiquido, totalTaxas, totalOperacoes: ops.length };
}

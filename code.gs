/**
 * SISTEMA APOLLO ENTERPRISE (SAE) - GESTÃO MINI ÍNDICE
 * Desenvolvido para: Anderson
 * Versão: 2.1 (Serviços + Health-check)
 */

const CONFIG = {
  TAXA_POR_CONTRATO: 0.25,
  VALOR_PONTO: 0.20,
  APP_ID: typeof __app_id !== 'undefined' ? __app_id : 'SaeMiniIndice'
};

const TABLES = {
  clientes: ['uuid', 'id_sequencial', 'data_cadastro', 'status', 'nome', 'idade', 'telefone', 'email', 'cidade', 'estado', 'redes_sociais', 'valor_mensalidade', 'vencimento_dia', 'capital_inicial_contrato'],
  operacoes: ['uuid', 'cliente_id', 'data_iso', 'contratos', 'valor_por_contrato', 'pontos_pos', 'pontos_neg', 'take', 'stop', 'lucro_bruto', 'taxas', 'lucro_liquido', 'percentual_ganho'],
  saldos: ['uuid', 'cliente_id', 'data_atualizacao', 'saldo_atual'],
  auditoria: ['uuid', 'data_iso', 'entidade', 'entidade_id', 'acao', 'payload_json']
};

function toNumber(value, fallback) {
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function normalizeText(value) {
  return String(value || '').trim();
}

function getSheetOrThrow(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Aba obrigatória não encontrada: ${name}. Execute setupDatabase().`);
  return sheet;
}



function toIsoNow() {
  return new Date().toISOString();
}

function writeAuditLog(entry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetOrThrow(ss, 'auditoria');
  sheet.appendRow([
    Utilities.getUuid(),
    toIsoNow(),
    normalizeText(entry.entidade),
    normalizeText(entry.entidade_id),
    normalizeText(entry.acao),
    JSON.stringify(entry.payload || {})
  ]);
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SAE - Gestão Mini Índice')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABLES).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(TABLES[name]);
      sheet.getRange(1, 1, 1, TABLES[name].length).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    } else {
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      TABLES[name].forEach(h => {
        if (currentHeaders.indexOf(h) === -1) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
        }
      });
    }
  });
  return 'Estrutura SAE sincronizada com sucesso!';
}

function healthCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const details = Object.keys(TABLES).map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return { entidade: name, ok: false, erro: 'Aba não encontrada' };
    const headers = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
    const missing = TABLES[name].filter(h => headers.indexOf(h) === -1);
    return { entidade: name, ok: missing.length === 0, erro: missing.length ? `Colunas ausentes: ${missing.join(', ')}` : '' };
  });

  const invalid = details.filter(d => !d.ok);
  return {
    success: invalid.length === 0,
    timestamp_iso: toIsoNow(),
    app_id: CONFIG.APP_ID,
    details,
    message: invalid.length === 0 ? 'Health-check OK' : 'Health-check encontrou inconsistências'
  };
}

const ClientesService = {
  list() {
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
  },

  save(clienteDTO) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheetOrThrow(ss, 'clientes');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const dto = {
      uuid: clienteDTO.uuid || null,
      id_sequencial: toNumber(clienteDTO.id_sequencial, 0),
      data_cadastro: clienteDTO.data_cadastro || toIsoNow(),
      status: clienteDTO.status || 'Ativo',
      nome: normalizeText(clienteDTO.nome),
      idade: toNumber(clienteDTO.idade, 0),
      telefone: normalizeText(clienteDTO.telefone),
      email: normalizeText(clienteDTO.email),
      cidade: normalizeText(clienteDTO.cidade),
      estado: normalizeText(clienteDTO.estado).toUpperCase(),
      redes_sociais: normalizeText(clienteDTO.redes_sociais),
      valor_mensalidade: toNumber(clienteDTO.valor_mensalidade, 0),
      vencimento_dia: toNumber(clienteDTO.vencimento_dia, 10),
      capital_inicial_contrato: toNumber(clienteDTO.capital_inicial_contrato, 500)
    };

    if (!dto.nome || !dto.telefone || !dto.email) {
      throw new Error('Nome, telefone e e-mail são obrigatórios.');
    }

    const colTelefone = headers.indexOf('telefone');
    const colEmail = headers.indexOf('email');
    const colUuid = headers.indexOf('uuid');

    for (let i = 1; i < data.length; i++) {
      if (dto.uuid && data[i][colUuid] === dto.uuid) continue;
      if (data[i][colTelefone] == dto.telefone) throw new Error('Telefone já cadastrado no sistema.');
      if (data[i][colEmail] == dto.email) throw new Error('E-mail já cadastrado no sistema.');
    }

    let idSequencial = dto.id_sequencial;
    if (!idSequencial) {
      const ids = data.slice(1).map(r => Number(r[headers.indexOf('id_sequencial')]) || 0);
      idSequencial = ids.length > 0 ? Math.max(...ids) + 1 : 1;
    }

    const payload = [
      dto.uuid || Utilities.getUuid(),
      idSequencial,
      dto.data_cadastro,
      dto.status,
      dto.nome,
      dto.idade,
      dto.telefone,
      dto.email,
      dto.cidade,
      dto.estado,
      dto.redes_sociais,
      dto.valor_mensalidade,
      dto.vencimento_dia,
      dto.capital_inicial_contrato
    ];

    if (dto.uuid) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][colUuid] === dto.uuid) {
          sheet.getRange(i + 1, 1, 1, payload.length).setValues([payload]);
          break;
        }
      }
    } else {
      sheet.appendRow(payload);
    }

    writeAuditLog({
      entidade: 'clientes',
      entidade_id: String(payload[0]),
      acao: dto.uuid ? 'update' : 'create',
      payload: { id_sequencial: idSequencial, status: dto.status, nome: dto.nome, email: dto.email }
    });

    return { success: true, id: idSequencial };
  }
};

const OperacoesService = {
  list() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('operacoes');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    data.shift();
    return data.map(r => ({ data_iso: r[2], lucro_liquido: r[11], cliente_id: r[1] }));
  },

  save(opDTO) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheetOrThrow(ss, 'operacoes');

    const dto = {
      cliente_id: normalizeText(opDTO.cliente_id),
      contratos: toNumber(opDTO.contratos, 0),
      valor_por_contrato: toNumber(opDTO.valor_por_contrato, CONFIG.VALOR_PONTO),
      pontos_pos: toNumber(opDTO.pontos_pos, 0),
      pontos_neg: toNumber(opDTO.pontos_neg, 0),
      take: opDTO.take,
      stop: opDTO.stop,
      capital_base: toNumber(opDTO.capital_base, 0)
    };

    if (!dto.cliente_id) throw new Error('Cliente é obrigatório.');
    if (dto.contratos <= 0) throw new Error('Quantidade de contratos deve ser maior que zero.');

    const pontosSaldo = dto.pontos_pos - dto.pontos_neg;
    const lucroBruto = pontosSaldo * dto.contratos * dto.valor_por_contrato;
    const taxas = dto.contratos * CONFIG.TAXA_POR_CONTRATO;
    const lucroLiquido = lucroBruto - taxas;
    const percentual = dto.capital_base > 0 ? (lucroBruto / (dto.capital_base * dto.contratos)) * 100 : 0;

    sheet.appendRow([
      Utilities.getUuid(),
      dto.cliente_id,
      toIsoNow(),
      dto.contratos,
      dto.valor_por_contrato,
      dto.pontos_pos,
      dto.pontos_neg,
      dto.take,
      dto.stop,
      lucroBruto,
      taxas,
      lucroLiquido,
      percentual.toFixed(2) + '%'
    ]);

    writeAuditLog({
      entidade: 'operacoes',
      entidade_id: dto.cliente_id,
      acao: 'create',
      payload: { contratos: dto.contratos, lucro_bruto: lucroBruto, taxas, lucro_liquido: lucroLiquido }
    });

    return { success: true };
  }
};

const DashboardService = {
  getStats() {
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
};

// Wrappers para google.script.run (contrato estável com frontend)
function getClientes() { return ClientesService.list(); }
function salvarCliente(clienteDTO) { return ClientesService.save(clienteDTO); }
function getOperacoes() { return OperacoesService.list(); }
function registrarOperacao(opDTO) { return OperacoesService.save(opDTO); }
function getDashboardData() { return DashboardService.getStats(); }

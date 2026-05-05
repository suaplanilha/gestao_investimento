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
  operacoes: ['uuid', 'id_usuario', 'cliente_id', 'data_iso', 'status', 'contratos', 'valor_por_contrato', 'pontos_pos', 'pontos_neg', 'take', 'stop', 'lucro_bruto', 'taxas', 'lucro_liquido', 'percentual_ganho', 'capital_base'],
  saldos: ['uuid', 'cliente_id', 'data_atualizacao', 'saldo_atual'],
  auditoria: ['uuid', 'data_iso', 'entidade', 'entidade_id', 'acao', 'payload_json'],
  anotacoes: ['uuid','id','data','nome','tel','email','cidade','estado','status','pagamento_valor','data_pagamento','data_desligamento']
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



function repairOperacoesRowIfNeeded(row, headers) {
  const idx = name => headers.indexOf(name);
  const iContratos = idx('contratos');
  const iStatus = idx('status');
  const iCapitalBase = idx('capital_base');
  if (iContratos < 0 || iStatus < 0) return row;

  const contratosVal = row[iContratos];
  const statusVal = normalizeText(row[iStatus]);
  const looksBroken = (typeof contratosVal === 'string' && ['ATIVO', 'INATIVO'].includes(contratosVal.toUpperCase())) ||
    (statusVal && !['ATIVO', 'INATIVO'].includes(statusVal.toUpperCase()));

  if (!looksBroken) return row;

  // Correção para layout legado corrompido observado em produção.
  const fixed = row.slice();
  const broken = {
    contratos: row[idx('contratos')],
    pontos: row[idx('pontos')],
    lucro_bruto: row[idx('lucro_bruto')],
    taxas: row[idx('taxas')],
    lucro_liquido: row[idx('lucro_liquido')],
    valor_por_contrato: row[idx('valor_por_contrato')],
    pontos_pos: row[idx('pontos_pos')],
    pontos_neg: row[idx('pontos_neg')],
    take: row[idx('take')],
    stop: row[idx('stop')],
    percentual_ganho: row[idx('percentual_ganho')],
    status: row[idx('status')],
    capital_base: row[idx('capital_base')]
  };

  if (idx('status') >= 0) fixed[idx('status')] = normalizeText(broken.contratos || 'Ativo');
  if (idx('contratos') >= 0) fixed[idx('contratos')] = toNumber(broken.pontos, 0);
  if (idx('valor_por_contrato') >= 0) fixed[idx('valor_por_contrato')] = toNumber(broken.lucro_bruto, CONFIG.VALOR_PONTO);
  if (idx('pontos_pos') >= 0) fixed[idx('pontos_pos')] = toNumber(broken.taxas, 0);
  if (idx('pontos_neg') >= 0) fixed[idx('pontos_neg')] = toNumber(broken.lucro_liquido, 0);
  if (idx('take') >= 0) fixed[idx('take')] = toNumber(broken.valor_por_contrato, 0);
  if (idx('stop') >= 0) fixed[idx('stop')] = toNumber(broken.pontos_pos, 0);
  if (idx('capital_base') >= 0) fixed[idx('capital_base')] = toNumber(broken.status, toNumber(broken.capital_base, 0));

  const pontosSaldo = toNumber(fixed[idx('pontos_pos')], 0) - toNumber(fixed[idx('pontos_neg')], 0);
  const lucroBruto = pontosSaldo * toNumber(fixed[idx('contratos')], 0) * toNumber(fixed[idx('valor_por_contrato')], CONFIG.VALOR_PONTO);
  const taxas = toNumber(fixed[idx('contratos')], 0) * CONFIG.TAXA_POR_CONTRATO;
  const liquido = lucroBruto - taxas;
  const capital = toNumber(fixed[idx('capital_base')], 0);
  const perc = capital > 0 && toNumber(fixed[idx('contratos')], 0) > 0 ? (lucroBruto / (capital * toNumber(fixed[idx('contratos')], 0))) * 100 : 0;

  if (idx('lucro_bruto') >= 0) fixed[idx('lucro_bruto')] = lucroBruto;
  if (idx('taxas') >= 0) fixed[idx('taxas')] = taxas;
  if (idx('lucro_liquido') >= 0) fixed[idx('lucro_liquido')] = liquido;
  if (idx('percentual_ganho') >= 0) fixed[idx('percentual_ganho')] = perc.toFixed(2) + '%';

  return fixed;
}

function repairOperacoesSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('operacoes');
  if (!sheet || sheet.getLastRow() < 2) return { repaired: 0 };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let repaired = 0;
  for (let i = 1; i < data.length; i++) {
    const fixed = repairOperacoesRowIfNeeded(data[i], headers);
    if (JSON.stringify(fixed) !== JSON.stringify(data[i])) {
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([fixed]);
      repaired++;
    }
  }
  return { repaired };
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
  const repaired = repairOperacoesSheetData();
  return `Estrutura SAE sincronizada com sucesso! Linhas operacionais reparadas: ${repaired.repaired}.`;
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


function validateInfrastructureOrThrow() {
  const report = healthCheck();
  if (!report.success) {
    const details = report.details.filter(d => !d.ok).map(d => `${d.entidade}: ${d.erro}`).join(' | ');
    throw new Error(`Infraestrutura inválida: ${details}`);
  }
  return report;
}

function runPreReleaseHealthCheck() {
  const report = healthCheck();
  return {
    success: report.success,
    checked_at_iso: toIsoNow(),
    release_ready: report.success,
    issues: report.details.filter(d => !d.ok)
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

    const entity = {
      uuid: dto.uuid || Utilities.getUuid(),
      id_sequencial: idSequencial,
      data_cadastro: dto.data_cadastro,
      status: dto.status,
      nome: dto.nome,
      idade: dto.idade,
      telefone: dto.telefone,
      email: dto.email,
      cidade: dto.cidade,
      estado: dto.estado,
      redes_sociais: dto.redes_sociais,
      valor_mensalidade: dto.valor_mensalidade,
      vencimento_dia: dto.vencimento_dia,
      capital_inicial_contrato: dto.capital_inicial_contrato
    };
    const payload = headers.map(h => Object.prototype.hasOwnProperty.call(entity, h) ? entity[h] : '');

    if (dto.uuid) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][colUuid] === dto.uuid) {
          sheet.getRange(i + 1, 1, 1, headers.length).setValues([payload]);
          break;
        }
      }
    } else {
      sheet.appendRow(payload);
    }

    writeAuditLog({
      entidade: 'clientes',
      entidade_id: String(entity.uuid),
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
    const headers = data.shift();
    const clientes = ClientesService.list();
    const mapNome = {};
    clientes.forEach(c => mapNome[c.uuid] = c.nome);

    return data.map(r => {
      const o = {};
      headers.forEach((h, i) => o[h] = r[i]);
      o.nome_cliente = mapNome[o.cliente_id] || 'Não encontrado';
      return o;
    });
  },

  save(opDTO) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheetOrThrow(ss, 'operacoes');

    const dto = {
      uuid: normalizeText(opDTO.uuid),
      id_usuario: normalizeText(opDTO.id_usuario),
      cliente_id: normalizeText(opDTO.cliente_id),
      status: normalizeText(opDTO.status || 'Ativo'),
      contratos: toNumber(opDTO.contratos, 0),
      valor_por_contrato: toNumber(opDTO.valor_por_contrato, CONFIG.VALOR_PONTO),
      pontos_pos: toNumber(opDTO.pontos_pos, 0),
      pontos_neg: toNumber(opDTO.pontos_neg, 0),
      take: toNumber(opDTO.take, 0),
      stop: toNumber(opDTO.stop, 0),
      capital_base: toNumber(opDTO.capital_base, 0)
    };

    if (!dto.cliente_id) throw new Error('Cliente é obrigatório.');
    if (dto.contratos <= 0) throw new Error('Quantidade de contratos deve ser maior que zero.');

    const pontosSaldo = dto.pontos_pos - dto.pontos_neg;
    const lucroBruto = pontosSaldo * dto.contratos * dto.valor_por_contrato;
    const taxas = dto.contratos * CONFIG.TAXA_POR_CONTRATO;
    const lucroLiquido = lucroBruto;
    const percentual = dto.capital_base > 0 ? (lucroBruto / (dto.capital_base * dto.contratos)) * 100 : 0;

    const all = sheet.getDataRange().getValues();
    const headers = all[0] || [];
    const entity = {
      uuid: dto.uuid || Utilities.getUuid(),
      id_usuario: dto.id_usuario || '',
      cliente_id: dto.cliente_id,
      data_iso: toIsoNow(),
      status: dto.status,
      contratos: dto.contratos,
      valor_por_contrato: dto.valor_por_contrato,
      pontos_pos: dto.pontos_pos,
      pontos_neg: dto.pontos_neg,
      take: dto.take,
      stop: dto.stop,
      lucro_bruto: lucroBruto,
      taxas: taxas,
      lucro_liquido: lucroLiquido,
      percentual_ganho: percentual.toFixed(2) + '%',
      capital_base: dto.capital_base,
      pontos: dto.pontos_pos - dto.pontos_neg
    };
    const payload = headers.map(h => Object.prototype.hasOwnProperty.call(entity, h) ? entity[h] : '');
    const colUuid = headers.indexOf('uuid');
    if (dto.uuid && colUuid >= 0) {
      let updated = false;
      for (let i = 1; i < all.length; i++) {
        if (all[i][colUuid] === dto.uuid) {
          sheet.getRange(i + 1, 1, 1, headers.length).setValues([payload]);
          updated = true;
          break;
        }
      }
      if (!updated) sheet.appendRow(payload);
    } else {
      sheet.appendRow(payload);
    }

    writeAuditLog({
      entidade: 'operacoes',
      entidade_id: entity.uuid,
      acao: dto.uuid ? 'update' : 'create',
      payload: { id_usuario: dto.id_usuario || '',
      cliente_id: dto.cliente_id, contratos: dto.contratos, lucro_liquido: lucroLiquido, status: dto.status }
    });

    return { success: true, uuid: payload[0] };
  }
};

const DashboardService = {
  getStats() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const opSheet = ss.getSheetByName('operacoes');
    if (!opSheet || opSheet.getLastRow() < 2) return { totalLiquido: 0, totalTaxas: 0, totalOperacoes: 0 };

    const data = opSheet.getDataRange().getValues();
    const headers = data.shift();
    const colTaxas = headers.indexOf('taxas');
    const colLiquido = headers.indexOf('lucro_liquido');

    let totalLiquido = 0;
    let totalTaxas = 0;
    data.forEach(row => {
      totalTaxas += Number(row[colTaxas] || 0);
      totalLiquido += Number(row[colLiquido] || 0);
    });

    return { totalLiquido, totalTaxas, totalOperacoes: data.length };
  }
};


function getDashboardByCliente(clienteId) {
  validateInfrastructureOrThrow();
  const clientes = ClientesService.list();
  const ops = OperacoesService.list().filter(o => !clienteId || o.cliente_id === clienteId);
  const labels = [];
  const metas = [];
  const pontos = [];
  const valores = [];
  clientes.filter(c => !clienteId || c.uuid === clienteId).forEach(c => {
    const cOps = ops.filter(o => o.cliente_id === c.uuid && (o.status || 'Ativo') !== 'Inativo');
    const sumPontos = cOps.reduce((a,b)=>a + toNumber(b.pontos_pos,0) - toNumber(b.pontos_neg,0),0);
    const sumValores = cOps.reduce((a,b)=>a + toNumber(b.lucro_liquido,0),0);
    const meta = toNumber(c.capital_inicial_contrato,0) * Math.max(cOps.length,1);
    labels.push(c.nome);
    metas.push(meta);
    pontos.push(sumPontos);
    valores.push(sumValores);
  });
  return { labels, metas, pontos, valores };
}

// Wrappers para google.script.run (contrato estável com frontend)
function getClientes() { validateInfrastructureOrThrow(); return ClientesService.list(); }
function salvarCliente(clienteDTO) { validateInfrastructureOrThrow(); return ClientesService.save(clienteDTO); }
function getOperacoes() { validateInfrastructureOrThrow(); return OperacoesService.list(); }
function registrarOperacao(opDTO) { validateInfrastructureOrThrow(); return OperacoesService.save(opDTO); }
function getDashboardData() { validateInfrastructureOrThrow(); return DashboardService.getStats(); }

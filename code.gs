/**
 * SAE - Sistema Apollo Enterprise / Eficiente
 * Gestão de Investimentos / Índice Pessoal
 * Stack: Google Apps Script V8 + HtmlService + Google Sheets
 */

const SAE_CONFIG = {
  VALOR_PONTO_PADRAO: 0.20,
  META_PONTOS_PADRAO: 750,
  STATUS: ['Ativo', 'Inativo', 'Potencial']
};

const SCHEMA = {
  tbl_clientes: [
    'uuid', 'cliente_id', 'data_cadastro', 'status', 'nome', 'idade', 'telefone', 'email', 'cidade', 'estado',
    'redes_sociais', 'valor_mensalidade', 'vencimento_dia', 'capital_inicial_contrato', 'pagamento_valor',
    'data_pagamento', 'data_desligamento'
  ],
  tbl_operacoes: [
    'uuid', 'cliente_id', 'data_operacao', 'capital_inicial_contrato', 'n_contratos', 'valor_por_contrato',
    'pontos_pos', 'pontos_neg', 'percentual_ganho', 'take', 'stop'
  ],
  config: ['uuid', 'data_cadastro', 'status', 'meta_pontos', 'observacao'],
  auditoria: ['uuid', 'data_iso', 'entidade', 'entidade_id', 'acao', 'payload_json']
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SAE - Gestão de Investimentos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(name => ensureSheet(ss, name));
  ensureDefaultMeta_();
  return { success: true, message: 'Banco SAE sincronizado com sucesso.', schema: SCHEMA };
}

function ensureSheet(ss, name) {
  const headers = SCHEMA[name];
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    headers.forEach(h => {
      if (current.indexOf(h) === -1) sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
    });
  }
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
  return sheet;
}

function getSheetOrThrow_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error(`Aba obrigatória ausente: ${name}. Execute setupDatabase().`);
  return sheet;
}

function getRows_(name) {
  const sheet = getSheetOrThrow_(name);
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.map((row, rowIndex) => {
    const obj = { __row: rowIndex + 2 };
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function writeEntity_(name, entity, uuid) {
  const sheet = getSheetOrThrow_(name);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const payload = headers.map(h => Object.prototype.hasOwnProperty.call(entity, h) ? entity[h] : '');
  const rows = getRows_(name);
  const found = uuid ? rows.find(r => r.uuid === uuid) : null;
  if (found) {
    sheet.getRange(found.__row, 1, 1, headers.length).setValues([payload]);
  } else {
    sheet.appendRow(payload);
  }
  return entity;
}

function toIso_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime()) ? '' : date.toISOString();
}

function todayIso_() {
  return new Date().toISOString();
}

function toNumber_(value, fallback) {
  if (value === '' || value === null || value === undefined) return fallback;
  let normalized = value;
  if (typeof value === 'string') {
    const raw = value.trim();
    normalized = raw.indexOf(',') >= 0 ? raw.replace(/\./g, '').replace(',', '.') : raw;
  }
  const num = Number(normalized);
  return Number.isFinite(num) ? num : fallback;
}

function text_(value) {
  return String(value || '').trim();
}

function audit_(entidade, entidadeId, acao, payload) {
  const entity = {
    uuid: Utilities.getUuid(),
    data_iso: todayIso_(),
    entidade,
    entidade_id: entidadeId,
    acao,
    payload_json: JSON.stringify(payload || {})
  };
  writeEntity_('auditoria', entity);
}

function validateInfra_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const details = Object.keys(SCHEMA).map(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return { entidade: name, ok: false, erro: 'Aba ausente' };
    const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    const missing = SCHEMA[name].filter(h => headers.indexOf(h) === -1);
    return { entidade: name, ok: missing.length === 0, erro: missing.join(', ') };
  });
  const ok = details.every(d => d.ok);
  return { success: ok, timestamp_iso: todayIso_(), details };
}

function healthCheck() {
  return validateInfra_();
}

function runPreReleaseHealthCheck() {
  const report = validateInfra_();
  return { success: report.success, release_ready: report.success, checked_at_iso: todayIso_(), issues: report.details.filter(d => !d.ok) };
}

function assertInfra_() {
  const report = validateInfra_();
  if (!report.success) throw new Error('Infraestrutura inválida: ' + report.details.filter(d => !d.ok).map(d => `${d.entidade}: ${d.erro}`).join(' | '));
}

const ClientesService = {
  list() {
    return getRows_('tbl_clientes').map(c => normalizeClienteOut_(c));
  },

  save(input) {
    const dto = normalizeClienteIn_(input || {});
    if (!dto.nome) throw new Error('Nome é obrigatório.');
    if (!dto.cliente_id) throw new Error('ID do cliente é obrigatório.');
    if (!dto.telefone && !dto.email) throw new Error('Telefone ou e-mail é obrigatório para validação de duplicidade.');

    const clientes = getRows_('tbl_clientes');
    const duplicate = clientes.find(c => c.uuid !== dto.uuid && (
      (dto.email && text_(c.email).toLowerCase() === dto.email.toLowerCase()) ||
      (dto.telefone && onlyDigits_(c.telefone) === onlyDigits_(dto.telefone)) ||
      (dto.cliente_id && text_(c.cliente_id).toLowerCase() === dto.cliente_id.toLowerCase())
    ));
    if (duplicate) throw new Error('Cliente duplicado por ID, telefone ou e-mail.');

    const entity = {
      uuid: dto.uuid || Utilities.getUuid(),
      cliente_id: dto.cliente_id,
      data_cadastro: dto.data_cadastro || todayIso_(),
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
      capital_inicial_contrato: dto.capital_inicial_contrato,
      pagamento_valor: dto.pagamento_valor,
      data_pagamento: dto.data_pagamento,
      data_desligamento: dto.data_desligamento
    };
    writeEntity_('tbl_clientes', entity, dto.uuid);
    audit_('tbl_clientes', entity.uuid, dto.uuid ? 'update' : 'create', { cliente_id: entity.cliente_id, nome: entity.nome, status: entity.status });
    return { success: true, cliente: normalizeClienteOut_(entity) };
  },

  softDelete(uuid) {
    const cliente = this.list().find(c => c.uuid === uuid);
    if (!cliente) throw new Error('Cliente não encontrado.');
    cliente.status = cliente.status === 'Inativo' ? 'Ativo' : 'Inativo';
    return this.save(cliente);
  }
};

function normalizeClienteIn_(input) {
  return {
    uuid: text_(input.uuid),
    cliente_id: text_(input.cliente_id),
    data_cadastro: toIso_(input.data_cadastro) || todayIso_(),
    status: SAE_CONFIG.STATUS.indexOf(input.status) >= 0 ? input.status : 'Ativo',
    nome: text_(input.nome),
    idade: toNumber_(input.idade, 0),
    telefone: text_(input.telefone),
    email: text_(input.email).toLowerCase(),
    cidade: text_(input.cidade),
    estado: text_(input.estado).toUpperCase(),
    redes_sociais: text_(input.redes_sociais),
    valor_mensalidade: toNumber_(input.valor_mensalidade, 0),
    vencimento_dia: toNumber_(input.vencimento_dia, 10),
    capital_inicial_contrato: toNumber_(input.capital_inicial_contrato, 0),
    pagamento_valor: toNumber_(input.pagamento_valor, 0),
    data_pagamento: toIso_(input.data_pagamento),
    data_desligamento: toIso_(input.data_desligamento)
  };
}

function normalizeClienteOut_(c) {
  return {
    uuid: c.uuid,
    cliente_id: c.cliente_id,
    data_cadastro: toIso_(c.data_cadastro),
    status: c.status || 'Ativo',
    nome: c.nome || '',
    idade: toNumber_(c.idade, 0),
    telefone: c.telefone || '',
    email: c.email || '',
    cidade: c.cidade || '',
    estado: c.estado || '',
    redes_sociais: c.redes_sociais || '',
    valor_mensalidade: toNumber_(c.valor_mensalidade, 0),
    vencimento_dia: toNumber_(c.vencimento_dia, 10),
    capital_inicial_contrato: toNumber_(c.capital_inicial_contrato, 0),
    pagamento_valor: toNumber_(c.pagamento_valor, 0),
    data_pagamento: toIso_(c.data_pagamento),
    data_desligamento: toIso_(c.data_desligamento)
  };
}

function onlyDigits_(value) {
  return String(value || '').replace(/\D/g, '');
}

const OperacoesService = {
  list(filters) {
    const clientes = ClientesService.list();
    const map = {};
    clientes.forEach(c => map[c.cliente_id] = c);
    return getRows_('tbl_operacoes')
      .map(o => normalizeOperacaoOut_(o, map[o.cliente_id]))
      .filter(o => filterByPeriodo_(o.data_operacao, filters && filters.de, filters && filters.ate))
      .filter(o => !filters || !filters.cliente_id || o.cliente_id === filters.cliente_id);
  },

  save(input) {
    const cliente = ClientesService.list().find(c => c.cliente_id === text_(input && input.cliente_id));
    if (!cliente) throw new Error('Cliente não encontrado para lançamento de operação.');
    const dto = normalizeOperacaoIn_(input || {}, cliente);
    const entity = calculateOperacao_(dto);
    writeEntity_('tbl_operacoes', entity, dto.uuid);
    audit_('tbl_operacoes', entity.uuid, dto.uuid ? 'update' : 'create', { cliente_id: entity.cliente_id, take: entity.take, stop: entity.stop });
    return { success: true, operacao: normalizeOperacaoOut_(entity, cliente) };
  }
};

function normalizeOperacaoIn_(input, cliente) {
  return {
    uuid: text_(input.uuid),
    cliente_id: text_(input.cliente_id),
    data_operacao: toIso_(input.data_operacao) || todayIso_(),
    capital_inicial_contrato: toNumber_(input.capital_inicial_contrato, cliente.capital_inicial_contrato || 0),
    n_contratos: toNumber_(input.n_contratos, 1),
    valor_por_contrato: toNumber_(input.valor_por_contrato, SAE_CONFIG.VALOR_PONTO_PADRAO),
    pontos_pos: toNumber_(input.pontos_pos, 0),
    pontos_neg: toNumber_(input.pontos_neg, 0)
  };
}

function calculateOperacao_(dto) {
  const take = dto.pontos_pos * dto.n_contratos * dto.valor_por_contrato;
  const stop = dto.pontos_neg * dto.n_contratos * dto.valor_por_contrato;
  const saldo = take - stop;
  const base = dto.capital_inicial_contrato * dto.n_contratos;
  const percentual = base > 0 ? (saldo / base) * 100 : 0;
  return {
    uuid: dto.uuid || Utilities.getUuid(),
    cliente_id: dto.cliente_id,
    data_operacao: dto.data_operacao,
    capital_inicial_contrato: dto.capital_inicial_contrato,
    n_contratos: dto.n_contratos,
    valor_por_contrato: dto.valor_por_contrato,
    pontos_pos: dto.pontos_pos,
    pontos_neg: dto.pontos_neg,
    percentual_ganho: percentual,
    take,
    stop
  };
}

function normalizeOperacaoOut_(o, cliente) {
  const take = toNumber_(o.take, 0);
  const stop = toNumber_(o.stop, 0);
  return {
    uuid: o.uuid,
    cliente_id: o.cliente_id,
    nome: cliente ? cliente.nome : 'Cliente não encontrado',
    data_operacao: toIso_(o.data_operacao),
    capital_inicial_contrato: toNumber_(o.capital_inicial_contrato, 0),
    n_contratos: toNumber_(o.n_contratos, 0),
    valor_por_contrato: toNumber_(o.valor_por_contrato, SAE_CONFIG.VALOR_PONTO_PADRAO),
    pontos_pos: toNumber_(o.pontos_pos, 0),
    pontos_neg: toNumber_(o.pontos_neg, 0),
    percentual_ganho: toNumber_(o.percentual_ganho, 0),
    take,
    stop,
    saldo: take - stop
  };
}

const ConfigService = {
  list() {
    return getRows_('config').map(c => ({ uuid: c.uuid, data_cadastro: toIso_(c.data_cadastro), status: c.status || 'Ativo', meta_pontos: toNumber_(c.meta_pontos, SAE_CONFIG.META_PONTOS_PADRAO), observacao: c.observacao || '' }));
  },

  metaAtual() {
    const metas = this.list().filter(m => m.status === 'Ativo');
    return metas.length ? metas[metas.length - 1] : { meta_pontos: SAE_CONFIG.META_PONTOS_PADRAO };
  },

  saveMeta(input) {
    const meta = toNumber_(input && input.meta_pontos, SAE_CONFIG.META_PONTOS_PADRAO);
    if (meta <= 0) throw new Error('Meta deve ser maior que zero.');
    const entity = { uuid: Utilities.getUuid(), data_cadastro: todayIso_(), status: 'Ativo', meta_pontos: meta, observacao: text_(input && input.observacao) };
    writeEntity_('config', entity);
    audit_('config', entity.uuid, 'create_meta', { meta_pontos: meta });
    return { success: true, meta: entity };
  }
};

function ensureDefaultMeta_() {
  const rows = getRows_('config');
  if (rows.length === 0) ConfigService.saveMeta({ meta_pontos: SAE_CONFIG.META_PONTOS_PADRAO, observacao: 'Meta inicial padrão SAE' });
}

const DashboardService = {
  get(filters) {
    const operacoes = OperacoesService.list(filters || {});
    const clientes = ClientesService.list().filter(c => !filters || !filters.cliente_id || c.cliente_id === filters.cliente_id);
    const metaAtual = ConfigService.metaAtual().meta_pontos;
    const linhaMap = {};
    operacoes.forEach(o => {
      const dia = (o.data_operacao || '').slice(0, 10);
      if (!linhaMap[dia]) linhaMap[dia] = { data: dia, pontos_pos: 0, pontos_neg: 0 };
      linhaMap[dia].pontos_pos += o.pontos_pos;
      linhaMap[dia].pontos_neg += o.pontos_neg;
    });
    const barras = clientes.map(c => {
      const ops = operacoes.filter(o => o.cliente_id === c.cliente_id);
      return {
        cliente_id: c.cliente_id,
        nome: c.nome,
        pontos: ops.reduce((acc, o) => acc + o.pontos_pos - o.pontos_neg, 0),
        meta: metaAtual
      };
    });
    return {
      linha: Object.keys(linhaMap).sort().map(k => linhaMap[k]),
      barras,
      resumo: buildResumo_(operacoes)
    };
  }
};

function buildResumo_(operacoes) {
  return {
    total_take: operacoes.reduce((acc, o) => acc + o.take, 0),
    total_stop: operacoes.reduce((acc, o) => acc + o.stop, 0),
    saldo_total: operacoes.reduce((acc, o) => acc + o.saldo, 0),
    total_operacoes: operacoes.length
  };
}

const CarteiraService = {
  get(filters) {
    const clientes = ClientesService.list().filter(c => !filters || !filters.cliente_id || c.cliente_id === filters.cliente_id);
    const operacoes = OperacoesService.list(filters || {});
    return clientes.map(c => {
      const ops = operacoes.filter(o => o.cliente_id === c.cliente_id);
      return {
        cliente_id: c.cliente_id,
        nome: c.nome,
        status: c.status,
        saldo_total: ops.reduce((acc, o) => acc + o.saldo, 0),
        pontos_pos: ops.reduce((acc, o) => acc + o.pontos_pos, 0),
        pontos_neg: ops.reduce((acc, o) => acc + o.pontos_neg, 0),
        operacoes: ops
      };
    });
  }
};

function filterByPeriodo_(iso, de, ate) {
  const day = (iso || '').slice(0, 10);
  if (de && day < de) return false;
  if (ate && day > ate) return false;
  return true;
}

function getAppData(filters) {
  assertInfra_();
  const safeFilters = filters || {};
  return {
    clientes: ClientesService.list(),
    operacoes: OperacoesService.list(safeFilters),
    dashboard: DashboardService.get(safeFilters),
    carteira: CarteiraService.get(safeFilters),
    metas: ConfigService.list(),
    metaAtual: ConfigService.metaAtual()
  };
}

function salvarCliente(cliente) { assertInfra_(); return ClientesService.save(cliente); }
function softDeleteCliente(uuid) { assertInfra_(); return ClientesService.softDelete(uuid); }
function registrarOperacao(op) { assertInfra_(); return OperacoesService.save(op); }
function salvarMeta(meta) { assertInfra_(); return ConfigService.saveMeta(meta); }
function getDashboardData(filters) { assertInfra_(); return DashboardService.get(filters || {}); }
function getCarteira(filters) { assertInfra_(); return CarteiraService.get(filters || {}); }

// ============================================================
// MNTPREV DASHBOARD — Google Apps Script Backend
// NorteTech Energia | Manutenção Preventiva
// ============================================================

// ---- CONFIGURAÇÃO DE ABAS ----
const SHEETS = {
  CADASTRO: '🛻 Cadastro',
  HISTORICO: '📂 Histórico MNTPREV',
  ACOMPANHAMENTO: '📋 Acompanhamento por Item',
  RESUMO: '🚗 Resumo por Veículo',
  MESES_VENC: 'Meses Vencimento',
  PLANO: '⚙️ Plano Oficial',
  SEM_PLANO: '⛔ Sem Plano'
};

// ---- PONTO DE ENTRADA: Servir o Dashboard HTML ----
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Dashboard')
    .setTitle('MNTPREV — Dashboard NorteTech')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// FUNÇÃO PRINCIPAL: Carrega todos os dados do dashboard
// ============================================================
function getDashboardData(filters) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    filters = filters || {};

    const cadastro = readCadastro(ss);
    const resumo = readResumo(ss);
    const acompanhamento = readAcompanhamento(ss);
    const historico = readHistorico(ss);
    const plano = readPlano(ss);
    const semPlano = readSemPlano(ss);
    const mesesVenc = readMesesVencimento(ss);

    // Listas para filtros
    const filterOptions = buildFilterOptions(cadastro, resumo);

    // Aplicar filtros
    const resumoFiltrado = applyFilters(resumo, filters);
    const acompFiltrado = applyFilters(acompanhamento, filters);

    // KPIs principais
    const kpis = computeKPIs(resumoFiltrado, acompFiltrado);

    // Distribuição por status
    const statusDist = computeStatusDistribution(resumoFiltrado);

    // Distribuição por tipo de veículo
    const tiposDist = computeTiposDistribution(resumoFiltrado);

    // Top veículos mais críticos
    const topCriticos = getTopCriticos(resumoFiltrado, 15);

    // Histórico mensal (últimos 12 meses)
    const historicoMensal = computeHistoricoMensal(historico);

    // Acompanhamento por item (agrupado)
    const itensSummary = computeItensSummary(acompFiltrado);

    // Sem plano resumo
    const semPlanoSummary = {
      total: semPlano.length,
      porTipo: groupBy(semPlano, 'Tipo Veículo'),
      lista: semPlano.slice(0, 50)
    };

    // Meses vencimento
    const vencimentoSummary = computeVencimentoSummary(mesesVenc, resumoFiltrado);

    return {
      success: true,
      kpis,
      statusDist,
      tiposDist,
      topCriticos,
      historicoMensal,
      itensSummary,
      semPlanoSummary,
      vencimentoSummary,
      filterOptions,
      totalVeiculos: resumoFiltrado.length,
      timestamp: new Date().toISOString()
    };
  } catch (e) {
    return { success: false, error: e.message, stack: e.stack };
  }
}

// ============================================================
// LEITURA DAS ABAS
// ============================================================
function readSheet(ss, sheetName, headerRow) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  headerRow = headerRow || 1;
  const data = sh.getDataRange().getValues();
  if (data.length <= headerRow) return [];
  const headers = data[headerRow - 1].map(h => String(h).trim());
  const rows = [];
  for (let i = headerRow; i < data.length; i++) {
    const row = data[i];
    if (row.every(c => c === '' || c === null || c === undefined)) continue;
    const obj = {};
    headers.forEach((h, j) => { if (h) obj[h] = row[j]; });
    rows.push(obj);
  }
  return rows;
}

function readCadastro(ss) {
  return readSheet(ss, SHEETS.CADASTRO, 1);
}

function readResumo(ss) {
  const rows = readSheet(ss, SHEETS.RESUMO, 3);
  return rows.map(r => ({
    ...r,
    'KM Atual': toNum(r['KM Atual']),
    'KM Últ.MNT': toNum(r['KM Últ.MNT']),
    'KM p/ Próxima': toNum(r['KM p/ Próxima']),
    'Total Itens': toNum(r['Total Itens']),
    'Atrasados': toNum(r['Atrasados']),
    'Urgentes': toNum(r['Urgentes']),
    'Próximos': toNum(r['Próximos']),
    'OK': toNum(r['OK']),
    'Total Deveria': toNum(r['Total Deveria']),
    'Total Realizado': toNum(r['Total Realizado']),
    'Total Pendente': toNum(r['Total Pendente']),
    'Data Últ.MNT': toDateStr(r['Data Últ.MNT'])
  }));
}

function readAcompanhamento(ss) {
  const rows = readSheet(ss, SHEETS.ACOMPANHAMENTO, 3);
  return rows.map(r => ({
    ...r,
    'KM Atual': toNum(r['KM Atual']),
    'KM Últ.MNT': toNum(r['KM Últ.MNT']),
    'KM Desde Últ.': toNum(r['KM Desde Últ.']),
    'KM p/ Próxima': toNum(r['KM p/ Próxima']),
    'Intervalo(km)': toNum(r['Intervalo(km)']),
    'Qtd Deveria': toNum(r['Qtd Deveria']),
    'Qtd Realizada': toNum(r['Qtd Realizada']),
    'Qtd Pendente': toNum(r['Qtd Pendente']),
    'Data Últ.MNT': toDateStr(r['Data Últ.MNT'])
  }));
}

function readHistorico(ss) {
  const rows = readSheet(ss, SHEETS.HISTORICO, 3);
  return rows.map(r => ({
    ...r,
    'KM Manutenção': toNum(r['KM Manutenção']),
    'Data Manutenção': toDateStr(r['Data Manutenção'])
  }));
}

function readPlano(ss) {
  return readSheet(ss, SHEETS.PLANO, 2);
}

function readSemPlano(ss) {
  const rows = readSheet(ss, SHEETS.SEM_PLANO, 2);
  return rows.map(r => ({ ...r, 'KM Atual': toNum(r['KM Atual']) }));
}

function readMesesVencimento(ss) {
  const sh = ss.getSheetByName(SHEETS.MESES_VENC);
  if (!sh) return {};
  const data = sh.getDataRange().getValues();
  const result = {};
  // Layout em múltiplas colunas: Placa | Vencimento | (vazia) | Placa | Vencimento ...
  for (let r = 1; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c += 3) {
      const placa = String(data[r][c] || '').trim();
      const venc = String(data[r][c + 1] || '').trim();
      if (placa && venc && placa !== 'Placa' && placa !== 'Total Geral') {
        result[placa] = venc;
      }
    }
  }
  return result;
}

// ============================================================
// FILTROS
// ============================================================
function applyFilters(rows, filters) {
  if (!filters || Object.keys(filters).length === 0) return rows;
  return rows.filter(r => {
    if (filters.placas && filters.placas.length > 0) {
      if (!filters.placas.includes(r['Placa'])) return false;
    }
    if (filters.modelos && filters.modelos.length > 0) {
      if (!filters.modelos.includes(r['Modelo'])) return false;
    }
    if (filters.centrosCusto && filters.centrosCusto.length > 0) {
      const cc = r['Centro de Custo'] || r['Centro Custo'] || '';
      if (!filters.centrosCusto.includes(cc)) return false;
    }
    if (filters.tipos && filters.tipos.length > 0) {
      if (!filters.tipos.includes(r['Tipo'])) return false;
    }
    if (filters.status && filters.status.length > 0) {
      const st = r['Status Geral'] || r['Status'] || '';
      if (!filters.status.includes(st)) return false;
    }
    return true;
  });
}

function buildFilterOptions(cadastro, resumo) {
  const unique = (arr, key) => [...new Set(arr.map(r => r[key]).filter(Boolean))].sort();
  return {
    placas: unique(resumo, 'Placa'),
    modelos: unique(resumo, 'Modelo'),
    centrosCusto: unique(resumo, 'Centro de Custo'),
    tipos: unique(resumo, 'Tipo'),
    status: ['OK', 'ATRASADO', 'URGENTE', 'PRÓXIMO']
  };
}

// ============================================================
// CÁLCULOS / KPIs
// ============================================================
function computeKPIs(resumo, acomp) {
  const total = resumo.length;
  const atrasados = resumo.filter(r => (r['Status Geral'] || '').toUpperCase() === 'ATRASADO').length;
  const urgentes = resumo.filter(r => (r['Status Geral'] || '').toUpperCase() === 'URGENTE').length;
  const proximos = resumo.filter(r => (r['Status Geral'] || '').toUpperCase() === 'PRÓXIMO' || (r['Status Geral'] || '').toUpperCase() === 'PROXIMO').length;
  const ok = resumo.filter(r => (r['Status Geral'] || '').toUpperCase() === 'OK').length;
  const semDado = total - atrasados - urgentes - proximos - ok;

  const totalDeveria = resumo.reduce((s, r) => s + (r['Total Deveria'] || 0), 0);
  const totalRealizado = resumo.reduce((s, r) => s + (r['Total Realizado'] || 0), 0);
  const totalPendente = resumo.reduce((s, r) => s + (r['Total Pendente'] || 0), 0);

  const aderencia = totalDeveria > 0
    ? Math.round((totalRealizado / totalDeveria) * 100)
    : 0;

  // Veículos vencidos (Vencimento !== '' e não é NaN)
  const vencidos = resumo.filter(r => {
    const v = String(r['Unnamed: 22'] || r['Vencimento'] || '').trim();
    return v === 'Vencida';
  }).length;

  // Total de itens monitorados
  const totalItens = resumo.reduce((s, r) => s + (r['Total Itens'] || 0), 0);

  return {
    total, atrasados, urgentes, proximos, ok, semDado,
    totalDeveria, totalRealizado, totalPendente,
    aderencia, vencidos, totalItens,
    pctAtrasados: total > 0 ? Math.round((atrasados / total) * 100) : 0,
    pctOK: total > 0 ? Math.round((ok / total) * 100) : 0
  };
}

function computeStatusDistribution(resumo) {
  const dist = {};
  resumo.forEach(r => {
    const s = String(r['Status Geral'] || 'SEM DADO').toUpperCase().trim();
    dist[s] = (dist[s] || 0) + 1;
  });
  return dist;
}

function computeTiposDistribution(resumo) {
  return groupBy(resumo, 'Tipo');
}

function getTopCriticos(resumo, n) {
  return resumo
    .filter(r => r['Atrasados'] > 0 || r['Urgentes'] > 0)
    .sort((a, b) => {
      const scoreA = (a['Urgentes'] || 0) * 3 + (a['Atrasados'] || 0);
      const scoreB = (b['Urgentes'] || 0) * 3 + (b['Atrasados'] || 0);
      return scoreB - scoreA;
    })
    .slice(0, n)
    .map(r => ({
      placa: r['Placa'],
      modelo: r['Modelo'],
      tipo: r['Tipo'],
      centroCusto: r['Centro de Custo'],
      kmAtual: r['KM Atual'],
      kmProxima: r['KM p/ Próxima'],
      itemCritico: r['Item Mais Crítico'],
      atrasados: r['Atrasados'],
      urgentes: r['Urgentes'],
      proximos: r['Próximos'],
      ok: r['OK'],
      status: r['Status Geral'],
      vencimento: r['Unnamed: 22'] || ''
    }));
}

function computeHistoricoMensal(historico) {
  const meses = {};
  const hoje = new Date();
  const limite = new Date(hoje.getFullYear() - 1, hoje.getMonth(), 1);

  historico.forEach(r => {
    const d = r['Data Manutenção'];
    if (!d) return;
    const dt = new Date(d);
    if (isNaN(dt) || dt < limite) return;
    const key = dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0');
    if (!meses[key]) meses[key] = { preventiva: 0, corretiva: 0, total: 0 };
    const tipo = String(r['Tipo'] || '').toUpperCase();
    if (tipo === 'PREVENTIVA') meses[key].preventiva++;
    else if (tipo === 'CORRETIVA') meses[key].corretiva++;
    meses[key].total++;
  });

  return Object.entries(meses)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([mes, v]) => ({ mes, ...v }));
}

function computeItensSummary(acomp) {
  const itens = {};
  acomp.forEach(r => {
    const item = r['Item MNT'] || 'Desconhecido';
    if (!itens[item]) itens[item] = { atrasados: 0, ok: 0, total: 0, pendente: 0 };
    itens[item].total++;
    const st = String(r['Status'] || '').toUpperCase();
    if (st === 'ATRASADO') itens[item].atrasados++;
    else if (st === 'OK') itens[item].ok++;
    itens[item].pendente += (r['Qtd Pendente'] || 0);
  });
  return Object.entries(itens)
    .map(([item, v]) => ({ item, ...v }))
    .sort((a, b) => b.atrasados - a.atrasados)
    .slice(0, 20);
}

function computeVencimentoSummary(mesesVenc, resumo) {
  const grupos = {
    '<2025': [], 'janeiro': [], 'fevereiro': [], 'março': [],
    'abril': [], 'maio': [], 'junho': [], 'julho': [], 'agosto': [],
    'setembro': [], 'outubro': [], 'novembro': [], 'dezembro': [], 'Outros': []
  };

  resumo.forEach(r => {
    const v = String(r['Vencimento'] || '').trim().toLowerCase();
    if (grupos[v] !== undefined) grupos[v].push(r['Placa']);
    else if (v) grupos['Outros'].push(r['Placa']);
  });

  return Object.entries(grupos)
    .map(([periodo, placas]) => ({ periodo, qtd: placas.length }))
    .filter(e => e.qtd > 0);
}

// ============================================================
// FUNÇÕES PARA ABAS DETALHADAS (chamadas sob demanda)
// ============================================================

// Detalhes de um veículo específico
function getVehicleDetail(placa) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const acomp = readAcompanhamento(ss).filter(r => r['Placa'] === placa);
    const resumo = readResumo(ss).find(r => r['Placa'] === placa) || {};
    const hist = readHistorico(ss)
      .filter(r => r['Placa'] === placa)
      .sort((a, b) => new Date(b['Data Manutenção']) - new Date(a['Data Manutenção']))
      .slice(0, 30);
    return { success: true, placa, resumo, itens: acomp, historico: hist };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Tabela completa Acompanhamento por Item (paginada)
function getAcompanhamentoTable(filters, page, pageSize) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    page = page || 1;
    pageSize = pageSize || 50;
    const all = applyFilters(readAcompanhamento(ss), filters || {});
    const total = all.length;
    const start = (page - 1) * pageSize;
    const items = all.slice(start, start + pageSize);
    return {
      success: true, items, total, page, pageSize,
      totalPages: Math.ceil(total / pageSize)
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Tabela Resumo por Veículo (paginada)
function getResumoTable(filters, page, pageSize) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    page = page || 1;
    pageSize = pageSize || 50;
    const all = applyFilters(readResumo(ss), filters || {});
    const total = all.length;
    const start = (page - 1) * pageSize;
    const items = all.slice(start, start + pageSize).map(r => ({
      placa: r['Placa'],
      modelo: r['Modelo'],
      tipo: r['Tipo'],
      cidade: r['Cidade'],
      centroCusto: r['Centro de Custo'],
      kmAtual: r['KM Atual'],
      kmUltMnt: r['KM Últ.MNT'],
      dataUltMnt: r['Data Últ.MNT'],
      itemCritico: r['Item Mais Crítico'],
      kmProxima: r['KM p/ Próxima'],
      proxMntKm: r['Próx.MNT(km)'],
      totalItens: r['Total Itens'],
      atrasados: r['Atrasados'],
      urgentes: r['Urgentes'],
      proximos: r['Próximos'],
      ok: r['OK'],
      totalDeveria: r['Total Deveria'],
      totalReal: r['Total Realizado'],
      totalPend: r['Total Pendente'],
      status: r['Status Geral'],
      vencimento: r['Unnamed: 22'] || ''
    }));
    return {
      success: true, items, total, page, pageSize,
      totalPages: Math.ceil(total / pageSize)
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Sem Plano completo
function getSemPlanoFull() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return { success: true, items: readSemPlano(ss) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// UTILITÁRIOS
// ============================================================
function toNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function toDateStr(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Manaus', 'dd/MM/yyyy');
  const d = new Date(v);
  if (!isNaN(d)) return Utilities.formatDate(d, 'America/Manaus', 'dd/MM/yyyy');
  return String(v);
}

function groupBy(arr, key) {
  const result = {};
  arr.forEach(r => {
    const k = r[key] || 'Sem Categoria';
    result[k] = (result[k] || 0) + 1;
  });
  return result;
}
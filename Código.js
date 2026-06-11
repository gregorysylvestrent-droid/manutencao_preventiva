// ============================================================
// MNTPREV DASHBOARD — Google Apps Script Backend
// NorteTech Energia | Manutenção Preventiva
// Versão corrigida: calcula tudo das bases brutas
// ============================================================

// ---- NOMES DAS ABAS (aceita variações com/sem emoji) ----
const SHEETS = {
  CADASTRO:    ['🛻 Cadastro',                'Cadastro'],
  HISTORICO:   ['📂 Histórico MNTPREV',       'Histórico MNTPREV',   'Historico MNTPREV'],
  ACOMP:       ['📋 Acompanhamento por Item',  'Acompanhamento por Item'],
  RESUMO:      ['🚗 Resumo por Veículo',       'Resumo por Veículo',  'Resumo por Veiculo'],
  PLANO:       ['⚙️ Plano Oficial',            'Plano Oficial',       'Plano Mnt', 'Plano MNT'],
  SEM_PLANO:   ['⛔ Sem Plano',               'Sem Plano'],
  MESES_VENC:  ['Meses Vencimento']
};

// ---- PONTO DE ENTRADA ----
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Dashboard')
    .setTitle('MNTPREV — Dashboard NorteTech')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// FUNÇÃO PRINCIPAL — chamada pelo frontend
// ============================================================
function getDashboardData(filters) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    filters = filters || {};

    // ── 1. Lê as bases brutas (sem fórmulas) ──
    const cadastro    = readCadastro(ss);
    const plano       = readPlano(ss);
    const historico   = readHistorico(ss);
    const semPlanoSh  = readSemPlano(ss);

    // ── 2. Calcula visões computadas a partir das bases ──
    const computed    = buildComputedViews(cadastro, plano, historico);

    // ── 3. Tenta ler abas calculadas do Sheet; se vazias/com fórmulas, usa computed ──
    const resumoSheet = readResumoSheet(ss);
    const acompSheet  = readAcompSheet(ss);

    const resumoBase  = resumoSheet.length  ? resumoSheet  : computed.resumo;
    const acompBase   = acompSheet.length   ? acompSheet   : computed.acomp;
    const semPlanoBase = semPlanoSh.length  ? semPlanoSh   : computed.semPlano;

    // ── 4. Aplica filtros ──
    const resumoFilt  = applyFilters(resumoBase, filters);
    const acompFilt   = applyFilters(acompBase, filters);

    // ── 5. Agrega KPIs e métricas ──
    const kpis           = computeKPIs(resumoFilt);
    const statusDist     = computeStatusDist(resumoFilt);
    const tiposDist      = groupBy(resumoFilt, 'Tipo');
    const topCriticos    = getTopCriticos(resumoFilt, 15);
    const historicoMensal = computeHistoricoMensal(historico);
    const itensSummary   = computeItensSummary(acompFilt);
    const semPlanoSum    = {
      total: semPlanoBase.length,
      porTipo: groupBy(semPlanoBase, 'Tipo Veículo'),
      lista: semPlanoBase.slice(0, 100)
    };
    const mesesVenc      = readMesesVencimento(ss);
    const vencimentoSum  = computeVencimentoSummary(mesesVenc, resumoFilt);
    const filterOptions  = buildFilterOptions(resumoBase);

    return {
      success: true,
      kpis,
      statusDist,
      tiposDist,
      topCriticos,
      historicoMensal,
      itensSummary,
      semPlanoSummary: semPlanoSum,
      vencimentoSummary: vencimentoSum,
      filterOptions,
      totalVeiculos: resumoFilt.length,
      timestamp: new Date().toISOString()
    };
  } catch (e) {
    return { success: false, error: e.message, stack: e.stack };
  }
}

// ============================================================
// LEITURA DAS ABAS BRUTAS
// ============================================================

function getSheet(ss, names) {
  for (let i = 0; i < names.length; i++) {
    const s = ss.getSheetByName(names[i]);
    if (s) return s;
  }
  return null;
}

/** Converte dados brutos [linha, coluna] em array de objetos usando o cabeçalho da linha headerIdx (base-0) */
function sheetToObjects(data, headerIdx) {
  if (!data || data.length <= headerIdx) return [];
  const headers = data[headerIdx].map(h => String(h || '').trim());
  const result = [];
  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    // Ignora linhas completamente vazias
    if (!row || row.every(c => c === '' || c === null || c === undefined)) continue;
    // Ignora linhas cujo primeiro campo parece ser uma fórmula
    if (typeof row[0] === 'string' && row[0].startsWith('=')) continue;
    const obj = {};
    headers.forEach((h, j) => { if (h) obj[h] = row[j]; });
    result.push(obj);
  }
  return result;
}

function readCadastro(ss) {
  const sh = getSheet(ss, SHEETS.CADASTRO);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  return sheetToObjects(data, 0)
    .filter(r => r['Placa'])
    .map(r => ({
      Placa:            String(r['Placa'] || '').trim(),
      Modelo:           String(r['Descrição Modelo'] || r['Modelo'] || '').trim(),
      Tipo:             String(r['Tipo veiculo'] || r['Tipo Veículo'] || r['Tipo'] || '').trim(),
      Cidade:           String(r['Cidade'] || '').trim(),
      'Centro de Custo': String(r['Centro Custo'] || r['Centro de Custo'] || '').trim(),
      'Cod Centro Custo': String(r['Cod Centro Custo'] || '').trim(),
      Estado:           String(r['Estado'] || '').trim(),
      'Ano Fabricação': r['Ano Fabricação'] || ''
    }));
}

function readPlano(ss) {
  const sh = getSheet(ss, SHEETS.PLANO);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  // Cabeçalho na linha 2 (índice 1); linha 1 é título
  const rows = sheetToObjects(data, 1);
  return rows
    .filter(r => r['Modelo'] && (r['Item MNT (MNTPREV)'] || r['Item MNT']))
    .map(r => ({
      Modelo:          String(r['Modelo'] || '').trim(),
      'Item MNT':      String(r['Item MNT (MNTPREV)'] || r['Item MNT'] || '').trim(),
      'Intervalo(km)': toNum(r['Intervalo (km)'] || r['Intervalo(km)'] || 0),
      'Vida Útil Meses': toNum(r['Vida Útil Meses'] || 0),
      Tipo:            String(r['Tipo'] || '').trim()
    }));
}

function readHistorico(ss) {
  const sh = getSheet(ss, SHEETS.HISTORICO);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  // Cabeçalho na linha 3 (índice 2); linhas 1-2 são título/aviso
  const rows = sheetToObjects(data, 2);
  return rows
    .filter(r => r['Placa'])
    .map(r => ({
      Placa:              String(r['Placa'] || '').trim(),
      Modelo:             String(r['Modelo'] || '').trim(),
      'Centro de Custo':  String(r['Centro de Custo'] || '').trim(),
      'Item MNT':         String(r['Item MNT'] || '').trim(),
      'KM Manutenção':    toNum(r['KM Manutenção']),
      'Data Manutenção':  toDateStr(r['Data Manutenção']),
      Tipo:               String(r['Tipo'] || '').trim(),
      'Situação Manut.':  String(r['Situação Manut.'] || '').trim(),
      Oficina:            String(r['Oficina'] || '').trim()
    }));
}

function readSemPlano(ss) {
  const sh = getSheet(ss, SHEETS.SEM_PLANO);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  // Cabeçalho na linha 2 (índice 1); linha 1 é título
  return sheetToObjects(data, 1)
    .filter(r => r['Placa'])
    .map(r => ({
      Placa:             String(r['Placa'] || '').trim(),
      Modelo:            String(r['Modelo'] || '').trim(),
      'Tipo Veículo':    String(r['Tipo Veículo'] || '').trim(),
      Cidade:            String(r['Cidade'] || '').trim(),
      'Centro de Custo': String(r['Centro de Custo'] || '').trim(),
      'KM Atual':        toNum(r['KM Atual']),
      Ativo:             String(r['Ativo'] || '').trim()
    }));
}

/** Tenta ler a aba de Resumo por Veículo, descartando linhas que contenham fórmulas */
function readResumoSheet(ss) {
  const sh = getSheet(ss, SHEETS.RESUMO);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  // Cabeçalho linha 3 (índice 2)
  const rows = sheetToObjects(data, 2);
  // Filtra linhas que ainda têm fórmulas como string nos campos numéricos-chave
  return rows.filter(r =>
    r['Placa'] &&
    typeof r['KM Atual'] !== 'string' &&
    typeof r['Total Itens'] !== 'string'
  ).map(mapResumoRow);
}

/** Tenta ler a aba de Acompanhamento, descartando linhas com fórmulas */
function readAcompSheet(ss) {
  const sh = getSheet(ss, SHEETS.ACOMP);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const rows = sheetToObjects(data, 2);
  return rows.filter(r =>
    r['Placa'] &&
    typeof r['KM Últ.MNT'] !== 'string' &&
    !(typeof r['KM Últ.MNT'] === 'string' && r['KM Últ.MNT'].startsWith('='))
  ).map(r => ({
    ...r,
    'KM Atual':        toNum(r['KM Atual']),
    'KM Últ.MNT':     toNum(r['KM Últ.MNT']),
    'KM Desde Últ.':  toNum(r['KM Desde Últ.']),
    'KM p/ Próxima':  toNum(r['KM p/ Próxima']),
    'Próx.MNT(km)':   toNum(r['Próx.MNT(km)']),
    'Intervalo(km)':  toNum(r['Intervalo(km)']),
    'Qtd Deveria':    toNum(r['Qtd Deveria']),
    'Qtd Realizada':  toNum(r['Qtd Realizada']),
    'Qtd Pendente':   toNum(r['Qtd Pendente']),
    'Data Últ.MNT':   toDateStr(r['Data Últ.MNT'])
  }));
}

function readMesesVencimento(ss) {
  const sh = getSheet(ss, SHEETS.MESES_VENC);
  if (!sh) return {};
  const data = sh.getDataRange().getValues();
  const result = {};
  for (let r = 2; r < data.length; r++) {
    const row = data[r];
    for (let c = 0; c < row.length; c += 3) {
      const placa = String(row[c] || '').trim();
      const venc  = String(row[c + 1] || '').trim();
      if (placa && venc && placa !== 'Placa' && placa !== 'Total Geral') {
        result[placa] = venc;
      }
    }
  }
  return result;
}

// ============================================================
// CÁLCULO DAS VISÕES A PARTIR DAS BASES BRUTAS
// ============================================================
function buildComputedViews(cadastro, plano, historico) {
  // Índice: modelo → itens do plano
  const planoPorModelo = {};
  plano.forEach(p => {
    const m = p.Modelo;
    if (!planoPorModelo[m]) planoPorModelo[m] = [];
    planoPorModelo[m].push(p);
  });

  // Índice: "placa|item" → registro mais recente do histórico (maior KM)
  const ultimoHistorico = {};
  // Índice: "placa|item" → contagem de realizações
  const countHistorico = {};
  // Índice: placa → KM máximo encontrado no histórico
  const kmMaxHistorico = {};

  historico.forEach(h => {
    const key = h.Placa + '|' + h['Item MNT'];
    const km  = toNum(h['KM Manutenção']);
    const prev = ultimoHistorico[key];
    if (!prev || km > toNum(prev['KM Manutenção'])) {
      ultimoHistorico[key] = h;
    }
    countHistorico[key] = (countHistorico[key] || 0) + 1;
    if (!kmMaxHistorico[h.Placa] || km > kmMaxHistorico[h.Placa]) {
      kmMaxHistorico[h.Placa] = km;
    }
  });

  const acomp   = [];
  const resumo  = [];
  const semPlano = [];

  cadastro.forEach(v => {
    const itensPlano = planoPorModelo[v.Modelo] || [];
    // KM atual: melhor estimativa possível
    const kmAtual = kmMaxHistorico[v.Placa] || toNum(v['KM Atual']) || 0;

    if (!itensPlano.length) {
      semPlano.push({
        Placa:             v.Placa,
        Modelo:            v.Modelo,
        'Tipo Veículo':    v.Tipo,
        Cidade:            v.Cidade,
        'Centro de Custo': v['Centro de Custo'],
        'KM Atual':        kmAtual
      });
      return;
    }

    const linhasAcomp = itensPlano.map(item => {
      const key     = v.Placa + '|' + item['Item MNT'];
      const ult     = ultimoHistorico[key] || {};
      const kmUlt   = toNum(ult['KM Manutenção']) || 0;
      const intv    = toNum(item['Intervalo(km)']) || 0;
      const kmDesde = kmAtual - kmUlt;
      const proxKm  = intv ? kmUlt + intv : 0;
      const kmParaProx = proxKm ? proxKm - kmAtual : null;
      const qtdDev  = intv && kmAtual > 0 ? Math.max(0, Math.floor(kmAtual / intv)) : 0;
      const qtdReal = countHistorico[key] || 0;
      const qtdPend = Math.max(0, qtdDev - qtdReal);
      const status  = calcStatus(kmParaProx, intv, qtdPend);

      return {
        Placa:             v.Placa,
        Modelo:            v.Modelo,
        Tipo:              v.Tipo,
        Cidade:            v.Cidade,
        'Centro de Custo': v['Centro de Custo'],
        'Item MNT':        item['Item MNT'],
        'Intervalo(km)':   intv,
        'KM Atual':        kmAtual,
        'KM Últ.MNT':     kmUlt,
        'KM Desde Últ.':  kmDesde,
        'Próx.MNT(km)':   proxKm,
        'KM p/ Próxima':  kmParaProx !== null ? kmParaProx : '',
        'Qtd Deveria':     qtdDev,
        'Qtd Realizada':   qtdReal,
        'Qtd Pendente':    qtdPend,
        'Data Últ.MNT':   ult['Data Manutenção'] || '',
        Status:            status
      };
    });

    acomp.push(...linhasAcomp);

    // Linha de resumo por veículo
    const counts = { ATRASADO: 0, URGENTE: 0, PRÓXIMO: 0, OK: 0, 'SEM DADO': 0 };
    linhasAcomp.forEach(l => { counts[l.Status] = (counts[l.Status] || 0) + 1; });

    // Item mais crítico = pior status e menor KM para próxima
    const sorted = [...linhasAcomp].sort((a, b) =>
      statusScore(b.Status) - statusScore(a.Status) ||
      (toNum(a['KM p/ Próxima']) || 99999) - (toNum(b['KM p/ Próxima']) || 99999)
    );
    const critico = sorted[0] || {};
    const statusGeral = critico.Status || 'SEM DADO';

    const totalDev  = linhasAcomp.reduce((s, l) => s + l['Qtd Deveria'], 0);
    const totalReal = linhasAcomp.reduce((s, l) => s + l['Qtd Realizada'], 0);
    const totalPend = linhasAcomp.reduce((s, l) => s + l['Qtd Pendente'], 0);

    resumo.push(mapResumoRow({
      Placa:              v.Placa,
      Modelo:             v.Modelo,
      Tipo:               v.Tipo,
      Cidade:             v.Cidade,
      'Centro de Custo':  v['Centro de Custo'],
      'KM Atual':         kmAtual,
      'KM Últ.MNT':      toNum(critico['KM Últ.MNT']),
      'Data Últ.MNT':    critico['Data Últ.MNT'] || '',
      'Item Mais Crítico': critico['Item MNT'] || '',
      'KM p/ Próxima':   toNum(critico['KM p/ Próxima']) || 0,
      'Próx.MNT(km)':    toNum(critico['Próx.MNT(km)']) || 0,
      'Total Itens':      linhasAcomp.length,
      Atrasados:          counts.ATRASADO,
      Urgentes:           counts.URGENTE,
      Próximos:           counts.PRÓXIMO,
      OK:                 counts.OK,
      'Sem Dado':         counts['SEM DADO'],
      'Total Deveria':    totalDev,
      'Total Realizado':  totalReal,
      'Total Pendente':   totalPend,
      'Status Geral':     statusGeral,
      Vencimento:         ''  // preenchido depois via Meses Vencimento
    }));
  });

  return { resumo, acomp, semPlano };
}

function mapResumoRow(r) {
  return {
    ...r,
    'KM Atual':        toNum(r['KM Atual']),
    'KM Últ.MNT':     toNum(r['KM Últ.MNT']),
    'KM p/ Próxima':  toNum(r['KM p/ Próxima']),
    'Próx.MNT(km)':   toNum(r['Próx.MNT(km)']),
    'Total Itens':     toNum(r['Total Itens']),
    Atrasados:         toNum(r['Atrasados']),
    Urgentes:          toNum(r['Urgentes']),
    Próximos:          toNum(r['Próximos']),
    OK:                toNum(r['OK']),
    'Total Deveria':   toNum(r['Total Deveria']),
    'Total Realizado': toNum(r['Total Realizado']),
    'Total Pendente':  toNum(r['Total Pendente']),
    'Data Últ.MNT':   toDateStr(r['Data Últ.MNT'])
  };
}

function calcStatus(kmParaProx, intervalo, qtdPendente) {
  if (qtdPendente > 0) return 'ATRASADO';
  if (kmParaProx === null || kmParaProx === '') return 'SEM DADO';
  const kp = toNum(kmParaProx);
  if (kp < 0) return 'ATRASADO';
  const limUrg  = Math.max(1000, toNum(intervalo) * 0.1);
  const limProx = Math.max(3000, toNum(intervalo) * 0.2);
  if (kp <= limUrg)  return 'URGENTE';
  if (kp <= limProx) return 'PRÓXIMO';
  return 'OK';
}

function statusScore(s) {
  const u = String(s || '').toUpperCase();
  if (u === 'ATRASADO') return 4;
  if (u === 'URGENTE')  return 3;
  if (u === 'PRÓXIMO')  return 2;
  if (u === 'OK')       return 1;
  return 0;
}

// ============================================================
// FILTROS
// ============================================================
function applyFilters(rows, filters) {
  if (!filters || !Object.keys(filters).some(k => (filters[k] || []).length > 0)) return rows;
  return rows.filter(r => {
    if (filters.placas && filters.placas.length > 0 &&
        !filters.placas.includes(r.Placa)) return false;
    if (filters.modelos && filters.modelos.length > 0 &&
        !filters.modelos.includes(r.Modelo)) return false;
    if (filters.centrosCusto && filters.centrosCusto.length > 0) {
      const cc = r['Centro de Custo'] || '';
      if (!filters.centrosCusto.includes(cc)) return false;
    }
    if (filters.tipos && filters.tipos.length > 0 &&
        !filters.tipos.includes(r.Tipo)) return false;
    if (filters.status && filters.status.length > 0) {
      const st = r['Status Geral'] || r.Status || '';
      if (!filters.status.includes(st)) return false;
    }
    return true;
  });
}

function buildFilterOptions(resumo) {
  const unique = (key) => [...new Set(resumo.map(r => r[key]).filter(Boolean))].sort();
  return {
    placas:      unique('Placa'),
    modelos:     unique('Modelo'),
    centrosCusto: unique('Centro de Custo'),
    tipos:       unique('Tipo'),
    status:      ['OK', 'PRÓXIMO', 'URGENTE', 'ATRASADO', 'SEM DADO']
  };
}

// ============================================================
// KPIs e MÉTRICAS
// ============================================================
function computeKPIs(resumo) {
  const total      = resumo.length;
  const atrasados  = resumo.filter(r => String(r['Status Geral'] || '').toUpperCase() === 'ATRASADO').length;
  const urgentes   = resumo.filter(r => String(r['Status Geral'] || '').toUpperCase() === 'URGENTE').length;
  const proximos   = resumo.filter(r => String(r['Status Geral'] || '').toUpperCase() === 'PRÓXIMO').length;
  const ok         = resumo.filter(r => String(r['Status Geral'] || '').toUpperCase() === 'OK').length;
  const semDado    = total - atrasados - urgentes - proximos - ok;
  const totalDev   = resumo.reduce((s, r) => s + toNum(r['Total Deveria']), 0);
  const totalReal  = resumo.reduce((s, r) => s + toNum(r['Total Realizado']), 0);
  const totalPend  = resumo.reduce((s, r) => s + toNum(r['Total Pendente']), 0);
  const aderencia  = totalDev > 0 ? Math.round((totalReal / totalDev) * 100) : 0;
  const totalItens = resumo.reduce((s, r) => s + toNum(r['Total Itens']), 0);

  return {
    total, atrasados, urgentes, proximos, ok, semDado,
    totalDev, totalReal, totalPend, aderencia, totalItens,
    pctAtrasados: total > 0 ? Math.round((atrasados / total) * 100) : 0,
    pctOK:        total > 0 ? Math.round((ok / total) * 100) : 0
  };
}

function computeStatusDist(resumo) {
  const d = {};
  resumo.forEach(r => {
    const s = String(r['Status Geral'] || 'SEM DADO').toUpperCase().trim();
    d[s] = (d[s] || 0) + 1;
  });
  return d;
}

function getTopCriticos(resumo, n) {
  return resumo
    .filter(r => toNum(r.Atrasados) > 0 || toNum(r.Urgentes) > 0)
    .sort((a, b) =>
      (toNum(b.Urgentes) * 3 + toNum(b.Atrasados)) -
      (toNum(a.Urgentes) * 3 + toNum(a.Atrasados))
    )
    .slice(0, n)
    .map(r => ({
      placa:        r.Placa,
      modelo:       r.Modelo,
      tipo:         r.Tipo,
      cidade:       r.Cidade,
      centroCusto:  r['Centro de Custo'],
      kmAtual:      toNum(r['KM Atual']),
      kmProxima:    toNum(r['KM p/ Próxima']),
      itemCritico:  r['Item Mais Crítico'],
      atrasados:    toNum(r.Atrasados),
      urgentes:     toNum(r.Urgentes),
      proximos:     toNum(r.Próximos),
      ok:           toNum(r.OK),
      status:       r['Status Geral'],
      vencimento:   r.Vencimento || ''
    }));
}

function computeHistoricoMensal(historico) {
  const meses = {};
  const limite = new Date();
  limite.setFullYear(limite.getFullYear() - 1);

  historico.forEach(r => {
    const dt = parseDate(r['Data Manutenção']);
    if (!dt || isNaN(dt) || dt < limite) return;
    const key = dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0');
    if (!meses[key]) meses[key] = { preventiva: 0, corretiva: 0, total: 0 };
    const tipo = String(r.Tipo || '').toUpperCase();
    if (tipo === 'PREVENTIVA')  meses[key].preventiva++;
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
    if (!itens[item]) itens[item] = { atrasados: 0, ok: 0, urgentes: 0, total: 0, pendente: 0 };
    itens[item].total++;
    const st = String(r.Status || '').toUpperCase();
    if (st === 'ATRASADO') itens[item].atrasados++;
    else if (st === 'URGENTE') itens[item].urgentes++;
    else if (st === 'OK') itens[item].ok++;
    itens[item].pendente += toNum(r['Qtd Pendente']);
  });
  return Object.entries(itens)
    .map(([item, v]) => ({ item, ...v }))
    .sort((a, b) => b.atrasados - a.atrasados || b.urgentes - a.urgentes)
    .slice(0, 20);
}

function computeVencimentoSummary(mesesVenc, resumo) {
  const grupos = {};
  resumo.forEach(r => {
    const v = String(mesesVenc[r.Placa] || r.Vencimento || 'Sem Info').trim();
    grupos[v] = (grupos[v] || 0) + 1;
  });
  return Object.entries(grupos)
    .map(([periodo, qtd]) => ({ periodo, qtd }))
    .filter(e => e.qtd > 0)
    .sort((a, b) => b.qtd - a.qtd);
}

function groupBy(arr, key) {
  const r = {};
  arr.forEach(o => {
    const k = o[key] || 'Sem Categoria';
    r[k] = (r[k] || 0) + 1;
  });
  return r;
}

// ============================================================
// FUNÇÕES CHAMADAS SOB DEMANDA (tabelas paginadas, detalhe)
// ============================================================
function getResumoTable(filters, page, pageSize) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    page = page || 1;
    pageSize = pageSize || 50;
    const cadastro = readCadastro(ss);
    const plano    = readPlano(ss);
    const historico = readHistorico(ss);
    const computed = buildComputedViews(cadastro, plano, historico);
    const resumoSh = readResumoSheet(ss);
    const base     = resumoSh.length ? resumoSh : computed.resumo;
    const all      = applyFilters(base, filters || {});
    const total    = all.length;
    const start    = (page - 1) * pageSize;
    const items    = all.slice(start, start + pageSize).map(r => ({
      placa:       r.Placa,
      modelo:      r.Modelo,
      tipo:        r.Tipo,
      cidade:      r.Cidade,
      centroCusto: r['Centro de Custo'],
      kmAtual:     toNum(r['KM Atual']),
      kmUltMnt:    toNum(r['KM Últ.MNT']),
      dataUltMnt:  r['Data Últ.MNT'] || '',
      itemCritico: r['Item Mais Crítico'] || '',
      kmProxima:   toNum(r['KM p/ Próxima']),
      totalItens:  toNum(r['Total Itens']),
      atrasados:   toNum(r.Atrasados),
      urgentes:    toNum(r.Urgentes),
      proximos:    toNum(r.Próximos),
      ok:          toNum(r.OK),
      totalDev:    toNum(r['Total Deveria']),
      totalReal:   toNum(r['Total Realizado']),
      totalPend:   toNum(r['Total Pendente']),
      status:      r['Status Geral'] || '',
      vencimento:  r.Vencimento || ''
    }));
    return { success: true, items, total, page, pageSize, totalPages: Math.ceil(total / pageSize) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getAcompanhamentoTable(filters, page, pageSize) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    page = page || 1;
    pageSize = pageSize || 50;
    const cadastro  = readCadastro(ss);
    const plano     = readPlano(ss);
    const historico = readHistorico(ss);
    const computed  = buildComputedViews(cadastro, plano, historico);
    const acompSh   = readAcompSheet(ss);
    const base      = acompSh.length ? acompSh : computed.acomp;
    const all       = applyFilters(base, filters || {});
    const total     = all.length;
    const start     = (page - 1) * pageSize;
    const items     = all.slice(start, start + pageSize);
    return { success: true, items, total, page, pageSize, totalPages: Math.ceil(total / pageSize) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getVehicleDetail(placa) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cadastro  = readCadastro(ss);
    const plano     = readPlano(ss);
    const historico = readHistorico(ss);
    const computed  = buildComputedViews(cadastro, plano, historico);
    const acompSh   = readAcompSheet(ss);
    const resumoSh  = readResumoSheet(ss);
    const acompBase = acompSh.length ? acompSh : computed.acomp;
    const resumoBase = resumoSh.length ? resumoSh : computed.resumo;

    const itens   = acompBase.filter(r => r.Placa === placa);
    const resumo  = resumoBase.find(r => r.Placa === placa) || {};
    const hist    = historico
      .filter(r => r.Placa === placa)
      .sort((a, b) => parseDate(b['Data Manutenção']) - parseDate(a['Data Manutenção']))
      .slice(0, 50);

    return { success: true, placa, resumo, itens, historico: hist };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

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
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  if (typeof v === 'string' && v.startsWith('=')) return 0; // fórmula não calculada
  const clean = String(v).replace(/\./g, '').replace(',', '.').replace(/[^0-9.-]/g, '');
  const n = Number(clean);
  return isNaN(n) ? 0 : n;
}

function toDateStr(v) {
  if (!v) return '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, 'America/Manaus', 'dd/MM/yyyy');
  }
  if (typeof v === 'string' && v.startsWith('=')) return ''; // fórmula
  const d = parseDate(v);
  if (d && !isNaN(d.getTime())) return Utilities.formatDate(d, 'America/Manaus', 'dd/MM/yyyy');
  return String(v);
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const s = String(v).trim();
  if (s.startsWith('=')) return null;
  // dd/MM/yyyy
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m1) return new Date(Number(m1[3]), Number(m1[2]) - 1, Number(m1[1]));
  // yyyy-MM-dd
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

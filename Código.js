// ============================================================
// MNTPREV DASHBOARD — Google Apps Script Backend
// NorteTech Energia | Manutenção Preventiva
// ============================================================

// ---- CONFIGURAÇÃO DE ABAS ----
// Os nomes abaixo aceitam a nomenclatura original da planilha e os nomes
// gerados em versões anteriores com emojis, para facilitar a implantação.
const SHEETS = {
  CADASTRO: ['Cadastro', '🛻 Cadastro'],
  HISTORICO: ['Histórico MNTPREV', 'Historico MNTPREV', '📂 Histórico MNTPREV'],
  ACOMPANHAMENTO: ['Acompanhamento por Item', '📋 Acompanhamento por Item'],
  RESUMO: ['Resumo por Veículo', 'Resumo por Veiculo', '🚗 Resumo por Veículo'],
  ABASTECIMENTOS: ['Registro de Abastecimentos', 'Abastecimentos'],
  MESES_VENC: ['Meses Vencimento'],
  PLANO: ['Plano Mnt', 'Plano MNT', 'Plano Manutenção', 'Plano Manutencao', '⚙️ Plano Oficial'],
  SEM_PLANO: ['Sem Plano', '⛔ Sem Plano']
};

const HEADER_ROW_CANDIDATES = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];

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
    const historico = readHistorico(ss);
    const plano = readPlano(ss);
    const abastecimentos = readAbastecimentos(ss);

    // Se as abas calculadas existirem, elas são reaproveitadas. Se não existirem
    // ou estiverem vazias, o painel calcula tudo diretamente das bases coladas.
    const computed = buildComputedViews(cadastro, plano, historico, abastecimentos);
    const resumo = readResumo(ss);
    const acompanhamento = readAcompanhamento(ss);
    const semPlanoSheet = readSemPlano(ss);
    const mesesVenc = readMesesVencimento(ss);

    const resumoBase = hasRowsWithPlaca(resumo) ? resumo : computed.resumo;
    const acompanhamentoBase = hasRowsWithPlaca(acompanhamento) ? acompanhamento : computed.acompanhamento;
    const semPlanoBase = hasRowsWithPlaca(semPlanoSheet) ? semPlanoSheet : computed.semPlano;

    const filterOptions = buildFilterOptions(cadastro, resumoBase);
    const resumoFiltrado = applyFilters(resumoBase, filters);
    const acompFiltrado = applyFilters(acompanhamentoBase, filters);

    const kpis = computeKPIs(resumoFiltrado, acompFiltrado);
    const statusDist = computeStatusDistribution(resumoFiltrado);
    const tiposDist = computeTiposDistribution(resumoFiltrado);
    const topCriticos = getTopCriticos(resumoFiltrado, 15);
    const historicoMensal = computeHistoricoMensal(historico);
    const itensSummary = computeItensSummary(acompFiltrado);

    const semPlanoSummary = {
      total: semPlanoBase.length,
      porTipo: groupBy(semPlanoBase, 'Tipo Veículo'),
      lista: semPlanoBase.slice(0, 50)
    };

    const vencimentoSummary = Object.keys(mesesVenc).length
      ? computeVencimentoSummary(mesesVenc, resumoFiltrado)
      : computeVencimentoSummary({}, resumoFiltrado);

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
function getSheetByNameAny(ss, names) {
  const list = Array.isArray(names) ? names : [names];
  for (let i = 0; i < list.length; i++) {
    const sh = ss.getSheetByName(list[i]);
    if (sh) return sh;
  }
  return null;
}

function readSheet(ss, sheetNames, headerRow) {
  return readSheetFallback(ss, sheetNames, headerRow ? [headerRow] : HEADER_ROW_CANDIDATES);
}

function readSheetFallback(ss, sheetNames, headerRows) {
  const sh = getSheetByNameAny(ss, sheetNames);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const rowsToTry = uniqueHeaderRows((headerRows || []).concat(HEADER_ROW_CANDIDATES));
  for (let i = 0; i < rowsToTry.length; i++) {
    const rows = rowsFromData(data, rowsToTry[i]);
    if (rows.length) return rows;
  }
  return [];
}

function uniqueHeaderRows(rows) {
  const seen = {};
  return rows.filter(r => {
    const n = Number(r);
    if (!n || seen[n]) return false;
    seen[n] = true;
    return true;
  });
}

function rowsFromData(data, headerRow) {
  if (data.length < headerRow) return [];
  const headers = data[headerRow - 1].map(h => String(h).trim());
  if (headers.every(h => !h) || !looksLikeHeader(headers)) return [];
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

function normalizeRow(r) {
  const out = { ...r };
  Object.keys(r).forEach(k => {
    const nk = normKey(k);
    if (out[nk] === undefined) out[nk] = r[k];
  });
  return out;
}

function readCadastro(ss) {
  return readSheetFallback(ss, SHEETS.CADASTRO, HEADER_ROW_CANDIDATES).map(r => {
    const n = normalizeRow(r);
    const placa = firstVal(n, ['Placa', 'PLACA', 'placa']);
    const modelo = firstVal(n, ['Modelo', 'MODELO', 'modelo']);
    const centro = firstVal(n, ['Centro de Custo', 'Centro Custo', 'CC', 'centrodecusto', 'centrocusto']);
    const tipo = firstVal(n, ['Tipo Veículo', 'Tipo Veiculo', 'Tipo', 'tipoveiculo', 'tipo']);
    return {
      ...r,
      Placa: placa,
      Modelo: modelo,
      'Centro de Custo': centro,
      Tipo: tipo,
      'Tipo Veículo': tipo,
      'KM Atual': toNum(firstVal(n, ['KM Atual', 'KM', 'km_atual', 'kmatual']))
    };
  }).filter(r => r.Placa);
}

function readResumo(ss) {
  const rows = readSheetFallback(ss, SHEETS.RESUMO, [3, 1]);
  return rows.map(mapResumoRow);
}

function mapResumoRow(r) {
  const n = normalizeRow(r);
  const venc = firstVal(n, ['Vencimento', 'Mês Vencimento', 'Mes Vencimento', 'Unnamed: 22', 'vencimento']) || '';
  return {
    ...r,
    Placa: firstVal(n, ['Placa', 'PLACA', 'placa']),
    Modelo: firstVal(n, ['Modelo', 'MODELO', 'modelo']),
    Tipo: firstVal(n, ['Tipo', 'Tipo Veículo', 'Tipo Veiculo', 'tipo', 'tipoveiculo']),
    Cidade: firstVal(n, ['Cidade', 'Municipio', 'Município', 'cidade']),
    'Centro de Custo': firstVal(n, ['Centro de Custo', 'Centro Custo', 'CC', 'centrodecusto', 'centrocusto']),
    'Item Mais Crítico': firstVal(n, ['Item Mais Crítico', 'Item Mais Critico', 'Item Crítico', 'Item Critico', 'itemmaiscritico']),
    'Status Geral': firstVal(n, ['Status Geral', 'Status', 'statusgeral', 'status']),
    'KM Atual': toNum(firstVal(n, ['KM Atual', 'KM', 'Km Atual', 'Hodômetro', 'Hodometro', 'kmatual'])),
    'KM Últ.MNT': toNum(firstVal(n, ['KM Últ.MNT', 'KM Ult.MNT', 'KM Última MNT', 'KM Ultima MNT', 'KM Última Manutenção', 'KM Ultima Manutencao', 'kmultmnt'])),
    'KM p/ Próxima': toNum(firstVal(n, ['KM p/ Próxima', 'KM p/ Proxima', 'KM Para Próxima', 'KM Para Proxima', 'kmpProxima', 'kmpproxima'])),
    'Próx.MNT(km)': toNum(firstVal(n, ['Próx.MNT(km)', 'Prox.MNT(km)', 'Próxima MNT KM', 'Proxima MNT KM', 'proxmntkm'])),
    'Total Itens': toNum(firstVal(n, ['Total Itens', 'Itens', 'totalitens'])),
    Atrasados: toNum(firstVal(n, ['Atrasados', 'Atrasado', 'atrasados'])),
    Urgentes: toNum(firstVal(n, ['Urgentes', 'Urgente', 'urgentes'])),
    Próximos: toNum(firstVal(n, ['Próximos', 'Proximos', 'Próximo', 'Proximo', 'proximos'])),
    OK: toNum(firstVal(n, ['OK', 'Ok', 'ok'])),
    'Total Deveria': toNum(firstVal(n, ['Total Deveria', 'Deveria', 'totaldeveria'])),
    'Total Realizado': toNum(firstVal(n, ['Total Realizado', 'Realizado', 'totalrealizado'])),
    'Total Pendente': toNum(firstVal(n, ['Total Pendente', 'Pendente', 'totalpendente'])),
    'Data Últ.MNT': toDateStr(firstVal(n, ['Data Últ.MNT', 'Data Ult.MNT', 'Data Última MNT', 'Data Ultima MNT', 'Data Última Manutenção', 'Data Ultima Manutencao', 'dataultmnt'])),
    Vencimento: venc,
    'Unnamed: 22': venc
  };
}
function readAcompanhamento(ss) {
  const rows = readSheetFallback(ss, SHEETS.ACOMPANHAMENTO, [3, 1]);
  return rows.map(r => {
    const n = normalizeRow(r);
    return {
      ...r,
      Placa: firstVal(n, ['Placa', 'PLACA', 'placa']),
      Modelo: firstVal(n, ['Modelo', 'MODELO', 'modelo']),
      Tipo: firstVal(n, ['Tipo', 'Tipo Veículo', 'Tipo Veiculo', 'tipo', 'tipoveiculo']),
      'Centro de Custo': firstVal(n, ['Centro de Custo', 'Centro Custo', 'CC', 'centrodecusto', 'centrocusto']),
      'Item MNT': firstVal(n, ['Item MNT', 'Item', 'Serviço', 'Servico', 'Descrição', 'Descricao', 'itemmnt', 'item']),
      Status: firstVal(n, ['Status', 'Status Geral', 'status', 'statusgeral']),
      'KM Atual': toNum(firstVal(n, ['KM Atual', 'KM', 'Km Atual', 'Hodômetro', 'Hodometro', 'kmatual'])),
      'KM Últ.MNT': toNum(firstVal(n, ['KM Últ.MNT', 'KM Ult.MNT', 'KM Última MNT', 'KM Ultima MNT', 'kmultmnt'])),
      'KM Desde Últ.': toNum(firstVal(n, ['KM Desde Últ.', 'KM Desde Ult.', 'KM Desde Última', 'KM Desde Ultima', 'kmdesdeult'])),
      'KM p/ Próxima': toNum(firstVal(n, ['KM p/ Próxima', 'KM p/ Proxima', 'KM Para Próxima', 'KM Para Proxima', 'kmpproxima'])),
      'Próx.MNT(km)': toNum(firstVal(n, ['Próx.MNT(km)', 'Prox.MNT(km)', 'Próxima MNT KM', 'Proxima MNT KM', 'proxmntkm'])),
      'Intervalo(km)': toNum(firstVal(n, ['Intervalo(km)', 'Intervalo KM', 'Periodicidade KM', 'KM Intervalo', 'intervalokm'])),
      'Qtd Deveria': toNum(firstVal(n, ['Qtd Deveria', 'Quantidade Deveria', 'Deveria', 'qtddeveria'])),
      'Qtd Realizada': toNum(firstVal(n, ['Qtd Realizada', 'Quantidade Realizada', 'Realizada', 'qtdrealizada'])),
      'Qtd Pendente': toNum(firstVal(n, ['Qtd Pendente', 'Quantidade Pendente', 'Pendente', 'qtdpendente'])),
      'Data Últ.MNT': toDateStr(firstVal(n, ['Data Últ.MNT', 'Data Ult.MNT', 'Data Última MNT', 'Data Ultima MNT', 'dataultmnt']))
    };
  }).filter(r => r.Placa);
}
function readHistorico(ss) {
  const rows = readSheetFallback(ss, SHEETS.HISTORICO, [3, 1]);
  return rows.map(r => {
    const n = normalizeRow(r);
    return {
      ...r,
      Placa: firstVal(n, ['Placa', 'placa']),
      Modelo: firstVal(n, ['Modelo', 'modelo']),
      'Item MNT': firstVal(n, ['Item MNT', 'Item', 'Serviço', 'Servico', 'Descrição', 'Descricao', 'Manutenção', 'Manutencao', 'itemmnt', 'item']),
      Tipo: firstVal(n, ['Tipo', 'tipo']),
      'KM Manutenção': toNum(firstVal(n, ['KM Manutenção', 'KM Manutencao', 'KM', 'Km', 'Hodômetro', 'Hodometro', 'kmmanutencao'])),
      'Data Manutenção': toDateStr(firstVal(n, ['Data Manutenção', 'Data Manutencao', 'Data', 'datamanutencao']))
    };
  }).filter(r => r.Placa || r.Modelo || r['Item MNT']);
}

function readPlano(ss) {
  return readSheetFallback(ss, SHEETS.PLANO, [2, 1]).map(r => {
    const n = normalizeRow(r);
    return {
      ...r,
      Modelo: firstVal(n, ['Modelo', 'modelo']),
      'Item MNT': firstVal(n, ['Item MNT', 'Item', 'Serviço', 'Servico', 'Descrição', 'Descricao', 'Manutenção', 'Manutencao', 'itemmnt', 'item']),
      'Intervalo(km)': toNum(firstVal(n, ['Intervalo(km)', 'Intervalo KM', 'Periodicidade KM', 'KM Intervalo', 'KM', 'Quilometragem', 'intervalokm', 'periodicidadekm'])),
      'Intervalo(meses)': toNum(firstVal(n, ['Intervalo(meses)', 'Intervalo Meses', 'Periodicidade Meses', 'Meses', 'Prazo Meses', 'intervalomeses', 'periodicidademeses']))
    };
  }).filter(r => r.Modelo && r['Item MNT']);
}

function readAbastecimentos(ss) {
  return readSheetFallback(ss, SHEETS.ABASTECIMENTOS, HEADER_ROW_CANDIDATES).map(r => {
    const n = normalizeRow(r);
    return {
      ...r,
      Placa: firstVal(n, ['Placa', 'placa']),
      Data: toDateStr(firstVal(n, ['Data', 'Data Abastecimento', 'dataabastecimento'])),
      KM: toNum(firstVal(n, ['KM', 'Km', 'Hodômetro', 'Hodometro', 'Odometro', 'km']))
    };
  }).filter(r => r.Placa);
}

function readSemPlano(ss) {
  const rows = readSheetFallback(ss, SHEETS.SEM_PLANO, [2, 1]);
  return rows.map(r => {
    const n = normalizeRow(r);
    const tipo = firstVal(n, ['Tipo Veículo', 'Tipo Veiculo', 'Tipo', 'tipoveiculo', 'tipo']);
    return {
      ...r,
      Placa: firstVal(n, ['Placa', 'PLACA', 'placa']),
      Modelo: firstVal(n, ['Modelo', 'MODELO', 'modelo']),
      Tipo: tipo,
      'Tipo Veículo': tipo,
      'Centro de Custo': firstVal(n, ['Centro de Custo', 'Centro Custo', 'CC', 'centrodecusto', 'centrocusto']),
      'KM Atual': toNum(firstVal(n, ['KM Atual', 'KM', 'Km Atual', 'Hodômetro', 'Hodometro', 'kmatual']))
    };
  }).filter(r => r.Placa);
}
function readMesesVencimento(ss) {
  const sh = getSheetByNameAny(ss, SHEETS.MESES_VENC);
  if (!sh) return {};
  const data = sh.getDataRange().getValues();
  const result = {};
  for (let r = 1; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c += 3) {
      const placa = String(data[r][c] || '').trim();
      const venc = String(data[r][c + 1] || '').trim();
      if (placa && venc && placa !== 'Placa' && placa !== 'Total Geral') result[placa] = venc;
    }
  }
  return result;
}

// ============================================================
// CÁLCULOS A PARTIR DAS BASES COLADAS
// ============================================================
function buildComputedViews(cadastro, plano, historico, abastecimentos) {
  const planoPorModelo = {};
  plano.forEach(p => {
    const modelo = String(p.Modelo || '').trim();
    if (!planoPorModelo[modelo]) planoPorModelo[modelo] = [];
    planoPorModelo[modelo].push(p);
  });

  const kmAtualPorPlaca = latestKmByPlate(abastecimentos, cadastro);
  const histPorPlacaItem = latestMaintenanceByPlateItem(historico);
  const realizadosPorPlacaItem = countMaintenanceByPlateItem(historico);
  const acompanhamento = [];
  const resumo = [];
  const semPlano = [];

  cadastro.forEach(v => {
    const itensPlano = planoPorModelo[String(v.Modelo || '').trim()] || [];
    const kmAtual = kmAtualPorPlaca[v.Placa] || toNum(v['KM Atual']);
    if (!itensPlano.length) {
      semPlano.push({
        Placa: v.Placa,
        Modelo: v.Modelo,
        'Tipo Veículo': v['Tipo Veículo'] || v.Tipo,
        'Centro de Custo': v['Centro de Custo'],
        'KM Atual': kmAtual
      });
      return;
    }

    const rows = itensPlano.map(item => computeItemRow(v, item, kmAtual, histPorPlacaItem, realizadosPorPlacaItem));
    acompanhamento.push(...rows);
    resumo.push(computeResumoRow(v, kmAtual, rows));
  });

  return { resumo, acompanhamento, semPlano };
}

function computeItemRow(veiculo, item, kmAtual, histPorPlacaItem, realizadosPorPlacaItem) {
  const key = veiculo.Placa + '|' + item['Item MNT'];
  const last = histPorPlacaItem[key] || {};
  const kmUlt = toNum(last['KM Manutenção']);
  const intervaloKm = toNum(item['Intervalo(km)']);
  const kmDesdeUlt = kmUlt ? kmAtual - kmUlt : kmAtual;
  const proxKm = intervaloKm ? kmUlt + intervaloKm : 0;
  const kmParaProx = proxKm ? proxKm - kmAtual : 0;
  const qtdDeveria = intervaloKm ? Math.max(0, Math.floor(kmAtual / intervaloKm)) : 0;
  const qtdRealizada = realizadosPorPlacaItem[key] || 0;
  const qtdPendente = Math.max(0, qtdDeveria - qtdRealizada);
  const status = statusByKm(kmParaProx, intervaloKm, qtdPendente);

  return {
    Placa: veiculo.Placa,
    Modelo: veiculo.Modelo,
    Tipo: veiculo.Tipo || veiculo['Tipo Veículo'],
    Cidade: veiculo.Cidade || '',
    'Centro de Custo': veiculo['Centro de Custo'],
    'Item MNT': item['Item MNT'],
    'KM Atual': kmAtual,
    'KM Últ.MNT': kmUlt,
    'KM Desde Últ.': kmDesdeUlt,
    'Próx.MNT(km)': proxKm,
    'KM p/ Próxima': kmParaProx,
    'Intervalo(km)': intervaloKm,
    'Qtd Deveria': qtdDeveria,
    'Qtd Realizada': qtdRealizada,
    'Qtd Pendente': qtdPendente,
    'Data Últ.MNT': last['Data Manutenção'] || '',
    Status: status
  };
}

function computeResumoRow(veiculo, kmAtual, itens) {
  const counts = { ATRASADO: 0, URGENTE: 0, 'PRÓXIMO': 0, OK: 0 };
  itens.forEach(i => { counts[i.Status] = (counts[i.Status] || 0) + 1; });
  const criticos = itens.slice().sort((a, b) => statusScore(b.Status) - statusScore(a.Status) || a['KM p/ Próxima'] - b['KM p/ Próxima']);
  const itemCritico = criticos[0] || {};
  const statusGeral = itemCritico.Status || 'SEM DADO';
  const totalDeveria = itens.reduce((s, i) => s + (i['Qtd Deveria'] || 0), 0);
  const totalRealizado = itens.reduce((s, i) => s + (i['Qtd Realizada'] || 0), 0);
  const totalPendente = itens.reduce((s, i) => s + (i['Qtd Pendente'] || 0), 0);
  const vencimento = statusGeral === 'ATRASADO' ? 'Vencida' : nextMonthLabel(itemCritico['KM p/ Próxima']);

  return mapResumoRow({
    Placa: veiculo.Placa,
    Modelo: veiculo.Modelo,
    Tipo: veiculo.Tipo || veiculo['Tipo Veículo'],
    Cidade: veiculo.Cidade || '',
    'Centro de Custo': veiculo['Centro de Custo'],
    'KM Atual': kmAtual,
    'KM Últ.MNT': itemCritico['KM Últ.MNT'] || 0,
    'Data Últ.MNT': itemCritico['Data Últ.MNT'] || '',
    'Item Mais Crítico': itemCritico['Item MNT'] || '',
    'KM p/ Próxima': itemCritico['KM p/ Próxima'] || 0,
    'Próx.MNT(km)': itemCritico['Próx.MNT(km)'] || 0,
    'Total Itens': itens.length,
    Atrasados: counts.ATRASADO || 0,
    Urgentes: counts.URGENTE || 0,
    Próximos: counts['PRÓXIMO'] || 0,
    OK: counts.OK || 0,
    'Total Deveria': totalDeveria,
    'Total Realizado': totalRealizado,
    'Total Pendente': totalPendente,
    'Status Geral': statusGeral,
    Vencimento: vencimento
  });
}

function statusByKm(kmParaProx, intervaloKm, qtdPendente) {
  if (qtdPendente > 0 || (intervaloKm && kmParaProx < 0)) return 'ATRASADO';
  if (intervaloKm && kmParaProx <= Math.max(1000, intervaloKm * 0.1)) return 'URGENTE';
  if (intervaloKm && kmParaProx <= Math.max(3000, intervaloKm * 0.2)) return 'PRÓXIMO';
  return 'OK';
}

function statusScore(status) {
  const s = String(status || '').toUpperCase();
  if (s === 'ATRASADO') return 4;
  if (s === 'URGENTE') return 3;
  if (s === 'PRÓXIMO' || s === 'PROXIMO') return 2;
  if (s === 'OK') return 1;
  return 0;
}

function nextMonthLabel(kmParaProx) {
  if (kmParaProx < 0) return 'Vencida';
  const monthNames = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  const d = new Date();
  d.setMonth(d.getMonth() + Math.min(11, Math.max(0, Math.ceil(kmParaProx / 3000))));
  return monthNames[d.getMonth()];
}

function latestKmByPlate(abastecimentos, cadastro) {
  const latest = {};
  abastecimentos.forEach(a => {
    const t = parseDate(a.Data).getTime() || 0;
    if (!latest[a.Placa] || t >= latest[a.Placa].t) latest[a.Placa] = { km: toNum(a.KM), t };
  });
  cadastro.forEach(v => {
    if (!latest[v.Placa] && toNum(v['KM Atual'])) latest[v.Placa] = { km: toNum(v['KM Atual']), t: 0 };
  });
  const out = {};
  Object.keys(latest).forEach(p => { out[p] = latest[p].km; });
  return out;
}

function latestMaintenanceByPlateItem(historico) {
  const latest = {};
  historico.forEach(h => {
    const key = h.Placa + '|' + h['Item MNT'];
    const t = parseDate(h['Data Manutenção']).getTime() || 0;
    if (!latest[key] || t >= latest[key].t) latest[key] = { ...h, t };
  });
  return latest;
}

function countMaintenanceByPlateItem(historico) {
  const counts = {};
  historico.forEach(h => {
    const key = h.Placa + '|' + h['Item MNT'];
    counts[key] = (counts[key] || 0) + 1;
  });
  return counts;
}

// ============================================================
// FILTROS
// ============================================================
function applyFilters(rows, filters) {
  if (!filters || Object.keys(filters).length === 0) return rows;
  return rows.filter(r => {
    if (filters.placas && filters.placas.length > 0 && !filters.placas.includes(r.Placa)) return false;
    if (filters.modelos && filters.modelos.length > 0 && !filters.modelos.includes(r.Modelo)) return false;
    if (filters.centrosCusto && filters.centrosCusto.length > 0) {
      const cc = r['Centro de Custo'] || r['Centro Custo'] || '';
      if (!filters.centrosCusto.includes(cc)) return false;
    }
    if (filters.tipos && filters.tipos.length > 0 && !filters.tipos.includes(r.Tipo)) return false;
    if (filters.status && filters.status.length > 0) {
      const st = r['Status Geral'] || r.Status || '';
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
  const atrasados = resumo.filter(r => String(r['Status Geral'] || '').toUpperCase() === 'ATRASADO').length;
  const urgentes = resumo.filter(r => String(r['Status Geral'] || '').toUpperCase() === 'URGENTE').length;
  const proximos = resumo.filter(r => ['PRÓXIMO', 'PROXIMO'].includes(String(r['Status Geral'] || '').toUpperCase())).length;
  const ok = resumo.filter(r => String(r['Status Geral'] || '').toUpperCase() === 'OK').length;
  const semDado = total - atrasados - urgentes - proximos - ok;
  const totalDeveria = resumo.reduce((s, r) => s + (r['Total Deveria'] || 0), 0);
  const totalRealizado = resumo.reduce((s, r) => s + (r['Total Realizado'] || 0), 0);
  const totalPendente = resumo.reduce((s, r) => s + (r['Total Pendente'] || 0), 0);
  const aderencia = totalDeveria > 0 ? Math.round((totalRealizado / totalDeveria) * 100) : 0;
  const vencidos = resumo.filter(r => String(r.Vencimento || r['Unnamed: 22'] || '').trim().toUpperCase() === 'VENCIDA').length;
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
    .filter(r => r.Atrasados > 0 || r.Urgentes > 0)
    .sort((a, b) => {
      const scoreA = (a.Urgentes || 0) * 3 + (a.Atrasados || 0);
      const scoreB = (b.Urgentes || 0) * 3 + (b.Atrasados || 0);
      return scoreB - scoreA;
    })
    .slice(0, n)
    .map(r => ({
      placa: r.Placa,
      modelo: r.Modelo,
      tipo: r.Tipo,
      centroCusto: r['Centro de Custo'],
      kmAtual: r['KM Atual'],
      kmProxima: r['KM p/ Próxima'],
      itemCritico: r['Item Mais Crítico'],
      atrasados: r.Atrasados,
      urgentes: r.Urgentes,
      proximos: r.Próximos,
      ok: r.OK,
      status: r['Status Geral'],
      vencimento: r.Vencimento || r['Unnamed: 22'] || ''
    }));
}

function computeHistoricoMensal(historico) {
  const meses = {};
  const hoje = new Date();
  const limite = new Date(hoje.getFullYear() - 1, hoje.getMonth(), 1);

  historico.forEach(r => {
    const dt = parseDate(r['Data Manutenção']);
    if (isNaN(dt) || dt < limite) return;
    const key = dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0');
    if (!meses[key]) meses[key] = { preventiva: 0, corretiva: 0, total: 0 };
    const tipo = String(r.Tipo || '').toUpperCase();
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
    const st = String(r.Status || '').toUpperCase();
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
    '<2025': [], janeiro: [], fevereiro: [], março: [], abril: [], maio: [], junho: [], julho: [], agosto: [], setembro: [], outubro: [], novembro: [], dezembro: [], Vencida: [], Outros: []
  };

  resumo.forEach(r => {
    const placa = r.Placa;
    const raw = mesesVenc[placa] || r.Vencimento || r['Unnamed: 22'] || '';
    const v = String(raw).trim().toLowerCase();
    if (v === 'vencida') grupos.Vencida.push(placa);
    else if (grupos[v] !== undefined) grupos[v].push(placa);
    else if (v) grupos.Outros.push(placa);
  });

  return Object.entries(grupos)
    .map(([periodo, placas]) => ({ periodo, qtd: placas.length }))
    .filter(e => e.qtd > 0);
}

// ============================================================
// FUNÇÕES PARA ABAS DETALHADAS (chamadas sob demanda)
// ============================================================
function loadBaseTables(ss) {
  const cadastro = readCadastro(ss);
  const historico = readHistorico(ss);
  const plano = readPlano(ss);
  const abastecimentos = readAbastecimentos(ss);
  const computed = buildComputedViews(cadastro, plano, historico, abastecimentos);
  const resumo = readResumo(ss);
  const acompanhamento = readAcompanhamento(ss);
  const semPlano = readSemPlano(ss);
  return {
    cadastro,
    historico,
    plano,
    abastecimentos,
    resumo: hasRowsWithPlaca(resumo) ? resumo : computed.resumo,
    acompanhamento: hasRowsWithPlaca(acompanhamento) ? acompanhamento : computed.acompanhamento,
    semPlano: hasRowsWithPlaca(semPlano) ? semPlano : computed.semPlano
  };
}

function getVehicleDetail(placa) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = loadBaseTables(ss);
    const acomp = data.acompanhamento.filter(r => r.Placa === placa);
    const resumo = data.resumo.find(r => r.Placa === placa) || {};
    const hist = data.historico
      .filter(r => r.Placa === placa)
      .sort((a, b) => parseDate(b['Data Manutenção']) - parseDate(a['Data Manutenção']))
      .slice(0, 30);
    return { success: true, placa, resumo, itens: acomp, historico: hist };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getAcompanhamentoTable(filters, page, pageSize) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    page = page || 1;
    pageSize = pageSize || 50;
    const all = applyFilters(loadBaseTables(ss).acompanhamento, filters || {});
    const total = all.length;
    const start = (page - 1) * pageSize;
    const items = all.slice(start, start + pageSize);
    return { success: true, items, total, page, pageSize, totalPages: Math.ceil(total / pageSize) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getResumoTable(filters, page, pageSize) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    page = page || 1;
    pageSize = pageSize || 50;
    const all = applyFilters(loadBaseTables(ss).resumo, filters || {});
    const total = all.length;
    const start = (page - 1) * pageSize;
    const items = all.slice(start, start + pageSize).map(r => ({
      placa: r.Placa,
      modelo: r.Modelo,
      tipo: r.Tipo,
      cidade: r.Cidade,
      centroCusto: r['Centro de Custo'],
      kmAtual: r['KM Atual'],
      kmUltMnt: r['KM Últ.MNT'],
      dataUltMnt: r['Data Últ.MNT'],
      itemCritico: r['Item Mais Crítico'],
      kmProxima: r['KM p/ Próxima'],
      proxMntKm: r['Próx.MNT(km)'],
      totalItens: r['Total Itens'],
      atrasados: r.Atrasados,
      urgentes: r.Urgentes,
      proximos: r.Próximos,
      ok: r.OK,
      totalDeveria: r['Total Deveria'],
      totalReal: r['Total Realizado'],
      totalPend: r['Total Pendente'],
      status: r['Status Geral'],
      vencimento: r.Vencimento || r['Unnamed: 22'] || ''
    }));
    return { success: true, items, total, page, pageSize, totalPages: Math.ceil(total / pageSize) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getSemPlanoFull() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return { success: true, items: loadBaseTables(ss).semPlano };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// UTILITÁRIOS
// ============================================================
function toNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const clean = String(v).replace(/\./g, '').replace(',', '.').replace(/[^0-9.-]/g, '');
  const n = Number(clean);
  return isNaN(n) ? 0 : n;
}

function toDateStr(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Manaus', 'dd/MM/yyyy');
  const d = parseDate(v);
  if (!isNaN(d)) return Utilities.formatDate(d, 'America/Manaus', 'dd/MM/yyyy');
  return String(v);
}

function parseDate(v) {
  if (!v) return new Date('');
  if (v instanceof Date) return v;
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return new Date(s);
}

function groupBy(arr, key) {
  const result = {};
  arr.forEach(r => {
    const k = r[key] || 'Sem Categoria';
    result[k] = (result[k] || 0) + 1;
  });
  return result;
}

function hasRowsWithPlaca(rows) {
  return Array.isArray(rows) && rows.some(r => r && r.Placa);
}

function looksLikeHeader(headers) {
  const known = {
    placa: true, modelo: true, item: true, itemmnt: true, centrodecusto: true, centrocusto: true, status: true, statusgeral: true,
    km: true, kmatual: true, kmmanutencao: true, datamanutencao: true, data: true, tipo: true, tipoveiculo: true,
    intervalokm: true, intervalomeses: true, periodicidadekm: true, periodicidademeses: true, vencimento: true, totalitens: true,
    descricao: true, servico: true, manutencao: true, hodometro: true, odometro: true
  };
  let score = 0;
  headers.forEach(h => {
    const k = normKey(h);
    if (known[k] || k.indexOf('placa') >= 0 || k.indexOf('status') >= 0 || k.indexOf('intervalo') >= 0 || k.indexOf('periodicidade') >= 0) score++;
  });
  return score >= 2;
}

function normKey(k) {
  return String(k || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function firstVal(obj, keys) {
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const nk = normKey(k);
    if (obj[k] !== undefined && obj[k] !== '') return obj[k];
    if (obj[nk] !== undefined && obj[nk] !== '') return obj[nk];
  }

  const objKeys = Object.keys(obj);
  for (let i = 0; i < keys.length; i++) {
    const nk = normKey(keys[i]);
    if (nk.length < 4) continue;
    for (let j = 0; j < objKeys.length; j++) {
      const ok = normKey(objKeys[j]);
      if ((ok.indexOf(nk) >= 0 || nk.indexOf(ok) >= 0) && obj[objKeys[j]] !== undefined && obj[objKeys[j]] !== '') {
        return obj[objKeys[j]];
      }
    }
  }
  return '';
}

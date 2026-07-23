function novoLancamento_(dados) {
  Logger.log("novoLancamento_ recebeu: recorrente=" + dados.recorrente + " numMeses=" + dados.numMeses + " dataInicial=" + dados.dataInicial + " manual=" + dados.manual);
  inicializarContas_(); // garantir que CONTAS está populado
  if (!dados) throw new Error("Dados nao informados.");
  var valor = Math.abs(fixNum_(dados.valor));
  var descricao = String(dados.descricao || "").trim();
  var contaNome = String(dados.conta || "").trim();
  var direcao = String(dados.direcao || "Saida").trim();
  var tipo = String(dados.tipo || "DV").trim();
  // Todos os lançamentos entram obrigatoriamente com SIM em lancamentos
  var recorrente = dados.recorrente === true || dados.recorrente === "true";
  var numMeses = recorrente ? (parseInt(dados.numMeses) || 1) : 1;
  // Data informada pelo usuário (default = hoje)
  var dataInicial = dados.dataInicial
    ? new Date(dados.dataInicial + "T12:00:00")
    : new Date();

  if (!valor || valor <= 0) throw new Error("Valor deve ser maior que zero.");
  if (!descricao) throw new Error("Descricao e obrigatoria.");
  if (!contaNome) throw new Error("Conta e obrigatoria.");
  var contaCod = CONTAS[contaNome];
  if (!contaCod) throw new Error("Conta invalida: " + contaNome);
  var sinal = direcao === "Entrada" ? 1 : -1;
  var valorAjustado = valor * sinal;
  var agora = new Date();
  var tz = getTZ_();
  var dataFormatada = Utilities.formatDate(dataInicial, tz, "dd/MM/yyyy");
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var registros = 0;

  // Sempre vai para lancamentos com SIM (aguardando confirmação)
  var lancSheet = ss.getSheetByName("lancamentos");
  if (!lancSheet) throw new Error("Aba lancamentos nao encontrada.");

  var novasLinhas = [];
  var ultimoDiaRef = new Date(dataInicial.getFullYear(), dataInicial.getMonth() + 1, 0).getDate();
  var isUltimoDia = (dataInicial.getDate() === ultimoDiaRef);

  for (var m = 0; m < numMeses; m++) {
    var dataLanc;
    if (isUltimoDia) {
      dataLanc = new Date(dataInicial.getFullYear(), dataInicial.getMonth() + m + 1, 0);
    } else {
      dataLanc = new Date(dataInicial.getFullYear(), dataInicial.getMonth() + m, dataInicial.getDate());
    }
    novasLinhas.push([dataLanc, descricao, valorAjustado, contaCod, "SIM", tipo]);
    registros++;
  }

  var proximaLinha = lancSheet.getLastRow() + 1;
  if (proximaLinha < 2) proximaLinha = 2;
  lancSheet.getRange(proximaLinha, 1, novasLinhas.length, 6).setValues(novasLinhas);
  var totalLinhas = lancSheet.getLastRow() - 1;
  if (totalLinhas > 1) lancSheet.getRange(2, 1, totalLinhas, lancSheet.getLastColumn()).sort({ column: 1, ascending: true });
  SpreadsheetApp.flush();
  CacheService.getScriptCache().removeAll(["getDadosLancamentos", "getDadosDV", "getDashboard"]);

  return { ok: true, registros: registros, registro: {
    data: dataFormatada, descricao: descricao, conta: contaCod,
    contaNome: contaNome, direcao: direcao, valor: valor,
    valorAjustado: valorAjustado, tipo: tipo }};
}


function novoLancamentoTransf_(params) {
  var valor = Math.abs(fixNum_(params.valor));
  var descricao = String(params.descricao || "").trim();
  var origem = String(params.origem || "").trim();
  var destino = String(params.destino || "").trim();
  var manual = (params.manual === "true" || params.manual === true || params.manual === "SIM") ? "SIM" : "";
  if (!valor) throw new Error("Valor e obrigatorio.");
  if (!descricao) throw new Error("Descricao e obrigatoria.");
  if (!origem && !destino) throw new Error("Informe ao menos origem ou destino.");
  if (origem && !CONTAS[origem]) throw new Error("Conta de origem invalida: " + origem);
  if (destino && !CONTAS[destino]) throw new Error("Conta de destino invalida: " + destino);
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var agora = new Date();
  var tz = getTZ_();
  var registros = [];
  if (manual === "SIM") {
    var lancSheet = ss.getSheetByName("lancamentos");
    if (!lancSheet) throw new Error("Aba lancamentos nao encontrada.");
    if (origem) {
      var codOrig = CONTAS[origem];
      lancSheet.getRange(lancSheet.getLastRow() + 1, 1, 1, 6).setValues([[agora, descricao, -valor, codOrig, "SIM", "NA"]]);
      registros.push({ conta: codOrig, valor: -valor });
    }
    if (destino) {
      var codDest = CONTAS[destino];
      lancSheet.getRange(lancSheet.getLastRow() + 1, 1, 1, 6).setValues([[agora, descricao, valor, codDest, "SIM", "NA"]]);
      registros.push({ conta: codDest, valor: valor });
    }
    var totalLinhas = lancSheet.getLastRow() - 1;
    if (totalLinhas > 1) { lancSheet.getRange(2, 1, totalLinhas, lancSheet.getLastColumn()).sort({ column: 1, ascending: true }); }
    SpreadsheetApp.flush();
    CacheService.getScriptCache().removeAll(["getDadosLancamentos", "getDadosDV"]);
  } else {
    var bkpSheet2 = ss.getSheetByName("bkp");
    var saldosSheet2 = ss.getSheetByName("saldos");
    if (!bkpSheet2) throw new Error("Aba bkp nao encontrada.");
    if (!saldosSheet2) throw new Error("Aba saldos nao encontrada.");
    var lastSaldo2 = saldosSheet2.getLastRow();
    var codigos2 = lastSaldo2 >= 2 ? saldosSheet2.getRange(2, 1, lastSaldo2 - 1, 1).getValues() : [];
    var dataHoraExec2 = Utilities.formatDate(agora, tz, "dd/MM/yyyy HH:mm:ss");
    function atualizarSaldoT(contaCod, val) {
      for (var i = 0; i < codigos2.length; i++) {
        if (String(codigos2[i][0]).trim() === contaCod) {
          var cell = saldosSheet2.getRange(i + 2, 4);
          cell.setValue(Number(cell.getValue()) + val);
          saldosSheet2.getRange(i + 2, 2).setValue(agora);
          break;
        }
      }
      bkpSheet2.getRange(bkpSheet2.getLastRow() + 1, 1, 1, 6).setValues([[agora, descricao, val, contaCod, "NA", dataHoraExec2]]);
      registros.push({ conta: contaCod, valor: val });
    }
    if (origem) { atualizarSaldoT(CONTAS[origem], -valor); }
    if (destino) { atualizarSaldoT(CONTAS[destino], valor); }
    SpreadsheetApp.flush();
    CacheService.getScriptCache().removeAll(["getDadosLancamentos", "getDeltaBkp", "getDadosDV"]);
  }
  return { ok: true, registros: registros };
}


// Marca ou desmarca aguardando (SIM) na col E
// marcar=true → grava SIM; marcar=false → limpa col E
function desmarcarAguardando(params) {
  try {
    var linha  = parseInt(params.linha);
    var marcar = params.marcar === true || params.marcar === 'true';
    if (!linha || linha < 2) return { ok: false, erro: 'Linha inválida.' };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('lancamentos');
    if (!sheet) return { ok: false, erro: 'Aba lancamentos não encontrada.' };
    sheet.getRange(linha, 5).setValue(marcar ? 'SIM' : '');
    SpreadsheetApp.flush();
    CacheService.getScriptCache().removeAll(['getDadosLancamentos', 'getDadosDV', 'getDashboard', 'getDadosLancPendentes']);
    return { ok: true, linha: linha };
  } catch(e) {
    return { ok: false, erro: e.message };
  }
}
function confirmarLancamento_(params) {
  var linha = parseInt(params.linha);
  if (!linha || linha < 2) throw new Error("Linha invalida.");
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var lancSheet = ss.getSheetByName("lancamentos");
  var bkpSheet = ss.getSheetByName("bkp");
  var saldosSheet = ss.getSheetByName("saldos");
  if (!lancSheet) throw new Error("Aba lancamentos nao encontrada.");
  if (!bkpSheet) throw new Error("Aba bkp nao encontrada.");
  if (!saldosSheet) throw new Error("Aba saldos nao encontrada.");
  var reg = lancSheet.getRange(linha, 1, 1, 6).getValues()[0];
  var data = reg[0];
  var descricao = String(reg[1] || "").trim();
  var valor = fixNum_(reg[2]);
  var conta = String(reg[3] || "").trim();
  var tipo = String(reg[5] || "").trim();
  if (!descricao || !conta) throw new Error("Registro invalido na linha " + linha);
  var tz = getTZ_();
  var ultimaLinhaSaldos = saldosSheet.getLastRow();
  var dadosSaldos = saldosSheet.getRange(2, 1, ultimaLinhaSaldos - 1, 1).getValues();
  var encontrou = false;
  for (var i = 0; i < dadosSaldos.length; i++) {
    if (String(dadosSaldos[i][0]).trim() === conta) {
      var saldoCell = saldosSheet.getRange(i + 2, 4);
      saldoCell.setValue(fixNum_(saldoCell.getValue()) + valor);
      encontrou = true;
      break;
    }
  }
  if (!encontrou) throw new Error("Conta '" + conta + "' nao encontrada em saldos.");
  var agora = new Date();
  var dataHoraExec = Utilities.formatDate(agora, tz, "dd/MM/yyyy HH:mm:ss");
  bkpSheet.getRange(bkpSheet.getLastRow() + 1, 1, 1, 6).setValues([[data, descricao, valor, conta, tipo, dataHoraExec]]);
  lancSheet.getRange(linha, 5).setValue("");
  SpreadsheetApp.flush();
  CacheService.getScriptCache().removeAll(["getDadosLancamentos", "getDeltaBkp", "getDadosDV"]);
  return { ok: true, linha: linha };
}

function processarLancamentos() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var lancSheet = ss.getSheetByName("lancamentos");
  var bkpSheet  = ss.getSheetByName("bkp");
  var saldosSheet = ss.getSheetByName("saldos");
  if (!lancSheet) { return { erro: "Aba lancamentos nao encontrada." }; }
  var lastRow = lancSheet.getLastRow();
  if (lastRow < 2) { return { ok: true, processados: 0 }; }

  // Sort pela col A usando o sort nativo do Sheet (todas as colunas juntas, sem risco de desalinhamento)
  lancSheet.getRange(2, 1, lastRow - 1, lancSheet.getLastColumn()).sort({ column: 1, ascending: true });
  SpreadsheetApp.flush();

  lastRow = lancSheet.getLastRow();
  if (lastRow < 2) { return { ok: true, processados: 0 }; }

  // Ler 7 colunas (A=data B=desc C=valor D=conta E=aguard F=tipo G=idFatura)
  var numCols = lancSheet.getLastColumn();
  var dados = lancSheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  var tz = getTZ_();
  var processados = 0;
  var linhasRemover = [];
  var ultimaLinhaSaldos = saldosSheet.getLastRow();
  var codigos = saldosSheet.getRange(2, 1, ultimaLinhaSaldos - 1, 4).getValues();

  dados.forEach(function(l, i) {
    var data      = fixDate_(l[0]);
    var descricao = String(l[1] || "").trim();
    var valor     = Number(l[2]) || 0;
    var conta     = String(l[3] || "").trim();
    var aguardando = String(l[4] || "").trim().toUpperCase();
    var tipo      = String(l[5] || "").trim();
    if (!(data instanceof Date) || isNaN(data)) { return; }
    if (aguardando === "SIM") { return; }
    var dataLanc = new Date(data); dataLanc.setHours(0, 0, 0, 0);
    if (dataLanc > hoje) { return; }

    // Atualizar saldo
    for (var j = 0; j < codigos.length; j++) {
      if (String(codigos[j][0]).trim() === conta) {
        var saldoCell  = saldosSheet.getRange(j + 2, 4);
        var saldoAtual = Number(saldoCell.getValue()) || 0;
        var novoSaldo  = saldoAtual + valor;
        saldoCell.setValue(novoSaldo);
        saldosSheet.getRange(j + 2, 2).setValue(new Date()); // atualizar data
        SpreadsheetApp.flush();
        codigos[j][3] = novoSaldo;
        break;
      }
    }

    // Mover para bkp (6 colunas fixas)
    var dataHoraExec = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
    bkpSheet.getRange(bkpSheet.getLastRow() + 1, 1, 1, 6)
      .setValues([[data, descricao, valor, conta, tipo, dataHoraExec]]);
    linhasRemover.push(i + 2);
    processados++;
  });

  // Deletar de baixo para cima para não deslocar índices
  linhasRemover.sort(function(a, b) { return b - a; });
  linhasRemover.forEach(function(linha) { lancSheet.deleteRow(linha); });
  SpreadsheetApp.flush();

  CacheService.getScriptCache().removeAll([
    "getDadosLancamentos", "getDeltaBkp", "getDadosDV",
    "getDashboard", "getDadosLancPendentes"
  ]);
  return { ok: true, processados: processados };
}

function getDadosLancamentos() {
  var cached = cacheGet_('getDadosLancamentos');
  if (cached) { return cached; }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetLanc = ss.getSheetByName("lancamentos");
  if (!sheetLanc) throw new Error("Aba lancamentos nao encontrada.");
  var lastLanc = sheetLanc.getLastRow();
  if (lastLanc < 2) { return { meses: [], categorias: [], saldos: [] }; }
  var dados = sheetLanc.getRange(2, 1, lastLanc - 1, 6).getValues();
  var mensais = {}, detalhes = {}, cats = {};
  dados.forEach(function(l, i) {
    var data = fixDate_(l[0]);
    var descricao = l[1] ? String(l[1]).trim() : "Sem descricao";
    var valor = fixNum_(l[2]);
    var conta = l[3] ? String(l[3]).trim() : "Sem Conta";
    var tipo = l[5] ? String(l[5]).trim() : "";
    if (!(data instanceof Date) || isNaN(data)) { return; }
    var ano = data.getFullYear();
    var mes = pad_(data.getMonth() + 1);
    var chave = ano + "-" + mes;
    cats[conta] = true;
    if (!mensais[chave]) { mensais[chave] = {}; }
    mensais[chave][conta] = (mensais[chave][conta] || 0) + valor;
    if (!detalhes[chave]) { detalhes[chave] = []; }
    detalhes[chave].push({ linhaPlanilha: "L" + (i + 2), data: pad_(data.getDate()) + "/" + mes + "/" + ano, descricao: descricao, categoria: conta, conta: conta, tipo: tipo, valor: valor, aguardando: l[4] ? String(l[4]).trim().toUpperCase() : "" });
  });
  var saldos = [];
  var sheetSaldos = ss.getSheetByName("saldos");
  var tz = getTZ_();
  if (sheetSaldos && sheetSaldos.getLastRow() >= 2) {
    sheetSaldos.getRange(2, 1, sheetSaldos.getLastRow() - 1, 4).getValues().forEach(function(s) {
      var codigo = s[0] ? String(s[0]).trim() : "";
      var dataSaldo = fixDate_(s[1]);
      var desc = s[2] ? String(s[2]).trim() : "";
      var valor = fixNum_(s[3]);
      if (!codigo) { return; }
      var dataAtual = (dataSaldo instanceof Date && !isNaN(dataSaldo)) ? Utilities.formatDate(dataSaldo, tz, "dd/MM/yyyy") : "";
      saldos.push({ conta: codigo, descricao: desc, saldo: valor, dataAtual: dataAtual });
      if (dataSaldo instanceof Date && !isNaN(dataSaldo)) {
        // Saldos sempre agrupados no mês corrente (independente da data de atualização)
        var hoje = new Date();
        var ano  = hoje.getFullYear();
        var mes  = pad_(hoje.getMonth() + 1);
        var chave = ano + "-" + mes;
        cats[codigo] = true;
        if (!mensais[chave]) { mensais[chave] = {}; }
        mensais[chave][codigo] = (mensais[chave][codigo] || 0) + valor;
        if (!detalhes[chave]) { detalhes[chave] = []; }
        detalhes[chave].push({ linhaPlanilha: "S" + codigo, data: pad_(dataSaldo.getDate()) + "/" + pad_(dataSaldo.getMonth()+1) + "/" + dataSaldo.getFullYear(), descricao: "Saldo " + codigo + " (" + desc + ")", categoria: codigo, tipo: "NA", valor: valor });
      }
    });
  }
  var chaves = Object.keys(mensais).sort();
  var catsArray = Object.keys(cats).sort();
  var result = {
    meses: chaves.map(function(chave) {
      var p = chave.split('-'), porCat = {};
      catsArray.forEach(function(c) { porCat[c] = mensais[chave][c] || 0; });
      return { chaveOrdenacao: chave, mesAno: p[1] + "/" + p[0], valoresPorCategoria: porCat, detalhes: detalhes[chave] || [] };
    }),
    categorias: catsArray,
    saldos: saldos
  };
  cachePut_('getDadosLancamentos', result, 300); // 5 min
  return result;
}

function getDadosCartaoC() {
  var cached = cacheGet_('getDadosCartaoC');
  if (cached) { return cached; }
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("CartaoC");
  if (!sheet) throw new Error("Aba CartaoC nao encontrada.");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { return { meses: [], categorias: [] }; }
  var hoje = new Date(), limite = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  var mensais = {}, detalhes = {}, cats = {}, tiposSet = {};
  sheet.getRange(2, 2, lastRow - 1, 5).getValues().forEach(function(l, i) {
    var data = fixDate_(l[0]);
    if (!(data instanceof Date) || isNaN(data) || data < limite) { return; }
    // Colunas (a partir de col B): l[0]=DataFatura l[1]=Desc l[2]=Valor l[3]=Tipo l[4]=Cartao
    var val     = fixNum_(l[2]);
    var tipo    = l[3] ? String(l[3]).trim() : "";
    var cartao  = l[4] ? String(l[4]).trim() : "Sem Cartão";
    var desc    = l[1] ? String(l[1]).trim() : "Sem descrição";
    var ano = data.getFullYear(), mes = pad_(data.getMonth() + 1), chave = ano + "-" + mes;
    // Categoria = nome do cartão (para chips e colunas da tabela)
    cats[cartao] = true;
    tiposSet[tipo] = true;
    if (!mensais[chave]) { mensais[chave] = {}; }
    mensais[chave][cartao] = (mensais[chave][cartao] || 0) + val;
    if (!detalhes[chave]) { detalhes[chave] = []; }
    detalhes[chave].push({
      linhaPlanilha: i + 2,
      data: pad_(data.getDate()) + "/" + mes + "/" + ano,
      descricao: desc,
      categoria: cartao,
      tipo: tipo,
      valor: val
    });
  });
  var chaves = Object.keys(mensais).sort().slice(0, 12);
  var catsArray = Object.keys(cats).sort();
  var result2 = {
    meses: chaves.map(function(chave) {
      var p = chave.split('-'), porCat = {};
      catsArray.forEach(function(c) { porCat[c] = mensais[chave][c] || 0; });
      return { chaveOrdenacao: chave, mesAno: p[1] + "/" + p[0], valoresPorCategoria: porCat, detalhes: detalhes[chave] || [] };
    }),
    categorias: catsArray,
    tipos: Object.keys(tiposSet).sort()
  };
  cachePut_('getDadosCartaoC', result2);
  return result2;
}

function getDadosResumo() {
  var cached = cacheGet_('getDadosResumo');
  if (cached) { return cached; }
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("resumo");
  if (!sheet) throw new Error("Aba resumo nao encontrada.");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { return { despesas: [], rendas: [] }; }
  var despesas = [], rendas = [];
  sheet.getRange(2, 1, lastRow - 1, 9).getValues().forEach(function(l, i) {
    var linhaReal = i + 2;
    var dCat = l[2] ? String(l[2]).trim() : "", dVal = fixNum_(l[3]);
    if (dCat !== "" || dVal > 0) {
      despesas.push({ linhaPlanilha: linhaReal, tipo: l[0] ? String(l[0]).trim().toUpperCase() : "NAO DEFINIDO", descricao: l[1] ? String(l[1]).trim() : "Sem descricao", categoria: dCat || "Sem Categoria", valor: dVal });
    }
    var rCat = l[7] ? String(l[7]).trim() : "", rVal = fixNum_(l[8]);
    if (rCat !== "" || rVal > 0) {
      rendas.push({ linhaPlanilha: linhaReal, tipo: l[5] ? String(l[5]).trim().toUpperCase() : "RENDA", descricao: l[6] ? String(l[6]).trim() : "Sem descricao", categoria: rCat || "Sem Categoria", valor: rVal });
    }
  });
  var result3 = { despesas: despesas, rendas: rendas };
  cachePut_('getDadosResumo', result3);
  return result3;
}

function getDadosDV() {
  var cached = cacheGet_('getDadosDV');
  if (cached) { return cached; }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), hoje = new Date();
  var mA = hoje.getMonth(), aA = hoje.getFullYear();
  var mS = mA === 11 ? 0 : mA + 1, aS = mA === 11 ? aA + 1 : aA;
  var mS2 = mS === 11 ? 0 : mS + 1, aS2 = mS === 11 ? aS + 1 : aS;
  var totA = 0, totS = 0, totS2 = 0;
  function proc(dc, vc) {
    var d = fixDate_(dc instanceof Date ? dc : new Date(dc));
    if (!(d instanceof Date) || isNaN(d)) { return; }
    var v = fixNum_(vc);
    if (d.getMonth() === mA && d.getFullYear() === aA)   { totA  += v; }
    else if (d.getMonth() === mS  && d.getFullYear() === aS)  { totS  += v; }
    else if (d.getMonth() === mS2 && d.getFullYear() === aS2) { totS2 += v; }
  }
  var sBkp = ss.getSheetByName("bkp");
  if (sBkp && sBkp.getLastRow() >= 2) {
    sBkp.getRange(2, 1, sBkp.getLastRow() - 1, 5).getValues().forEach(function(l) {
      if (String(l[4]||'').trim().toUpperCase() === 'DV') { proc(l[0], l[2]); }
    });
  }
  var sLanc = ss.getSheetByName("lancamentos");
  if (sLanc && sLanc.getLastRow() >= 2) {
    sLanc.getRange(2, 1, sLanc.getLastRow() - 1, 7).getValues().forEach(function(l) {
      var idFatura = String(l[6]||'').trim();
      if (idFatura) return;
      if (String(l[5]||'').trim().toUpperCase() === 'DV') { proc(l[0], l[2]); }
    });
  }
  var sCart = ss.getSheetByName("CartaoC");
  if (sCart && sCart.getLastRow() >= 2) {
    sCart.getRange(2, 2, sCart.getLastRow() - 1, 5).getValues().forEach(function(l) {
      if (String(l[3] || "").trim().toUpperCase() === "DV") { proc(l[0], l[2]); }
    });
  }
  var dvPlanejado = 0;
  var shResumo = ss.getSheetByName('resumo');
  if (shResumo && shResumo.getLastRow() >= 2) {
    shResumo.getRange(2,1,shResumo.getLastRow()-1,9).getValues().forEach(function(r){
      var tipo = String(r[0]||'').trim().toUpperCase();
      if (/vari|^dv$/i.test(tipo)) dvPlanejado += Math.abs(fixNum_(r[3]));
    });
  }
  var result4 = { mesAtual: totA, mesSeguinte: totS, maisUm: totS2, totalPrevisto: totA + totS, dvPlanejado: dvPlanejado };
  cachePut_('getDadosDV', result4);
  return result4;
}

function getDeltaBkp() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("bkp");
  if (!sheet) throw new Error("Aba bkp nao encontrada.");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { return []; }
  var startRow = Math.max(2, lastRow - 99);
  var numRows = lastRow - startRow + 1;
  var dados = sheet.getRange(startRow, 1, numRows, 6).getValues();
  var result = [];
  for (var i = dados.length - 1; i >= 0; i--) {
    var l = dados[i];
    var data = fixDate_(l[0]);
    if (!(data instanceof Date) || isNaN(data)) { continue; }
    result.push({ linha: startRow + i, data: pad_(data.getDate()) + "/" + pad_(data.getMonth() + 1) + "/" + data.getFullYear(), descricao: l[1] ? String(l[1]).trim() : "", valor: fixNum_(l[2]), conta: l[3] ? String(l[3]).trim() : "", tipo: l[4] ? String(l[4]).trim() : "" });
  }
  return result;
}

function getDeltaCartao() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("CartaoC");
  if (!sheet) throw new Error("Aba CartaoC nao encontrada.");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { return []; }
  var startRow = Math.max(2, lastRow - 99);
  var numRows = lastRow - startRow + 1;
  var dados = sheet.getRange(startRow, 2, numRows, 5).getValues();
  var result = [];
  for (var i = dados.length - 1; i >= 0; i--) {
    var l = dados[i];
    var data = fixDate_(l[0]); // l[0]=DataFatura (col B)
    // Nova estrutura: l[0]=DataFatura l[1]=Desc l[2]=Valor l[3]=Tipo l[4]=Cartao
    result.push({ linha: startRow + i, data: (data instanceof Date && !isNaN(data)) ? pad_(data.getDate()) + "/" + pad_(data.getMonth() + 1) + "/" + data.getFullYear() : "", descricao: l[1] ? String(l[1]).trim() : "", valor: fixNum_(l[2]), tipo: l[3] ? String(l[3]).trim() : "", cartao: l[4] ? String(l[4]).trim() : "" });
  }
  return result;
}

// DRIVE_FOLDER_ID lido da aba config (chave: drive_folder_id)
function getDriveFolderId_() {
  return getConfig_()['drive_folder_id'] || '';
}
var NOME_JSON_BKP = "historico_bkp.json";
var NOME_JSON_CARTAO = "historico_cartao.json";

function normalizarBkp_(r) {
  // JSON antigo (2237 registros):
  // r.data      = data correta "dd/MM/yyyy"
  // r.conta     = valor numerico ex: "-5.4"
  // r.descricao = codigo da conta ex: "BK2"
  // r.valor     = sempre 0 (descricao textual perdida)
  // r.tipo      = sempre vazio
  var isAntigo = (Number(r.valor) === 0 && r.conta && !isNaN(Number(r.conta)));
  if (isAntigo) {
    return {
      linha:     r.linha || 0,
      data:      String(r.data || ""),
      descricao: "(historico antigo)",
      valor:     Number(r.conta) || 0,
      conta:     String(r.descricao || ""),
      tipo:      ""
    };
  }
  return r;
}

function normalizarCartao_(r) {
  if (r.descricao !== undefined) { return r; }
  var data = r.data || r.Data || r[0] || "";
  var descricao = r.descricao || r.Descricao || r[1] || "";
  var valor = r.valor !== undefined ? r.valor : (r.Valor !== undefined ? r.Valor : (r[4] !== undefined ? r[4] : 0));
  var tipo = r.tipo || r.Tipo || r[5] || "";
  var cartao = r.cartao || r.Cartao || r[6] || "";
  if (typeof data === 'object' && data !== null) {
    try { var d = new Date(data); data = pad_(d.getDate()) + "/" + pad_(d.getMonth()+1) + "/" + d.getFullYear(); } catch(e) { data = String(data); }
  }
  return { linha: r.linha || 0, data: String(data), descricao: String(descricao), valor: Number(valor) || 0, tipo: String(tipo), cartao: String(cartao) };
}

function getHistoricoBkp() {
  var historico = lerJsonDrive_(NOME_JSON_BKP).map(normalizarBkp_);
  var delta = getDeltaBkp();
  var linhasDelta = {};
  delta.forEach(function(r) { linhasDelta[r.linha] = true; });
  var resultado = delta.concat(historico.filter(function(r) { return !linhasDelta[r.linha]; }));
  // Atualiza o JSON no Drive com os dados mais recentes
  try { salvarJsonDrive_(NOME_JSON_BKP, resultado); } catch(err) { Logger.log("Erro ao salvar JSON bkp: " + err.message); }
  return resultado;
}

function getHistoricoCartao() {
  var historico = lerJsonDrive_(NOME_JSON_CARTAO).map(normalizarCartao_);
  var delta = getDeltaCartao();
  var linhasDelta = {};
  delta.forEach(function(r) { linhasDelta[r.linha] = true; });
  var resultado = delta.concat(historico.filter(function(r) { return !linhasDelta[r.linha]; }));
  // Atualiza o JSON no Drive com os dados mais recentes
  try { salvarJsonDrive_(NOME_JSON_CARTAO, resultado); } catch(err) { Logger.log("Erro ao salvar JSON cartao: " + err.message); }
  return resultado;
}

function lerJsonDrive_(nomeArquivo) {
  try {
    var folder = DriveApp.getFolderById(getDriveFolderId_());
    var files = folder.getFilesByName(nomeArquivo);
    if (!files.hasNext()) { return []; }
    return JSON.parse(files.next().getBlob().getDataAsString());
  } catch(err) {
    Logger.log("Erro ao ler JSON do Drive: " + err.message);
    return [];
  }
}

function salvarJsonDrive_(nomeArquivo, dados) {
  var folder = DriveApp.getFolderById(getDriveFolderId_());
  var content = JSON.stringify(dados);
  var files = folder.getFilesByName(nomeArquivo);
  if (files.hasNext()) {
    files.next().setContent(content);
  } else {
    folder.createFile(nomeArquivo, content, MimeType.PLAIN_TEXT);
  }
}

function testeRecorrente() {
  // Simula 3 meses a partir de 01/07/2026
  var resultado = novoLancamento_({
    valor: 100,
    descricao: "Teste recorrente",
    conta: "Itau",
    direcao: "Saida",
    tipo: "DF",
    manual: "SIM",
    recorrente: true,
    dataInicial: "2026-07-01",
    numMeses: 3
  });
  Logger.log("Resultado: " + JSON.stringify(resultado));
}

function diagnosticoHistorico() {
  var historico = lerJsonDrive_(NOME_JSON_BKP);
  if (!historico.length) { Logger.log("JSON vazio ou nao encontrado"); return; }
  Logger.log("Total: " + historico.length);
  // Mostrar 5 registros do meio para ver variedade
  [0, 1, 2, 100, 500].forEach(function(i) {
    if (historico[i]) Logger.log("R["+i+"]: " + JSON.stringify(historico[i]));
  });
}

function buscarLancamentos(termo) {
  try {
    inicializarContas_(); // garantir que CONTAS está populado
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("lancamentos");
    if (!sheet || sheet.getLastRow() < 2) { return []; }
    var dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    var tz = getTZ_();
    var termo_lower = termo.toLowerCase();
    var resultado = [];
    dados.forEach(function(l, i) {
      var descricao = String(l[1] || "").trim();
      if (descricao.toLowerCase().indexOf(termo_lower) === -1) { return; }
      var data = fixDate_(l[0]);
      var valor = fixNum_(l[2]);
      var contaCod = String(l[3] || "").trim();
      var tipo = String(l[5] || "").trim();
      // Descobrir nome da conta pelo codigo
      var contaNome = contaCod;
      Object.keys(CONTAS).forEach(function(k) { if (CONTAS[k] === contaCod) { contaNome = k; } });
      resultado.push({
        linha:     i + 2,
        data:      (data instanceof Date && !isNaN(data)) ? Utilities.formatDate(data, tz, "dd/MM/yyyy") : String(l[0]),
        descricao: descricao,
        valor:     valor,
        conta:     contaCod,
        contaNome: contaNome,
        tipo:      tipo
      });
    });
    // Ordenar por data desc (mais recente primeiro)
    resultado.sort(function(a, b) {
      var da = a.data.split("/").reverse().join("-");
      var db = b.data.split("/").reverse().join("-");
      return db.localeCompare(da);
    });
    return resultado.slice(0, 50); // max 50 resultados
  } catch(err) {
    throw new Error("Erro ao buscar: " + err.message);
  }
}

// Retorna sugestões para Projetar Recorrente:
// descrições únicas da aba lancamentos, ordenadas por frequência (mais recorrentes primeiro)
function getSugestoesProjetar() {
  try {
    inicializarContas_(); // garantir que CONTAS está populado
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("lancamentos");
    if (!sheet || sheet.getLastRow() < 2) { return []; }
    var dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    var tz    = getTZ_();
    var hoje  = new Date(); hoje.setHours(0,0,0,0);

    // Agrupar por descrição — guardar frequência total e a última ocorrência REAL (até hoje)
    var mapa = {};
    dados.forEach(function(l) {
      var desc     = String(l[1] || "").trim();
      var data     = fixDate_(l[0]);
      var valor    = fixNum_(l[2]);
      var contaCod = String(l[3] || "").trim();
      var tipo     = String(l[5] || "").trim();
      if (!desc) return;
      if (!(data instanceof Date) || isNaN(data)) return;
      data.setHours(0,0,0,0);

      if (!mapa[desc]) {
        mapa[desc] = {
          descricao: desc, valor: valor, conta: contaCod,
          contaNome: contaCod, tipo: tipo,
          dataObj: null, data: "", freq: 0, freqPassada: 0
        };
        Object.keys(CONTAS).forEach(function(k) {
          if (CONTAS[k] === contaCod) mapa[desc].contaNome = k;
        });
      }
      mapa[desc].freq++;
      // Conta como "recorrência real" apenas se já ocorreu (até hoje)
      if (data <= hoje) {
        mapa[desc].freqPassada++;
        // Manter a ocorrência mais recente DENTRE AS JÁ OCORRIDAS
        if (!mapa[desc].dataObj || data > mapa[desc].dataObj) {
          mapa[desc].dataObj = data;
          mapa[desc].data    = Utilities.formatDate(data, tz, "dd/MM/yyyy");
          mapa[desc].valor   = valor;
          mapa[desc].conta   = contaCod;
          mapa[desc].tipo    = tipo;
        }
      }
    });

    // Só sugerir descrições com pelo menos 1 ocorrência passada (já aconteceram)
    // e idealmente recorrentes (freqPassada >= 2) — mas mantemos as com freqPassada=1 no fim da lista
    var lista = Object.values(mapa).filter(function(m){ return m.freqPassada > 0; });

    lista.sort(function(a, b) {
      // Mais recorrentes primeiro (baseado em ocorrências passadas)
      if (b.freqPassada !== a.freqPassada) return b.freqPassada - a.freqPassada;
      // Empate: mais recente primeiro (comparação de Date real, não string)
      return b.dataObj - a.dataObj;
    });

    // Renomear freqPassada para freq no retorno (mantém compatibilidade com frontend)
    var resultado = lista.slice(0, 30).map(function(m) {
      return {
        descricao: m.descricao, valor: m.valor, conta: m.conta,
        contaNome: m.contaNome, tipo: m.tipo, data: m.data, freq: m.freqPassada
      };
    });
    return resultado;
  } catch(err) {
    Logger.log("getSugestoesProjetar erro: " + err.message);
    return [];
  }
}

function buscarCartaoC(termo) {
  try {
    inicializarCartoes_(); // garantir que CARTOES_POR_EMAIL está populado
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("CartaoC");
    if (!sheet || sheet.getLastRow() < 2) { return []; }
    var dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    var tz = getTZ_();
    var termo_lower = termo.toLowerCase();
    var resultado = [];
    dados.forEach(function(l, i) {
      var descricao = String(l[2] || "").trim();
      if (descricao.toLowerCase().indexOf(termo_lower) === -1) { return; }
      var dataFatura = fixDate_(l[1]);
      var valor = fixNum_(l[3]);
      var tipo = String(l[4] || "").trim();
      var cartaoCodigo = String(l[5] || "").trim();
      var cartaoNome = cartaoCodigo;
      Object.keys(CARTOES_POR_EMAIL).forEach(function(email) {
        (CARTOES_POR_EMAIL[email] || []).forEach(function(c) {
          if (c.codigo === cartaoCodigo) { cartaoNome = c.nome; }
        });
      });
      resultado.push({
        linha:        i + 2,
        dataFatura:   (dataFatura instanceof Date && !isNaN(dataFatura)) ? Utilities.formatDate(dataFatura, tz, "dd/MM/yyyy") : String(l[1]),
        descricao:    descricao,
        valor:        valor,
        tipo:         tipo,
        cartaoCodigo: cartaoCodigo,
        cartao:       cartaoNome
      });
    });
    resultado.sort(function(a, b) {
      var da = a.dataFatura.split("/").reverse().join("-");
      var db = b.dataFatura.split("/").reverse().join("-");
      return db.localeCompare(da);
    });
    return resultado.slice(0, 50);
  } catch(err) {
    throw new Error("Erro ao buscar CartaoC: " + err.message);
  }
}

// Sugestões para Projetar Cartão — descrições únicas ordenadas por frequência
function getSugestoesCartaoC() {
  try {
    inicializarCartoes_(); // garantir que CARTOES_POR_EMAIL está populado
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("CartaoC");
    if (!sheet || sheet.getLastRow() < 2) { return []; }
    var dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    var tz    = getTZ_();
    var hoje  = new Date(); hoje.setHours(0,0,0,0);
    var mapa  = {};
    dados.forEach(function(l) {
      var desc         = String(l[2] || "").trim();
      var dataFatura   = fixDate_(l[1]);
      var valor        = fixNum_(l[3]);
      var tipo         = String(l[4] || "").trim();
      var cartaoCodigo = String(l[5] || "").trim();
      if (!desc) return;
      if (!(dataFatura instanceof Date) || isNaN(dataFatura)) return;
      dataFatura.setHours(0,0,0,0);

      var cartaoNome = cartaoCodigo;
      Object.keys(CARTOES_POR_EMAIL).forEach(function(email) {
        (CARTOES_POR_EMAIL[email] || []).forEach(function(c) {
          if (c.codigo === cartaoCodigo) { cartaoNome = c.nome; }
        });
      });
      if (!mapa[desc]) {
        mapa[desc] = {
          descricao: desc, valor: valor, tipo: tipo,
          cartaoCodigo: cartaoCodigo, cartao: cartaoNome,
          dataObj: null, dataFatura: "", freq: 0, freqPassada: 0
        };
      }
      mapa[desc].freq++;
      // Conta como recorrência real apenas se a fatura já venceu (até hoje)
      if (dataFatura <= hoje) {
        mapa[desc].freqPassada++;
        if (!mapa[desc].dataObj || dataFatura > mapa[desc].dataObj) {
          mapa[desc].dataObj      = dataFatura;
          mapa[desc].dataFatura   = Utilities.formatDate(dataFatura, tz, "dd/MM/yyyy");
          mapa[desc].valor        = valor;
          mapa[desc].cartaoCodigo = cartaoCodigo;
          mapa[desc].cartao       = cartaoNome;
          mapa[desc].tipo         = tipo;
        }
      }
    });

    var lista = Object.values(mapa).filter(function(m){ return m.freqPassada > 0; });
    lista.sort(function(a, b) {
      if (b.freqPassada !== a.freqPassada) return b.freqPassada - a.freqPassada;
      return b.dataObj - a.dataObj;
    });

    var resultado = lista.slice(0, 30).map(function(m) {
      return {
        descricao: m.descricao, valor: m.valor, tipo: m.tipo,
        cartaoCodigo: m.cartaoCodigo, cartao: m.cartao,
        dataFatura: m.dataFatura, freq: m.freqPassada
      };
    });
    return resultado;
  } catch(err) {
    Logger.log("getSugestoesCartaoC erro: " + err.message);
    return [];
  }
}

function projetarCartaoC(dados) {
  inicializarCartoes_();
  try {
    var valor = Math.abs(fixNum_(dados.valor));
    var descricao = String(dados.descricao || "").trim();
    var cartao = String(dados.cartao || "").trim();
    var tipo = String(dados.tipo || "DV").trim();
    var direcao = String(dados.direcao || "Saida").trim();
    var numMeses = parseInt(dados.numParcelas) || 1;
    var dataFaturaInicial = dados.dataFaturaInicial ? new Date(dados.dataFaturaInicial + "T12:00:00") : new Date();
    var sinal = direcao === "Entrada" ? 1 : -1;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("CartaoC");
    if (!sheet) throw new Error("Aba CartaoC nao encontrada.");
    var agora = new Date();
    var novasLinhas = [];
    for (var m = 0; m < numMeses; m++) {
      var dataFatura = new Date(dataFaturaInicial.getFullYear(), dataFaturaInicial.getMonth() + m, dataFaturaInicial.getDate());
      var row = new Array(6).fill("");
      row[0] = agora;
      row[1] = dataFatura;
      row[2] = descricao;
      row[3] = valor * sinal;
      row[4] = tipo;
      row[5] = cartao;
      novasLinhas.push(row);
    }
    var ultimaLinha = sheet.getLastRow() + 1;
    sheet.getRange(ultimaLinha, 1, novasLinhas.length, 6).setValues(novasLinhas);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().removeAll(["getDadosCartaoC", "getDeltaCartao"]);
    return { ok: true, registros: novasLinhas.length };
  } catch(err) {
    return { erro: err.message };
  }
}

function diagnosticoFaturas() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var lancSheet = ss.getSheetByName("lancamentos");
  var fatSheet  = ss.getSheetByName("faturas");

  // Ver col G de lancamentos
  Logger.log("=== COL G de lancamentos ===");
  if (lancSheet && lancSheet.getLastRow() >= 2) {
    var dadosL = lancSheet.getRange(2, 1, lancSheet.getLastRow()-1, 7).getValues();
    dadosL.forEach(function(r, i) {
      var idFat = String(r[6]||"").trim();
      if (idFat) { Logger.log("Linha " + (i+2) + ": " + idFat + " | desc=" + r[1]); }
    });
  }

  // Ver aba faturas
  Logger.log("=== ABA FATURAS ===");
  if (fatSheet && fatSheet.getLastRow() >= 2) {
    var dadosF = fatSheet.getRange(2, 1, fatSheet.getLastRow()-1, 8).getValues();
    dadosF.forEach(function(r, i) {
      if (String(r[6]||"") !== "config") {
        Logger.log("L" + (i+2) + ": " + r[0] + " | mes=" + r[3] + " | calc=" + r[4] + " | linhaLanc=" + r[7]);
      }
    });
  }
}


function getVencCC() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("VencCC");
  if (!sheet || sheet.getLastRow() < 2) { return []; }
  var tz = getTZ_();
  var numCols = Math.max(sheet.getLastColumn(), 8);
  var dados = sheet.getRange(2, 1, sheet.getLastRow()-1, numCols).getValues();
  var resultado = [];
  dados.forEach(function(r, i) {
    var ultimos  = String(r[0]||"").trim();
    var dia      = Number(r[1]) || 0;
    var proxVenc = fixDate_(r[2]);
    var nome     = String(r[3]||"").trim();
    var codigo   = String(r[4]||"").trim();
    var emails   = String(r[5]||"").trim();
    var chave    = String(r[6]||"").trim();
    var conta    = String(r[7]||"").trim() || "BK1";
    if (!ultimos) { return; }
    var proxVencStr = (proxVenc instanceof Date && !isNaN(proxVenc))
      ? Utilities.formatDate(proxVenc, tz, "dd/MM/yyyy") : "";
    var proxVencISO = (proxVenc instanceof Date && !isNaN(proxVenc))
      ? Utilities.formatDate(proxVenc, tz, "yyyy-MM-dd") : "";
    resultado.push({
      linha:       i + 2,
      ultimos:     ultimos,
      dia:         dia,
      proxVenc:    proxVencStr,
      proxVencISO: proxVencISO,
      nome:        nome,
      codigo:      codigo,
      emails:      emails,
      chave:       chave,
      conta:       conta
    });
  });
  return resultado;
}

function salvarVencCC(dados) {
  try {
    var ultimos  = String(dados.ultimos||"").trim();
    var dia      = parseInt(dados.dia) || 0;
    var proxVenc = dados.proxVenc ? new Date(dados.proxVenc + "T12:00:00") : null;
    var nome     = String(dados.nome||"").trim();
    var codigo   = String(dados.codigo||"").trim();
    var emails   = String(dados.emails||"").trim();
    var chave    = String(dados.chave||"").trim();
    var conta    = String(dados.conta||"").trim() || "BK1";
    var linha    = dados.linha ? parseInt(dados.linha) : 0;
    if (!ultimos) throw new Error("Ultimos digitos obrigatorios.");
    if (!dia || dia < 1 || dia > 31) throw new Error("Dia invalido.");
    if (!proxVenc) throw new Error("Data de vencimento obrigatoria.");
    if (!nome)   throw new Error("Nome obrigatorio.");
    if (!codigo) throw new Error("Codigo do cartao obrigatorio.");
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("VencCC");
    if (!sheet) throw new Error("Aba VencCC nao encontrada.");
    var row = [ultimos, dia, proxVenc, nome, codigo, emails, chave, conta];
    if (linha >= 2) {
      sheet.getRange(linha, 1, 1, 8).setValues([row]);
    } else {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, 8).setValues([row]);
    }
    SpreadsheetApp.flush();
    // Reinicializar cartões para refletir mudanças
    CARTOES_POR_EMAIL = {};
    FINAL_CARTAO = {};
    CHAVES_CARTAO = {};
    CHAVES_CARTAO_CODIGO = {};
    CONTA_POR_CARTAO = {};
    return { ok: true, dados: getVencCC() };
  } catch(err) {
    return { erro: err.message };
  }
}


// ── CONTAS ────────────────────────────────────────────────────────
function getContas() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("contas");
  if (!sheet || sheet.getLastRow() < 2) {
    // Retornar contas hardcoded se aba nao existe
    return [
      { linha:0, codigo:"BK1", nome:"Itau",           tipo:"banco" },
      { linha:0, codigo:"BK2", nome:"Mercado Pago",    tipo:"banco" },
      { linha:0, codigo:"BK3", nome:"Banco do Brasil", tipo:"banco" },
      { linha:0, codigo:"CX",  nome:"Dinheiro",        tipo:"carteira" },
      { linha:0, codigo:"INV1",nome:"Investimento 1",  tipo:"investimento" },
      { linha:0, codigo:"INV2",nome:"Investimento 2",  tipo:"investimento" },
      { linha:0, codigo:"X",   nome:"A confirmar",     tipo:"outro" }
    ];
  }
  var dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 4).getValues();
  return dados
    .filter(function(r) { return String(r[0]||"").trim(); })
    .map(function(r, i) {
      return {
        linha:  i + 2,
        codigo: String(r[0]||"").trim(),
        nome:   String(r[1]||"").trim(),
        tipo:   String(r[2]||"banco").trim(),
        ativo:  String(r[3]||"SIM").trim().toUpperCase() !== "NAO"
      };
    });
}

function salvarConta(dados) {
  try {
    var codigo = String(dados.codigo||"").trim().toUpperCase();
    var nome   = String(dados.nome||"").trim();
    var tipo   = String(dados.tipo||"banco").trim();
    var ativo  = dados.ativo !== false ? "SIM" : "NAO";
    var linha  = dados.linha ? parseInt(dados.linha) : 0;
    if (!codigo) throw new Error("Codigo obrigatorio.");
    if (!nome)   throw new Error("Nome obrigatorio.");
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("contas");
    if (!sheet) {
      sheet = ss.insertSheet("contas");
      var headers = ["Codigo","Nome","Tipo","Ativo"];
      sheet.getRange(1,1,1,4).setValues([headers])
        .setBackground("#1e3a5f").setFontColor("white").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    var row = [codigo, nome, tipo, ativo];
    if (linha >= 2) {
      sheet.getRange(linha, 1, 1, 4).setValues([row]);
    } else {
      // Verificar se codigo ja existe
      if (sheet.getLastRow() >= 2) {
        var existing = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
        for (var i = 0; i < existing.length; i++) {
          if (String(existing[i][0]||"").trim().toUpperCase() === codigo) {
            throw new Error("Codigo '" + codigo + "' ja existe.");
          }
        }
      }
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues([row]);
    }
    SpreadsheetApp.flush();
    // Reinicializar CONTAS
    inicializarContas_();
    return { ok: true, dados: getContas() };
  } catch(err) {
    return { erro: err.message };
  }
}

function inicializarContas_() {
  try {
    var lista = getContas();
    // Resetar CONTAS
    Object.keys(CONTAS).forEach(function(k) { delete CONTAS[k]; });
    lista.forEach(function(c) {
      if (c.ativo !== false) { CONTAS[c.nome] = c.codigo; }
    });
  } catch(e) {
    Logger.log("inicializarContas_ erro: " + e.message);
  }
}

// ── CARTÕES DISPONÍVEIS PARA O PAINEL ─────────────────────────────
function getCartoesParaModal() {
  inicializarCartoes_();
  inicializarChaves_();
  var BASE_URL = ScriptApp.getService().getUrl();
  var resultado = [];
  Object.keys(CHAVES_CARTAO_CODIGO).forEach(function(chave) {
    var codigo = CHAVES_CARTAO_CODIGO[chave];
    var nome   = "";
    Object.keys(CARTOES_POR_EMAIL).forEach(function(email) {
      (CARTOES_POR_EMAIL[email]||[]).forEach(function(c) {
        if (c.codigo === codigo && !nome) { nome = c.nome; }
      });
    });
    resultado.push({
      chave: chave,
      codigo: codigo,
      nome:  nome || codigo,
      url:   BASE_URL + "?page=cartao&key=" + chave
    });
  });
  // Ordenar por nome
  resultado.sort(function(a,b) { return a.nome.localeCompare(b.nome); });
  return resultado;
}
// ── HELPERS DE CACHE ─────────────────────────────────────────────
var CACHE_TTL = 21600; // 6 horas (máximo do GAS)

function cacheGet_(key) {
  try {
    var val = CacheService.getScriptCache().get(key);
    return val ? JSON.parse(val) : null;
  } catch(e) { return null; }
}

function cachePut_(key, data, ttl) {
  try {
    var json = JSON.stringify(data);
    // GAS limita cache a 100KB por entrada
    if (json.length < 90000) {
      CacheService.getScriptCache().put(key, json, ttl || CACHE_TTL);
    }
  } catch(e) {}
}


// ── EDITAR REGISTRO ───────────────────────────────────────────────
function editarRegistro(dados) {
  try {
    var linha   = parseInt(dados.linha) || 0;
    var origem  = String(dados.origem||"").trim();
    var desc    = String(dados.descricao||"").trim();
    var tipo    = String(dados.tipo||"DV").trim();
    if (!linha || linha < 2) throw new Error("Linha inválida: " + linha);
    if (!desc) throw new Error("Descrição obrigatória.");
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tz = getTZ_();

    if (origem === "lanc") {
      // Editar em lancamentos: A=Data B=Desc C=Valor D=Conta E=Aguardando F=Tipo
      var sheet = ss.getSheetByName("lancamentos");
      if (!sheet) throw new Error("Aba lancamentos não encontrada.");
      var row = sheet.getRange(linha, 1, 1, 6).getValues()[0];
      // Data
      if (dados.data) {
        var partes = dados.data.split('-');
        row[0] = new Date(parseInt(partes[0]), parseInt(partes[1])-1, parseInt(partes[2]), 12);
      }
      row[1] = desc;
      // Valor — manter sinal original
      if (dados.valor !== undefined && !isNaN(parseFloat(dados.valor))) {
        var sinal = row[2] < 0 ? -1 : 1;
        row[2] = Math.abs(parseFloat(dados.valor)) * sinal;
      }
      // Conta
      if (dados.conta) { row[3] = dados.conta; }
      // Aguardando — desmarcar apenas remove o SIM (processamento ocorre automaticamente na data)
      row[4] = dados.aguardando ? "SIM" : "";
      row[5] = tipo;
      sheet.getRange(linha, 1, 1, 6).setValues([row]);
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosLancamentos","getDadosDV","getDeltaBkp"]);

    } else if (origem === "bkp") {
      // Editar em bkp: A=Data B=Desc C=Valor D=Conta E=Tipo F=DataHoraExec
      var sheet2 = ss.getSheetByName("bkp");
      if (!sheet2) throw new Error("Aba bkp não encontrada.");
      sheet2.getRange(linha, 2).setValue(desc);
      sheet2.getRange(linha, 5).setValue(tipo);
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosLancamentos","getDadosDV","getDeltaBkp"]);

    } else if (origem === "cartao") {
      // Editar em CartaoC: A=DataReg B=DataFatura C=Desc F=Valor G=Tipo H=Cartao
      var sheet3 = ss.getSheetByName("CartaoC");
      if (!sheet3) throw new Error("Aba CartaoC não encontrada.");
      var rowC = sheet3.getRange(linha, 1, 1, 6).getValues()[0];
      // Data fatura
      if (dados.data) {
        var p2 = dados.data.split('-');
        rowC[1] = new Date(parseInt(p2[0]), parseInt(p2[1])-1, parseInt(p2[2]), 12);
      }
      rowC[2] = desc;
      // Valor — manter sinal
      if (dados.valor !== undefined && !isNaN(parseFloat(dados.valor))) {
        var sinalC = rowC[3] < 0 ? -1 : 1;
        rowC[3] = Math.abs(parseFloat(dados.valor)) * sinalC;
      }
      rowC[4] = tipo;
      sheet3.getRange(linha, 1, 1, 6).setValues([rowC]);
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosCartaoC","getDeltaCartao"]);
      // Recalcular fatura afetada
      try { calcularFaturas(); } catch(e) {}
    } else {
      throw new Error("Origem desconhecida: " + origem);
    }
    return { ok: true };
  } catch(err) {
    return { erro: err.message };
  }
}

// ── EXCLUIR REGISTRO ──────────────────────────────────────────────
function excluirRegistro(dados) {
  try {
    var linha  = parseInt(dados.linha) || 0;
    var origem = String(dados.origem||"").trim();
    if (!linha || linha < 2) throw new Error("Linha inválida.");
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (origem === "lanc") {
      var sheet = ss.getSheetByName("lancamentos");
      if (!sheet) throw new Error("Aba lancamentos não encontrada.");
      sheet.deleteRow(linha);
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosLancamentos","getDadosDV"]);

    } else if (origem === "bkp") {
      // Estornar saldo antes de excluir
      var bkpSheet = ss.getSheetByName("bkp");
      var saldoSheet = ss.getSheetByName("saldos");
      if (!bkpSheet) throw new Error("Aba bkp não encontrada.");
      var rowB = bkpSheet.getRange(linha, 1, 1, 6).getValues()[0];
      var valor = fixNum_(rowB[2]);
      var conta = String(rowB[3]||"").trim();
      // Estornar: inverter o sinal
      if (saldoSheet && conta) {
        var lastS = saldoSheet.getLastRow();
        if (lastS >= 2) {
          var codigos = saldoSheet.getRange(2, 1, lastS-1, 1).getValues();
          for (var i = 0; i < codigos.length; i++) {
            if (String(codigos[i][0]).trim() === conta) {
              var cell = saldoSheet.getRange(i+2, 4);
              cell.setValue(fixNum_(cell.getValue()) - valor);
              saldoSheet.getRange(i+2, 2).setValue(new Date());
              break;
            }
          }
        }
      }
      bkpSheet.deleteRow(linha);
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosLancamentos","getDadosDV","getDeltaBkp"]);

    } else if (origem === "cartao") {
      var cartSheet = ss.getSheetByName("CartaoC");
      if (!cartSheet) throw new Error("Aba CartaoC não encontrada.");
      cartSheet.deleteRow(linha);
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosCartaoC","getDeltaCartao"]);
      // Recalcular faturas
      try { calcularFaturas(); } catch(e) {}
    } else {
      throw new Error("Origem desconhecida: " + origem);
    }
    return { ok: true };
  } catch(err) {
    return { erro: err.message };
  }
}

function ordenarCartaoC_() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("CartaoC");
  if (!sheet || sheet.getLastRow() < 3) { return; }
  var lastRow = sheet.getLastRow();
  var range   = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  // Ordenar por col B (data fatura) ASC, depois col A (data registro) ASC
  range.sort([{ column: 2, ascending: true }, { column: 1, ascending: true }]);
}


function keepAlive() {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getName();
    inicializarCartoes_();
    inicializarChaves_();
    // Pre-aquecer cache das funcoes principais
    try { getDadosLancamentos(); }   catch(e) {}
    try { getDadosCartaoC(); }       catch(e) {}
    try { getDadosDV(); }            catch(e) {}
    try { getDadosLancPendentes(); } catch(e) {}
    CacheService.getScriptCache().put("keepAlive_ts", String(Date.now()), 600);
    Logger.log("keepAlive OK: " + new Date().toLocaleTimeString());
  } catch(err) {
    Logger.log("keepAlive erro: " + err.message);
  }
}

// ── NOMES DOS CARTÕES ─────────────────────────────────────────────
// NOMES_CARTAO é resolvido dinamicamente a partir da aba VencCC
// via inicializarCartoes_() — não há hardcode de nomes aqui
var NOMES_CARTAO = {};
function resolverNomeCartao_(codigo) {
  // Tenta via NOMES_CARTAO já populado
  if (NOMES_CARTAO[codigo]) return NOMES_CARTAO[codigo];
  // Tenta via CARTOES_POR_EMAIL
  var encontrado = '';
  Object.keys(CARTOES_POR_EMAIL).forEach(function(email) {
    (CARTOES_POR_EMAIL[email]||[]).forEach(function(c) {
      if (c.codigo === codigo && !encontrado) encontrado = c.nome;
    });
  });
  return encontrado || codigo;
}

// ── CRIAR ABA FATURAS SE NÃO EXISTIR ─────────────────────────────
function criarAbaFaturas_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("faturas");
  if (!sheet) {
    sheet = ss.insertSheet("faturas");
    var headers = ["Cartao_Codigo","Cartao_Nome","Conta_Debito","Mes_Fatura",
                   "Valor_Calculado","Valor_Ajustado","Status","Linha_Lancamento"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#1e3a5f").setFontColor("white").setFontWeight("bold");
    // Inserir linhas de exemplo — substitua pelos cartões reais do cliente
    var configs = [
      ["CARTAO-EXEMPLO-0001","Cartão Exemplo 1","BK1","","","","config",""],
      ["CARTAO-EXEMPLO-0002","Cartão Exemplo 2","BK1","","","","config",""]
    ];
    sheet.getRange(2, 1, configs.length, 8).setValues(configs);
    sheet.setFrozenRows(1);
    Logger.log("Aba faturas criada com sucesso.");
  }
  return sheet;
}

// ── CALCULAR E SINCRONIZAR FATURAS ────────────────────────────────
// Extrai o agrupador (últimos 4 dígitos) de um código de cartão
function getAgrupador_(codigo) {
  var m = String(codigo || '').match(/(\d{4})$/);
  return m ? m[1] : codigo;
}

// Código genérico para lançamentos de ajuste
function getCodigoAjuste_(agrupador) {
  return 'CARTAO-' + agrupador;
}

function calcularFaturas() {
  inicializarCartoes_();
  var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  var cartaoC   = ss.getSheetByName('CartaoC');
  var lancSheet = ss.getSheetByName('lancamentos');
  var fatSheet  = criarAbaFaturas_();
  if (!cartaoC || !lancSheet) { return { erro: 'Abas necessarias nao encontradas.' }; }

  var hoje = new Date(); hoje.setHours(0,0,0,0);
  var tz   = getTZ_();

  // 1. Ler CartaoC e agrupar por agrupador (últimos 4 dígitos) + mês
  var lastC = cartaoC.getLastRow();
  if (lastC < 2) { return { ok: true, criadas: 0, atualizadas: 0 }; }
  var dadosC = cartaoC.getRange(2, 1, lastC-1, 6).getValues();

  // Mapa de nomes por código de cartão (da aba VencCC, col D)
  var nomesPorCodigo = {};
  var shVencCC = ss.getSheetByName('VencCC');
  if (shVencCC && shVencCC.getLastRow() >= 2) {
    shVencCC.getRange(2, 1, shVencCC.getLastRow()-1, 5).getValues().forEach(function(v) {
      var cod  = String(v[4] || '').trim(); // col E = código
      var nome = String(v[3] || '').trim(); // col D = nome do titular
      if (cod && nome) nomesPorCodigo[cod] = nome;
    });
  }

  var grupos = {};
  dadosC.forEach(function(l) {
    var dataFatura = fixDate_(l[1]);
    if (!(dataFatura instanceof Date) || isNaN(dataFatura)) { return; }
    dataFatura.setHours(0,0,0,0);
    if (dataFatura < hoje) { return; }
    var cartaoCod   = String(l[5] || '').trim();
    if (!cartaoCod) { return; }
    var agrupador   = getAgrupador_(cartaoCod);
    var codGenerico = getCodigoAjuste_(agrupador);
    var anoMes = dataFatura.getFullYear() + '-' + pad_(dataFatura.getMonth()+1);
    var chave  = codGenerico + '_' + anoMes;
    if (!grupos[chave]) {
      var contaDebito = CONTA_POR_CARTAO[cartaoCod] || 'BK1';
      grupos[chave] = {
        codGenerico: codGenerico,
        agrupador:   agrupador,
        nomeGrupo:   '', // preenchido abaixo
        contaDebito: contaDebito,
        dataFatura:  dataFatura,
        anoMes:      anoMes,
        mesFatura:   pad_(dataFatura.getMonth()+1) + '/' + dataFatura.getFullYear(),
        total:       0,
        titulares:   [] // coleta nomes únicos dos titulares
      };
    }
    // Acumular nome do titular (sem duplicatas)
    var nomeTitular = nomesPorCodigo[cartaoCod] || '';
    if (nomeTitular && grupos[chave].titulares.indexOf(nomeTitular) < 0) {
      grupos[chave].titulares.push(nomeTitular);
    }
    grupos[chave].total += fixNum_(l[3]);
  });

  // Montar nomeGrupo com nomes dos titulares
  Object.keys(grupos).forEach(function(chave) {
    var g = grupos[chave];
    var nomes = g.titulares;
    var prefixo = nomes.length > 0 ? nomes.join(' / ') : 'Cartão ' + g.agrupador;
    g.nomeGrupo = prefixo + ' ' + g.mesFatura;
  });

  // 2. Ler aba faturas
  var lastF = fatSheet.getLastRow();
  var faturasExist = {};
  if (lastF >= 2) {
    var dadosF = fatSheet.getRange(2, 1, lastF-1, 8).getValues();
    dadosF.forEach(function(r, i) {
      var cod  = String(r[0] || '').trim();
      var anoM = String(r[3] || '').trim();
      var stat = String(r[6] || '').trim();
      if (!cod || !anoM || stat === 'config') { return; }
      var chave = cod + '_' + anoM;
      faturasExist[chave] = {
        linhaFat:      i+2,
        valorAjustado: fixNum_(r[5]),
        idLanc:        String(r[7] || '').trim()
      };
    });
  }

  // 3. Ler lancamentos — mapear id_fatura → linha
  var lastL = lancSheet.getLastRow();
  var lancMap = {};
  if (lastL >= 2) {
    var dadosL = lancSheet.getRange(2, 1, lastL-1, 7).getValues();
    dadosL.forEach(function(r, i) {
      var idFat = String(r[6] || '').trim();
      if (idFat) { lancMap[idFat] = i+2; }
    });
  }

  var criadas = 0, atualizadas = 0, removidas = 0;

  // 4. Detectar se há registros com códigos antigos (pré-agrupamento)
  // Se a aba faturas usa códigos individuais (ex: VISA-MMO-4478) ao invés de
  // genéricos (CARTAO-4478), limpar os registros antigos de lancamentos e recriar
  var temCodigoAntigo = false;
  Object.keys(faturasExist).forEach(function(ch) {
    if (ch.indexOf('CARTAO-') !== 0) { temCodigoAntigo = true; }
  });

  if (temCodigoAntigo) {
    // Remover lançamentos antigos de faturas (identificados por chave no formato CODIGO_YYYY-MM)
    Logger.log('calcularFaturas: detectados códigos antigos — limpando registros migrados');
    if (lancSheet.getLastRow() >= 2) {
      var dadosLMig = lancSheet.getRange(2, 1, lancSheet.getLastRow()-1, 7).getValues();
      var linhasMig = [];
      dadosLMig.forEach(function(r, i) {
        var idFat = String(r[6] || '').trim();
        // Lançamento de fatura antigo: tem idFat com código individual OU col F=DF e col G vazia
        // mas col B contém nome de cartão conhecido
        var desc = String(r[1] || '').trim();
        if (idFat && idFat.indexOf('CARTAO-') < 0 && idFat.indexOf('_20') > 0) {
          linhasMig.push(i + 2);
        }
      });
      linhasMig.sort(function(a,b){ return b-a; });
      linhasMig.forEach(function(l){ lancSheet.deleteRow(l); });
      SpreadsheetApp.flush();
      removidas += linhasMig.length;
      Logger.log('calcularFaturas: removidas ' + linhasMig.length + ' linhas com códigos antigos');
    }
    // Limpar também a aba faturas de registros com códigos antigos
    if (fatSheet.getLastRow() >= 2) {
      var dadosFMig = fatSheet.getRange(2, 1, fatSheet.getLastRow()-1, 8).getValues();
      var linhasFMig = [];
      dadosFMig.forEach(function(r, i) {
        var cod = String(r[0]||'').trim();
        var stat = String(r[6]||'').trim();
        if (stat === 'config') return;
        if (cod && cod.indexOf('CARTAO-') < 0) { linhasFMig.push(i + 2); }
      });
      linhasFMig.sort(function(a,b){ return b-a; });
      linhasFMig.forEach(function(l){ fatSheet.deleteRow(l); });
      SpreadsheetApp.flush();
      Logger.log('calcularFaturas: removidas ' + linhasFMig.length + ' linhas antigas da aba faturas');
    }
    // Reconstruir faturasExist e lancMap com dados limpos
    faturasExist = {};
    lancMap = {};
    if (fatSheet.getLastRow() >= 2) {
      fatSheet.getRange(2, 1, fatSheet.getLastRow()-1, 8).getValues().forEach(function(r, i) {
        var cod = String(r[0]||'').trim(), anoM = String(r[3]||'').trim(), stat = String(r[6]||'').trim();
        if (!cod || !anoM || stat === 'config') return;
        faturasExist[cod + '_' + anoM] = { linhaFat: i+2, valorAjustado: fixNum_(r[5]), idLanc: String(r[7]||'').trim() };
      });
    }
    if (lancSheet.getLastRow() >= 2) {
      lancSheet.getRange(2, 1, lancSheet.getLastRow()-1, 7).getValues().forEach(function(r, i) {
        var idFat = String(r[6]||'').trim();
        if (idFat) lancMap[idFat] = i+2;
      });
    }
  }

  // 5. Processar cada grupo — criar ou atualizar
  Object.keys(grupos).forEach(function(chave) {
    var g     = grupos[chave];
    var total = Math.round(g.total * 100) / 100;

    if (faturasExist[chave]) {
      // Atualizar valor calculado e nome na aba faturas
      var linhaF = faturasExist[chave].linhaFat;
      fatSheet.getRange(linhaF, 2).setValue(g.nomeGrupo); // col B = nome
      fatSheet.getRange(linhaF, 5).setValue(total);       // col E = valor calculado
      // Garantir que col H (idLanc) está preenchida
      if (!faturasExist[chave].idLanc) {
        fatSheet.getRange(linhaF, 8).setValue(chave);
        faturasExist[chave].idLanc = chave;
      }
      // Atualizar lançamento correspondente (se não ajustado manualmente)
      var linhaL = lancMap[chave];
      if (linhaL && !faturasExist[chave].valorAjustado) {
        lancSheet.getRange(linhaL, 2).setValue(g.nomeGrupo);  // col B = descrição
        lancSheet.getRange(linhaL, 3).setValue(-Math.abs(total)); // col C = valor
        // Garantir que col G do lançamento está preenchida
        lancSheet.getRange(linhaL, 7).setValue(chave);
      } else if (!linhaL) {
        // Lançamento não existe — criar
        var novaLinhaL2 = lancSheet.getLastRow() + 1;
        lancSheet.getRange(novaLinhaL2, 1, 1, 7).setValues([[
          g.dataFatura, g.nomeGrupo, -Math.abs(total), g.contaDebito, 'SIM', 'NA', chave
        ]]);
      }
      atualizadas++;
    } else {
      // Criar novo registro em faturas e lançamentos
      var novaLinhaF = fatSheet.getLastRow() + 1;
      fatSheet.getRange(novaLinhaF, 1, 1, 8).setValues([[
        g.codGenerico, g.nomeGrupo, g.contaDebito,
        g.anoMes, total, '', 'pendente', chave
      ]]);
      fatSheet.getRange(novaLinhaF, 4).setNumberFormat('@');
      var novaLinhaL = lancSheet.getLastRow() + 1;
      lancSheet.getRange(novaLinhaL, 1, 1, 7).setValues([[
        g.dataFatura, g.nomeGrupo, -Math.abs(total), g.contaDebito, 'SIM', 'NA', chave
      ]]);
      criadas++;
    }
  });

  // 5. Remover faturas e lançamentos cujos grupos sumiram
  SpreadsheetApp.flush();
  var lastF2 = fatSheet.getLastRow();
  var linhasRemoverFat = [];
  if (lastF2 >= 2) {
    var dadosF2 = fatSheet.getRange(2, 1, lastF2-1, 8).getValues();
    dadosF2.forEach(function(r, i) {
      var stat = String(r[6] || '').trim();
      if (stat === 'config' || stat === 'ajustado') { return; }
      var anoM = String(r[3] || '').trim();
      var cod  = String(r[0] || '').trim();
      if (!cod || !anoM) { return; }
      var chave = cod + '_' + anoM;
      if (!grupos[chave]) {
        linhasRemoverFat.push(i + 2);
        var linhaL = lancMap[chave];
        if (linhaL) { lancSheet.getRange(linhaL, 1, 1, 7).clearContent(); removidas++; }
      }
    });
  }

  if (removidas > 0) {
    SpreadsheetApp.flush();
    var lastL2 = lancSheet.getLastRow();
    if (lastL2 >= 2) {
      var dadosL2 = lancSheet.getRange(2, 1, lastL2-1, 1).getValues();
      var linhasVazias = [];
      dadosL2.forEach(function(r, i) { if (!r[0]) { linhasVazias.push(i + 2); } });
      linhasVazias.sort(function(a,b){ return b-a; });
      linhasVazias.forEach(function(l){ lancSheet.deleteRow(l); });
    }
  }

  linhasRemoverFat.sort(function(a,b){ return b-a; });
  linhasRemoverFat.forEach(function(l){ fatSheet.deleteRow(l); });

  var totalLinhas = lancSheet.getLastRow() - 1;
  if (totalLinhas > 1) {
    lancSheet.getRange(2, 1, totalLinhas, lancSheet.getLastColumn()).sort({ column: 1, ascending: true });
  }

  SpreadsheetApp.flush();
  CacheService.getScriptCache().removeAll(['getDadosLancamentos','getDadosDV','getDashboard', 'getDadosLancPendentes']);
  Logger.log('calcularFaturas: criadas=' + criadas + ' atualizadas=' + atualizadas + ' removidas=' + removidas);
  return { ok: true, criadas: criadas, atualizadas: atualizadas, removidas: removidas };
}

// Retorna faturas agrupadas para a tela de ajuste
function getFaturasParaAjuste() {
  inicializarCartoes_();
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var fatSheet = ss.getSheetByName('faturas');
  if (!fatSheet || fatSheet.getLastRow() < 2) { return []; }
  var dados = fatSheet.getRange(2, 1, fatSheet.getLastRow()-1, 8).getValues();
  var resultado = [];
  dados.forEach(function(r) {
    var cod      = String(r[0] || '').trim();
    var nome     = String(r[1] || '').trim();
    var anoMes   = String(r[3] || '').trim();
    var valCalc  = fixNum_(r[4]);
    var valAjust = fixNum_(r[5]);
    var status   = String(r[6] || '').trim();
    if (!cod || !anoMes || status === 'config') { return; }
    var partes = anoMes.split('-');
    var mesAno = partes.length === 2 ? partes[1] + '/' + partes[0] : anoMes;
    resultado.push({
      codigo: cod, nome: nome, anoMes: anoMes, mesAno: mesAno,
      valorCalculado: valCalc, valorAjustado: valAjust,
      valorEfetivo: valAjust || valCalc, status: status
    });
  });
  // Ordenar: faturas do mês atual e próximos primeiro, depois futuras, depois passadas
  var hoje = new Date();
  var anoMesHoje = hoje.getFullYear() + '-' + pad_(hoje.getMonth() + 1);
  resultado.sort(function(a, b) {
    // Faturas a partir do mês atual vêm primeiro, em ordem crescente
    var aFuturo = a.anoMes >= anoMesHoje;
    var bFuturo = b.anoMes >= anoMesHoje;
    if (aFuturo && bFuturo) return a.anoMes.localeCompare(b.anoMes); // ambas futuras: crescente
    if (!aFuturo && !bFuturo) return b.anoMes.localeCompare(a.anoMes); // ambas passadas: decrescente
    return aFuturo ? -1 : 1; // futuras antes das passadas
  });
  return resultado;
}

// Ajusta uma fatura: lança diferença na CartaoC e atualiza aba faturas
function ajustarFatura(params) {
  inicializarCartoes_();
  try {
    var codigo    = String(params.codigo   || '').trim();
    var anoMes    = String(params.anoMes   || '').trim();
    var valorReal = Math.abs(fixNum_(params.valorReal));
    var dataVenc  = params.dataVenc ? fixDate_(params.dataVenc) : null;
    if (!codigo || !anoMes || !valorReal) { return { ok: false, erro: 'Parâmetros inválidos.' }; }

    var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    var fatSheet  = ss.getSheetByName('faturas');
    var cartaoC   = ss.getSheetByName('CartaoC');
    var lancSheet = ss.getSheetByName('lancamentos');
    var tz        = getTZ_();
    var chave     = codigo + '_' + anoMes;

    // Encontrar fatura
    var linhaFat = -1, valCalc = 0, idLanc = '';
    if (fatSheet && fatSheet.getLastRow() >= 2) {
      fatSheet.getRange(2, 1, fatSheet.getLastRow()-1, 8).getValues().forEach(function(r, i) {
        if (String(r[0]||'').trim() === codigo && String(r[3]||'').trim() === anoMes) {
          linhaFat = i + 2; valCalc = fixNum_(r[4]); idLanc = String(r[7]||'').trim();
        }
      });
    }
    if (linhaFat < 0) { return { ok: false, erro: 'Fatura não encontrada: ' + chave }; }

    // Diferença: valorReal é o que vai ser pago (positivo no input → negativo no lançamento)
    var valorRealNeg = -valorReal;
    var diferenca    = Math.round((valorRealNeg - valCalc) * 100) / 100;

    // Data do ajuste
    var partes   = anoMes.split('-');
    var dataLanc = (dataVenc instanceof Date && !isNaN(dataVenc))
      ? dataVenc
      : new Date(parseInt(partes[0]), parseInt(partes[1])-1, 1);
    dataLanc.setHours(12,0,0,0);

    // Lançar diferença na CartaoC
    if (Math.abs(diferenca) > 0.009) {
      var agrupador = getAgrupador_(codigo);
      var desc = 'Ajuste fatura ' + partes[1] + '/' + partes[0];
      cartaoC.appendRow([
        Utilities.formatDate(dataLanc, tz, 'yyyy-MM-dd'),
        dataLanc, desc, diferenca, 'DF',
        getCodigoAjuste_(agrupador)
      ]);
    }

    // Atualizar aba faturas
    fatSheet.getRange(linhaFat, 6).setValue(valorRealNeg);
    fatSheet.getRange(linhaFat, 7).setValue('ajustado');

    // Atualizar lançamento correspondente
    if (idLanc && lancSheet && lancSheet.getLastRow() >= 2) {
      lancSheet.getRange(2, 1, lancSheet.getLastRow()-1, 7).getValues().forEach(function(r, i) {
        if (String(r[6]||'').trim() === idLanc) {
          lancSheet.getRange(i+2, 3).setValue(valorRealNeg);
        }
      });
    }

    SpreadsheetApp.flush();
    CacheService.getScriptCache().removeAll(
      ['getDadosLancamentos','getDadosDV','getDashboard','getDeltaBkp']
    );
    Logger.log('ajustarFatura: ' + chave + ' valCalc=' + valCalc +
      ' valorReal=' + valorRealNeg + ' diferenca=' + diferenca);
    return {
      ok: true, chave: chave,
      valorCalc: valCalc, valorReal: valorRealNeg,
      diferenca: diferenca, lancado: Math.abs(diferenca) > 0.009
    };
  } catch(e) {
    Logger.log('ajustarFatura erro: ' + e.message);
    return { ok: false, erro: e.message };
  }
}

// ── EDIÇÃO EM LOTE ────────────────────────────────────────────────
function editarEmLote(dados) {
  try {
    var linhas  = dados.linhas  || [];
    var origem  = String(dados.origem  || "").trim();
    var campos  = dados.campos  || {}; // { descricao, valor, tipo }
    if (!linhas.length) throw new Error("Nenhuma linha selecionada.");
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (origem === "cartao") {
      var sheet = ss.getSheetByName("CartaoC");
      if (!sheet) throw new Error("Aba CartaoC não encontrada.");
      linhas.forEach(function(linha) {
        linha = parseInt(linha);
        if (!linha || linha < 2) return;
        var row = sheet.getRange(linha, 1, 1, 6).getValues()[0];
        if (campos.descricao !== undefined && campos.descricao !== "") row[2] = campos.descricao;
        if (campos.valor     !== undefined) {
          var sinal = row[3] < 0 ? -1 : 1;
          row[3] = Math.abs(parseFloat(campos.valor)) * sinal;
        }
        if (campos.tipo      !== undefined && campos.tipo !== "") row[4] = campos.tipo;
        sheet.getRange(linha, 1, 1, 6).setValues([row]);
      });
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosCartaoC","getDeltaCartao"]);
      try { calcularFaturas(); } catch(e) {}

    } else if (origem === "lanc") {
      var sheet2 = ss.getSheetByName("lancamentos");
      if (!sheet2) throw new Error("Aba lancamentos não encontrada.");
      linhas.forEach(function(linha) {
        linha = parseInt(linha);
        if (!linha || linha < 2) return;
        var row = sheet2.getRange(linha, 1, 1, 6).getValues()[0];
        if (campos.descricao !== undefined && campos.descricao !== "") row[1] = campos.descricao;
        if (campos.valor     !== undefined) {
          var sinal = row[2] < 0 ? -1 : 1;
          row[2] = Math.abs(parseFloat(campos.valor)) * sinal;
        }
        if (campos.tipo !== undefined && campos.tipo !== "") row[5] = campos.tipo;
        sheet2.getRange(linha, 1, 1, 6).setValues([row]);
      });
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosLancamentos","getDadosDV"]);

    } else if (origem === "bkp") {
      var sheet3 = ss.getSheetByName("bkp");
      if (!sheet3) throw new Error("Aba bkp não encontrada.");
      linhas.forEach(function(linha) {
        linha = parseInt(linha);
        if (!linha || linha < 2) return;
        if (campos.descricao !== undefined && campos.descricao !== "") sheet3.getRange(linha, 2).setValue(campos.descricao);
        if (campos.tipo      !== undefined && campos.tipo !== "")      sheet3.getRange(linha, 5).setValue(campos.tipo);
      });
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosLancamentos","getDadosDV"]);
    }

    return { ok: true, alterados: linhas.length };
  } catch(err) {
    return { erro: err.message };
  }
}

function excluirEmLote(dados) {
  try {
    var linhas  = (dados.linhas || []).map(function(l){ return parseInt(l); }).filter(function(l){ return l >= 2; });
    var origem  = String(dados.origem || "").trim();
    if (!linhas.length) throw new Error("Nenhuma linha selecionada.");
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Ordenar de baixo para cima para não deslocar índices ao deletar
    linhas.sort(function(a,b){ return b-a; });

    if (origem === "lanc") {
      var sheet = ss.getSheetByName("lancamentos");
      if (!sheet) throw new Error("Aba lancamentos não encontrada.");
      linhas.forEach(function(l){ sheet.deleteRow(l); });
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosLancamentos","getDadosDV"]);

    } else if (origem === "bkp") {
      var bkpSheet = ss.getSheetByName("bkp");
      var saldoSheet = ss.getSheetByName("saldos");
      if (!bkpSheet) throw new Error("Aba bkp não encontrada.");
      // Estornar cada linha antes de deletar
      linhas.forEach(function(l) {
        var row = bkpSheet.getRange(l, 1, 1, 5).getValues()[0];
        var valor = fixNum_(row[2]);
        var conta = String(row[3]||"").trim();
        if (saldoSheet && conta) {
          var lastS = saldoSheet.getLastRow();
          if (lastS >= 2) {
            var cods = saldoSheet.getRange(2, 1, lastS-1, 1).getValues();
            for (var i = 0; i < cods.length; i++) {
              if (String(cods[i][0]).trim() === conta) {
                var cell = saldoSheet.getRange(i+2, 4);
                cell.setValue(fixNum_(cell.getValue()) - valor);
                break;
              }
            }
          }
        }
        bkpSheet.deleteRow(l);
      });
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosLancamentos","getDadosDV","getDeltaBkp"]);

    } else if (origem === "cartao") {
      var cartSheet = ss.getSheetByName("CartaoC");
      if (!cartSheet) throw new Error("Aba CartaoC não encontrada.");
      linhas.forEach(function(l){ cartSheet.deleteRow(l); });
      SpreadsheetApp.flush();
      CacheService.getScriptCache().removeAll(["getDadosCartaoC","getDeltaCartao"]);
      try { calcularFaturas(); } catch(e) {}

    } else {
      throw new Error("Origem desconhecida: " + origem);
    }

    return { ok: true, excluidos: linhas.length };
  } catch(err) {
    return { erro: err.message };
  }
}

// ── LANÇAMENTOS PENDENTES ─────────────────────────────────────────
function getDadosLancPendentes() {
  var cached = cacheGet_('getDadosLancPendentes');
  if (cached) { return cached; }
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("lancamentos");
  if (!sheet || sheet.getLastRow() < 2) { return []; }
  var tz    = getTZ_();
  var dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues();
  var resultado = [];
  dados.forEach(function(r, i) {
    var data  = fixDate_(r[0]);
    var desc  = String(r[1]||"").trim();
    if (!desc && !r[2]) { return; }
    var valor = fixNum_(r[2]);
    var conta = String(r[3]||"").trim();
    var aguard = String(r[4]||"").trim();
    var tipo  = String(r[5]||"").trim();
    var idFat = String(r[6]||"").trim();
    resultado.push({
      linha:      i + 2,
      data:       (data instanceof Date && !isNaN(data)) ? Utilities.formatDate(data, tz, "dd/MM/yyyy") : String(r[0]||""),
      descricao:  desc,
      valor:      valor,
      conta:      conta,
      aguardando: aguard,
      tipo:       tipo,
      idFatura:   idFat
    });
  });
  cachePut_('getDadosLancPendentes', resultado, 300);
  return resultado;
}

// ── RESUMO CRUD ───────────────────────────────────────────────────
function getDadosResumoEdit() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("resumo");
  if (!sheet || sheet.getLastRow() < 2) { return { despesas: [], rendas: [] }; }
  var dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 9).getValues();
  var despesas = [], rendas = [];
  dados.forEach(function(r, i) {
    var tipoD = String(r[0]||"").trim();
    var descD = String(r[1]||"").trim();
    var catD  = String(r[2]||"").trim();
    var valD  = fixNum_(r[3]);
    var tipoR = String(r[5]||"").trim();
    var descR = String(r[6]||"").trim();
    var catR  = String(r[7]||"").trim();
    var valR  = fixNum_(r[8]);
    if (tipoD || descD) {
      despesas.push({ linha: i+2, tipo: tipoD, descricao: descD, categoria: catD, mensal: valD, lado: 'desp' });
    }
    if (tipoR || descR) {
      rendas.push({ linha: i+2, tipo: tipoR, descricao: descR, categoria: catR, mensal: valR, lado: 'renda' });
    }
  });
  return { despesas: despesas, rendas: rendas };
}

function _ultimaLinhaResumoLado_(sheet, colStart) {
  var lastRowSheet = sheet.getLastRow();
  if (lastRowSheet < 2) return 1; // só cabeçalho — próxima linha de dados é a 2
  var colDesc = colStart + 1; // B=2 (desp) ou G=7 (renda) — coluna de descrição
  var vals = sheet.getRange(2, colDesc, lastRowSheet - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return i + 2; // linha real (1-indexed, offset de 2)
  }
  return 1; // nenhuma linha usada nesse lado — próxima é a 2
}

function salvarResumoLinha(dados) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("resumo");
    if (!sheet) throw new Error("Aba resumo não encontrada.");
    var linha = parseInt(dados.linha) || 0;
    var lado  = String(dados.lado||"desp");
    var tipo  = String(dados.tipo||"").trim();
    var desc  = String(dados.descricao||"").trim();
    var cat   = String(dados.categoria||"").trim();
    var val   = parseFloat(dados.mensal) || 0;
    if (!desc) throw new Error("Descrição obrigatória.");
    var colStart = lado === 'renda' ? 6 : 1; // F=6 para renda, A=1 para desp
    if (linha >= 2) {
      sheet.getRange(linha, colStart, 1, 4).setValues([[tipo, desc, cat, val]]);
    } else {
      // Nova linha — próxima linha livre considerando apenas as colunas do lado (desp ou renda)
      var novaLinha = _ultimaLinhaResumoLado_(sheet, colStart) + 1;
      sheet.getRange(novaLinha, colStart, 1, 4).setValues([[tipo, desc, cat, val]]);
    }
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove("getDadosResumo");
    return { ok: true, dados: getDadosResumoEdit() };
  } catch(err) {
    return { erro: err.message };
  }
}

function excluirResumoLinha(dados) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("resumo");
    if (!sheet) throw new Error("Aba resumo não encontrada.");
    var linha = parseInt(dados.linha) || 0;
    var lado  = String(dados.lado||"desp");
    if (!linha || linha < 2) throw new Error("Linha inválida.");
    // Limpar apenas as colunas do lado, não deletar a linha inteira
    // (linha pode ter dados no outro lado)
    var colStart = lado === 'renda' ? 6 : 1;
    sheet.getRange(linha, colStart, 1, 4).clearContent();
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove("getDadosResumo");
    return { ok: true, dados: getDadosResumoEdit() };
  } catch(err) {
    return { erro: err.message };
  }
}

// ══════════════════════════════════════════════════════════
// DASHBOARD EXECUTIVO
// ══════════════════════════════════════════════════════════
function getDashboard() {
  var cached = cacheGet_('getDashboard');
  if (cached) { return cached; }

  var ssId = SPREADSHEET_ID || getSpreadsheetId_();
  if (!ssId) { return { erro: 'SPREADSHEET_ID não configurado. Configure nas Script Properties.' }; }
  var ss   = SpreadsheetApp.openById(ssId);
  var tz   = getTZ_();
  var hoje = new Date();
  hoje.setHours(0,0,0,0);

  // ── 0. Config do cliente ──────────────────────────────────
  var limiteDias = 7;
  try {
    var cfg = getConfig_();
    limiteDias = parseInt(cfg['limite_pendentes_dias']) || 7;
  } catch(eCfg) {
    Logger.log('getDashboard: getConfig_ erro: ' + eCfg.message);
  }

  // ── 0. Mapa código → tipo e nome (aba contas: A=código, B=nome, C=tipo) ──
  var tiposPorCodigo = {}, nomesPorCodigo = {};
  var shContas = ss.getSheetByName('contas');
  if (shContas && shContas.getLastRow() >= 2) {
    shContas.getRange(2, 1, shContas.getLastRow() - 1, 3).getValues().forEach(function(r) {
      var cod  = String(r[0] || '').trim();
      var nome = String(r[1] || '').trim();
      var tipo = String(r[2] || '').trim().toLowerCase();
      if (cod) { tiposPorCodigo[cod] = tipo; nomesPorCodigo[cod] = nome; }
    });
  }
  function ehInvestimento(cod) {
    if (tiposPorCodigo[cod]) return tiposPorCodigo[cod] === 'investimento';
    return /^INV/i.test(cod);
  }
  function nomeConta(cod) { return nomesPorCodigo[cod] || cod; }

  // ── 1. Saldos: operacional vs investimentos ───────────────
  // Aba saldos: A=código, B=data, C=descrição, D=valor
  var shSaldos = ss.getSheetByName('saldos');
  var saldosPorConta = [], saldosInv = [], totalSaldo = 0, totalInv = 0;
  if (shSaldos && shSaldos.getLastRow() >= 2) {
    shSaldos.getRange(2, 1, shSaldos.getLastRow() - 1, 4).getValues().forEach(function(r) {
      var cod  = String(r[0] || '').trim();
      if (!cod) return;
      var data = fixDate_(r[1]);
      var desc = String(r[2] || '').trim();
      var val  = fixNum_(r[3]);
      var obj  = {
        codigo: cod, nome: nomeConta(cod), desc: desc, valor: val,
        data: (data instanceof Date && !isNaN(data))
          ? Utilities.formatDate(data, tz, 'dd/MM/yyyy') : ''
      };
      if (ehInvestimento(cod)) { saldosInv.push(obj); totalInv += val; }
      else                     { saldosPorConta.push(obj); totalSaldo += val; }
    });
  }

  // Calcular saldo previsto para o mês corrente
  // Saldo previsto = saldo base + TODOS os lançamentos da aba lancamentos até fim do mês
  // Calcular por conta para mostrar projeção individual
  var shLancSaldo = ss.getSheetByName('lancamentos');
  var fimMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);
  fimMes.setHours(23,59,59,0);
  var ajustePorConta = {}; // soma de lançamentos do mês por código de conta
  if (shLancSaldo && shLancSaldo.getLastRow() >= 2) {
    shLancSaldo.getRange(2, 1, shLancSaldo.getLastRow()-1, 4).getValues().forEach(function(l) {
      var data = fixDate_(l[0]);
      if (!(data instanceof Date) || isNaN(data)) return;
      data.setHours(0,0,0,0);
      if (data > fimMes) return;
      var cod = String(l[3] || '').trim();
      var val = fixNum_(l[2]);
      if (!cod) return;
      ajustePorConta[cod] = (ajustePorConta[cod] || 0) + val;
    });
  }
  // Aplicar ajuste por conta e calcular totais previstos
  var totalPrevisto = 0, totalPrevInv = 0;
  saldosPorConta.forEach(function(c) {
    var ajuste = ajustePorConta[c.codigo] || 0;
    c.valorPrevisto = c.valor + ajuste;
    totalPrevisto += c.valorPrevisto;
  });
  saldosInv.forEach(function(c) {
    var ajuste = ajustePorConta[c.codigo] || 0;
    c.valorPrevisto = c.valor + ajuste;
    totalPrevInv += c.valorPrevisto;
  });
  // Contas que estão em lancamentos mas não em saldos (ex: BK3 sem linha em saldos)
  Object.keys(ajustePorConta).forEach(function(cod) {
    var jaIncluso = saldosPorConta.some(function(c){ return c.codigo === cod; }) ||
                    saldosInv.some(function(c){ return c.codigo === cod; });
    if (!jaIncluso) {
      if (ehInvestimento(cod)) totalPrevInv += ajustePorConta[cod];
      else totalPrevisto += ajustePorConta[cod];
    }
  });
  var shVenc = ss.getSheetByName('VencCC');
  var shFat  = ss.getSheetByName('faturas');
  var faturasProximas = [];
  if (shVenc && shVenc.getLastRow() >= 2 && shFat && shFat.getLastRow() >= 2) {
    var vencRows = shVenc.getRange(2, 1, shVenc.getLastRow() - 1, 8).getValues();
    var fatRows  = shFat.getRange(2, 1, shFat.getLastRow() - 1, 8).getValues();
    var mesAtualStr   = Utilities.formatDate(hoje, tz, 'yyyy-MM');
    var proxMes       = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 1);
    var mesProximoStr = Utilities.formatDate(proxMes, tz, 'yyyy-MM');

    // Montar mapa de faturas por codGenerico_anoMes para lookup rápido
    // A aba faturas agora usa CARTAO-4478 (código genérico) como chave
    var fatMap = {};
    fatRows.forEach(function(f) {
      var cod  = String(f[0] || '').trim();
      var anoM = String(f[3] || '').trim();
      var stat = String(f[6] || '').trim();
      if (!cod || !anoM || stat === 'config') return;
      var chave = cod + '_' + anoM;
      if (!fatMap[chave]) {
        fatMap[chave] = { valCalc: fixNum_(f[4]), valAdj: fixNum_(f[5]), status: stat };
      }
    });

    // Agrupar cartões do VencCC por agrupador — coletar nomes dos titulares
    var agrupMap = {}; // agrupador → { nomes[], dia }
    vencRows.forEach(function(v) {
      var codigo = String(v[4] || '').trim();
      var nome   = String(v[3] || '').trim();
      var dia    = Number(v[1]) || 0;
      if (!codigo || !dia) return;
      var agrup = getAgrupador_(codigo);
      if (!agrupMap[agrup]) {
        agrupMap[agrup] = { nomes: [], dia: dia };
      }
      if (nome && agrupMap[agrup].nomes.indexOf(nome) < 0) {
        agrupMap[agrup].nomes.push(nome);
      }
    });

    // Para cada agrupador, buscar faturas nos próximos 3 meses
    var proxMes2      = new Date(hoje.getFullYear(), hoje.getMonth() + 2, 1);
    var mesDoisStr    = Utilities.formatDate(proxMes2, tz, 'yyyy-MM');

    // Faturas pendentes em lancamentos (col G preenchida = fatura cartão)
    // Se a fatura NÃO está em lancamentos, já foi paga (migrou para bkp)
    // Nota: shLanc é definido abaixo na seção de pendentes — reler aqui
    var shLancFat = ss.getSheetByName('lancamentos');
    var faturasEmLanc = {}; // chave codGenerico_anoMes → true se pendente
    if (shLancFat && shLancFat.getLastRow() >= 2) {
      shLancFat.getRange(2,1,shLancFat.getLastRow()-1,7).getValues().forEach(function(r){
        var colG = String(r[6]||'').trim();
        if (colG) faturasEmLanc[colG] = true; // col G = chave da fatura
      });
    }

    Object.keys(agrupMap).forEach(function(agrup) {
      var info        = agrupMap[agrup];
      var codGenerico = getCodigoAjuste_(agrup);
      var nomeExib = info.nomes.length > 0 ? info.nomes.join(' / ') : 'Cartão ' + agrup;

      [mesAtualStr, mesProximoStr, mesDoisStr].forEach(function(mes) {
        var chave = codGenerico + '_' + mes;
        var fat   = fatMap[chave];
        if (!fat) return;
        var valor = fat.valAdj !== 0 ? fat.valAdj : fat.valCalc;
        if (valor === 0) return;
        // Pular se fatura não está mais em lancamentos (já foi paga)
        if (!faturasEmLanc[chave]) return;
        var p        = mes.split('-');
        var vencDate = new Date(parseInt(p[0]), parseInt(p[1]) - 1, info.dia);
        var diasAte  = Math.round((vencDate - hoje) / 86400000);
        faturasProximas.push({
          cartao:   nomeExib,
          codigo:   codGenerico,
          mes:      mes,
          valor:    valor,
          status:   fat.status,
          vencDate: Utilities.formatDate(vencDate, tz, 'dd/MM/yyyy'),
          diasAte:  diasAte
        });
      });
    });
    faturasProximas.sort(function(a, b) { return a.diasAte - b.diasAte; });
  }

  // ── 3. Lançamentos pendentes — próximos N dias (config) ──
  // Compara por número AAAAMMDD para evitar problemas de hora/timezone
  function toNum(d) {
    return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
  }
  // Limite só para registros FUTUROS — passado exibe tudo sem corte
  var limiteData = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + limiteDias);
  var limiteNum  = toNum(limiteData);
  var hojeNum    = toNum(hoje);

  var shLanc = ss.getSheetByName('lancamentos');
  var pendentes = [], totalPendente = 0;
  if (shLanc && shLanc.getLastRow() >= 2) {
    shLanc.getRange(2, 1, shLanc.getLastRow() - 1, 6).getValues().forEach(function(l, i) {
      var aguard = String(l[4] || '').trim().toUpperCase();
      var data = fixDate_(l[0]);
      if (!(data instanceof Date) || isNaN(data)) return;
      var dNum = toNum(data);
      if (dNum > limiteNum) return; // futuro além do limite: pular
      // passado (dNum < hojeNum): sempre inclui — com ou sem SIM
      var val   = fixNum_(l[2]);
      var desc  = String(l[1] || '').trim();
      var conta = String(l[3] || '').trim();
      var tipo  = String(l[5] || '').trim();
      pendentes.push({
        linha: i + 2,
        data:  Utilities.formatDate(data, tz, 'dd/MM/yyyy'),
        descricao: desc, valor: val, conta: conta, tipo: tipo,
        aguardando: aguard === 'SIM'
      });
      totalPendente += val;
    });
    pendentes.sort(function(a, b) {
      // Converter dd/MM/yyyy para número comparável YYYYMMDD
      function toOrd(s) {
        var p = (s||'').split('/');
        return p.length === 3 ? parseInt(p[2])*10000 + parseInt(p[1])*100 + parseInt(p[0]) : 0;
      }
      return toOrd(a.data) - toOrd(b.data);
    });
  }

  // ── 4. Fluxo mensal M, M+1, M+2 (bkp + lancamentos + CartaoC) ──
  var shBkp     = ss.getSheetByName('bkp');
  var shCartaoC = ss.getSheetByName('CartaoC');
  function somarMes(sh, nCols, mes, ano) {
    var ent = 0, sai = 0;
    if (!sh || sh.getLastRow() < 2) return { entradas: 0, saidas: 0, saldo: 0 };
    sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues().forEach(function(l) {
      var d = fixDate_(l[0]);
      if (!(d instanceof Date) || isNaN(d)) return;
      if (d.getMonth() !== mes || d.getFullYear() !== ano) return;
      var v = fixNum_(l[2]);
      if (v > 0) ent += v; else sai += Math.abs(v);
    });
    return { entradas: ent, saidas: sai, saldo: ent - sai };
  }
  var mesesRes = [];
  for (var mi = 0; mi < 3; mi++) {
    var dM = new Date(hoje.getFullYear(), hoje.getMonth() + mi, 1);
    var mm = dM.getMonth(), aa = dM.getFullYear();
    var eBkp  = somarMes(shBkp,     5, mm, aa);
    var eLanc = somarMes(shLanc,    6, mm, aa);
    var eCart = somarMes(shCartaoC, 5, mm, aa);
    mesesRes.push({
      mes:      pad_(mm + 1) + '/' + aa,
      entradas: eBkp.entradas + eLanc.entradas + eCart.entradas,
      saidas:   eBkp.saidas   + eLanc.saidas   + eCart.saidas,
      saldo:    eBkp.saldo    + eLanc.saldo     + eCart.saldo
    });
  }

  var result = {
    saldoOperacional:   totalPrevisto,
    saldoAtual:         totalSaldo,
    saldoInvestimentos: totalInv,
    saldoInvestimentos: totalInv,
    saldosPorConta:     saldosPorConta,
    saldosInv:          saldosInv,
    faturasProximas:    faturasProximas,
    pendentes:          pendentes,
    totalPendente:      totalPendente,
    limiteDias:         limiteDias,
    mesAtual:    mesesRes[0],
    mesSeguinte: mesesRes[1],
    mesDois:     mesesRes[2]
  };
  cachePut_('getDashboard', result, 300); // cache de 5 min
  return result;
}


// ── PROJEÇÃO DE SALDO (Planejado II) ────────────────────────────────────────
function getProjecao(nMeses) {
  try {
  nMeses = parseInt(nMeses) || 12;
  nMeses = Math.min(Math.max(nMeses, 6), 36); // entre 6 e 36
  var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tz   = getTZ_();
  var hoje = new Date(); hoje.setHours(0,0,0,0);

  // 1. Saldo atual por tipo (operacional e investimentos)
  var shSaldos = ss.getSheetByName('saldos');
  var saldoOp = 0, saldoInv = 0;
  if (shSaldos && shSaldos.getLastRow() >= 2) {
    shSaldos.getRange(2,1,shSaldos.getLastRow()-1,4).getValues().forEach(function(r) {
      var cod = String(r[0]||'').trim(); if (!cod) return;
      var val = fixNum_(r[3]);
      if (/^INV/i.test(cod)) saldoInv += val; else saldoOp += val;
    });
  }

  // 2. DF e DV mensais do planejado (aba resumo)
  var shResumo = ss.getSheetByName('resumo');
  var dfMensal = 0, dvMensal = 0;
  if (shResumo && shResumo.getLastRow() >= 2) {
    shResumo.getRange(2,1,shResumo.getLastRow()-1,9).getValues().forEach(function(r) {
      var tipo = String(r[0]||'').trim().toUpperCase();
      var val  = Math.abs(fixNum_(r[3]));
      if (/fix|^df$/i.test(tipo))  dfMensal += val;
      if (/vari|^dv$/i.test(tipo)) dvMensal += val;
    });
  }

  // 3. Entradas previstas em lancamentos por mês (apenas valores positivos)
  var entradasPorMes = {};
  var shLanc = ss.getSheetByName('lancamentos');
  if (shLanc && shLanc.getLastRow() >= 2) {
    shLanc.getRange(2,1,shLanc.getLastRow()-1,3).getValues().forEach(function(r) {
      var data = fixDate_(r[0]);
      if (!(data instanceof Date)||isNaN(data)) return;
      var val = fixNum_(r[2]); if (val <= 0) return;
      var chave = data.getFullYear()+'-'+String(data.getMonth()+1).padStart(2,'0');
      entradasPorMes[chave] = (entradasPorMes[chave]||0) + val;
    });
  }

  // 4. Projeção mês a mês — saldo acumulado
  var meses = [];
  var saldo = saldoOp;

  for (var i = 0; i < nMeses; i++) {
    var d2    = new Date(hoje.getFullYear(), hoje.getMonth()+i, 1);
    var chave = d2.getFullYear()+'-'+String(d2.getMonth()+1).padStart(2,'0');
    var label = Utilities.formatDate(d2, tz, 'MM/yyyy');
    var entradas = entradasPorMes[chave] || 0;
    saldo += entradas - dfMensal - dvMensal;
    meses.push({
      mes:       label,
      base:      Math.round(saldo*100)/100,
      patrimonio:Math.round((saldo+saldoInv)*100)/100,
      entradas:  Math.round(entradas*100)/100,
      df:        Math.round(dfMensal*100)/100,
      dv:        Math.round(dvMensal*100)/100,
    });
  }

  return {
    saldoOp:   Math.round(saldoOp*100)/100,
    saldoInv:  Math.round(saldoInv*100)/100,
    dfMensal:  Math.round(dfMensal*100)/100,
    dvMensal:  Math.round(dvMensal*100)/100,
    meses:     meses,
  };
  } catch(e) { return { erro: e.message }; }
}

// ── FLUXO UNIFICADO (Fluxo de Caixa + DV Estimada) ──────────────────────────
// Usa exatamente a mesma lógica do getDadosLancamentos para o saldo
// e adiciona DV e DV Estimada por mês
function getFluxoUnificado(nMeses) {
  try {
  nMeses = Math.min(Math.max(parseInt(nMeses)||12, 6), 36);
  var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tz  = getTZ_();
  var hoje = new Date(); hoje.setHours(0,0,0,0);
  var mesAtualChave = hoje.getFullYear()+'-'+pad_(hoje.getMonth()+1);

  // 1. Saldo investimentos (para patrimônio)
  var shSaldos = ss.getSheetByName('saldos');
  var saldoInv = 0;
  if (shSaldos && shSaldos.getLastRow() >= 2) {
    shSaldos.getRange(2,1,shSaldos.getLastRow()-1,4).getValues().forEach(function(r){
      var cod = String(r[0]||'').trim(); if(!cod) return;
      if(/^INV/i.test(cod)) saldoInv += fixNum_(r[3]);
    });
  }

  // 2. DV mensal planejado (aba resumo)
  var shResumo = ss.getSheetByName('resumo');
  var dvMensal = 0, dfMensal = 0;
  if (shResumo && shResumo.getLastRow() >= 2) {
    shResumo.getRange(2,1,shResumo.getLastRow()-1,9).getValues().forEach(function(r){
      var tipo = String(r[0]||'').trim().toUpperCase();
      var val  = Math.abs(fixNum_(r[3]));
      if(/vari|^dv$/i.test(tipo)) dvMensal += val;
      if(/fix|^df$/i.test(tipo))  dfMensal += val;
    });
  }

  // 3. Saldo acumulado por mês — mesma lógica do getDadosLancamentos
  // inclui saldos das contas no mês atual + todos os lançamentos
  var saldoPorMes = {}; // chave -> saldo acumulado do mês (fluxo líquido)
  // 3a. Saldos das contas (entram no mês atual)
  if (shSaldos && shSaldos.getLastRow() >= 2) {
    shSaldos.getRange(2,1,shSaldos.getLastRow()-1,4).getValues().forEach(function(r){
      var cod = String(r[0]||'').trim(); if(!cod) return;
      if(/^INV/i.test(cod)) return; // investimentos não entram no fluxo operacional
      var val = fixNum_(r[3]);
      saldoPorMes[mesAtualChave] = (saldoPorMes[mesAtualChave]||0) + val;
    });
  }
  // 3b. Lançamentos (todos, incluindo faturas col G)
  var shLanc = ss.getSheetByName('lancamentos');
  if (shLanc && shLanc.getLastRow() >= 2) {
    shLanc.getRange(2,1,shLanc.getLastRow()-1,3).getValues().forEach(function(r){
      var data = fixDate_(r[0]);
      if(!(data instanceof Date)||isNaN(data)) return;
      var chave = data.getFullYear()+'-'+pad_(data.getMonth()+1);
      saldoPorMes[chave] = (saldoPorMes[chave]||0) + fixNum_(r[2]);
    });
  }

  // 4. DV por mês: lancamentos tipo DV (sem col G) + CartaoC tipo DV
  var dvPorMes = {};
  if (shLanc && shLanc.getLastRow() >= 2) {
    shLanc.getRange(2,1,shLanc.getLastRow()-1,7).getValues().forEach(function(r){
      if(String(r[6]||'').trim()) return; // fatura cartão — ignorar
      if(!/vari|^dv$/i.test(String(r[5]||'').trim().toUpperCase())) return;
      var data = fixDate_(r[0]);
      if(!(data instanceof Date)||isNaN(data)) return;
      var chave = data.getFullYear()+'-'+pad_(data.getMonth()+1);
      dvPorMes[chave] = (dvPorMes[chave]||0) + fixNum_(r[2]);
    });
  }
  var shCartao = ss.getSheetByName('CartaoC');
  if (shCartao && shCartao.getLastRow() >= 2) {
    shCartao.getRange(2,2,shCartao.getLastRow()-1,4).getValues().forEach(function(r){
      if(!/vari|^dv$/i.test(String(r[3]||'').trim().toUpperCase())) return;
      var data = fixDate_(r[0]);
      if(!(data instanceof Date)||isNaN(data)) return;
      var chave = data.getFullYear()+'-'+pad_(data.getMonth()+1);
      dvPorMes[chave] = (dvPorMes[chave]||0) + fixNum_(r[2]);
    });
  }

  // 5. Projeção mês a mês
  var meses = [];
  var sFluxo = 0; // acumula o saldo do Fluxo de Caixa
  var sHibr  = 0; // acumula o saldo Híbrido (Fluxo - dvEstimada)

  for (var i = 0; i < nMeses; i++) {
    var d2    = new Date(hoje.getFullYear(), hoje.getMonth()+i, 1);
    var chave = d2.getFullYear()+'-'+pad_(d2.getMonth()+1);
    var label = Utilities.formatDate(d2, tz, 'MM/yyyy');

    // Saldo do mês (idêntico ao Fluxo de Caixa)
    var fluxoMes = saldoPorMes[chave] || 0;

    // DV do mês para dvEstimada
    var dvMes      = Math.abs(dvPorMes[chave] || 0);
    var dvEstimada = Math.max(dvMensal - dvMes, 0);

    sFluxo += fluxoMes;
    sHibr  += fluxoMes - dvEstimada;

    meses.push({
      mes:        label,
      chave:      chave,
      fluxo:      Math.round(sFluxo*100)/100,
      hibrido:    Math.round(sHibr*100)/100,
      projecao:   Math.round(sHibr*100)/100, // mantido para compatibilidade
      patrimHibr: Math.round((sHibr+saldoInv)*100)/100,
      entradas:   Math.round(sFluxo*100)/100, // saldo acumulado = "entradas líquidas"
      dvMes:      Math.round(dvMes*100)/100,
      dvEstimada: Math.round(dvEstimada*100)/100,
      fluxoLiq:   Math.round(sHibr*100)/100,
    });
  }

  return {
    saldoOp:       Math.round(sFluxo*100)/100,
    saldoInv:      Math.round(saldoInv*100)/100,
    dfMensal:      Math.round(dfMensal*100)/100,
    dvMensal:      Math.round(dvMensal*100)/100,
    dvBkpMesAtual: 0,
    meses:         meses,
  };
  } catch(e) { return {erro: e.message}; }
}

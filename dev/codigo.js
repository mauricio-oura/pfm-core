function novoLancamento_(dados) {
  Logger.log("novoLancamento_ recebeu: recorrente=" + dados.recorrente + " numMeses=" + dados.numMeses + " dataInicial=" + dados.dataInicial + " manual=" + dados.manual);
  if (!dados) throw new Error("Dados nao informados.");
  var valor = Math.abs(fixNum_(dados.valor));
  var descricao = String(dados.descricao || "").trim();
  var contaNome = String(dados.conta || "").trim();
  var direcao = String(dados.direcao || "Saida").trim();
  var tipo = String(dados.tipo || "DV").trim();
  var manual = (dados.manual === true || dados.manual === "true" || dados.manual === "SIM") ? "SIM" : "";
  var recorrente = dados.recorrente === true || dados.recorrente === "true";
  var numMeses = recorrente ? (parseInt(dados.numMeses) || 1) : 1;
  var dataInicial = recorrente && dados.dataInicial ? new Date(dados.dataInicial + "T12:00:00") : new Date();

  if (!valor || valor <= 0) throw new Error("Valor deve ser maior que zero.");
  if (!descricao) throw new Error("Descricao e obrigatoria.");
  if (!contaNome) throw new Error("Conta e obrigatoria.");
  var contaCod = CONTAS[contaNome];
  if (!contaCod) throw new Error("Conta invalida: " + contaNome);
  var sinal = direcao === "Entrada" ? 1 : -1;
  var valorAjustado = valor * sinal;
  var agora = new Date();
  var tz = getTZ_();
  var dataFormatada = Utilities.formatDate(agora, tz, "dd/MM/yyyy");
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var registros = 0;

  if (recorrente || manual === "SIM") {
    // Recorrente OU aguardando confirmacao: entra em lancamentos com SIM
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
      var descRec = descricao;  // sem numeracao em contas — diferente de cartao parcelado
      novasLinhas.push([dataLanc, descRec, valorAjustado, contaCod, "SIM", tipo]);
      registros++;
    }

    var proximaLinha = lancSheet.getLastRow() + 1;
    if (proximaLinha < 2) { proximaLinha = 2; }
    lancSheet.getRange(proximaLinha, 1, novasLinhas.length, 6).setValues(novasLinhas);
    var totalLinhas = lancSheet.getLastRow() - 1;
    if (totalLinhas > 1) { lancSheet.getRange(2, 1, totalLinhas, 6).sort({ column: 1, ascending: true }); }
    SpreadsheetApp.flush();
    CacheService.getScriptCache().removeAll(["getDadosLancamentos", "getDadosDV"]);

  } else {
    // Sem aguardando e sem recorrente: vai direto para bkp
    var bkpSheet = ss.getSheetByName("bkp");
    var saldosSheet = ss.getSheetByName("saldos");
    if (!bkpSheet) throw new Error("Aba bkp nao encontrada.");
    if (!saldosSheet) throw new Error("Aba saldos nao encontrada.");
    var lastSaldo = saldosSheet.getLastRow();
    if (lastSaldo >= 2) {
      var codigos = saldosSheet.getRange(2, 1, lastSaldo - 1, 1).getValues();
      for (var i = 0; i < codigos.length; i++) {
        if (String(codigos[i][0]).trim() === contaCod) {
          var saldoCell = saldosSheet.getRange(i + 2, 4);
          saldoCell.setValue(Number(saldoCell.getValue()) + valorAjustado);
          saldosSheet.getRange(i + 2, 2).setValue(agora);
          break;
        }
      }
    }
    var dataHoraExec = Utilities.formatDate(agora, tz, "dd/MM/yyyy HH:mm:ss");
    bkpSheet.getRange(bkpSheet.getLastRow() + 1, 1, 1, 6).setValues([[agora, descricao, valorAjustado, contaCod, tipo, dataHoraExec]]);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().removeAll(["getDadosLancamentos", "getDeltaBkp", "getDadosDV"]);
    registros = 1;
  }
  return { ok: true, registros: registros, registro: { data: dataFormatada, descricao: descricao, conta: contaCod, contaNome: contaNome, direcao: direcao, valor: valor, valorAjustado: valorAjustado, tipo: tipo, manual: manual }};
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
    if (totalLinhas > 1) { lancSheet.getRange(2, 1, totalLinhas, 6).sort({ column: 1, ascending: true }); }
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
  var bkpSheet = ss.getSheetByName("bkp");
  var saldosSheet = ss.getSheetByName("saldos");
  if (!lancSheet) { return { erro: "Aba lancamentos nao encontrada." }; }
  var lastRow = lancSheet.getLastRow();
  if (lastRow < 2) { return { ok: true, processados: 0 }; }
  var totalLinhas = lastRow - 1;
  lancSheet.getRange(2, 1, totalLinhas, 6).sort({ column: 1, ascending: true });
  SpreadsheetApp.flush();
  lastRow = lancSheet.getLastRow();
  if (lastRow < 2) { return { ok: true, processados: 0 }; }
  var dados = lancSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  var tz = getTZ_();
  var processados = 0;
  var linhasRemover = [];
  var ultimaLinhaSaldos = saldosSheet.getLastRow();
  var codigos = saldosSheet.getRange(2, 1, ultimaLinhaSaldos - 1, 4).getValues();
  dados.forEach(function(l, i) {
    var data = fixDate_(l[0]);
    var descricao = String(l[1] || "").trim();
    var valor = Number(l[2]) || 0;
    var conta = String(l[3] || "").trim();
    var aguardando = String(l[4] || "").trim().toUpperCase();
    var tipo = String(l[5] || "").trim();
    if (!(data instanceof Date) || isNaN(data)) { return; }
    if (aguardando === "SIM") { return; }
    var dataLanc = new Date(data);
    dataLanc.setHours(0, 0, 0, 0);
    if (dataLanc > hoje) { return; }
    for (var j = 0; j < codigos.length; j++) {
      if (String(codigos[j][0]).trim() === conta) {
        var saldoCell = saldosSheet.getRange(j + 2, 4);
        var saldoAtual = Number(saldoCell.getValue()) || 0;
        var novoSaldo = saldoAtual + valor;
        saldoCell.setValue(novoSaldo);
        SpreadsheetApp.flush();
        codigos[j][3] = novoSaldo;
        Logger.log("Saldo " + conta + ": " + saldoAtual + " + " + valor + " = " + novoSaldo);
        break;
      }
    }
    var dataHoraExec = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
    bkpSheet.getRange(bkpSheet.getLastRow() + 1, 1, 1, 6).setValues([[data, descricao, valor, conta, tipo, dataHoraExec]]);
    linhasRemover.push(i + 2);
    processados++;
  });
  linhasRemover.sort(function(a, b) { return b - a; });
  linhasRemover.forEach(function(linha) { lancSheet.deleteRow(linha); });
  SpreadsheetApp.flush();
  CacheService.getScriptCache().removeAll(["getDadosLancamentos", "getDeltaBkp", "getDadosDV"]);
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
    detalhes[chave].push({ linhaPlanilha: "L" + (i + 2), data: pad_(data.getDate()) + "/" + mes + "/" + ano, descricao: descricao, categoria: conta, tipo: tipo, valor: valor });
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
        var ano = dataSaldo.getFullYear();
        var mes = pad_(dataSaldo.getMonth() + 1);
        var chave = ano + "-" + mes;
        cats[codigo] = true;
        if (!mensais[chave]) { mensais[chave] = {}; }
        mensais[chave][codigo] = (mensais[chave][codigo] || 0) + valor;
        if (!detalhes[chave]) { detalhes[chave] = []; }
        detalhes[chave].push({ linhaPlanilha: "S" + codigo, data: pad_(dataSaldo.getDate()) + "/" + mes + "/" + ano, descricao: "Saldo " + codigo + " (" + desc + ")", categoria: codigo, tipo: "NA", valor: valor });
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
  cachePut_('getDadosLancamentos', result);
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
  var mensais = {}, detalhes = {}, cats = {};
  sheet.getRange(2, 2, lastRow - 1, 7).getValues().forEach(function(l, i) {
    var data = fixDate_(l[0]);
    if (!(data instanceof Date) || isNaN(data) || data < limite) { return; }
    var val = fixNum_(l[4]), cat = l[6] ? String(l[6]).trim() : "Sem Categoria", desc = l[1] ? String(l[1]).trim() : "Sem descricao";
    var ano = data.getFullYear(), mes = pad_(data.getMonth() + 1), chave = ano + "-" + mes;
    cats[cat] = true;
    if (!mensais[chave]) { mensais[chave] = {}; }
    mensais[chave][cat] = (mensais[chave][cat] || 0) + val;
    if (!detalhes[chave]) { detalhes[chave] = []; }
    detalhes[chave].push({ linhaPlanilha: i + 2, data: pad_(data.getDate()) + "/" + mes + "/" + ano, descricao: desc, categoria: cat, valor: val });
  });
  var chaves = Object.keys(mensais).sort().slice(0, 12), catsArray = Object.keys(cats).sort();
  var result2 = {
    meses: chaves.map(function(chave) {
      var p = chave.split('-'), porCat = {};
      catsArray.forEach(function(c) { porCat[c] = mensais[chave][c] || 0; });
      return { chaveOrdenacao: chave, mesAno: p[1] + "/" + p[0], valoresPorCategoria: porCat, detalhes: detalhes[chave] || [] };
    }),
    categorias: catsArray
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
  var totA = 0, totS = 0;
  function proc(dc, vc) {
    var d = fixDate_(dc instanceof Date ? dc : new Date(dc));
    if (!(d instanceof Date) || isNaN(d)) { return; }
    var v = fixNum_(vc);
    if (d.getMonth() === mA && d.getFullYear() === aA) { totA += v; }
    else if (d.getMonth() === mS && d.getFullYear() === aS) { totS += v; }
  }
  var sBkp = ss.getSheetByName("bkp");
  if (sBkp && sBkp.getLastRow() >= 2) {
    sBkp.getRange(2, 1, sBkp.getLastRow() - 1, 10).getValues().forEach(function(l) { if (String(l[9] || "").trim().toUpperCase() === "DV") { proc(l[0], l[6]); } });
  }
  var sLanc = ss.getSheetByName("lancamentos");
  if (sLanc && sLanc.getLastRow() >= 4) {
    sLanc.getRange(4, 2, sLanc.getLastRow() - 3, 7).getValues().forEach(function(l) { if (String(l[6] || "").trim().toUpperCase() === "DV") { proc(l[0], l[2]); } });
  }
  var sCart = ss.getSheetByName("CartaoC");
  if (sCart && sCart.getLastRow() >= 2) {
    sCart.getRange(2, 2, sCart.getLastRow() - 1, 6).getValues().forEach(function(l) { if (String(l[5] || "").trim().toUpperCase() === "DV") { proc(l[0], l[4]); } });
  }
  var result4 = { mesAtual: totA, mesSeguinte: totS, totalPrevisto: totA + totS };
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
  var dados = sheet.getRange(startRow, 2, numRows, 7).getValues();
  var result = [];
  for (var i = dados.length - 1; i >= 0; i--) {
    var l = dados[i];
    var data = fixDate_(l[0]);
    result.push({ linha: startRow + i, data: (data instanceof Date && !isNaN(data)) ? pad_(data.getDate()) + "/" + pad_(data.getMonth() + 1) + "/" + data.getFullYear() : "", descricao: l[1] ? String(l[1]).trim() : "", valor: fixNum_(l[4]), tipo: l[5] ? String(l[5]).trim() : "", cartao: l[6] ? String(l[6]).trim() : "" });
  }
  return result;
}

var DRIVE_FOLDER_ID = "1RZoRIVw8KETxOaxKZ-ZyFkBdbY0w9dnw";
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
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var files = folder.getFilesByName(nomeArquivo);
    if (!files.hasNext()) { return []; }
    return JSON.parse(files.next().getBlob().getDataAsString());
  } catch(err) {
    Logger.log("Erro ao ler JSON do Drive: " + err.message);
    return [];
  }
}

function salvarJsonDrive_(nomeArquivo, dados) {
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
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

function buscarCartaoC(termo) {
  try {
    inicializarCartoes_(); // garantir que CARTOES_POR_EMAIL está populado
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("CartaoC");
    if (!sheet || sheet.getLastRow() < 2) { return []; }
    var dados = sheet.getRange(2, 1, sheet.getLastRow()-1, 8).getValues();
    var tz = getTZ_();
    var termo_lower = termo.toLowerCase();
    var resultado = [];
    dados.forEach(function(l, i) {
      var descricao = String(l[2] || "").trim();
      if (descricao.toLowerCase().indexOf(termo_lower) === -1) { return; }
      var dataFatura = fixDate_(l[1]);
      var valor = fixNum_(l[5]);
      var tipo = String(l[6] || "").trim();
      var cartaoCodigo = String(l[7] || "").trim();
      // Descobrir nome do cartao pelo codigo
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
      var row = new Array(8).fill("");
      row[0] = agora;
      row[1] = dataFatura;
      row[2] = descricao;
      row[5] = valor * sinal;
      row[6] = tipo;
      row[7] = cartao;
      novasLinhas.push(row);
    }
    var ultimaLinha = sheet.getLastRow() + 1;
    sheet.getRange(ultimaLinha, 1, novasLinhas.length, 8).setValues(novasLinhas);
    SpreadsheetApp.flush();
    ordenarCartaoC_();
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

function cachePut_(key, data) {
  try {
    var json = JSON.stringify(data);
    // GAS limita cache a 100KB por entrada
    if (json.length < 90000) {
      CacheService.getScriptCache().put(key, json, CACHE_TTL);
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
      // Aguardando — se desmarcado, confirmar
      row[4] = dados.aguardando ? "SIM" : "";
      row[5] = tipo;
      sheet.getRange(linha, 1, 1, 6).setValues([row]);
      // Se confirmado, processar
      if (!dados.aguardando) {
        SpreadsheetApp.flush();
        var lancRow = sheet.getRange(linha, 1, 1, 6).getValues()[0];
        confirmarLancamento_({ linha: String(linha) });
      }
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
      var rowC = sheet3.getRange(linha, 1, 1, 8).getValues()[0];
      // Data fatura
      if (dados.data) {
        var p2 = dados.data.split('-');
        rowC[1] = new Date(parseInt(p2[0]), parseInt(p2[1])-1, parseInt(p2[2]), 12);
      }
      rowC[2] = desc;
      // Valor — manter sinal
      if (dados.valor !== undefined && !isNaN(parseFloat(dados.valor))) {
        var sinalC = rowC[5] < 0 ? -1 : 1;
        rowC[5] = Math.abs(parseFloat(dados.valor)) * sinalC;
      }
      rowC[6] = tipo;
      sheet3.getRange(linha, 1, 1, 8).setValues([rowC]);
      SpreadsheetApp.flush();
      ordenarCartaoC_();
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
      ordenarCartaoC_();
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

// ── SETUP TRIGGER KEEP-ALIVE ─────────────────────────────────────
function setupKeepAlive() {
  // Deletar triggers antigos
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'keepAlive') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Criar novo a cada 4 minutos (abaixo do cold start de 5min)
  ScriptApp.newTrigger('keepAlive')
    .timeBased()
    .everyMinutes(4)
    .create();
  Logger.log('keepAlive trigger criado: a cada 4 minutos');
}

function keepAlive() {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getName();
    inicializarCartoes_();
    inicializarChaves_();
    // Pre-aquecer cache das funcoes principais
    try { getDadosLancamentos(); } catch(e) {}
    try { getDadosCartaoC(); }    catch(e) {}
    try { getDadosDV(); }         catch(e) {}
    CacheService.getScriptCache().put("keepAlive_ts", String(Date.now()), 600);
    Logger.log("keepAlive OK: " + new Date().toLocaleTimeString());
  } catch(err) {
    Logger.log("keepAlive erro: " + err.message);
  }
}

// ── NOMES DOS CARTÕES ─────────────────────────────────────────────
var NOMES_CARTAO = {
  "MASTERCARD-MMO-1676": "Mastercard Mauricio",
  "VISA-HENRIQUE-4478":  "Visa Henrique",
  "VISA-JANE-4478":      "Visa Jane",
  "VISA-MMO-4478":       "Visa Mauricio"
};

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
    // Inserir config padrão dos cartões (linhas de config ficam com Mes_Fatura vazio)
    var configs = [
      ["MASTERCARD-MMO-1676","Mastercard Mauricio","BK1","","","","config",""],
      ["VISA-HENRIQUE-4478", "Visa Henrique",      "BK1","","","","config",""],
      ["VISA-JANE-4478",     "Visa Jane",           "BK1","","","","config",""],
      ["VISA-MMO-4478",      "Visa Mauricio",       "BK1","","","","config",""]
    ];
    sheet.getRange(2, 1, configs.length, 8).setValues(configs);
    sheet.setFrozenRows(1);
    Logger.log("Aba faturas criada com sucesso.");
  }
  return sheet;
}

// ── CALCULAR E SINCRONIZAR FATURAS ────────────────────────────────
function calcularFaturas() {
  inicializarCartoes_();
  var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  var cartaoC   = ss.getSheetByName("CartaoC");
  var lancSheet = ss.getSheetByName("lancamentos");
  var fatSheet  = criarAbaFaturas_();
  if (!cartaoC || !lancSheet) { return { erro: "Abas necessarias nao encontradas." }; }

  var hoje = new Date(); hoje.setHours(0,0,0,0);
  var tz   = getTZ_();

  // 1. Ler CartaoC e agrupar por cartao+mes — apenas datas >= hoje
  var lastC = cartaoC.getLastRow();
  if (lastC < 2) { return { ok: true, criadas: 0, atualizadas: 0 }; }
  var dadosC = cartaoC.getRange(2, 1, lastC-1, 8).getValues();

  var grupos = {};
  dadosC.forEach(function(l) {
    var dataFatura = fixDate_(l[1]);
    if (!(dataFatura instanceof Date) || isNaN(dataFatura)) { return; }
    dataFatura.setHours(0,0,0,0);
    if (dataFatura < hoje) { return; }
    var cartaoCod = String(l[7]||"").trim();
    if (!cartaoCod) { return; }
    // Chave única: CODIGO_YYYY-MM
    var anoMes = dataFatura.getFullYear() + "-" + pad_(dataFatura.getMonth()+1);
    var chave  = cartaoCod + "_" + anoMes;
    if (!grupos[chave]) {
      grupos[chave] = {
        cartaoCod:   cartaoCod,
        cartaoNome:  NOMES_CARTAO[cartaoCod] || cartaoCod,
        contaDebito: CONTA_POR_CARTAO[cartaoCod] || "BK1",
        dataFatura:  dataFatura,
        anoMes:      anoMes,
        mesFatura:   pad_(dataFatura.getMonth()+1) + "/" + dataFatura.getFullYear(),
        total:       0
      };
    }
    grupos[chave].total += fixNum_(l[5]);
  });

  // 2. Ler aba faturas — usar col D como string YYYY-MM (nunca Date)
  var lastF = fatSheet.getLastRow();
  var faturasExist = {}; // chave → { linhaFat, valorAjustado, idLanc }
  if (lastF >= 2) {
    var dadosF = fatSheet.getRange(2, 1, lastF-1, 8).getValues();
    dadosF.forEach(function(r, i) {
      var cod  = String(r[0]||"").trim();
      var anoM = String(r[3]||"").trim(); // deve ser "YYYY-MM"
      var stat = String(r[6]||"").trim();
      if (!cod || !anoM || stat === "config") { return; }
      var chave = cod + "_" + anoM;
      faturasExist[chave] = {
        linhaFat:      i+2,
        valorAjustado: fixNum_(r[5]),
        idLanc:        String(r[7]||"").trim()  // ID da chave, nao linha
      };
    });
  }

  // 3. Ler lancamentos — mapear id_fatura (col G) → linha atual
  var lastL = lancSheet.getLastRow();
  var lancMap = {}; // chave → linha atual na planilha
  if (lastL >= 2) {
    var dadosL = lancSheet.getRange(2, 1, lastL-1, 7).getValues();
    dadosL.forEach(function(r, i) {
      var idFat = String(r[6]||"").trim();
      if (idFat) { lancMap[idFat] = i+2; }
    });
  }

  var criadas = 0, atualizadas = 0, removidas = 0;

  // 4. Processar cada grupo — criar ou atualizar
  Object.keys(grupos).forEach(function(chave) {
    var g     = grupos[chave];
    var total = Math.round(g.total * 100) / 100;
    var desc  = g.cartaoNome + " " + g.mesFatura;

    if (faturasExist[chave]) {
      var linhaF = faturasExist[chave].linhaFat;
      fatSheet.getRange(linhaF, 5).setValue(total);
      var linhaL = lancMap[chave];
      if (linhaL && !faturasExist[chave].valorAjustado) {
        lancSheet.getRange(linhaL, 3).setValue(total);
      }
      atualizadas++;
    } else {
      var novaLinhaF = fatSheet.getLastRow() + 1;
      fatSheet.getRange(novaLinhaF, 1, 1, 8).setValues([[
        g.cartaoCod, g.cartaoNome, g.contaDebito,
        g.anoMes, total, "", "pendente", chave
      ]]);
      fatSheet.getRange(novaLinhaF, 4).setNumberFormat("@");
      var novaLinhaL = lancSheet.getLastRow() + 1;
      lancSheet.getRange(novaLinhaL, 1, 1, 7).setValues([[
        g.dataFatura, desc, total, g.contaDebito, "SIM", "DF", chave
      ]]);
      criadas++;
    }
  });

  // 5. Remover faturas e lancamentos cujos grupos sumiram do CartaoC
  SpreadsheetApp.flush();
  var lastF2 = fatSheet.getLastRow();
  var linhasRemoverFat = [];
  if (lastF2 >= 2) {
    var dadosF2 = fatSheet.getRange(2, 1, lastF2-1, 8).getValues();
    dadosF2.forEach(function(r, i) {
      var stat = String(r[6]||"").trim();
      if (stat === "config") { return; }
      var anoM = String(r[3]||"").trim();
      var cod  = String(r[0]||"").trim();
      if (!cod || !anoM) { return; }
      var chave = cod + "_" + anoM;
      if (!grupos[chave]) {
        linhasRemoverFat.push(i + 2);
        // Limpar lancamento correspondente
        var linhaL = lancMap[chave];
        if (linhaL) {
          lancSheet.getRange(linhaL, 1, 1, 7).clearContent();
          removidas++;
        }
      }
    });
  }

  // Deletar linhas vazias de lancamentos (de baixo para cima)
  if (removidas > 0) {
    SpreadsheetApp.flush();
    var lastL2 = lancSheet.getLastRow();
    if (lastL2 >= 2) {
      var dadosL2 = lancSheet.getRange(2, 1, lastL2-1, 1).getValues();
      var linhasVazias = [];
      dadosL2.forEach(function(r, i) {
        if (!r[0]) { linhasVazias.push(i + 2); }
      });
      linhasVazias.sort(function(a,b){ return b-a; });
      linhasVazias.forEach(function(l){ lancSheet.deleteRow(l); });
    }
  }

  // Deletar linhas de faturas (de baixo para cima)
  linhasRemoverFat.sort(function(a,b){ return b-a; });
  linhasRemoverFat.forEach(function(l){ fatSheet.deleteRow(l); });

  // 6. Ordenar lancamentos
  var totalLinhas = lancSheet.getLastRow() - 1;
  if (totalLinhas > 1) {
    lancSheet.getRange(2, 1, totalLinhas, 7).sort({ column: 1, ascending: true });
  }

  SpreadsheetApp.flush();
  CacheService.getScriptCache().removeAll(["getDadosLancamentos", "getDadosDV"]);
  Logger.log("calcularFaturas: criadas=" + criadas + " atualizadas=" + atualizadas + " removidas=" + removidas);
  return { ok: true, criadas: criadas, atualizadas: atualizadas, removidas: removidas };
}

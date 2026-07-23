// ── CONFIG ──────────────────────────────────────────────────────────
// Lê a aba "config" da planilha (A=chave, B=valor) e cacha em memória.
// Permite personalizar o app por cliente sem alterar código.
var _config = null;
function getConfig_() {
  if (_config) return _config;
  _config = {
    app_nome:              'Painel Financeiro',
    app_cor_primaria:      '#1e3a5f',
    drive_folder_id:       '',
    fuso_horario:          'America/Sao_Paulo',
    email_admin:           '',
    moeda:                 'BRL',
    moeda_simbolo:         'R$',
    limite_pendentes_dias: '7',
    cache_ttl_segundos:    '300',
    cartao_dias_antes_venc: '7'
  };
  try {
    var ss    = SpreadsheetApp.openById(getSpreadsheetId_());
    var sheet = ss.getSheetByName('config');
    if (sheet && sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function(r) {
        var chave = String(r[0] || '').trim();
        var valor = String(r[1] || '').trim();
        if (chave) _config[chave] = valor;
      });
    }
  } catch(e) {
    Logger.log('getConfig_ erro: ' + e.message);
  }
  return _config;
}

// Emails autorizados a acessar o painel principal — lidos da config
var EMAILS_AUTORIZADOS = (function() {
  try {
    var cfg = getConfig_();
    var email = cfg['email_admin'] || '';
    return email ? email.split(',').map(function(e){ return e.trim(); }) : [];
  } catch(e) { return []; }
})();

// Chaves de acesso por pessoa — inclua no link: ?page=cartao&key=CHAVE
// Cada pessoa salva o link com a chave no celular como favorito
var CHAVES_CARTAO = {};
var CHAVES_CARTAO_CODIGO = {};
function inicializarChaves_() {
  // Ler da aba VencCC — col G=chave, col E=codigo, col F=emails
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("VencCC");
    if (sheet && sheet.getLastRow() >= 2) {
      var numCols = Math.max(sheet.getLastColumn(), 8);
      var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, numCols).getValues();
      rows.forEach(function(r) {
        var codigo = String(r[4]||"").trim();
        var emails = String(r[5]||"").trim();
        var chave  = String(r[6]||"").trim();
        if (!chave || !codigo) { return; }
        // Mapear chave para o primeiro email da lista
        var emailList = emails ? emails.split(",") : [];
        if (emailList.length > 0) {
          CHAVES_CARTAO[chave] = emailList[0].trim();
        }
        CHAVES_CARTAO_CODIGO[chave] = codigo;
      });
      return;
    }
  } catch(e) {
    Logger.log("inicializarChaves_ erro: " + e.message);
  }
  // Sem fallback hardcoded — dados de acesso devem estar na aba VencCC
  Logger.log("inicializarChaves_: aba VencCC sem dados de chaves. Configure a coluna G (chave de acesso).");
}

var CARTOES_POR_EMAIL = {};
var FINAL_CARTAO = {};

var CONTA_POR_CARTAO = {};

function inicializarCartoes_() {
  // Ler da aba VencCC — fonte de verdade para cartoes
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("VencCC");
    if (sheet && sheet.getLastRow() >= 2) {
      var numCols = Math.max(sheet.getLastColumn(), 8);
      var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, numCols).getValues();
      rows.forEach(function(r) {
        var ultimos = String(r[0]||"").trim();
        var nome    = String(r[3]||"").trim();
        var codigo  = String(r[4]||"").trim();
        var emails  = String(r[5]||"").trim();
        var conta   = String(r[7]||"").trim() || "BK1";
        if (!ultimos || !codigo) { return; }
        // Icone baseado no codigo
        var icone = codigo.indexOf("MASTER") >= 0 ? "[MC]"
                  : codigo.indexOf("VISA")   >= 0 ? "[VI]" : "[CC]";
        var cartaoObj = { codigo: codigo, nome: nome, ultimos: ultimos, icone: icone };
        // Registrar por cada email dono
        var emailList = emails ? emails.split(",") : [];
        emailList.forEach(function(email) {
          email = email.trim();
          if (!email) { return; }
          if (!CARTOES_POR_EMAIL[email]) { CARTOES_POR_EMAIL[email] = []; }
          // Evitar duplicata
          var jaExiste = CARTOES_POR_EMAIL[email].some(function(c) { return c.codigo === codigo; });
          if (!jaExiste) { CARTOES_POR_EMAIL[email].push(cartaoObj); }
        });
        FINAL_CARTAO[codigo]    = ultimos;
        CONTA_POR_CARTAO[codigo] = conta;
      });
      return; // sucesso — nao precisa fallback
    }
  } catch(e) {
    Logger.log("inicializarCartoes_ erro ao ler VencCC: " + e.message);
  }
  // Sem fallback hardcoded — dados de cartões devem estar na aba VencCC
  Logger.log("inicializarCartoes_: aba VencCC não encontrada ou sem dados. Configure os cartões.");
}

// SPREADSHEET_ID vem do Script Properties (Arquivo > Propriedades do projeto > Script properties)
// Chave: SPREADSHEET_ID  Valor: ID da planilha do cliente
// Isso evita que o ID fique exposto no código-fonte.
var _spreadsheetId = null;
function getSpreadsheetId_() {
  if (!_spreadsheetId) {
    _spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
      || SpreadsheetApp.getActiveSpreadsheet().getId(); // fallback se rodando dentro da planilha
  }
  return _spreadsheetId;
}
var SPREADSHEET_ID = getSpreadsheetId_();

// CONTAS é populado por inicializarContas_() que lê a aba "contas" da planilha.
// Não há fallback hardcoded — cada cliente configura suas próprias contas.
var CONTAS = {
};

var PWA_META_TAGS = [
  ["mobile-web-app-capable", "yes"],
  ["apple-mobile-web-app-capable", "yes"],
  ["apple-mobile-web-app-status-bar-style", "black-translucent"],
  ["theme-color", "#1e3a5f"]
];

function addPwaMeta_(out, titulo) {
  out.addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  return out;
}

function doGet(e) {
  inicializarCartoes_();
  inicializarContas_();
  var params = {};
  if (e && e.parameter) { params = e.parameter; }

  if (params.page === 'cartao') {
    inicializarChaves_();
    inicializarCartoes_();
    var key = params.key || '';
    var emailKey = key ? (CHAVES_CARTAO[key] || '') : '';
    var acessoKey = emailKey !== '';
    // Se chave tem cartão específico, mostrar só esse cartão
    var codigoEspecifico = CHAVES_CARTAO_CODIGO[key] || '';
    var todosCartoes = acessoKey ? (CARTOES_POR_EMAIL[emailKey] || []) : [];
    var cartoesKey = codigoEspecifico
      ? todosCartoes.filter(function(c) { return c.codigo === codigoEspecifico; })
      : todosCartoes;
    var tmplC = HtmlService.createTemplateFromFile('cartao');
    tmplC.emailUsuario = emailKey;
    tmplC.acesso = acessoKey;
    tmplC.cartoes = cartoesKey;
    tmplC.chave = key;
    var outC = tmplC.evaluate();
    outC.setTitle('Lancamento Cartao');
    addPwaMeta_(outC, 'Lançar Cartão');
    return outC;
  }

  if (!params.page || params.page === 'painel') {
    var emailPainel = '';
    try { emailPainel = Session.getActiveUser().getEmail(); } catch(err) { emailPainel = ''; }
    if (!emailPainel || !EMAILS_AUTORIZADOS.includes(emailPainel)) {
      return HtmlService.createHtmlOutput('<h2>Acesso negado</h2>');
    }
    var tmplP = HtmlService.createTemplateFromFile('index_gas');
    tmplP.emailUsuario = emailPainel;
    tmplP.appNome = getConfig_()['app_nome'] || 'Painel Financeiro';
    // PIN mobile — lido da aba config (coluna A = "pin", coluna B = valor)
    try {
      var cfgPin = getConfig_();
      tmplP.pinMobile = String(cfgPin['pin'] || '');
    } catch(ePin) { tmplP.pinMobile = ''; }
    var outP = tmplP.evaluate();
    outP.setTitle(tmplP.appNome);
    outP.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    return outP;
  }

  if (params.page === 'saldos') {
    var emailS = '';
    try { emailS = Session.getActiveUser().getEmail(); } catch(err) { emailS = ''; }
    if (!emailS || !EMAILS_AUTORIZADOS.includes(emailS)) {
      return HtmlService.createHtmlOutput('<h2>Acesso negado</h2>');
    }
    var tmplS = HtmlService.createTemplateFromFile('saldos');
    tmplS.emailUsuario = emailS;
    var outS = tmplS.evaluate();
    outS.setTitle('Saldos');
    addPwaMeta_(outS, 'Saldos');
    return outS;
  }

  if (params.page === 'lancar') {
    var emailL = '';
    try { emailL = Session.getActiveUser().getEmail(); } catch(err) { emailL = ''; }
    if (!emailL || !EMAILS_AUTORIZADOS.includes(emailL)) {
      return HtmlService.createHtmlOutput('<h2>Acesso negado</h2>');
    }
    inicializarContas_();
    var tmplL = HtmlService.createTemplateFromFile('lancar');
    tmplL.emailUsuario = emailL;
    tmplL.contasLista  = getContas().filter(function(c){ return c.ativo !== false; });
    var hojeDate = new Date();
    var tz = getTZ_();
    tmplL.hojeISO = hojeDate.getFullYear() + '-' +
      pad_(hojeDate.getMonth()+1) + '-' +
      pad_(hojeDate.getDate());
    var outL = tmplL.evaluate();
    outL.setTitle('Lancamento');
    addPwaMeta_(outL, 'Novo Lançamento');
    return outL;
  }

  if (params.page === 'projetar_cartao') {
    var emailPc = '';
    try { emailPc = Session.getActiveUser().getEmail(); } catch(err) { emailPc = ''; }
    if (!emailPc || !EMAILS_AUTORIZADOS.includes(emailPc)) {
      return HtmlService.createHtmlOutput('<h2>Acesso negado</h2>');
    }
    inicializarCartoes_();
    var tmplPc = HtmlService.createTemplateFromFile('projetar_cartao');
    tmplPc.emailUsuario = emailPc;
    tmplPc.cartoesLista = getCartoesParaModal();
    var outPc = tmplPc.evaluate();
    outPc.setTitle('Projetar Cartao');
    outPc.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    return outPc;
  }

  if (params.page === 'projetar') {
    var emailPr = '';
    try { emailPr = Session.getActiveUser().getEmail(); } catch(err) { emailPr = ''; }
    if (!emailPr || !EMAILS_AUTORIZADOS.includes(emailPr)) {
      return HtmlService.createHtmlOutput('<h2>Acesso negado</h2>');
    }
    var tmplPr = HtmlService.createTemplateFromFile('projetar');
    tmplPr.emailUsuario = emailPr;
    var outPr = tmplPr.evaluate();
    outPr.setTitle('Projetar Recorrente');
    outPr.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    return outPr;
  }

  if (params.page === 'ajuste-fatura') {
    var emailAj = '';
    try { emailAj = Session.getActiveUser().getEmail(); } catch(err) { emailAj = ''; }
    if (!emailAj || !EMAILS_AUTORIZADOS.includes(emailAj)) {
      return HtmlService.createHtmlOutput('<h2>Acesso negado</h2>');
    }
    inicializarCartoes_();
    var tmplAj = HtmlService.createTemplateFromFile('ajuste_fatura');
    tmplAj.emailUsuario = emailAj;
    tmplAj.appNome = getConfig_()['app_nome'] || 'Painel Financeiro';
    tmplAj.faturasLista = getFaturasParaAjuste();
    var outAj = tmplAj.evaluate();
    outAj.setTitle('Ajuste de Fatura');
    outAj.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    return outAj;
  }

  if (params.page === 'venccc') {
    var emailV = '';
    try { emailV = Session.getActiveUser().getEmail(); } catch(err) { emailV = ''; }
    if (!emailV || !EMAILS_AUTORIZADOS.includes(emailV)) {
      return HtmlService.createHtmlOutput('<h2>Acesso negado</h2>');
    }
    inicializarCartoes_();
    var tmplV = HtmlService.createTemplateFromFile('venccc');
    tmplV.emailUsuario = emailV;
    tmplV.vencDados = getVencCC();
    var outV = tmplV.evaluate();
    outV.setTitle('Vencimentos Cartao');
    outV.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    return outV;
  }

  if (params.page === 'contas') {
    var emailCo = '';
    try { emailCo = Session.getActiveUser().getEmail(); } catch(err) { emailCo = ''; }
    if (!emailCo || !EMAILS_AUTORIZADOS.includes(emailCo)) {
      return HtmlService.createHtmlOutput('<h2>Acesso negado</h2>');
    }
    var tmplCo = HtmlService.createTemplateFromFile('contas');
    tmplCo.emailUsuario = emailCo;
    var outCo = tmplCo.evaluate();
    outCo.setTitle('Contas');
    outCo.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    return outCo;
  }

  return HtmlService.createHtmlOutput('ok');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

var _tz = null;
function getTZ_() {
  if (!_tz) { _tz = SpreadsheetApp.openById(SPREADSHEET_ID).getSpreadsheetTimeZone(); }
  return _tz;
}

function pad_(n) { return n < 10 ? '0' + n : String(n); }

function fixNum_(v) {
  if (typeof v === 'number') { return v; }
  var s = String(v || '').replace(',', '.').trim();
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function fixDate_(v) {
  if (v instanceof Date) { return v; }
  if (typeof v === 'string' && v.trim()) {
    var parts = v.trim().split('/');
    if (parts.length === 3) {
      return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    }
  }
  return new Date(v);
}

function getBaseUrl() {
  return ScriptApp.getService().getUrl();
}

function limparCacheGAS_() {
  try {
    CacheService.getScriptCache().removeAll([
      'getDadosLancamentos','getDadosCartaoC','getDadosResumo',
      'getDadosDV','getDeltaBkp','getDeltaCartao','getHistoricoBkp','getHistoricoCartao',
      'getDashboard','getDadosLancPendentes'
    ]);
  } catch(err) {}
  _config = null;
  return { ok: true };
}

function limparCacheDashboard() {
  try { CacheService.getScriptCache().remove('getDashboard'); } catch(e) {}
  return { ok: true };
}



function chamarAcao(acao) {
  inicializarCartoes_();
  var ACOES = {
    'getDadosLancamentos': getDadosLancamentos,
    'getDadosCartaoC': getDadosCartaoC,
    'getDadosResumo': getDadosResumo,
    'getDadosDV': getDadosDV,
    'getDeltaBkp': getDeltaBkp,
    'getDeltaCartao': getDeltaCartao,
    'getHistoricoBkp': getHistoricoBkp,
    'getHistoricoCartao': getHistoricoCartao,
    'limparCache': limparCacheGAS_,
    'processarLancamentos': processarLancamentos,
    'calcularFaturas': calcularFaturas,
    'getCartoesParaModal': getCartoesParaModal,
    'getContas': getContas,
    'getDashboard': getDashboard,
    'getSugestoesProjetar': getSugestoesProjetar,
    'getSugestoesCartaoC': getSugestoesCartaoC,
    'getFaturasParaAjuste': getFaturasParaAjuste,
    'getBaseUrl': function() { return ScriptApp.getService().getUrl(); },
    'desmarcarAguardando': desmarcarAguardando,
    'getProjecao': getProjecao,
    'getFluxoUnificado': getFluxoUnificado,
  };
  var fn = ACOES[acao];
  if (!fn) { return { erro: 'Acao desconhecida: ' + acao }; }
  try { return fn(); } catch(err) { return { erro: err.message }; }
}

function confirmarLancamentoPublico(linha) {
  try { return confirmarLancamento_({ linha: String(linha) }); }
  catch(err) { return { erro: err.message }; }
}

// getCartoesUsuario removido — autenticacao agora por chave no URL

function registrarCartao(dados) {
  inicializarCartoes_();
  inicializarChaves_();
  try {
    var valor = Math.abs(fixNum_(dados.valor));
    var descricao = String(dados.descricao || '').trim();
    var cartao = String(dados.cartao || '').trim();
    var tipo = String(dados.tipo || 'DV').trim();
    var direcao = String(dados.direcao || 'Saida').trim();
    var chave = String(dados.chave || '').trim();

    // Valida pela chave — sem necessidade de login Google
    var emailR = chave ? (CHAVES_CARTAO[chave] || '') : '';
    if (!emailR) {
      return { erro: 'Acesso nao autorizado. Verifique o link.' };
    }
    var cartoesPermitidos = (CARTOES_POR_EMAIL[emailR] || []).map(function(c) { return c.codigo; });
    if (!cartoesPermitidos.includes(cartao)) {
      return { erro: 'Cartao nao autorizado para este usuario.' };
    }
    if (!valor) { return { erro: 'Valor obrigatorio.' }; }
    if (!descricao) { return { erro: 'Descricao obrigatoria.' }; }
    if (!cartao) { return { erro: 'Cartao nao informado.' }; }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var cartaoC = ss.getSheetByName('CartaoC');
    var vencCC = ss.getSheetByName('VencCC');
    if (!cartaoC) { return { erro: "Aba 'CartaoC' nao encontrada." }; }

    // Buscar dia de vencimento (col B) do cartão na VencCC
    var diaVencCartao = 5; // fallback
    if (vencCC && vencCC.getLastRow() >= 2) {
      var rowsVenc = vencCC.getRange(2, 1, vencCC.getLastRow()-1, 5).getValues();
      for (var iv = 0; iv < rowsVenc.length; iv++) {
        if (String(rowsVenc[iv][4]||'').trim() === cartao) {
          var dv = parseInt(rowsVenc[iv][1]) || 0;
          if (dv > 0) { diaVencCartao = dv; break; }
        }
      }
    }

    // Dias antes do vencimento = período de corte (da config, padrão 7)
    var diasCorte = parseInt(getConfig_()['cartao_dias_antes_venc']) || 7;

    // Calcular primeira fatura a partir de hoje + regra de corte
    // Corte = vencimento - diasCorte dias
    // Se hoje <= corte → fatura vence neste mês; senão → próximo mês
    var agora = new Date();
    agora.setHours(0,0,0,0);

    function proximaDataFatura(mesesOffset) {
      // Mês base = primeiro vencimento (considerando corte), + offset de parcelas
      var hoje = new Date(agora);
      // Data de corte do mês atual
      var corteMesAtual = new Date(hoje.getFullYear(), hoje.getMonth(), diaVencCartao - diasCorte);
      // Se diasCorte > diaVencCartao, corte cai no mês anterior
      if (diaVencCartao - diasCorte <= 0) {
        // Corte no mês anterior: ex venc=5, corte=7 → corte=28 de mês anterior
        var diasNoMes = new Date(hoje.getFullYear(), hoje.getMonth(), 0).getDate();
        corteMesAtual = new Date(hoje.getFullYear(), hoje.getMonth()-1, diasNoMes + (diaVencCartao - diasCorte));
      }
      // Mês base do vencimento
      var mesFatura = hoje <= corteMesAtual
        ? hoje.getMonth()       // ainda no corte — fatura este mês
        : hoje.getMonth() + 1;  // passou do corte — próxima fatura
      return new Date(hoje.getFullYear(), mesFatura + mesesOffset, diaVencCartao);
    }

    var sinal = direcao === 'Entrada' ? 1 : -1;
    var parcelado   = dados.parcelado === true || dados.parcelado === "true" || dados.parcelado === 1;
    var numParcelas = parcelado ? Math.max(1, parseInt(dados.numParcelas) || 1) : 1;
    var registros   = 0;
    Logger.log('registrarCartao: cartao=' + cartao + ' diaVenc=' + diaVencCartao + ' diasCorte=' + diasCorte + ' parcelado=' + parcelado + ' parcelas=' + numParcelas);

    var novasLinhas = [];
    var valorParcela = valor / numParcelas;

    for (var p = 0; p < numParcelas; p++) {
      var dataVencParcela = proximaDataFatura(p);
      var descParcela = parcelado ? descricao + " parc. " + (p+1) + " de " + numParcelas : descricao;
      var row = new Array(6).fill('');
      row[0] = agora;
      row[1] = dataVencParcela;
      row[2] = descParcela;
      row[3] = (parcelado ? valorParcela : valor) * sinal;
      row[4] = tipo;
      row[5] = cartao;
      novasLinhas.push(row);
      registros++;
    }

    // Inserir todas as linhas de uma vez
    var ultimaLinha = cartaoC.getLastRow() + 1;
    cartaoC.getRange(ultimaLinha, 1, novasLinhas.length, 6).setValues(novasLinhas);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().removeAll(['getDadosCartaoC', 'getDeltaCartao']);
    // Recalcular faturas automaticamente após novo lançamento de cartão
    try { calcularFaturas(); } catch(errF) { Logger.log("calcularFaturas erro: " + errF.message); }
    return { ok: true, registros: registros };
  } catch(err) {
    return { erro: err.message };
  }
}

function sincronizarFaturas() {
  inicializarCartoes_();
  inicializarChaves_();
  return calcularFaturas();
}

function registrarLancamento(dados) {
  Logger.log("registrarLancamento recebeu: " + JSON.stringify(dados));
  try {
    inicializarContas_(); // garantir que CONTAS está populado
    if (!dados) { return { erro: 'Dados nao informados.' }; }
    var op = String(dados.op || '').trim();
    var valor = Math.abs(fixNum_(dados.valor));
    var descricao = String(dados.descricao || '').trim();
    var manual = dados.manual;

    if (!valor) { return { erro: 'Valor e obrigatorio.' }; }
    if (!descricao) { return { erro: 'Descricao e obrigatoria.' }; }

    if (op === 'transf') {
      var origem = String(dados.origem || '').trim();
      var destino = String(dados.destino || '').trim();
      if (!CONTAS[origem]) { return { erro: 'Conta de origem invalida.' }; }
      if (!CONTAS[destino]) { return { erro: 'Conta de destino invalida.' }; }
      return novoLancamentoTransf_({ valor: valor, descricao: descricao, origem: origem, destino: destino, manual: manual });
    } else {
      var conta = String(dados.conta || '').trim();
      var tipo = String(dados.tipo || 'DV').trim();
      var direcao = op === 'saida' ? 'Saida' : 'Entrada';
      if (!CONTAS[conta]) { return { erro: 'Conta invalida: ' + conta }; }
      return novoLancamento_({ valor: valor, descricao: descricao, conta: conta, direcao: direcao, tipo: tipo, manual: manual,
        recorrente: dados.recorrente, dataInicial: dados.dataInicial, numMeses: dados.numMeses });
    }
  } catch(err) {
    return { erro: err.message };
  }
}

function getSaldos() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('saldos');
  if (!sheet || sheet.getLastRow() < 2) { return []; }
  var tz = getTZ_();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(function(l) { return l[0]; })
    .map(function(l) {
      var dataAtual = '';
      if (l[1] instanceof Date && !isNaN(l[1])) {
        dataAtual = Utilities.formatDate(l[1], tz, 'dd/MM/yyyy');
      }
      return { conta: String(l[0]).trim(), dataAtual: dataAtual, descricao: String(l[2]).trim(), saldo: fixNum_(l[3]) };
    });
}

function atualizarSaldo(dados) {
  try {
    var conta = String(dados.conta || '').trim();
    var novoSaldo = fixNum_(dados.novoSaldo);
    if (!conta) { return { erro: 'Conta nao informada.' }; }
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('saldos');
    if (!sheet) { return { erro: "Aba 'saldos' nao encontrada." }; }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { erro: 'Nenhuma conta cadastrada em saldos.' }; }
    var codigos = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var encontrou = false;
    for (var i = 0; i < codigos.length; i++) {
      if (String(codigos[i][0]).trim() === conta) {
        sheet.getRange(i + 2, 2).setValue(new Date());
        sheet.getRange(i + 2, 4).setValue(novoSaldo);
        encontrou = true;
        break;
      }
    }
    if (!encontrou) { return { erro: "Conta '" + conta + "' nao encontrada." }; }
    SpreadsheetApp.flush();
    CacheService.getScriptCache().removeAll(['getDadosLancamentos', 'getDadosDV']);
    return { ok: true };
  } catch(err) {
    return { erro: err.message };
  }
}

function keepAlive() {
  // Mantém o script ativo e pré-aquece as funções principais
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.getName();
  inicializarCartoes_();
  inicializarChaves_();
  try {
    var cache = CacheService.getScriptCache();
    cache.put("keepAlive_ts", String(Date.now()), 300);
    // Pre-aquecer getDadosLancamentos no cache
    try { getDadosLancamentos(); } catch(e) {}
  } catch(err) {}
}


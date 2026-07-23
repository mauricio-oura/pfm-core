// ══════════════════════════════════════════════════════════════════
// SETUP.GS — Funções de instalação e onboarding
// Execute estas funções UMA VEZ ao configurar o app para um cliente.
// Ordem recomendada:
//   1. configurarScriptProperties()  → define o SPREADSHEET_ID
//   2. criarAbaConfig()              → cria aba config na planilha
//   3. criarEstruturaPlanilha()      → cria todas as abas necessárias
//   4. instalarTriggers()            → instala triggers automáticos
// ══════════════════════════════════════════════════════════════════


// ── 1. SCRIPT PROPERTIES ─────────────────────────────────────────
// Define o SPREADSHEET_ID nas propriedades do projeto.
// Substitua o valor pelo ID da planilha do cliente antes de executar.
function configurarScriptProperties() {
  var SPREADSHEET_ID_CLIENTE = 'COLE_AQUI_O_ID_DA_PLANILHA';

  if (SPREADSHEET_ID_CLIENTE === 'COLE_AQUI_O_ID_DA_PLANILHA') {
    Logger.log('⚠️ Edite a variável SPREADSHEET_ID_CLIENTE antes de executar.');
    return { ok: false, msg: 'SPREADSHEET_ID não configurado.' };
  }

  PropertiesService.getScriptProperties()
    .setProperty('SPREADSHEET_ID', SPREADSHEET_ID_CLIENTE);

  Logger.log('✅ SPREADSHEET_ID configurado: ' + SPREADSHEET_ID_CLIENTE);
  return { ok: true, msg: 'SPREADSHEET_ID salvo nas Script Properties.' };
}


// ── 2. ABA CONFIG ────────────────────────────────────────────────
// Cria a aba "config" com todas as chaves e valores padrão.
// Se a aba já existir, não faz nada.
function criarAbaConfig() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('config');
  if (sheet) {
    Logger.log('ℹ️ Aba config já existe — nenhuma alteração feita.');
    return { ok: false, msg: 'Aba config já existe.' };
  }

  sheet = ss.insertSheet('config');
  var defaults = [
    ['chave',                 'valor',                   'descrição'],
    ['app_nome',              'Painel Financeiro',       'Nome exibido no app e no título da aba'],
    ['app_cor_primaria',      '#1e3a5f',                 'Cor principal do tema (hex)'],
    ['drive_folder_id',       '',                        'ID da pasta no Google Drive para os JSONs históricos'],
    ['fuso_horario',          'America/Sao_Paulo',       'Fuso horário para datas e horas'],
    ['email_admin',           '',                        'Email(s) do(s) administrador(es), separados por vírgula'],
    ['moeda',                 'BRL',                     'Código da moeda (ISO 4217)'],
    ['moeda_simbolo',         'R$',                      'Símbolo exibido nos valores'],
    ['limite_pendentes_dias', '7',                       'Quantos dias à frente mostrar no bloco Pendentes'],
    ['cache_ttl_segundos',    '300',                     'Tempo de cache server-side em segundos (máx 21600)'],
    ['cartao_dias_antes_venc','7',                       'Dias antes do vencimento usados como corte da fatura do cartão']
  ];

  sheet.getRange(1, 1, defaults.length, 3).setValues(defaults);

  // Formatar cabeçalho
  sheet.getRange(1, 1, 1, 3)
    .setBackground('#1e3a5f')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // Larguras
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);

  // Proteger coluna A (chaves não devem ser renomeadas)
  var prot = sheet.getRange('A:A').protect();
  prot.setDescription('Chaves de configuração — não renomeie');
  prot.setWarningOnly(true);

  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();

  Logger.log('✅ Aba config criada com sucesso.');
  return { ok: true, msg: 'Aba config criada. Preencha os valores na coluna B.' };
}


// ── 3. ESTRUTURA DA PLANILHA ──────────────────────────────────────
// Cria todas as abas necessárias para o app funcionar.
// Abas já existentes são ignoradas.
function criarEstruturaPlanilha() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var criadas = [], existentes = [];

  var abas = [
    {
      nome: 'lancamentos',
      headers: ['Data', 'Descricao', 'Valor', 'Conta', 'Aguardando', 'Tipo'],
      cor: '#1e3a5f'
    },
    {
      nome: 'bkp',
      headers: ['Data', 'Descricao', 'Valor', 'Conta', 'Tipo', 'DataHoraExec'],
      cor: '#065f46'
    },
    {
      nome: 'saldos',
      headers: ['Codigo', 'Data', 'Descricao', 'Saldo'],
      cor: '#7c2d12'
    },
    {
      nome: 'contas',
      headers: ['Codigo', 'Nome', 'Tipo', 'Ativo'],
      cor: '#1e40af'
    },
    {
      nome: 'CartaoC',
      headers: ['Data', 'Descricao', 'Valor', 'Tipo', 'Cartao'],
      cor: '#6b21a8'
    },
    {
      nome: 'VencCC',
      headers: ['Ultimos', 'DiaVenc', 'ProxVenc', 'Nome', 'Codigo', 'Emails', 'Chave', 'Conta'],
      cor: '#7c3aed'
    },
    {
      nome: 'faturas',
      headers: ['Cartao_Codigo', 'Cartao_Nome', 'Conta_Debito', 'Mes_Fatura',
                'Valor_Calculado', 'Valor_Ajustado', 'Status', 'Linha_Lancamento'],
      cor: '#0f766e'
    },
    {
      nome: 'resumo',
      headers: ['Categoria', 'Tipo', 'Valor', 'Conta', 'Ativo'],
      cor: '#7e22ce'
    }
  ];

  abas.forEach(function(aba) {
    var sheet = ss.getSheetByName(aba.nome);
    if (sheet) {
      existentes.push(aba.nome);
      return;
    }
    sheet = ss.insertSheet(aba.nome);
    sheet.getRange(1, 1, 1, aba.headers.length).setValues([aba.headers]);
    sheet.getRange(1, 1, 1, aba.headers.length)
      .setBackground(aba.cor)
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    criadas.push(aba.nome);
  });

  SpreadsheetApp.flush();

  var msg = 'Criadas: [' + criadas.join(', ') + '] | Já existiam: [' + existentes.join(', ') + ']';
  Logger.log('✅ ' + msg);
  return { ok: true, criadas: criadas, existentes: existentes };
}


// ── 4. TRIGGERS ───────────────────────────────────────────────────
// Instala os triggers automáticos do app.
// Pode ser executado novamente para resetar os triggers.
function instalarTriggers() {
  var GERENCIADOS = ['keepAlive', 'processarLancamentosHorario'];

  // Remover triggers gerenciados existentes
  var removidos = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (GERENCIADOS.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
      removidos++;
    }
  });

  // keepAlive: a cada 5 minutos — mantém o servidor GAS aquecido
  ScriptApp.newTrigger('keepAlive')
    .timeBased()
    .everyMinutes(5)
    .create();

  // Processar lançamentos: a cada 1 hora — executa lançamentos vencidos sem "SIM"
  ScriptApp.newTrigger('processarLancamentosHorario')
    .timeBased()
    .everyHours(1)
    .create();

  var msg = 'Triggers instalados: keepAlive (5min) + processarLancamentosHorario (1h). Removidos anteriores: ' + removidos;
  Logger.log('✅ ' + msg);
  return { ok: true, msg: msg };
}


// ── DIAGNÓSTICO ───────────────────────────────────────────────────
// Verifica se tudo está configurado corretamente.
// Execute para checar o estado do app num cliente.
function diagnosticar() {
  var resultado = { ok: true, itens: [] };

  function check(label, fn) {
    try {
      var r = fn();
      resultado.itens.push({ label: label, status: r.ok ? '✅' : '⚠️', msg: r.msg });
      if (!r.ok) resultado.ok = false;
    } catch(e) {
      resultado.itens.push({ label: label, status: '❌', msg: e.message });
      resultado.ok = false;
    }
  }

  // 1. Script Properties
  check('SPREADSHEET_ID', function() {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    return { ok: !!id, msg: id ? 'ID: ' + id : 'Não configurado — execute configurarScriptProperties()' };
  });

  // 2. Aba config
  check('Aba config', function() {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('config');
    if (!sheet) return { ok: false, msg: 'Não existe — execute criarAbaConfig()' };
    var cfg = getConfig_();
    var semEmail = !cfg['email_admin'];
    var semDrive = !cfg['drive_folder_id'];
    var avisos = [];
    if (semEmail) avisos.push('email_admin vazio');
    if (semDrive) avisos.push('drive_folder_id vazio');
    return { ok: avisos.length === 0, msg: avisos.length ? 'Atenção: ' + avisos.join(', ') : 'OK — app_nome: ' + cfg['app_nome'] };
  });

  // 3. Abas obrigatórias
  check('Abas da planilha', function() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var obrigatorias = ['lancamentos', 'bkp', 'saldos', 'contas', 'CartaoC', 'VencCC', 'faturas', 'resumo'];
    var faltando = obrigatorias.filter(function(nome) { return !ss.getSheetByName(nome); });
    return { ok: faltando.length === 0, msg: faltando.length ? 'Faltando: ' + faltando.join(', ') : 'Todas presentes' };
  });

  // 4. Triggers
  check('Triggers', function() {
    var triggers = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });
    var temKeepAlive = triggers.indexOf('keepAlive') >= 0;
    var temHorario   = triggers.indexOf('processarLancamentosHorario') >= 0;
    var faltando = [];
    if (!temKeepAlive) faltando.push('keepAlive');
    if (!temHorario)   faltando.push('processarLancamentosHorario');
    return { ok: faltando.length === 0, msg: faltando.length ? 'Faltando: ' + faltando.join(', ') + ' — execute instalarTriggers()' : 'OK (' + triggers.join(', ') + ')' };
  });

  // 5. Pasta Drive
  check('Pasta Drive', function() {
    var folderId = getConfig_()['drive_folder_id'];
    if (!folderId) return { ok: false, msg: 'drive_folder_id não configurado na aba config' };
    DriveApp.getFolderById(folderId).getName(); // lança erro se não existir
    return { ok: true, msg: 'Pasta acessível: ' + folderId };
  });

  // Log resumo
  resultado.itens.forEach(function(i) {
    Logger.log(i.status + ' ' + i.label + ': ' + i.msg);
  });
  Logger.log(resultado.ok ? '✅ Diagnóstico OK' : '⚠️ Há itens para corrigir');
  return resultado;
}


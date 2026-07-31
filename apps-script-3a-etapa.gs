/**
 * =====================================================================
 *  COPA FT BR SPORT DE ARRANCADA — 3ª ETAPA
 *  Google Apps Script da planilha de cadastros da Roleta de Brindes
 * =====================================================================
 *
 *  COMO INSTALAR
 *  -------------
 *  1. Crie uma planilha NOVA no Google Sheets (só para a 3ª etapa).
 *  2. Menu  Extensões → Apps Script.
 *  3. Apague o conteúdo do Code.gs e cole este arquivo inteiro.
 *  4. Salve. Depois rode a função `configurarPlanilha` uma única vez
 *     (menu de funções no topo → configurarPlanilha → Executar).
 *     Ela cria o cabeçalho das 10 colunas. Autorize quando pedir.
 *  5. Menu  Implantar → Nova implantação → tipo "App da Web":
 *        - Executar como:        Eu
 *        - Quem pode acessar:    Qualquer pessoa
 *     Implantar → copie a URL que termina em /exec.
 *  6. Cole essa URL na constante SHEETS_URL do index.html
 *     (fica no bloco "PAINEL DE CONTROLE DA ATIVAÇÃO", no topo do <script>).
 *
 *  ⚠️ Ao editar este script depois, é preciso criar uma NOVA VERSÃO da
 *     implantação (Implantar → Gerenciar implantações → editar → Nova versão),
 *     senão a URL continua servindo o código antigo.
 *
 *  COLUNAS DA PLANILHA
 *  -------------------
 *  A = DATA E HORA   F = ESTADO
 *  B = NOME          G = GÊNERO    (F = Moleca/feminino, M = Actvitta/masculino)
 *  C = E-MAIL        H = TAMANHO
 *  D = TELEFONE      I = VOUCHER
 *  E = CIDADE        J = PRÊMIO
 *
 *  ENDPOINTS USADOS PELA LANDING PAGE
 *  ----------------------------------
 *  ?action=check&email=&telefone=      → {exists: true|false}
 *  ?action=counts                      → {premios:{...}, clogs:{"F-36":n,...}}
 *  ?action=setprize&voucher=&premio=   → grava o prêmio na linha do voucher
 *  (sem action)                        → grava um novo cadastro
 * =====================================================================
 */

// Índices das colunas (1-based)
var COL = {
  DATA: 1, NOME: 2, EMAIL: 3, TELEFONE: 4, CIDADE: 5,
  ESTADO: 6, GENERO: 7, TAMANHO: 8, VOUCHER: 9, PREMIO: 10
};
var TOTAL_COLS = 10;

// Nome exato do brinde que consome a grade de tamanhos
var CLOG_PRIZE = 'Kit Clog + Pins FT';

// Valores da coluna PRÊMIO que NÃO contam como brinde entregue
var NAO_CONTA = ['', 'Esgotado'];


/** Cria o cabeçalho da planilha. Rode uma vez, manualmente. */
function configurarPlanilha() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cab = ['DATA E HORA', 'NOME', 'E-MAIL', 'TELEFONE', 'CIDADE',
             'ESTADO', 'GÊNERO', 'TAMANHO', 'VOUCHER', 'PRÊMIO'];
  sheet.getRange(1, 1, 1, cab.length)
       .setValues([cab])
       .setFontWeight('bold')
       .setBackground('#E12529')
       .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, cab.length);
}


function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || '';

  try {
    if (action === 'check')    return json(verificarDuplicata(params));
    if (action === 'counts')   return json(contarSaidas());
    if (action === 'setprize') return json(gravarPremio(params));
    return json(gravarCadastro(params));
  } catch (err) {
    return json({ error: String(err) });
  }
}

// A LP usa GET, mas deixamos POST funcionando pelo mesmo caminho.
function doPost(e) { return doGet(e); }


/* ---------------------------------------------------------------
   1. Anti-duplicata — e-mail ou telefone já cadastrado?
   --------------------------------------------------------------- */
function verificarDuplicata(params) {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { exists: false };

  var dados = sheet.getRange(2, COL.EMAIL, lastRow - 1, 2).getValues(); // C e D

  var email = String(params.email || '').toLowerCase().trim();
  var fone  = String(params.telefone || '').replace(/\D/g, '');

  for (var i = 0; i < dados.length; i++) {
    var eSheet = String(dados[i][0]).toLowerCase().trim();
    var fSheet = String(dados[i][1]).replace(/\D/g, '');

    if (email && eSheet && eSheet === email) return { exists: true };
    if (fone.length >= 10 && fSheet === fone) return { exists: true };
  }
  return { exists: false };
}


/* ---------------------------------------------------------------
   2. Contagem do que já saiu (alimenta o controle de estoque)
      premios → total por brinde
      clogs   → total do Kit Clog por GÊNERO-TAMANHO (ex.: "F-36")
   --------------------------------------------------------------- */
function contarSaidas() {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var premios = {}, clogs = {};
  if (lastRow < 2) return { premios: premios, clogs: clogs };

  // Lê da coluna GÊNERO até PRÊMIO de uma vez só (G..J)
  var largura = COL.PREMIO - COL.GENERO + 1;
  var dados   = sheet.getRange(2, COL.GENERO, lastRow - 1, largura).getValues();

  for (var i = 0; i < dados.length; i++) {
    var genero  = String(dados[i][0]).trim().toUpperCase();
    var tamanho = String(dados[i][1]).trim();
    var premio  = String(dados[i][3]).trim();

    if (NAO_CONTA.indexOf(premio) !== -1) continue;

    premios[premio] = (premios[premio] || 0) + 1;

    if (premio === CLOG_PRIZE && genero && tamanho) {
      // tamanho pode vir como "36" ou "36.0" se o Sheets tratar como número
      tamanho = tamanho.replace(/\.0+$/, '');
      var chave = genero + '-' + tamanho;
      clogs[chave] = (clogs[chave] || 0) + 1;
    }
  }
  return { premios: premios, clogs: clogs };
}


/* ---------------------------------------------------------------
   3. Novo cadastro (gravado já no envio do formulário, sem prêmio)
   --------------------------------------------------------------- */
function gravarCadastro(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.appendRow([
      params.datetime || '',   // A
      params.nome     || '',   // B
      params.email    || '',   // C
      params.telefone || '',   // D
      params.cidade   || '',   // E
      params.estado   || '',   // F
      params.genero   || '',   // G
      "'" + (params.tamanho || ''), // H — aspa força texto, preserva "36"
      params.voucher  || '',   // I
      params.premio   || ''    // J
    ]);
    return { status: 'ok' };
  } finally {
    lock.releaseLock();
  }
}


/* ---------------------------------------------------------------
   4. Grava o prêmio sorteado na linha do voucher
      Se o voucher não estiver na planilha (falha de rede no cadastro),
      cria a linha para não perder o registro.
   --------------------------------------------------------------- */
function gravarPremio(params) {
  var voucher = String(params.voucher || '').trim();
  var premio  = String(params.premio  || '').trim();
  if (!voucher) return { error: 'voucher ausente' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      var vouchers = sheet.getRange(2, COL.VOUCHER, lastRow - 1, 1).getValues();
      // De baixo para cima: pega a linha mais recente desse voucher
      for (var i = vouchers.length - 1; i >= 0; i--) {
        if (String(vouchers[i][0]).trim() === voucher) {
          var linha = i + 2;
          sheet.getRange(linha, COL.PREMIO).setValue(premio);
          if (params.genero)  sheet.getRange(linha, COL.GENERO).setValue(params.genero);
          if (params.tamanho) sheet.getRange(linha, COL.TAMANHO).setValue("'" + params.tamanho);
          return { status: 'ok', row: linha };
        }
      }
    }

    // Não achou: registra assim mesmo para não perder o prêmio entregue
    sheet.appendRow([
      new Date().toLocaleString('pt-BR'), '', '', '', '', '',
      params.genero || '', "'" + (params.tamanho || ''), voucher, premio
    ]);
    return { status: 'ok-appended' };
  } finally {
    lock.releaseLock();
  }
}


/* --------------------------------------------------------------- */
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Contador anônimo — PyAnalytics · Quiz "Trilha do Dado" (v10)
 *
 * Conta VISITANTES ÚNICOS POR DIA, não aberturas.
 * - O quiz envia um id anônimo de dispositivo (aleatório, sem nome nem e-mail)
 *   e conta no máximo 1 vez por dia por dispositivo.
 * - ping=1  -> só lê o total, nunca grava.
 * - sem id  -> não é pedido do quiz (alguém abriu o link à mão): só lê, não grava.
 * - debug=1 -> devolve também as últimas linhas (para diagnóstico).
 *
 * PRIVACIDADE: o "id do dispositivo" é um número aleatório gerado no navegador,
 * que não identifica a pessoa. Não guarda nome, e-mail nem localização.
 *
 * A aba "acessos" tem SÓ duas colunas: data_hora | id_dispositivo.
 * O dia é calculado a partir da data_hora — nada de guardar a mesma coisa duas vezes.
 * O resumo por dia é uma fórmula, criada pela função criarAbaPorDia().
 *
 * A aba "oficina" guarda o ranking ao vivo: uma linha por participante
 * (código da oficina + dispositivo), com apelido, níveis e acertos.
 */

var VERSAO = 'v10';
var ABA = 'acessos';
var CABECALHO = ['data_hora', 'id_dispositivo'];

// --- ranking ao vivo da oficina ---
var ABA_OFICINA = 'oficina';
var CAB_OFICINA = ['data_hora', 'codigo', 'apelido', 'id_dispositivo', 'niveis', 'acertos'];
var TOPO = 6;

function doGet(e) {
  var acao = (e && e.parameter && e.parameter.acao) ? String(e.parameter.acao) : '';
  if (acao === 'entrar' || acao === 'pontuar' || acao === 'ranking') return oficina(e, acao);

  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ABA);
    if (!sh) sh = ss.insertSheet(ABA);
    garantirCabecalho(sh);

    var id    = (e && e.parameter && e.parameter.id) ? String(e.parameter.id).trim().slice(0, 40) : '';
    var semId = (id === '');                                // pedido que não veio do quiz
    var ping  = !!(e && e.parameter && e.parameter.ping);   // só ler, não gravar
    var debug = !!(e && e.parameter && e.parameter.debug);  // diagnóstico
    var hoje  = diaTexto(new Date());

    // já existe este id hoje? o dia vem do carimbo de tempo, sempre uma data verdadeira
    var dados = sh.getDataRange().getValues();
    var jaContado = false;
    for (var i = 1; !semId && i < dados.length; i++) {
      if (String(dados[i][1]).trim() === id && diaTexto(dados[i][0]) === hoje) {
        jaContado = true;
        break;
      }
    }

    // grava só se: veio do quiz (tem id), ainda não contou hoje, e não é um simples ping
    if (!semId && !jaContado && !ping) {
      sh.appendRow([new Date(), id]);
      dados.push([new Date(), id]);
    }

    var total = Math.max(0, sh.getLastRow() - 1);
    var saida = { value: total, versao: VERSAO, hoje: hoje, jaContado: jaContado, semId: semId };

    if (debug) {
      var amostra = [];
      for (var j = Math.max(1, dados.length - 5); j < dados.length; j++) {
        amostra.push({ id: String(dados[j][1]), dia: diaTexto(dados[j][0]) });
      }
      saida.amostra = amostra;
    }

    return ContentService
      .createTextOutput(JSON.stringify(saida))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** Repõe o cabeçalho se ele faltar (ex.: depois de uma limpeza manual). */
function garantirCabecalho(sh) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(CABECALHO);
    return;
  }
  var a1 = String(sh.getRange(1, 1).getValue()).trim().toLowerCase();
  if (a1 !== 'data_hora') {
    sh.insertRowBefore(1);                                  // não perde a linha que lá estava
    sh.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]);
  }
}

/** Converte um carimbo de tempo em "AAAA-MM-DD" no fuso da planilha. */
function diaTexto(d) {
  if (!(d instanceof Date)) d = new Date(d);
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

// ===================== RANKING AO VIVO DA OFICINA =====================
// Uma linha por participante (código + dispositivo). Três ações:
//   acao=entrar   (cod, nome, id)              -> inscreve e devolve o ranking
//   acao=pontuar  (cod, id, niveis, acertos)   -> atualiza a pontuação
//   acao=ranking  (cod)                        -> só lê os 6 melhores

function oficina(e, acao) {
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ABA_OFICINA);
    if (!sh) sh = ss.insertSheet(ABA_OFICINA);
    if (sh.getLastRow() === 0) sh.appendRow(CAB_OFICINA);

    var cod = String(e.parameter.cod || '').trim().toUpperCase().slice(0, 12);
    var id  = String(e.parameter.id  || '').trim().slice(0, 40);
    if (!cod) return json({ erro: 'codigo em falta' });

    var dados = sh.getDataRange().getValues();

    if (acao === 'ranking') return json({ codigo: cod, ranking: topo(dados, cod) });
    if (!id) return json({ erro: 'id em falta' });

    var linha = -1;
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][1]).trim().toUpperCase() === cod && String(dados[i][3]).trim() === id) {
        linha = i + 1;
        break;
      }
    }

    if (acao === 'entrar') {
      var nome = String(e.parameter.nome || '').trim().slice(0, 24) || 'anónimo';
      if (linha > 0) {
        sh.getRange(linha, 3).setValue(nome);          // mudou de apelido
      } else {
        sh.appendRow([new Date(), cod, nome, id, 0, 0]);
      }
      return json({ ok: true, nome: nome, codigo: cod, ranking: topo(sh.getDataRange().getValues(), cod) });
    }

    // acao === 'pontuar' — a pontuação só sobe, nunca desce
    if (linha < 0) return json({ erro: 'nao inscrito' });
    var niveis  = Math.max(0, parseInt(e.parameter.niveis,  10) || 0);
    var acertos = Math.max(0, parseInt(e.parameter.acertos, 10) || 0);
    sh.getRange(linha, 5).setValue(Math.max(Number(sh.getRange(linha, 5).getValue()) || 0, niveis));
    sh.getRange(linha, 6).setValue(Math.max(Number(sh.getRange(linha, 6).getValue()) || 0, acertos));

    return json({ ok: true, ranking: topo(sh.getDataRange().getValues(), cod) });
  } finally {
    lock.releaseLock();
  }
}

/** Os TOPO melhores do código dado: mais níveis primeiro, empate decidido por acertos. */
function topo(dados, cod) {
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][1]).trim().toUpperCase() !== cod) continue;
    lista.push({
      nome: String(dados[i][2]),
      niveis: Number(dados[i][4]) || 0,
      acertos: Number(dados[i][5]) || 0
    });
  }
  lista.sort(function (a, b) {
    return (b.niveis - a.niveis) || (b.acertos - a.acertos) || a.nome.localeCompare(b.nome);
  });
  return lista.slice(0, TOPO);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===================== MANUTENÇÃO (correr à mão no editor) =====================
// Não são usadas pelo quiz. Corre pelo menu "Executar" do editor.
// Não é preciso reimplantar para elas funcionarem.

var IDS_TESTE = ['teste-claude', 'teste-dedupe-claude', 'sem-id'];

/**
 * Remove linhas repetidas (mesmo dispositivo no mesmo dia) e as linhas de teste.
 * Faz uma cópia de segurança da aba antes de mexer.
 */
function limparDuplicados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA);
  if (!sh) throw new Error('Aba "' + ABA + '" não encontrada.');

  var dados = sh.getDataRange().getValues();
  if (dados.length < 2) { Logger.log('Nada a limpar.'); return; }

  var carimbo = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd_HHmm');
  sh.copyTo(ss).setName('backup_' + ABA + '_' + carimbo);

  var vistos = {}, mantidas = [], removidas = 0;
  for (var i = 1; i < dados.length; i++) {
    var id = String(dados[i][1]).trim();
    if (IDS_TESTE.indexOf(id) >= 0) { removidas++; continue; }
    var chave = id + '|' + diaTexto(dados[i][0]);
    if (vistos[chave]) { removidas++; continue; }
    vistos[chave] = true;
    mantidas.push([dados[i][0], id]);
  }

  sh.clear();
  sh.appendRow(CABECALHO);
  if (mantidas.length) sh.getRange(2, 1, mantidas.length, 2).setValues(mantidas);

  Logger.log('Removidas: ' + removidas + ' linha(s). Acessos únicos restantes: ' + mantidas.length);
}

/** Zera o contador (ex.: para começar a oficina em 0), com cópia de segurança. */
function zerarContador() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA);
  if (!sh) { Logger.log('Aba "' + ABA + '" não existe; nada a zerar.'); return; }

  var carimbo = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd_HHmm');
  if (sh.getLastRow() > 1) sh.copyTo(ss).setName('backup_' + ABA + '_' + carimbo);

  sh.clear();
  sh.appendRow(CABECALHO);
  Logger.log('Contador zerado.');
}

/**
 * Cria (ou repõe) a aba de resumo por dia, com uma fórmula que se recalcula
 * sozinha. Corre à mão, uma vez. Não precisa de reimplantar.
 */
function criarAbaPorDia() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('por_dia') || ss.getSheetByName('por dia');
  if (!sh) sh = ss.insertSheet('por_dia');

  sh.clear();
  sh.getRange('A1').setFormula(
    '=QUERY(' + ABA + '!A2:B, "select toDate(A), count(B) ' +
    'where A is not null group by toDate(A) ' +
    'label toDate(A) \'dia\', count(B) \'acessos_unicos\'", 0)'
  );

  sh.getRange('A1:B1').setFontWeight('bold');
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 130);
  sh.setFrozenRows(1);

  Logger.log('Aba "' + sh.getName() + '" pronta. A fórmula atualiza-se sozinha.');
}

/* ---------------------------------------------------------------------------
 * RESUMO POR DIA — é uma fórmula, não uma aba mantida por código.
 * Corre criarAbaPorDia() uma vez; a fórmula recalcula-se sozinha a partir daí.
 * ------------------------------------------------------------------------- */

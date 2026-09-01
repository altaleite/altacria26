/***************************************
 * Alta CRIA 2026
 * Portal de validação + questionário completo
 *
 * Como usar:
 * 1) Crie um projeto no Google Apps Script.
 * 2) Cole este conteúdo em Code.gs.
 * 3) Crie um arquivo HTML chamado Index e cole o conteúdo de Index.html.
 * 4) Publique como Web App.
 ***************************************/

const CONFIG = {
  SPREADSHEET_ID: '1OmAP-jbZnD8_GpUUhXBOZb2T_cCzpS2JXv6BJKzNsKA',
  CADASTRO_GID: 2105652336,
  CADASTRO_SHEET_NAME: 'Base',
  RESPOSTAS_SHEET: 'Respostas_26',
  TENTATIVAS_SHEET: 'Tentativas_Acesso',
  ALTERACOES_SHEET: 'Alteracoes_Cadastro',
  RASCUNHOS_SHEET: 'Rascunhos_26',
  FORM_VERSION: 'Alta CRIA 2026 - questionário completo v4',
  MIN_NAME_SCORE: 0.70,
  MIN_CITY_SCORE: 0.90,
  MAX_ATTEMPTS_PER_CODE_PER_DAY: 20
};

/**
 * Ordem fixa e completa das colunas da aba de respostas.
 * Gerada a partir do esquema do formulário (9 campos de identificação +
 * 193 perguntas + 31 campos "Outro"). Manter esta ordem garante que o
 * dashboard sempre encontre cada pergunta na mesma coluna, independente de
 * quem enviar primeiro e de quais perguntas condicionais foram exibidas.
 */
const RESPOSTAS_COLUMNS = [
  'timestamp', 'codigoAcesso', 'nomeValidacao', 'cidadeValidacao',
  'statusResposta', 'validacaoAcesso', 'similaridadeNome', 'similaridadeCidade',
  'versaoFormulario', 'q004', 'q005', 'q006',
  'q007', 'q008', 'q009', 'q010',
  'q011', 'q012', 'q013', 'q014',
  'q015', 'q016', 'q017', 'q018',
  'q019', 'q020', 'q021', 'q022',
  'q023', 'q024', 'q025', 'q026',
  'q027', 'q028', 'q028_outro', 'q029',
  'q029_outro', 'q030', 'q030_outro', 'q031',
  'q032', 'q033', 'q034', 'q035',
  'q036', 'q037', 'q038', 'q039',
  'q040', 'q041', 'q042', 'q043',
  'q044', 'q044_outro', 'q045', 'q046',
  'q047', 'q048', 'q049', 'q050',
  'q051', 'q052', 'q052_outro', 'q053',
  'q054', 'q055', 'q056', 'q057',
  'q057_outro', 'q058', 'q059', 'q060',
  'q061', 'q062', 'q063', 'q064',
  'q065', 'q065_outro', 'q066', 'q067',
  'q068', 'q069', 'q069_outro', 'q070',
  'q071', 'q072', 'q073', 'q074',
  'q075', 'q076', 'q077', 'q078',
  'q079', 'q080', 'q081', 'q082',
  'q083', 'q084', 'q085', 'q086',
  'q086_outro', 'q087', 'q088', 'q089',
  'q090', 'q091', 'q092', 'q093',
  'q094', 'q095', 'q096', 'q097',
  'q098', 'q099', 'q100', 'q101',
  'q102', 'q103', 'q104', 'q104_outro',
  'q105', 'q106', 'q107', 'q108',
  'q109', 'q110', 'q110_outro', 'q111',
  'q112', 'q113', 'q114', 'q115',
  'q116', 'q117', 'q118', 'q119',
  'q120', 'q121', 'q122', 'q123',
  'q123_outro', 'q124', 'q125', 'q126',
  'q127', 'q128', 'q129', 'q130',
  'q131', 'q132', 'q132_outro', 'q133',
  'q133_outro', 'q134', 'q134_outro', 'q135',
  'q136', 'q137', 'q138', 'q138_outro',
  'q139', 'q140', 'q141', 'q141_outro',
  'q142', 'q143', 'q144', 'q145',
  'q145_outro', 'q146', 'q147', 'q148',
  'q148_outro', 'q149', 'q150', 'q150_outro',
  'q151', 'q151_outro', 'q152', 'q153',
  'q154', 'q155', 'q156', 'q156b',
  'q157', 'q157_outro', 'q158', 'q159',
  'q159_outro', 'q160', 'q160_outro', 'q161',
  'q162', 'q163', 'q164', 'q165',
  'q166', 'q167', 'q168', 'q169',
  'q170', 'q170_outro', 'q171', 'q171_outro',
  'q172', 'q173', 'q174', 'q174_outro',
  'q175', 'q176', 'q176b', 'q177',
  'q178', 'q179', 'q179_outro', 'q180',
  'q180_outro', 'q181', 'q182', 'q183',
  'q184', 'q185', 'q186', 'q187',
  'q187_outro', 'q188', 'q189', 'q189_outro',
  'q190', 'q191', 'q192', 'q193',
  'q194', 'q195',
];

/**
 * Rode UMA vez, manualmente, no editor do Apps Script (menu Executar >
 * setupPlanilha) antes de liberar o formulário. Cria as abas necessárias e
 * grava o cabeçalho fixo da aba de respostas, deixando o dashboard estável.
 */
function setupPlanilha() {
  var ss = getSpreadsheet_();
  ensureRespostasHeader_();
  [CONFIG.TENTATIVAS_SHEET, CONFIG.ALTERACOES_SHEET, CONFIG.RASCUNHOS_SHEET].forEach(function (name) {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
  var rasc = ss.getSheetByName(CONFIG.RASCUNHOS_SHEET);
  if (rasc && rasc.getLastRow() === 0) {
    rasc.appendRow(['codigo', 'atualizado_em', 'dados_json']);
  }
  return 'Planilha preparada: cabeçalho de respostas e abas auxiliares prontos.';
}

/**
 * Garante que a aba de respostas exista e tenha o cabeçalho fixo na ordem
 * canônica. Não sobrescreve dados já existentes.
 */
function ensureRespostasHeader_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.RESPOSTAS_SHEET) || ss.insertSheet(CONFIG.RESPOSTAS_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, RESPOSTAS_COLUMNS.length).setValues([RESPOSTAS_COLUMNS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  // O questionário continua sendo a resposta padrão. A rota abaixo serve
  // somente números agregados para o dashboard público, sem linhas cruas,
  // nomes, códigos ou textos digitados pelas fazendas.
  var params = e && e.parameter ? e.parameter : {};
  if (params.api === 'dados_publicos') {
    var json = gerarDadosPublicos();
    if (params.callback) {
      var callback = String(params.callback).trim();
      // JSONP é usado pelo dashboard para funcionar mesmo quando o navegador
      // bloquear uma chamada fetch entre domínios. Aceita apenas identificador
      // JavaScript simples, evitando inserir código arbitrário na resposta.
      if (!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
        return ContentService
          .createTextOutput(JSON.stringify({ erro: 'callback inválido' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService
        .createTextOutput(callback + '(' + json + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Questionário Alta CRIA 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Valida o acesso sem devolver dados cadastrais da fazenda.
 * payload esperado: { codigo, nome, cidade }
 */
function validateAccess(payload) {
  try {
    payload = payload || {};
    var codigo = String(payload.codigo || '').trim();
    var nome = String(payload.nome || '').trim();
    var cidade = String(payload.cidade || '').trim();

    if (!codigo || !nome || !cidade) {
      logAccessAttempt_(codigo, nome, cidade, 'FALHA_CAMPOS_OBRIGATORIOS', 0, 0, 'Campos obrigatórios não preenchidos');
      return publicValidationResult_(false);
    }

    if (tooManyAttempts_(codigo)) {
      logAccessAttempt_(codigo, nome, cidade, 'BLOQUEADO_TENTATIVAS', 0, 0, 'Limite diário de tentativas por código excedido');
      return { ok: false, message: 'Houve muitas tentativas para este código hoje. Fale com a equipe Alta CRIA.' };
    }

    var cadastro = getCadastroByCodigo_(codigo);
    if (!cadastro) {
      logAccessAttempt_(codigo, nome, cidade, 'FALHA_VALIDACAO', 0, 0, 'Código não localizado');
      return publicValidationResult_(false);
    }

    var validation = validateAgainstCadastro_(cadastro, nome, cidade);
    var status = validation.ok ? 'VALIDADO' : 'FALHA_VALIDACAO';
    logAccessAttempt_(codigo, nome, cidade, status, validation.nameScore, validation.cityScore, validation.reason);

    if (!validation.ok) return publicValidationResult_(false);

    return {
      ok: true,
      message: 'Acesso validado. Agora preencha o questionário Alta CRIA 2026.',
      draft: getServerDraft_(codigo)
    };
  } catch (err) {
    logAccessAttempt_(payload && payload.codigo, payload && payload.nome, payload && payload.cidade, 'ERRO', 0, 0, String(err));
    return publicValidationResult_(false);
  }
}

/**
 * Recebe o questionário completo.
 * Revalida código + nome + cidade antes de gravar.
 */
function submitQuestionario(dados) {
  try {
    dados = dados || {};
    var codigo = String(dados.codigoAcesso || '').trim();
    var nomeValidacao = String(dados.nomeValidacao || '').trim();
    var cidadeValidacao = String(dados.cidadeValidacao || '').trim();

    var cadastro = getCadastroByCodigo_(codigo);
    if (!cadastro) {
      logAccessAttempt_(codigo, nomeValidacao, cidadeValidacao, 'ENVIO_BLOQUEADO', 0, 0, 'Código não localizado no envio final');
      return { ok: false, message: 'Não foi possível validar o acesso. Confira os dados iniciais e tente novamente.' };
    }

    var validation = validateAgainstCadastro_(cadastro, nomeValidacao, cidadeValidacao);
    if (!validation.ok) {
      logAccessAttempt_(codigo, nomeValidacao, cidadeValidacao, 'ENVIO_BLOQUEADO', validation.nameScore, validation.cityScore, validation.reason);
      return { ok: false, message: 'Não foi possível validar o acesso. Confira os dados iniciais e tente novamente.' };
    }

    var normalized = normalizePayloadForSheet_(dados);
    normalized.timestamp = new Date();
    normalized.versaoFormulario = CONFIG.FORM_VERSION;
    // Marca reenvios para o dashboard poder manter só a resposta mais recente por código.
    normalized.statusResposta = jaExisteResposta_(codigo) ? 'REENVIO' : 'ENVIADO';
    normalized.validacaoAcesso = 'VALIDADO';
    normalized.similaridadeNome = validation.nameScore;
    normalized.similaridadeCidade = validation.cityScore;

    appendRespostaToSheet_(normalized);
    registrarAlteracoesCadastro_(cadastro, normalized);
    limparRascunhoServidor_(codigo);

    return {
      ok: true,
      message: 'Questionário enviado com sucesso. Obrigado por atualizar as informações do Alta CRIA 2026.'
    };
  } catch (err) {
    return { ok: false, message: 'Ocorreu um erro ao enviar. Tente novamente ou fale com a equipe Alta CRIA. Detalhe: ' + String(err) };
  }
}

function publicValidationResult_(ok) {
  return {
    ok: !!ok,
    message: ok
      ? 'Acesso validado. Agora preencha o questionário Alta CRIA 2026.'
      : 'Não foi possível validar as informações. Confira o código Alta CRIA, o nome do proprietário/grupo e a cidade antes de prosseguir.'
  };
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getCadastroSheet_() {
  var ss = getSpreadsheet_();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === CONFIG.CADASTRO_GID) return sheets[i];
  }
  // Fallback: localiza a aba pelo nome, caso o GID esteja desatualizado.
  var porNome = ss.getSheetByName(CONFIG.CADASTRO_SHEET_NAME);
  if (porNome) return porNome;
  var disp = sheets.map(function (s) { return s.getName() + ' (GID ' + s.getSheetId() + ')'; }).join(', ');
  throw new Error('Aba de cadastro não encontrada por GID (' + CONFIG.CADASTRO_GID + ') nem por nome ("' + CONFIG.CADASTRO_SHEET_NAME + '"). Abas disponíveis: ' + disp);
}

function getCadastroByCodigo_(codigo) {
  var sheet = getCadastroSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Aba de cadastro sem dados.');

  var headers = data[0];
  var headerMap = buildHeaderMap_(headers);
  var codeIndex = findFirstHeaderIndex_(headerMap, [
    'codigo', 'codigofazenda', 'codigoaltacria', 'codalta', 'id', 'cod', 'código', 'cód'
  ]);

  if (codeIndex === -1) {
    throw new Error('Não encontrei a coluna de código na base. Use um cabeçalho como Código, Codigo, Código Alta CRIA ou codigo_fazenda.');
  }

  var target = normalizeCode_(codigo);
  for (var r = 1; r < data.length; r++) {
    var rowCode = normalizeCode_(data[r][codeIndex]);
    if (rowCode && rowCode === target) {
      return {
        rowNumber: r + 1,
        headers: headers,
        row: data[r],
        headerMap: headerMap,
        codeIndex: codeIndex
      };
    }
  }
  return null;
}

function validateAgainstCadastro_(cadastro, nomeInformado, cidadeInformada) {
  var headerMap = cadastro.headerMap;
  var row = cadastro.row;

  var cityIndex = findFirstHeaderIndex_(headerMap, [
    'cidade', 'municipio', 'município', 'cidadefazenda'
  ]);

  var ownerIndexes = findAllHeaderIndexes_(headerMap, [
    'proprietario', 'proprietaria', 'nomeproprietario', 'nomedoproprietario',
    'grupo', 'grupoproprietario', 'nomegrupo', 'responsavel', 'responsavelfazenda',
    'nomeresponsavel', 'nomedoresponsavel', 'produtor', 'cliente', 'contato',
    'fazenda', 'nomefazenda', 'faturamento', 'nomefaturamento'
  ]);

  if (cityIndex === -1) {
    return { ok: false, nameScore: 0, cityScore: 0, reason: 'Coluna de cidade não encontrada na base.' };
  }
  if (!ownerIndexes.length) {
    return { ok: false, nameScore: 0, cityScore: 0, reason: 'Coluna de proprietário/grupo/responsável não encontrada na base.' };
  }

  var cidadeBase = String(row[cityIndex] || '');
  var cityScore = similarityScore_(cidadeInformada, cidadeBase, false);

  var bestNameScore = 0;
  var bestNameField = '';
  for (var i = 0; i < ownerIndexes.length; i++) {
    var idx = ownerIndexes[i];
    var value = String(row[idx] || '');
    var score = similarityScore_(nomeInformado, value, true);
    if (score > bestNameScore) {
      bestNameScore = score;
      bestNameField = cadastro.headers[idx];
    }
  }

  var ok = cityScore >= CONFIG.MIN_CITY_SCORE && bestNameScore >= CONFIG.MIN_NAME_SCORE;
  var reason = ok
    ? 'Validado por similaridade. Campo de nome usado: ' + bestNameField
    : 'Similaridade insuficiente. Nome=' + bestNameScore + '; Cidade=' + cityScore;

  return { ok: ok, nameScore: bestNameScore, cityScore: cityScore, reason: reason };
}

function buildHeaderMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = normalizeHeader_(headers[i]);
    if (key && map[key] === undefined) map[key] = i;
  }
  return map;
}

function findFirstHeaderIndex_(headerMap, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var key = normalizeHeader_(candidates[i]);
    if (headerMap[key] !== undefined) return headerMap[key];
  }
  return -1;
}

function findAllHeaderIndexes_(headerMap, candidates) {
  var indexes = [];
  var seen = {};
  for (var i = 0; i < candidates.length; i++) {
    var key = normalizeHeader_(candidates[i]);
    if (headerMap[key] !== undefined && !seen[headerMap[key]]) {
      indexes.push(headerMap[key]);
      seen[headerMap[key]] = true;
    }
  }
  return indexes;
}

function normalizeHeader_(value) {
  return removeAccents_(String(value || ''))
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function normalizeCode_(value) {
  var txt = String(value || '').trim().toUpperCase().replace(/\s+/g, '');
  // Aceita 2, 02 e 002 como o mesmo código quando for numérico.
  // Aceita também 2.0 / 2,0, caso a planilha devolva o número já formatado.
  if (/^0*\d+([.,]0+)?$/.test(txt)) {
    txt = String(parseInt(txt, 10));
  }
  return txt;
}

function normalizeText_(value, isName) {
  var txt = removeAccents_(String(value || ''))
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!isName) return txt;

  var stop = {
    'de': true, 'da': true, 'do': true, 'das': true, 'dos': true, 'e': true,
    'ltda': true, 'me': true, 'eireli': true, 'sa': true, 's/a': true,
    'agropecuaria': true, 'agropecuária': true, 'fazenda': true, 'sitio': true,
    'sítio': true, 'grupo': true, 'propriedade': true, 'produtor': true
  };
  var parts = txt.split(' ').filter(function (p) { return p && !stop[p]; });
  return parts.join(' ');
}

function removeAccents_(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function similarityScore_(a, b, isName) {
  var x = normalizeText_(a, isName);
  var y = normalizeText_(b, isName);
  if (!x || !y) return 0;
  if (x === y) return 1;

  var tokenScore = tokenSimilarity_(x, y);
  var levScore = levenshteinSimilarity_(x, y);
  var containmentScore = containmentSimilarity_(x, y);
  var strongTokenScore = isName ? strongTokenMatchScore_(x, y) : 0;

  var score = Math.max(tokenScore, levScore, containmentScore, strongTokenScore);
  return Math.round(score * 100) / 100;
}

function tokenSimilarity_(x, y) {
  var ax = x.split(' ').filter(Boolean);
  var ay = y.split(' ').filter(Boolean);
  if (!ax.length || !ay.length) return 0;

  var setY = {};
  ay.forEach(function (t) { setY[t] = true; });

  var matches = 0;
  ax.forEach(function (t) {
    if (setY[t]) matches++;
  });

  return matches / Math.max(ax.length, ay.length);
}

function strongTokenMatchScore_(x, y) {
  var ax = x.split(' ').filter(function (t) { return t.length >= 4; });
  var ay = y.split(' ').filter(function (t) { return t.length >= 4; });
  if (!ax.length || !ay.length) return 0;

  var setY = {};
  ay.forEach(function (t) { setY[t] = true; });

  var matches = 0;
  ax.forEach(function (t) {
    if (setY[t]) matches++;
  });

  // Caso de uso: "Agrindus" deve validar contra "Santa Rita-Agrindus" quando código e cidade conferem.
  if (matches >= 1 && ax.length === 1) return 0.92;
  if (matches / ax.length >= 0.60) return 0.85;
  return 0;
}

function containmentSimilarity_(x, y) {
  if (x.indexOf(y) !== -1 || y.indexOf(x) !== -1) {
    return Math.min(x.length, y.length) / Math.max(x.length, y.length);
  }
  return 0;
}

function levenshteinSimilarity_(a, b) {
  var distance = levenshteinDistance_(a, b);
  var maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - (distance / maxLen);
}

function levenshteinDistance_(a, b) {
  var matrix = [];
  for (var i = 0; i <= b.length; i++) matrix[i] = [i];
  for (var j = 0; j <= a.length; j++) matrix[0][j] = j;

  for (var i2 = 1; i2 <= b.length; i2++) {
    for (var j2 = 1; j2 <= a.length; j2++) {
      if (b.charAt(i2 - 1) === a.charAt(j2 - 1)) {
        matrix[i2][j2] = matrix[i2 - 1][j2 - 1];
      } else {
        matrix[i2][j2] = Math.min(
          matrix[i2 - 1][j2 - 1] + 1,
          matrix[i2][j2 - 1] + 1,
          matrix[i2 - 1][j2] + 1
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

function normalizePayloadForSheet_(obj) {
  var out = {};
  Object.keys(obj || {}).forEach(function (k) {
    if (k === '_access') return;
    var v = obj[k];
    if (Array.isArray(v)) {
      out[k] = v.join(' | ');
    } else if (v && typeof v === 'object') {
      out[k] = JSON.stringify(v);
    } else {
      out[k] = v === undefined || v === null ? '' : v;
    }
  });
  return out;
}

function appendObjectToSheet_(sheetName, obj) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  var headers = [];

  if (sheet.getLastRow() === 0) {
    headers = Object.keys(obj);
    sheet.appendRow(headers);
  } else {
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  var objKeys = Object.keys(obj);
  var newHeaders = [];
  objKeys.forEach(function (k) {
    if (headers.indexOf(k) === -1) newHeaders.push(k);
  });

  if (newHeaders.length) {
    sheet.getRange(1, headers.length + 1, 1, newHeaders.length).setValues([newHeaders]);
    headers = headers.concat(newHeaders);
  }

  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.appendRow(row);
}

/**
 * Grava uma resposta usando a ordem fixa de colunas (RESPOSTAS_COLUMNS).
 * Se o objeto trouxer alguma chave inesperada fora da lista, ela é anexada
 * ao final como rede de segurança, sem bagunçar as colunas conhecidas.
 */
function appendRespostaToSheet_(obj) {
  var sheet = ensureRespostasHeader_();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var extras = Object.keys(obj).filter(function (k) { return headers.indexOf(k) === -1; });
  if (extras.length) {
    sheet.getRange(1, headers.length + 1, 1, extras.length).setValues([extras]);
    headers = headers.concat(extras);
  }

  var row = headers.map(function (h) { return obj[h] !== undefined && obj[h] !== null ? obj[h] : ''; });
  sheet.appendRow(row);
}

/**
 * Retorna true se já existe alguma resposta gravada para este código.
 */
function jaExisteResposta_(codigo) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(CONFIG.RESPOSTAS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return false;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var col = headers.indexOf('codigoAcesso');
    if (col === -1) return false;
    var values = sheet.getRange(2, col + 1, sheet.getLastRow() - 1, 1).getValues();
    var alvo = normalizeCode_(codigo);
    for (var i = 0; i < values.length; i++) {
      if (normalizeCode_(values[i][0]) === alvo) return true;
    }
    return false;
  } catch (err) {
    return false;
  }
}

/**
 * Salva um rascunho no servidor (aba Rascunhos_26), um registro por código.
 * Revalida o acesso antes de gravar. Chamado pelo formulário a cada troca de
 * etapa. Falhas aqui não impedem o preenchimento: o rascunho local continua
 * valendo como reserva.
 */
function saveServerDraft(payload) {
  try {
    payload = payload || {};
    var codigo = String(payload.codigoAcesso || '').trim();
    var nome = String(payload.nomeValidacao || '').trim();
    var cidade = String(payload.cidadeValidacao || '').trim();
    if (!codigo || !nome || !cidade) return { ok: false };

    var cadastro = getCadastroByCodigo_(codigo);
    if (!cadastro) return { ok: false };
    var validation = validateAgainstCadastro_(cadastro, nome, cidade);
    if (!validation.ok) return { ok: false };

    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(CONFIG.RASCUNHOS_SHEET) || ss.insertSheet(CONFIG.RASCUNHOS_SHEET);
    if (sheet.getLastRow() === 0) sheet.appendRow(['codigo', 'atualizado_em', 'dados_json']);

    var alvo = normalizeCode_(codigo);
    var json = JSON.stringify(payload);
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var codes = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < codes.length; i++) {
        if (normalizeCode_(codes[i][0]) === alvo) {
          sheet.getRange(i + 2, 2, 1, 2).setValues([[new Date(), json]]);
          return { ok: true };
        }
      }
    }
    sheet.appendRow([codigo, new Date(), json]);
    return { ok: true };
  } catch (err) {
    return { ok: false };
  }
}

/**
 * Recupera o rascunho salvo no servidor para um código já validado.
 * Retorna o objeto de respostas ou null. Usado na validação de acesso para
 * o participante continuar de onde parou, mesmo em outro aparelho.
 */
function getServerDraft_(codigo) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(CONFIG.RASCUNHOS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getDataRange().getValues();
    var alvo = normalizeCode_(codigo);
    for (var i = 1; i < data.length; i++) {
      if (normalizeCode_(data[i][0]) === alvo && data[i][2]) {
        return JSON.parse(data[i][2]);
      }
    }
    return null;
  } catch (err) {
    return null;
  }
}

/**
 * Remove o rascunho do servidor após o envio definitivo.
 */
function limparRascunhoServidor_(codigo) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(CONFIG.RASCUNHOS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return;
    var codes = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    var alvo = normalizeCode_(codigo);
    for (var i = codes.length - 1; i >= 0; i--) {
      if (normalizeCode_(codes[i][0]) === alvo) sheet.deleteRow(i + 2);
    }
  } catch (err) {
    // silencioso: não bloquear o envio por causa da limpeza do rascunho
  }
}

function logAccessAttempt_(codigo, nome, cidade, resultado, nameScore, cityScore, detalhe) {
  var obj = {
    timestamp: new Date(),
    codigo_digitado: codigo || '',
    nome_digitado: nome || '',
    cidade_digitada: cidade || '',
    resultado: resultado || '',
    similaridade_nome: nameScore || 0,
    similaridade_cidade: cityScore || 0,
    detalhe: detalhe || ''
  };
  appendObjectToSheet_(CONFIG.TENTATIVAS_SHEET, obj);
}

function tooManyAttempts_(codigo) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(CONFIG.TENTATIVAS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return false;
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function (h) { return normalizeHeader_(h); });
    var codeIdx = headers.indexOf('codigodigitado');
    var timeIdx = headers.indexOf('timestamp');
    if (codeIdx === -1 || timeIdx === -1) return false;
    var target = normalizeCode_(codigo);
    var today = new Date();
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      var rowCode = normalizeCode_(data[i][codeIdx]);
      var t = data[i][timeIdx];
      if (rowCode !== target || !(t instanceof Date)) continue;
      if (t.getFullYear() === today.getFullYear() && t.getMonth() === today.getMonth() && t.getDate() === today.getDate()) count++;
      if (count >= CONFIG.MAX_ATTEMPTS_PER_CODE_PER_DAY) return true;
    }
    return false;
  } catch (err) {
    return false;
  }
}

function registrarAlteracoesCadastro_(cadastro, dados) {
  var mappings = [
    { form: 'q004', label: 'Nome da fazenda', candidates: ['fazenda', 'nomefazenda'] },
    { form: 'q013', label: 'Cidade', candidates: ['cidade', 'municipio'] },
    { form: 'q014', label: 'Estado', candidates: ['estado', 'uf'] },
    { form: 'q005', label: 'Proprietário/grupo', candidates: ['proprietario', 'responsavel', 'cliente', 'produtor', 'grupo'] },
    { form: 'q011', label: 'Nome de faturamento', candidates: ['faturamento', 'nomefaturamento'] },
    { form: 'q015', label: 'E-mail principal', candidates: ['emailfazenda', 'email'] },
    { form: 'q026', label: 'Litros de leite/dia', candidates: ['leitedia', 'leite'] },
    { form: 'q025', label: 'Vacas ordenhadas/dia', candidates: ['nvacas', 'vacas', 'vacaslactacao', 'vacasordenhadas'] },
    { form: 'q008', label: 'Modo de envio de dados', candidates: ['envio', 'formaenvio', 'mododeenvio'] }
  ];

  mappings.forEach(function (m) {
    var idx = findFirstHeaderIndex_(cadastro.headerMap, m.candidates);
    if (idx === -1) return;

    var valorAntigo = String(cadastro.row[idx] || '').trim();
    var valorNovo = String(dados[m.form] || '').trim();
    if (!valorNovo) return;

    var antigoNorm = normalizeText_(valorAntigo, false);
    var novoNorm = normalizeText_(valorNovo, false);
    if (antigoNorm !== novoNorm) {
      appendObjectToSheet_(CONFIG.ALTERACOES_SHEET, {
        timestamp: new Date(),
        codigo_fazenda: dados.codigoAcesso || '',
        fazenda_informada: dados.q004 || '',
        campo: m.label,
        valor_anterior: valorAntigo,
        valor_novo: valorNovo,
        status_revisao: 'PENDENTE'
      });
    }
  });
}

/* ===================================================================
   ACOMPANHAMENTO DA COLETA
   Monta a aba Acompanhamento_26 como uma fila de trabalho: quem tentou
   e não conseguiu entrar aparece primeiro, depois quem parou no meio,
   depois quem nem começou, e por último quem já respondeu.
   O total sai sempre da Base (nunca de um número fixo), então fazendas
   novas entram na conta sozinhas.
   Rode pelo menu "Alta CRIA > Atualizar acompanhamento" ou direto no editor.
   =================================================================== */

var ACOMP_SHEET = 'Acompanhamento_26';
var SIT = {
  BARRADO:  'Tentou e não entrou',
  RASCUNHO: 'Em preenchimento',
  NAO:      'Não iniciou',
  OK:       'Respondido'
};
var SIT_ORDEM = {};
SIT_ORDEM[SIT.BARRADO] = 1;
SIT_ORDEM[SIT.RASCUNHO] = 2;
SIT_ORDEM[SIT.NAO] = 3;
SIT_ORDEM[SIT.OK] = 4;

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Alta CRIA')
      .addItem('Atualizar acompanhamento', 'atualizarAcompanhamento')
      .addToUi();
  } catch (err) {
    // Projeto não vinculado à planilha: o menu não aparece, mas a função
    // continua funcionando pelo editor. Não é erro.
  }
}

function atualizarAcompanhamento() {
  var ss = getSpreadsheet_();

  // ---------- 1) Base: quem deve responder ----------
  var cad = getCadastroSheet_();
  var cadData = cad.getDataRange().getValues();
  if (cadData.length < 2) throw new Error('Aba de cadastro sem dados.');
  var hm = buildHeaderMap_(cadData[0]);
  var cCod    = findFirstHeaderIndex_(hm, ['codigo','codigofazenda','codigoaltacria','cod']);
  var cFaz    = findFirstHeaderIndex_(hm, ['fazenda','nomedafazenda']);
  var cCid    = findFirstHeaderIndex_(hm, ['cidade','municipio']);
  var cUF     = findFirstHeaderIndex_(hm, ['estado','uf']);
  var cResp   = findFirstHeaderIndex_(hm, ['responsavel']);
  var cWhats  = findFirstHeaderIndex_(hm, ['whats1','whats']);
  var cNum    = findFirstHeaderIndex_(hm, ['numerowhats1','numerowhats','telefone']);
  if (cCod === -1) throw new Error('Não encontrei a coluna de código na Base.');

  // ---------- 2) Respostas: quem terminou ----------
  var respMap = {};
  var rSheet = ss.getSheetByName(CONFIG.RESPOSTAS_SHEET);
  if (rSheet && rSheet.getLastRow() > 1) {
    var rH = rSheet.getRange(1, 1, 1, rSheet.getLastColumn()).getValues()[0];
    var iCod = rH.indexOf('codigoAcesso'), iTs = rH.indexOf('timestamp');
    if (iCod > -1) {
      var rVals = rSheet.getRange(2, 1, rSheet.getLastRow() - 1, rSheet.getLastColumn()).getValues();
      for (var i = 0; i < rVals.length; i++) {
        var k = normalizeCode_(rVals[i][iCod]);
        if (!k) continue;
        var ts = iTs > -1 ? rVals[i][iTs] : '';
        // guarda sempre a resposta mais recente do código
        if (!respMap[k] || (ts && respMap[k] && ts > respMap[k])) respMap[k] = ts;
      }
    }
  }

  // ---------- 3) Rascunhos: quem parou no meio ----------
  var rascMap = {};
  var dSheet = ss.getSheetByName(CONFIG.RASCUNHOS_SHEET);
  if (dSheet && dSheet.getLastRow() > 1) {
    var dVals = dSheet.getRange(2, 1, dSheet.getLastRow() - 1, 2).getValues();
    for (var j = 0; j < dVals.length; j++) {
      var dk = normalizeCode_(dVals[j][0]);
      if (dk) rascMap[dk] = dVals[j][1];
    }
  }

  // ---------- 4) Tentativas: quem bateu na porta e não entrou ----------
  var falhaMap = {}, orfas = 0;
  var tSheet = ss.getSheetByName(CONFIG.TENTATIVAS_SHEET);
  if (tSheet && tSheet.getLastRow() > 1) {
    var tH = tSheet.getRange(1, 1, 1, tSheet.getLastColumn()).getValues()[0];
    var jCod = tH.indexOf('codigo_digitado'), jRes = tH.indexOf('resultado');
    var jTs = tH.indexOf('timestamp'), jDet = tH.indexOf('detalhe');
    if (jCod > -1 && jRes > -1) {
      var tVals = tSheet.getRange(2, 1, tSheet.getLastRow() - 1, tSheet.getLastColumn()).getValues();
      var codigosBase = {};
      for (var b = 1; b < cadData.length; b++) {
        var bk = normalizeCode_(cadData[b][cCod]);
        if (bk) codigosBase[bk] = true;
      }
      for (var t = 0; t < tVals.length; t++) {
        if (String(tVals[t][jRes]).indexOf('SUCESSO') > -1) continue;
        var tk = normalizeCode_(tVals[t][jCod]);
        if (!tk) continue;
        if (!codigosBase[tk]) { orfas++; continue; } // digitou um código que não existe na Base
        if (!falhaMap[tk]) falhaMap[tk] = { n: 0, ts: '', det: '' };
        falhaMap[tk].n++;
        var tts = jTs > -1 ? tVals[t][jTs] : '';
        if (!falhaMap[tk].ts || (tts && tts > falhaMap[tk].ts)) {
          falhaMap[tk].ts = tts;
          falhaMap[tk].det = jDet > -1 ? tVals[t][jDet] : '';
        }
      }
    }
  }

  // ---------- 5) Montar a fila ----------
  var linhas = [], cont = {};
  cont[SIT.OK] = 0; cont[SIT.RASCUNHO] = 0; cont[SIT.BARRADO] = 0; cont[SIT.NAO] = 0;
  for (var r = 1; r < cadData.length; r++) {
    var row = cadData[r];
    var cod = normalizeCode_(row[cCod]);
    if (!cod) continue;
    var sit, quando = '', tent = '', det = '';
    if (respMap[cod] !== undefined) { sit = SIT.OK; quando = respMap[cod]; }
    else if (rascMap[cod] !== undefined) { sit = SIT.RASCUNHO; quando = rascMap[cod]; }
    else if (falhaMap[cod]) { sit = SIT.BARRADO; quando = falhaMap[cod].ts; tent = falhaMap[cod].n; det = falhaMap[cod].det; }
    else { sit = SIT.NAO; }
    cont[sit]++;
    linhas.push([
      sit, cod,
      cFaz > -1 ? row[cFaz] : '', cCid > -1 ? row[cCid] : '', cUF > -1 ? row[cUF] : '',
      cResp > -1 ? row[cResp] : '',
      cWhats > -1 ? row[cWhats] : '', cNum > -1 ? row[cNum] : '',
      quando, tent, det,
      SIT_ORDEM[sit]
    ]);
  }
  // ação primeiro; dentro da mesma situação, agrupa por responsável
  linhas.sort(function (a, b) {
    if (a[11] !== b[11]) return a[11] - b[11];
    var ra = String(a[5] || ''), rb = String(b[5] || '');
    if (ra !== rb) return ra < rb ? -1 : 1;
    return String(a[1]).localeCompare(String(b[1]), undefined, { numeric: true });
  });
  for (var L = 0; L < linhas.length; L++) linhas[L].pop(); // tira a coluna de ordenação

  // ---------- 6) Escrever a aba ----------
  var sh = ss.getSheetByName(ACOMP_SHEET) || ss.insertSheet(ACOMP_SHEET);
  sh.clear();
  var total = linhas.length;
  var pct = function (n) { return total ? Math.round(n * 1000 / total) / 10 + '%' : '0%'; };

  var topo = [
    ['ACOMPANHAMENTO DA COLETA — ALTA CRIA 2026', '', '', ''],
    ['Atualizado em', new Date(), '', ''],
    ['', '', '', ''],
    ['Fazendas na Base', total, '', 'Total sai da Base: fazenda nova entra sozinha'],
    ['Respondido', cont[SIT.OK], pct(cont[SIT.OK]), 'Já enviaram o questionário'],
    ['Em preenchimento', cont[SIT.RASCUNHO], pct(cont[SIT.RASCUNHO]), 'Começaram e pararam — cutucar'],
    ['Tentou e não entrou', cont[SIT.BARRADO], pct(cont[SIT.BARRADO]), 'Não passaram na validação — ligar'],
    ['Não iniciou', cont[SIT.NAO], pct(cont[SIT.NAO]), 'Nem abriram'],
    ['', '', '', ''],
    ['Tentativas com código fora da Base', orfas, '', 'Alguém digitou um código que não existe'],
    ['', '', '', '']
  ];
  sh.getRange(1, 1, topo.length, 4).setValues(topo);
  sh.getRange(1, 1, 1, 4).setFontSize(13).setFontWeight('bold').setFontColor('#0E2E4D');
  sh.getRange(4, 1, 7, 1).setFontWeight('bold');
  sh.getRange(2, 2).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(4, 4, 7, 1).setFontColor('#667085').setFontStyle('italic');

  var head = ['Situação', 'Código', 'Fazenda', 'Cidade', 'UF', 'Responsável', 'Whats 1', 'Número', 'Última atividade', 'Tentativas barradas', 'Detalhe da última falha'];
  var hr = topo.length + 1;
  sh.getRange(hr, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#0E2E4D').setFontColor('#FFFFFF');
  if (linhas.length) {
    sh.getRange(hr + 1, 1, linhas.length, head.length).setValues(linhas);
    sh.getRange(hr + 1, 9, linhas.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    var cores = { }; cores[SIT.BARRADO] = '#FCE8E6'; cores[SIT.RASCUNHO] = '#FEF7E0'; cores[SIT.OK] = '#E6F4EA'; cores[SIT.NAO] = '#FFFFFF';
    var faixaIni = 0;
    for (var k = 1; k <= linhas.length; k++) {
      if (k === linhas.length || linhas[k][0] !== linhas[faixaIni][0]) {
        sh.getRange(hr + 1 + faixaIni, 1, k - faixaIni, head.length).setBackground(cores[linhas[faixaIni][0]]);
        faixaIni = k;
      }
    }
  }
  try {
    sh.setFrozenRows(hr);
    sh.autoResizeColumns(1, head.length);
    sh.getRange(hr, 1, Math.max(linhas.length + 1, 1), head.length).createFilter();
  } catch (err) { /* formatação é cosmética: nunca bloquear o painel */ }

  return 'Acompanhamento atualizado: ' + cont[SIT.OK] + ' de ' + total + ' responderam.';
}

/* ===================================================================
   DADOS PÚBLICOS DO DASHBOARD
   Lê a aba de respostas, agrega TUDO aqui dentro (onde os nomes ficam)
   e devolve um JSON só com números — nenhuma linha crua, nenhum nome,
   nenhum texto digitado pela fazenda.

   Por que agregar aqui e não no navegador:
   - nome de fazenda nunca sai da planilha;
   - quando entrar o banco de bezerras, a conta já vem pronta e o
     dashboard continua leve no celular.

   Como usar: rode gerarDadosPublicos() e copie o JSON do log para o
   arquivo dados.json do repositório do dashboard.
   =================================================================== */

var SCHEMA_PUBLICO = [{"idx":"1","t":"Informações gerais e caracterização da fazenda","q":[{"id":"q008","n":8,"tt":"Modo de envio de dados para o programa Alta CRIA","t":"Escolha única"},{"id":"q014","n":14,"tt":"Estado","t":"Resposta curta / UF"},{"id":"q025","n":25,"tt":"Número médio de vacas ordenhadas por dia no ano","t":"Número"},{"id":"q026","n":26,"tt":"Litros de leite produzidos por dia, na média do ano","t":"Número"},{"id":"q027","n":27,"tt":"Produção média calculada por vaca/dia","t":"Campo calculado"},{"id":"q028","n":28,"tt":"Principal sistema de criação das vacas em lactação","t":"Escolha única","o":"q028_outro"},{"id":"q029","n":29,"tt":"Raça predominante no rebanho","t":"Escolha única","o":"q029_outro"},{"id":"q030","n":30,"tt":"Principal software de lançamento de dados zootécnicos das bezerras e novilhas da fazenda","t":"Escolha única","o":"q030_outro"},{"id":"q031","n":31,"tt":"Número total de funcionários na área de leite","t":"Número"},{"id":"q032","n":32,"tt":"Número total de funcionários na criação das bezerras, fase de aleitamento","t":"Número"},{"id":"q033","n":33,"tt":"Número total de funcionários na criação das novilhas, do pós-aleitamento ao primeiro parto","t":"Número"},{"id":"q034","n":34,"tt":"Idade do(a) principal colaborador(a) responsável pelo bezerreiro","t":"Número"},{"id":"q035","n":35,"tt":"Gênero do(a) principal colaborador(a) responsável pelo bezerreiro","t":"Escolha única"},{"id":"q036","n":36,"tt":"Nível de instrução do(a) principal colaborador(a) responsável pelo bezerreiro","t":"Escolha única"},{"id":"q037","n":37,"tt":"Último treinamento ou reciclagem presencial e/ou on-line dos colaboradores ligados diretamente à criação das bezerras","t":"Escolha única"}]},{"idx":"2","t":"Pré-parto","q":[{"id":"q038","n":38,"tt":"Quantos dias antes da data prevista do parto as vacas são secas, com lactação interrompida?","t":"Escolha única"},{"id":"q039","n":39,"tt":"Escore de Condição Corporal (ECC) médio das novilhas no pré-parto","t":"Escolha única"},{"id":"q040","n":40,"tt":"Escore de Condição Corporal (ECC) médio das vacas secas no pré-parto","t":"Escolha única"},{"id":"q041","n":41,"tt":"Período de resfriamento com ventilação + aspersão das novilhas gestantes","t":"Escolha única"},{"id":"q042","n":42,"tt":"Período de resfriamento com ventilação + aspersão das vacas secas gestantes e/ou do pré-parto","t":"Escolha única"},{"id":"q043","n":43,"tt":"A fazenda possui protocolo de vacinação para vacas secas gestantes e/ou novilhas gestantes?","t":"Escolha única"},{"id":"q044","n":44,"tt":"Quais vacinas são utilizadas nas vacas secas gestantes e/ou novilhas gestantes?","t":"Múltipla escolha","o":"q044_outro"},{"id":"q045","n":45,"tt":"Novilhas e vacas gestantes são criadas juntas no pré-parto?","t":"Escolha única"},{"id":"q046","n":46,"tt":"Você utiliza sais aniônicos na dieta das novilhas no pré-parto?","t":"Escolha única"},{"id":"q047","n":47,"tt":"Você utiliza sais aniônicos na dieta das vacas no pré-parto?","t":"Escolha única"}]},{"idx":"3","t":"Maternidade e manejo inicial da recém-nascida","q":[{"id":"q048","n":48,"tt":"Local da maioria dos partos da fazenda","t":"Escolha única"},{"id":"q049","n":49,"tt":"Caso utilize baias individuais para parto, informe a largura média da baia, em metros","t":"Número"},{"id":"q050","n":50,"tt":"Caso utilize baias individuais para parto, informe o comprimento médio da baia, em metros","t":"Número"},{"id":"q051","n":51,"tt":"Área média calculada da baia individual de parto, em m²","t":"Campo calculado"},{"id":"q052","n":52,"tt":"Se o principal local de parto for em baias individuais, qual é o principal tipo de cama utilizado?","t":"Escolha única","o":"q052_outro"},{"id":"q053","n":53,"tt":"Número de observações na maternidade entre 6 e 12 horas (manhã)","t":"Escolha única"},{"id":"q054","n":54,"tt":"Número de observações na maternidade entre 12 e 18 horas (tarde)","t":"Escolha única"},{"id":"q055","n":55,"tt":"Número de observações na maternidade entre 18 e 6 horas (noite e madrugada)","t":"Escolha única"},{"id":"q056","n":56,"tt":"Após o rompimento da bolsa, quanto tempo, em média, espera-se antes de intervir no parto, em novilhas e vacas?","t":"Escolha única"},{"id":"q057","n":57,"tt":"Tipo de intervenção mais comum nos partos","t":"Escolha única","o":"q057_outro"},{"id":"q058","n":58,"tt":"Tempo médio que a recém-nascida permanece na maternidade, local do parto, com a mãe após o nascimento","t":"Escolha única"},{"id":"q059","n":59,"tt":"Tipo de secagem da recém-nascida após o nascimento","t":"Escolha única"},{"id":"q060","n":60,"tt":"Campânula/luz para aquecer a recém-nascida, quando necessário","t":"Escolha única"},{"id":"q061","n":61,"tt":"Jaqueta térmica é utilizada para aquecer a recém-nascida, quando necessário?","t":"Escolha única"},{"id":"q062","n":62,"tt":"Utiliza meloxicam como rotina para alívio da dor em recém-nascidas de partos auxiliados ou complicados?","t":"Escolha única"},{"id":"q063","n":63,"tt":"Carrinho exclusivo para tirar a recém-nascida da maternidade, local do parto","t":"Escolha única"},{"id":"q064","n":64,"tt":"Como as bezerras são pesadas ao nascimento?","t":"Escolha única"},{"id":"q065","n":65,"tt":"Principal produto utilizado para a cura do umbigo","t":"Escolha única","o":"q065_outro"},{"id":"q066","n":66,"tt":"Número de vezes por dia que o umbigo é curado","t":"Escolha única"},{"id":"q067","n":67,"tt":"Por quantos dias o umbigo é curado?","t":"Escolha única"},{"id":"q068","n":68,"tt":"Inspeção e remoção das tetas extranumerárias","t":"Escolha única"},{"id":"q069","n":69,"tt":"Tipo de mochação","t":"Escolha única","o":"q069_outro"},{"id":"q070","n":70,"tt":"Protocolo de mochação","t":"Escolha única"},{"id":"q071","n":71,"tt":"Idade média, em dias, da mochação","t":"Número"}]},{"idx":"4","t":"Colostragem","q":[{"id":"q072","n":72,"tt":"Tempo médio para ordenhar o colostro após o parto","t":"Escolha única"},{"id":"q073","n":73,"tt":"Local de ordenha do colostro","t":"Escolha única"},{"id":"q074","n":74,"tt":"É utilizada ocitocina em vacas/novilhas recém-paridas para estimular a descida do colostro?","t":"Escolha única"},{"id":"q075","n":75,"tt":"Tempo médio para fornecimento da primeira alimentação do colostro em partos que ocorrem durante o dia, entre 6h e 18h","t":"Escolha única"},{"id":"q076","n":76,"tt":"Tempo médio para fornecimento da primeira alimentação do colostro em partos que ocorrem durante a noite e madrugada, entre 18h e 6h","t":"Escolha única"},{"id":"q077","n":77,"tt":"Principal tipo de colostro utilizado na primeira alimentação das bezerras","t":"Escolha única"},{"id":"q078","n":78,"tt":"Quantidade de colostro fornecida na primeira alimentação","t":"Escolha única"},{"id":"q079","n":79,"tt":"Tempo médio para fornecimento da segunda alimentação do colostro após o fornecimento da primeira alimentação","t":"Escolha única"},{"id":"q080","n":80,"tt":"Principal tipo de colostro utilizado na segunda alimentação das bezerras","t":"Escolha única"},{"id":"q081","n":81,"tt":"Quantidade de colostro fornecida na segunda alimentação","t":"Escolha única"},{"id":"q082","n":82,"tt":"De que forma o colostro em pó da Alta (SCCL®) é utilizado na sua fazenda?","t":"Múltipla escolha"},{"id":"q083","n":83,"tt":"Método de avaliação da qualidade imunológica do colostro","t":"Escolha única"},{"id":"q084","n":84,"tt":"Ponto de corte no Brix utilizado para considerar um colostro com qualidade","t":"Escolha única"},{"id":"q085","n":85,"tt":"Pasteurização do colostro","t":"Escolha única"},{"id":"q086","n":86,"tt":"Equipamento para pasteurização do colostro","t":"Escolha única","o":"q086_outro"},{"id":"q087","n":87,"tt":"Análise da contagem padrão em placas (CPP; antiga CBT) do colostro","t":"Escolha única"},{"id":"q088","n":88,"tt":"Tipo de análise da contagem padrão em placas (CPP) do colostro","t":"Escolha única"},{"id":"q089","n":89,"tt":"Número máximo de dias que o colostro fica armazenado de forma refrigerada, em geladeira a 2–8 °C","t":"Escolha única"},{"id":"q090","n":90,"tt":"Tipo de armazenamento do colostro no freezer","t":"Escolha única"},{"id":"q091","n":91,"tt":"Tipo de descongelamento do colostro","t":"Escolha única"},{"id":"q092","n":92,"tt":"Se possui banco de colostro congelado, qual a temperatura máxima, em °C, utilizada para descongelamento do colostro?","t":"Número"},{"id":"q093","n":93,"tt":"Tipo de fornecimento da primeira alimentação do colostro","t":"Escolha única"},{"id":"q094","n":94,"tt":"Tipo de fornecimento da segunda alimentação do colostro","t":"Escolha única"},{"id":"q095","n":95,"tt":"Método de avaliação da eficiência de colostragem","t":"Escolha única"},{"id":"q096","n":96,"tt":"Tempo após a colostragem da coleta de sangue para avaliação da eficiência da colostragem por proteína sérica total ou Brix","t":"Escolha única"}]},{"idx":"5","t":"Nutrição","q":[{"id":"q097","n":97,"tt":"Tempo de fornecimento de leite de transição após a colostragem","t":"Escolha única"},{"id":"q098","n":98,"tt":"Tipo de leite de transição fornecido","t":"Escolha única"},{"id":"q099","n":99,"tt":"Modo de fornecimento do leite e/ou sucedâneo na maior parte do aleitamento","t":"Escolha única"},{"id":"q100","n":100,"tt":"Dieta líquida predominante fornecida até os 30 dias de vida","t":"Escolha única"},{"id":"q101","n":101,"tt":"Se for oferecido leite não comercializável, descarte, até 30 dias de vida, ele é pasteurizado?","t":"Escolha única"},{"id":"q102","n":102,"tt":"Se o leite fornecido até os 30 dias for adensado com sucedâneo ou corretor, qual a % de sólidos totais final da mistura?","t":"Escolha única"},{"id":"q103","n":103,"tt":"Se o sucedâneo utilizado é fornecido e diluído em água até os 30 dias, qual a % de sólidos totais ou a proporção de diluição adotada?","t":"Escolha única"},{"id":"q104","n":104,"tt":"Se o leite é adensado ou a fazenda utiliza sucedâneo diluído em água até os 30 dias, qual sucedâneo, balanceador ou corretor é utilizado?","t":"Escolha única","o":"q104_outro"},{"id":"q105","n":105,"tt":"Volume médio de leite ou sucedâneo fornecido até os 30 dias de vida, em litros/dia","t":"Escolha única"},{"id":"q106","n":106,"tt":"Dieta líquida predominante fornecida após os 30 dias de vida, até o desaleitamento","t":"Escolha única"},{"id":"q107","n":107,"tt":"Se for oferecido leite não comercializável após 30 dias de vida, descarte, ele é pasteurizado?","t":"Escolha única"},{"id":"q108","n":108,"tt":"Se o leite fornecido após os 30 dias for adensado com sucedâneo ou corretor, qual a % de sólidos totais final da mistura?","t":"Escolha única"},{"id":"q109","n":109,"tt":"Se o sucedâneo utilizado é fornecido e diluído em água após os 30 dias, qual a % de sólidos totais ou a proporção de diluição adotada?","t":"Escolha única"},{"id":"q110","n":110,"tt":"Se o leite é adensado ou a fazenda utiliza sucedâneo diluído em água após os 30 dias, qual sucedâneo, balanceador ou corretor é utilizado?","t":"Escolha única","o":"q110_outro"},{"id":"q111","n":111,"tt":"Volume médio de leite ou sucedâneo fornecido dos 30 dias até o final do aleitamento, em litros/dia","t":"Escolha única"},{"id":"q112","n":112,"tt":"Frequência de fornecimento do leite ou sucedâneo","t":"Escolha única"},{"id":"q113","n":113,"tt":"Intervalo de tempo entre os fornecimentos do leite ou sucedâneo durante o dia","t":"Escolha única"},{"id":"q114","n":114,"tt":"Estratégia de aleitamento adotada","t":"Escolha única"},{"id":"q115","n":115,"tt":"Idade de início de fornecimento de água","t":"Escolha única"},{"id":"q116","n":116,"tt":"Modo de fornecimento de água na fase de aleitamento","t":"Escolha única"},{"id":"q117","n":117,"tt":"Idade de início do fornecimento de concentrado","t":"Escolha única"},{"id":"q118","n":118,"tt":"Procedência do concentrado da fase de aleitamento","t":"Escolha única"},{"id":"q119","n":119,"tt":"Tipo de concentrado da fase de aleitamento","t":"Escolha única"},{"id":"q120","n":120,"tt":"Percentual de proteína bruta (PB) do concentrado fornecido na fase de aleitamento","t":"Escolha única"},{"id":"q121","n":121,"tt":"O concentrado fornecido na fase de aleitamento contém coccidiostático?","t":"Escolha única"},{"id":"q122","n":122,"tt":"Idade de início do fornecimento de volumoso na fase de aleitamento","t":"Escolha única"},{"id":"q123","n":123,"tt":"Principal volumoso utilizado na fase de aleitamento","t":"Escolha única","o":"q123_outro"},{"id":"q124","n":124,"tt":"Modo de fornecimento do volumoso na fase de aleitamento","t":"Escolha única"},{"id":"q125","n":125,"tt":"Quantidade de volumoso oferecido na fase de aleitamento","t":"Escolha única"},{"id":"q126","n":126,"tt":"Tipo de desaleitamento","t":"Escolha única"},{"id":"q127","n":127,"tt":"Se o sistema de aleitamento for gradual, com quantos dias de vida se inicia a primeira redução de dieta líquida?","t":"Número"},{"id":"q128","n":128,"tt":"Quais critérios a fazenda considera para desaleitar as bezerras?","t":"Múltipla escolha"}]},{"idx":"6","t":"Biosseguridade, morbidade e mortalidade","q":[{"id":"q129","n":129,"tt":"Protocolo para controle da entrada de pessoas na maternidade (local onde ocorrem os partos)","t":"Escolha única"},{"id":"q130","n":130,"tt":"Protocolo para controle da entrada de pessoas no berçário, local onde as bezerras ficam nos primeiros dias/semanas de vida antes de irem para o bezerreiro","t":"Escolha única"},{"id":"q131","n":131,"tt":"Protocolo para controle da entrada de pessoas no bezerreiro","t":"Escolha única"},{"id":"q132","n":132,"tt":"Manejo e/ou produto para desinfecção do local onde ocorreu o parto","t":"Múltipla escolha","o":"q132_outro"},{"id":"q133","n":133,"tt":"Limpeza dos utensílios utilizados para alimentação das bezerras, como mamadeira, balde com bico e sonda","t":"Múltipla escolha","o":"q133_outro"},{"id":"q134","n":134,"tt":"Lavagem/desinfecção das instalações das bezerras, como gaiolas e pisos","t":"Múltipla escolha","o":"q134_outro"},{"id":"q135","n":135,"tt":"Qual é o tempo de intervalo, vazio sanitário, entre a saída de uma bezerra e a entrada de outra no mesmo local?","t":"Escolha única"},{"id":"q136","n":136,"tt":"Teste nas recém-nascidas para verificação da presença de animais persistentemente infectados (PI) para o vírus da diarreia viral bovina (BVD)","t":"Escolha única"},{"id":"q137","n":137,"tt":"Vacinas feitas em todas as bezerras até 12 meses de idade","t":"Múltipla escolha"},{"id":"q138","n":138,"tt":"Medicamentos e/ou produtos utilizados em todas as bezerras como protocolo preventivo durante a fase de aleitamento","t":"Múltipla escolha","o":"q138_outro"},{"id":"q139","n":139,"tt":"Doenças que mais acometem as bezerras na fase de aleitamento na fazenda","t":"Múltipla escolha"},{"id":"q140","n":140,"tt":"Doenças que mais acometem as novilhas na fase de pós-aleitamento na fazenda","t":"Múltipla escolha"},{"id":"q141","n":141,"tt":"Como a fazenda registra/lança os casos de diarreia das bezerras?","t":"Escolha única","o":"q141_outro"},{"id":"q142","n":142,"tt":"Quando a bezerra apresenta diarreia, o uso de antibiótico é feito apenas após a avaliação da temperatura?","t":"Escolha única"},{"id":"q143","n":143,"tt":"Durante episódios de diarreia, a dieta líquida da bezerra é geralmente","t":"Escolha única"},{"id":"q144","n":144,"tt":"Monitoramento das doenças respiratórias com uso de ultrassom torácico na cria e/ou recria","t":"Escolha única"},{"id":"q145","n":145,"tt":"Principal soro oral utilizado para reidratar bezerras com diarreia ou que apresentam desidratação","t":"Escolha única","o":"q145_outro"},{"id":"q146","n":146,"tt":"Como é feito o monitoramento da tristeza parasitária bovina na fase de cria e recria?","t":"Múltipla escolha"}]},{"idx":"7","t":"Ambiente e instalações das bezerras","q":[{"id":"q147","n":147,"tt":"Modo de criação nos primeiros 30 dias de vida","t":"Escolha única"},{"id":"q148","n":148,"tt":"Instalação até os 30 dias de vida, em média","t":"Escolha única","o":"q148_outro"},{"id":"q149","n":149,"tt":"As gaiolas ou baias utilizadas até os 30 dias de vida possuem divisórias laterais fechadas, evitando o contato entre animais?","t":"Escolha única"},{"id":"q150","n":150,"tt":"Tipo de piso onde as bezerras ficam até os 30 dias de vida","t":"Escolha única","o":"q150_outro"},{"id":"q151","n":151,"tt":"Tipo de cama acima do piso nas instalações até os 30 dias de vida","t":"Escolha única","o":"q151_outro"},{"id":"q152","n":152,"tt":"Ventiladores nas instalações até os 30 dias de vida","t":"Escolha única"},{"id":"q153","n":153,"tt":"Largura, em metros, do local onde as bezerras ficam individualizadas até os 30 dias de vida","t":"Número"},{"id":"q154","n":154,"tt":"Comprimento, em metros, do local onde as bezerras ficam individualizadas até os 30 dias de vida","t":"Número"},{"id":"q155","n":155,"tt":"Área calculada, em m², do local individualizado até os 30 dias de vida","t":"Campo calculado"},{"id":"q156","n":156,"tt":"Modo de criação após 30 dias até o desaleitamento","t":"Escolha única"},{"id":"q156b","n":"156b","tt":"As instalações físicas das bezerras após os 30 dias são as mesmas de antes dos 30 dias?","t":"Escolha única"},{"id":"q157","n":157,"tt":"Tipo de instalações após os primeiros 30 dias de vida até o desaleitamento","t":"Escolha única","o":"q157_outro"},{"id":"q158","n":158,"tt":"As gaiolas ou baias utilizadas após os 30 dias de vida possuem divisórias laterais fechadas, evitando o contato entre animais?","t":"Escolha única"},{"id":"q159","n":159,"tt":"Tipo de piso onde as bezerras ficam após os 30 dias de vida","t":"Escolha única","o":"q159_outro"},{"id":"q160","n":160,"tt":"Tipo de cama acima do piso nas instalações após os 30 dias de vida","t":"Escolha única","o":"q160_outro"},{"id":"q161","n":161,"tt":"Ventiladores nas instalações após os 30 dias de vida até o desaleitamento","t":"Escolha única"},{"id":"q162","n":162,"tt":"Largura, em metros, do local onde as bezerras ficam individualizadas após os 30 dias de vida até o desaleitamento","t":"Número"},{"id":"q163","n":163,"tt":"Comprimento, em metros, do local onde as bezerras ficam individualizadas após os 30 dias de vida até o desaleitamento","t":"Número"},{"id":"q164","n":164,"tt":"Área calculada, em m², do local individualizado após 30 dias de vida até o desaleitamento","t":"Campo calculado"}]},{"idx":"8","t":"Desempenho e manejo das bezerras e novilhas","q":[{"id":"q165","n":165,"tt":"Após o desaleitamento, por quantos dias as bezerras permanecem no mesmo ambiente de aleitamento?","t":"Escolha única"},{"id":"q166","n":166,"tt":"Modo de pesagem das bezerras após o nascimento","t":"Escolha única"},{"id":"q167","n":167,"tt":"Modo de pesagem das novilhas para reprodução","t":"Escolha única"},{"id":"q168","n":168,"tt":"Modo de pesagem das novilhas ao parto","t":"Escolha única"},{"id":"q169","n":169,"tt":"Número de novilhas no primeiro grupo de transição após o desaleitamento","t":"Escolha única"},{"id":"q170","n":170,"tt":"Modo de fornecimento da dieta para as novilhas após saírem da fase de aleitamento","t":"Escolha única","o":"q170_outro"},{"id":"q171","n":171,"tt":"Concentrado utilizado para as novilhas após saírem da fase de aleitamento","t":"Escolha única","o":"q171_outro"},{"id":"q172","n":172,"tt":"Percentual de proteína bruta (PB) do concentrado fornecido para as novilhas na fase de transição","t":"Escolha única"},{"id":"q173","n":173,"tt":"Número de novilhas no segundo grupo após a transição","t":"Escolha única"},{"id":"q174","n":174,"tt":"Sistema de criação das novilhas","t":"Escolha única","o":"q174_outro"},{"id":"q175","n":175,"tt":"Quais critérios são considerados para liberar as novilhas para a primeira cobertura/inseminação?","t":"Múltipla escolha"},{"id":"q176","n":176,"tt":"Peso médio, em kg, do animal adulto da fazenda","t":"Número"},{"id":"q176b","n":"176b","tt":"A fazenda utiliza inseminação artificial nas novilhas?","t":"Escolha única"},{"id":"q177","n":177,"tt":"Utiliza sêmen sexado nas novilhas?","t":"Escolha única"},{"id":"q178","n":178,"tt":"Qual o número de tentativas de sêmen sexado nas novilhas?","t":"Escolha única"},{"id":"q179","n":179,"tt":"Com qual(is) empresa(s) de inseminação você trabalha atualmente em sua fazenda?","t":"Múltipla escolha","o":"q179_outro"},{"id":"q180","n":180,"tt":"Quais características genéticas são consideradas na escolha dos touros para o melhoramento genético das bezerras?","t":"Múltipla escolha","o":"q180_outro"},{"id":"q181","n":181,"tt":"Você utiliza avaliação genômica nas bezerras?","t":"Escolha única"}]},{"idx":"9","t":"Custos, tecnologias e criação de machos","q":[{"id":"q182","n":182,"tt":"Custo médio total, em R$, de 1 bezerra desaleitada na fazenda","t":"Número"},{"id":"q183","n":183,"tt":"O valor do custo da bezerra foi","t":"Escolha única"},{"id":"q184","n":184,"tt":"Custo médio total, em R$, de 1 novilha pronta para parir na fazenda","t":"Número"},{"id":"q185","n":185,"tt":"O valor do custo da novilha foi","t":"Escolha única"},{"id":"q186","n":186,"tt":"A fazenda utiliza sêmen de corte, exemplo Angus, para produzir animais mestiços leite × corte para comercialização?","t":"Escolha única"},{"id":"q187","n":187,"tt":"Destino principal do bezerro macho","t":"Escolha única","o":"q187_outro"},{"id":"q188","n":188,"tt":"Bezerros machos são criados nas mesmas instalações das bezerras até o seu destino final?","t":"Escolha única"},{"id":"q189","n":189,"tt":"Principal produto utilizado na cura de umbigo dos bezerros machos","t":"Escolha única","o":"q189_outro"},{"id":"q190","n":190,"tt":"Colostragem dos bezerros machos","t":"Escolha única"}]},{"idx":"10","t":"Avaliação geral do Programa Alta CRIA","q":[{"id":"q191","n":191,"tt":"Nota geral do Programa Alta CRIA","t":"Escala 0 a 10"},{"id":"q192","n":192,"tt":"Nota para o grupo de WhatsApp do Alta CRIA","t":"Escala 0 a 10"},{"id":"q193","n":193,"tt":"Você já participou de alguma reunião virtual do Alta CRIA TA ON para discussão dos seus dados?","t":"Escolha única"},{"id":"q194","n":194,"tt":"Se você já participou de alguma reunião virtual do Alta CRIA TA ON para discussão dos seus dados, qual nota você daria para esse tipo de reunião?","t":"Escala 0 a 10"}]}];

// Tipos que podem ser publicados. O resto (nome, e-mail, telefone, código,
// texto longo e os campos "Outro") fica fora: é digitado e pode identificar.
var TIPOS_PUBLICAVEIS = {
  'Escolha única': 'categorico',
  'Múltipla escolha': 'multipla',
  'Número': 'numerico',
  'Campo calculado': 'numerico',
  'Escala 0 a 10': 'escala',
  'Resposta curta / UF': 'categorico'
};

// Regra de célula mínima: hoje desligada (0) porque o dashboard não tem
// filtros — sem recorte, não há célula pequena. Quando os filtros entrarem,
// suba para 5 ou 10 e a regra passa a valer sozinha.
var MIN_FAZENDAS_POR_CELULA = 0;

function gerarDadosPublicos() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.RESPOSTAS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    return JSON.stringify({ gerado_em: new Date(), n_fazendas: 0, capitulos: [], aviso: 'Nenhuma resposta ainda.' });
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var iCod = headers.indexOf('codigoAcesso');
  var iTs = headers.indexOf('timestamp');

  // Só a resposta mais recente de cada fazenda entra na conta (REENVIO substitui).
  var porCodigo = {};
  for (var i = 0; i < vals.length; i++) {
    var cod = normalizeCode_(vals[i][iCod]);
    if (!cod) continue;
    var ts = iTs > -1 ? vals[i][iTs] : i;
    if (!porCodigo[cod] || ts >= porCodigo[cod].ts) porCodigo[cod] = { ts: ts, row: vals[i] };
  }
  var linhas = [];
  for (var k in porCodigo) linhas.push(porCodigo[k].row);

  var MAPA_CUR = carregarMapaCuradoria_();
  var pendentes = 0;

  var capitulos = [];
  for (var c = 0; c < SCHEMA_PUBLICO.length; c++) {
    var cap = SCHEMA_PUBLICO[c];
    var perguntas = [];
    for (var p = 0; p < cap.q.length; p++) {
      var q = cap.q[p];
      var tipo = TIPOS_PUBLICAVEIS[q.t];
      if (!tipo) continue;
      var col = headers.indexOf(q.id);
      if (col === -1) continue;
      var brutos = [];
      for (var L = 0; L < linhas.length; L++) {
        var v = linhas[L][col];
        if (v !== '' && v !== null && v !== undefined) brutos.push(v);
      }
      var item = { id: q.id, num: q.n, titulo: q.tt, tipo: tipo, n: brutos.length };
      if (tipo === 'numerico' || tipo === 'escala') {
        var nums = [];
        for (var x = 0; x < brutos.length; x++) {
          var f = parseFloat(String(brutos[x]).replace(',', '.'));
          if (!isNaN(f)) nums.push(f);
        }
        item.n = nums.length;
        if (nums.length) {
          nums.sort(function (a, b) { return a - b; });
          item.media = arred_(soma_(nums) / nums.length);
          item.mediana = arred_(percentil_(nums, 50));
          item.p25 = arred_(percentil_(nums, 25));
          item.p75 = arred_(percentil_(nums, 75));
          item.hist = histograma_(nums);
        }
      } else {
        // Para Q029 e Q030, o texto de "Outro" foi liberado para integrar
        // automaticamente as respostas principais. Os demais campos mantêm
        // a curadoria antes de qualquer publicação.
        var mapaQ = q.o ? (MAPA_CUR[q.id] || {}) : null;
        var colOutro = q.o ? headers.indexOf(q.o) : -1;
        var cont = {}, nResp = 0;
        for (var y = 0; y < linhas.length; y++) {
          var cru = linhas[y][col];
          if (cru === '' || cru === null || cru === undefined) continue;
          nResp++;
          var partes = tipo === 'multipla' ? String(cru).split('|') : [String(cru)];
          for (var z = 0; z < partes.length; z++) {
            var rot = partes[z].trim();
            if (!rot) continue;
            if (ehOpcaoOutro_(rot) && colOutro > -1) {
              var rr = (q.id === 'q029' || q.id === 'q030')
                ? rotulosDiretosDoOutro_(linhas[y][colOutro])
                : rotulosDoOutro_(mapaQ, linhas[y][colOutro]);
              for (var w = 0; w < rr.length; w++) cont[rr[w]] = (cont[rr[w]] || 0) + 1;
            } else {
              cont[rot] = (cont[rot] || 0) + 1;
            }
          }
        }
        item.n = nResp;
        var ops = [];
        for (var r in cont) {
          if (MIN_FAZENDAS_POR_CELULA && cont[r] < MIN_FAZENDAS_POR_CELULA) continue;
          ops.push({ rotulo: r, n: cont[r], pct: arred_(cont[r] * 100 / (nResp || 1)) });
        }
        ops.sort(function (a, b) { return b.n - a.n; });
        item.opcoes = ops;
        for (var pp = 0; pp < ops.length; pp++) if (ops[pp].rotulo === ROTULO_PENDENTE) pendentes += ops[pp].n;
      }
      perguntas.push(item);
    }
    if (perguntas.length) capitulos.push({ idx: cap.idx, titulo: cap.t, perguntas: perguntas });
  }

  var saida = {
    gerado_em: new Date(),
    n_fazendas: linhas.length,
    curadoria_pendente: pendentes,
    capitulos: capitulos
  };
  var json = JSON.stringify(saida);
  Logger.log('=== COPIE DAQUI PARA O ARQUIVO dados.json ===');
  Logger.log(json);
  return json;
}

function soma_(a) { var s = 0; for (var i = 0; i < a.length; i++) s += a[i]; return s; }
function arred_(v) { return Math.round(v * 100) / 100; }
function percentil_(ord, p) {
  if (!ord.length) return 0;
  var pos = (ord.length - 1) * p / 100;
  var b = Math.floor(pos), resto = pos - b;
  return ord[b + 1] !== undefined ? ord[b] + resto * (ord[b + 1] - ord[b]) : ord[b];
}
/**
 * Histograma com as pontas fechadas de propósito: a primeira e a última
 * faixa são abertas ("< x" e "≥ y"), então o valor máximo de nenhuma
 * fazenda aparece no gráfico. Extremo identifica fazenda; distribuição não.
 */
function histograma_(ord) {
  var n = ord.length;
  if (n < 4) return [];
  var lo = percentil_(ord, 5), hi = percentil_(ord, 95);
  if (hi <= lo) return [];
  var nb = Math.min(8, Math.max(4, Math.round(Math.sqrt(n))));
  var larg = (hi - lo) / nb;
  var faixas = [];
  for (var i = 0; i < nb; i++) faixas.push({ ini: lo + i * larg, fim: lo + (i + 1) * larg, n: 0 });
  for (var j = 0; j < n; j++) {
    var v = ord[j];
    var idx = Math.floor((v - lo) / larg);
    if (idx < 0) idx = 0;
    if (idx >= nb) idx = nb - 1;
    faixas[idx].n++;
  }
  var out = [];
  for (var f = 0; f < nb; f++) {
    var rot;
    if (f === 0) rot = '< ' + fmt_(faixas[f].fim);
    else if (f === nb - 1) rot = '≥ ' + fmt_(faixas[f].ini);
    else rot = fmt_(faixas[f].ini) + ' a ' + fmt_(faixas[f].fim);
    out.push({ faixa: rot, n: faixas[f].n });
  }
  return out;
}
function fmt_(v) {
  if (Math.abs(v) >= 100) return String(Math.round(v));
  if (Math.abs(v) >= 10) return String(Math.round(v * 10) / 10);
  return String(Math.round(v * 100) / 100);
}

/* ===================================================================
   CURADORIA DOS CAMPOS "OUTRO"
   O que a fazenda digita não vai para o dashboard como veio. Passa por
   aqui primeiro: você vê o texto, quantas fazendas escreveram aquilo, e
   escreve ao lado o rótulo oficial. Só o que tem rótulo oficial é
   publicado — e entra como uma opção de verdade, com % calculado.
   Texto sem rótulo aparece agrupado como "Outro (aguardando
   classificação)", nunca com as palavras que a fazenda escreveu.
   =================================================================== */

var CURADORIA_SHEET = 'Curadoria_Outros';
var ROTULO_PENDENTE = 'Outro (aguardando classificação)';
var OPCOES_OUTRO = { 'outro': 1, 'outra': 1, 'outros': 1, 'outro(s)': 1, 'outra(s)': 1 };

function ehOpcaoOutro_(v) {
  return !!OPCOES_OUTRO[String(v || '').trim().toLowerCase()];
}
function chaveTexto_(v) {
  return removeAccents_(String(v || '')).toLowerCase().replace(/\s+/g, ' ').trim();
}

/**
 * Monta/atualiza a aba de curadoria. Pode rodar quantas vezes quiser:
 * os rótulos que você já escreveu são preservados. Textos novos entram
 * no fim, com o rótulo em branco esperando você.
 */
function gerarCuradoriaOutros() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.RESPOSTAS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 'Nenhuma resposta ainda — nada a curar.';

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  // preserva o que você já classificou
  var jaFeito = {};
  var cur = ss.getSheetByName(CURADORIA_SHEET);
  if (cur && cur.getLastRow() > 1) {
    var antigo = cur.getRange(2, 1, cur.getLastRow() - 1, 6).getValues();
    for (var a = 0; a < antigo.length; a++) {
      var rot = String(antigo[a][5] || '').trim();
      if (rot) jaFeito[antigo[a][0] + '||' + chaveTexto_(antigo[a][3])] = rot;
    }
  }

  var linhas = [], novos = 0, pendentes = 0;
  for (var c = 0; c < SCHEMA_PUBLICO.length; c++) {
    for (var p = 0; p < SCHEMA_PUBLICO[c].q.length; p++) {
      var q = SCHEMA_PUBLICO[c].q[p];
      if (!q.o) continue;
      var col = headers.indexOf(q.o);
      if (col === -1) continue;
      var cont = {}, orig = {};
      for (var i = 0; i < vals.length; i++) {
        var bruto = String(vals[i][col] || '').trim();
        if (!bruto) continue;
        var partes = bruto.split('|'); // "Outro" repetível grava separado por " | "
        for (var z = 0; z < partes.length; z++) {
          var txt = partes[z].trim();
          if (!txt) continue;
          var k = chaveTexto_(txt);
          cont[k] = (cont[k] || 0) + 1;
          if (!orig[k]) orig[k] = txt; // guarda a 1a grafia vista, para você ler
        }
      }
      var chaves = Object.keys(cont).sort(function (x, y) { return cont[y] - cont[x]; });
      for (var w = 0; w < chaves.length; w++) {
        var kk = chaves[w];
        var rotulo = jaFeito[q.id + '||' + kk] || '';
        if (!rotulo) { pendentes++; if (!(q.id + '||' + kk in jaFeito)) novos++; }
        linhas.push([q.id, q.n, q.tt, orig[kk], cont[kk], rotulo]);
      }
    }
  }

  var sh = cur || ss.insertSheet(CURADORIA_SHEET);
  sh.clear();
  var head = ['pergunta_id', 'num', 'pergunta', 'texto_digitado', 'fazendas', 'rotulo_oficial'];
  sh.getRange(1, 1, 1, 6).setValues([head]).setFontWeight('bold').setBackground('#0E2E4D').setFontColor('#FFFFFF');
  if (linhas.length) {
    sh.getRange(2, 1, linhas.length, 6).setValues(linhas);
    // a coluna que você preenche fica destacada
    sh.getRange(2, 6, linhas.length, 1).setBackground('#FFF6E0');
  }
  try {
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 320); sh.setColumnWidth(4, 260); sh.setColumnWidth(6, 260);
    sh.getRange(1, 1, linhas.length + 1, 6).createFilter();
  } catch (err) { /* formatação é cosmética */ }

  return 'Curadoria pronta: ' + linhas.length + ' textos distintos, ' + pendentes +
         ' aguardando rótulo. Escreva o rótulo oficial na última coluna. ' +
         'Para descartar um texto, escreva DESCARTAR.';
}

/**
 * Lê a curadoria: { q104: { "sprayfo delta": "Sprayfo Azul - Trouw" } }
 */
function carregarMapaCuradoria_() {
  var mapa = {};
  var sh = getSpreadsheet_().getSheetByName(CURADORIA_SHEET);
  if (!sh || sh.getLastRow() < 2) return mapa;
  var d = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < d.length; i++) {
    var rot = String(d[i][5] || '').trim();
    if (!rot) continue;
    var qid = String(d[i][0] || '').trim();
    if (!mapa[qid]) mapa[qid] = {};
    mapa[qid][chaveTexto_(d[i][3])] = rot;
  }
  return mapa;
}

/**
 * Devolve os rótulos que uma resposta "Outro" deve virar.
 * Sem rótulo aprovado -> vira o balde pendente (nunca o texto cru).
 */
function rotulosDoOutro_(mapaQ, textoBruto) {
  var out = [];
  var partes = String(textoBruto || '').split('|');
  for (var i = 0; i < partes.length; i++) {
    var t = partes[i].trim();
    if (!t) continue;
    var rot = mapaQ ? mapaQ[chaveTexto_(t)] : null;
    if (rot && rot.toUpperCase() === 'DESCARTAR') continue;
    out.push(rot || ROTULO_PENDENTE);
  }
  if (!out.length) out.push(ROTULO_PENDENTE);
  return out;
}

// Q029 (raça) e Q030 (software) são exceções aprovadas: os textos
// preenchidos em "Outro" são respostas analíticas válidas e entram
// diretamente na respectiva distribuição pública, sem curadoria prévia.
function rotulosDiretosDoOutro_(textoBruto) {
  var out = [];
  var partes = String(textoBruto || '').split('|');
  for (var i = 0; i < partes.length; i++) {
    var texto = partes[i].replace(/\s+/g, ' ').trim();
    if (texto) out.push(texto);
  }
  return out;
}

/**
 * Backend do Avaliador de Hackathon.
 *
 * Script Properties necessárias (Configurações do projeto > Propriedades do script):
 *   OPENROUTER_API_KEY - chave da API do OpenRouter (openrouter.ai)
 *   OPENROUTER_MODEL   - ex. "google/gemini-2.5-flash" (opcional, default abaixo)
 *   TOP_X               - quantos grupos selecionar (opcional, default 8)
 *   CRITERIA_DOC_ID   - ID do Google Doc com os critérios de avaliação
 *   DRIVE_FOLDER_ID   - ID da pasta do Drive onde os PDFs são salvos
 *   SHEET_ID          - ID da Planilha usada como banco de dados
 */

var HEADERS = ['Timestamp Envio', 'Grupo', 'Projeto', 'Drive File Id', 'Drive File Url',
  'Status', 'Score', 'Justificativa', 'Selecionado', 'Timestamp Avaliacao'];

function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ ok: false, error: 'JSON inválido no corpo da requisição' });
  }

  try {
    switch (payload.action) {
      case 'submit':
        return jsonResponse_(handleSubmit_(payload));
      case 'evaluate':
        return jsonResponse_(handleEvaluate_());
      case 'results':
        return jsonResponse_(handleResults_());
      default:
        return jsonResponse_({ ok: false, error: 'Ação desconhecida: ' + payload.action });
    }
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  return jsonResponse_({ ok: true, message: 'Avaliador Hackathon API ativo' });
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getProp_(key, defaultValue) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  return (value !== null && value !== '') ? value : defaultValue;
}

function normalize_(s) {
  return String(s || '').trim().toLowerCase();
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(getProp_('SHEET_ID'));
  var sheet = ss.getSheetByName('Submissoes');
  if (!sheet) {
    sheet = ss.insertSheet('Submissoes');
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRowByGrupo_(sheet, grupoNormalizado) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalize_(data[i][1]) === grupoNormalizado) return i + 1;
  }
  return -1;
}

function handleSubmit_(payload) {
  var grupo = String(payload.grupo || '').trim();
  var projeto = String(payload.projeto || '').trim();
  var filename = String(payload.filename || 'projeto.pdf');
  var data = payload.data;

  if (!grupo) throw new Error('Nome do grupo é obrigatório');
  if (!data) throw new Error('Arquivo PDF é obrigatório');

  var bytes = Utilities.base64Decode(data);
  if (bytes.length > 20 * 1024 * 1024) throw new Error('Arquivo excede o limite de 20MB');

  var folder = DriveApp.getFolderById(getProp_('DRIVE_FOLDER_ID'));
  var safeName = grupo.replace(/[^a-zA-Z0-9-_ ]/g, '').substring(0, 60);
  var blob = Utilities.newBlob(bytes, 'application/pdf', safeName + ' - ' + filename);
  var file = folder.createFile(blob);

  var sheet = getSheet_();
  var grupoNorm = normalize_(grupo);
  var rowIndex = findRowByGrupo_(sheet, grupoNorm);
  var now = new Date();

  if (rowIndex > 0) {
    var oldFileId = sheet.getRange(rowIndex, 4).getValue();
    if (oldFileId) {
      try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (err) { /* arquivo já removido */ }
    }
    sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([[
      now, grupo, projeto, file.getId(), file.getUrl(), 'pendente', '', '', '', ''
    ]]);
  } else {
    sheet.appendRow([now, grupo, projeto, file.getId(), file.getUrl(), 'pendente', '', '', '', '']);
  }

  return { ok: true, message: 'Recebido com sucesso' };
}

function handleEvaluate_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error('Já existe uma avaliação em andamento. Aguarde terminar e tente de novo.');
  }

  try {
    var criteriaText = DocumentApp.openById(getProp_('CRITERIA_DOC_ID')).getBody().getText();
    if (!criteriaText.trim()) throw new Error('O documento de critérios está vazio');

    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var topX = Number(getProp_('TOP_X', '8'));

    var results = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var grupo = row[1], projeto = row[2], fileId = row[3];
      if (!fileId) continue;

      var rowNumber = i + 1;
      try {
        var blob = DriveApp.getFileById(fileId).getBlob();
        var base64Pdf = Utilities.base64Encode(blob.getBytes());
        var evalResult = callOpenRouter_(base64Pdf, criteriaText, projeto, grupo);
        sheet.getRange(rowNumber, 6, 1, 5).setValues([[
          'avaliado', evalResult.score, evalResult.justificativa, false, new Date()
        ]]);
        results.push({ grupo: grupo, projeto: projeto, score: evalResult.score,
          justificativa: evalResult.justificativa, erro: false, _row: rowNumber });
      } catch (err) {
        sheet.getRange(rowNumber, 6, 1, 5).setValues([[
          'erro', '', String(err), false, new Date()
        ]]);
        results.push({ grupo: grupo, projeto: projeto, score: null,
          justificativa: 'Erro na avaliação: ' + String(err), erro: true, _row: rowNumber });
      }
      Utilities.sleep(1000);
    }

    results.sort(function (a, b) { return (b.score || -1) - (a.score || -1); });
    for (var r = 0; r < results.length; r++) {
      results[r].selecionado = r < topX && !results[r].erro;
      sheet.getRange(results[r]._row, 9).setValue(results[r].selecionado);
      delete results[r]._row;
    }

    return { ok: true, resultados: results };
  } finally {
    lock.releaseLock();
  }
}

function handleResults_() {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[1]) continue;
    results.push({
      grupo: row[1],
      projeto: row[2],
      status: row[5],
      score: row[6] === '' ? null : Number(row[6]),
      justificativa: row[7],
      selecionado: row[8] === true
    });
  }
  results.sort(function (a, b) { return (b.score || -1) - (a.score || -1); });
  return { ok: true, resultados: results };
}

function callOpenRouter_(base64Pdf, criteriaText, projectName, groupName) {
  var apiKey = getProp_('OPENROUTER_API_KEY');
  if (!apiKey) throw new Error('OPENROUTER_API_KEY não configurada nas Propriedades do Script');
  var model = getProp_('OPENROUTER_MODEL', 'google/gemini-2.5-flash');
  var url = 'https://openrouter.ai/api/v1/chat/completions';

  var prompt = [
    'Você é um avaliador de um hackathon corporativo. Avalie o projeto descrito no PDF anexado',
    'seguindo ESTRITAMENTE os critérios abaixo.',
    '',
    '=== CRITÉRIOS DE AVALIAÇÃO ===',
    criteriaText,
    '=== FIM DOS CRITÉRIOS ===',
    '',
    'Grupo: ' + groupName,
    'Projeto: ' + (projectName || '(não informado)'),
    '',
    'Responda APENAS com um JSON no formato exato, sem markdown:',
    '{"score": <numero de 0 a 100>, "justificativa": "<explicação em até 3 frases>"}'
  ].join('\n');

  var body = {
    model: model,
    messages: [{
      role: 'user',
      content: [
        { type: 'text', text: prompt },
        {
          type: 'file',
          file: { filename: 'projeto.pdf', file_data: 'data:application/pdf;base64,' + base64Pdf }
        }
      ]
    }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': 'Avaliador Hackathon'
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  };

  var lastError = null;
  for (var attempt = 0; attempt < 3; attempt++) {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code === 200) {
      var json = JSON.parse(response.getContentText());
      var text = json.choices[0].message.content;
      var cleaned = String(text).replace(/```json/gi, '').replace(/```/g, '').trim();
      var parsed = JSON.parse(cleaned);
      return { score: Number(parsed.score), justificativa: String(parsed.justificativa || '') };
    }

    if (code === 429 || code >= 500) {
      lastError = 'HTTP ' + code + ': ' + response.getContentText();
      Utilities.sleep(2000 * (attempt + 1));
      continue;
    }

    throw new Error('OpenRouter HTTP ' + code + ': ' + response.getContentText());
  }
  throw new Error(lastError || 'Falha ao chamar o OpenRouter após 3 tentativas');
}

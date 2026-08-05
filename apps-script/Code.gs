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
  'Status', 'Score Individual', 'Justificativa Individual', 'Score Final',
  'Comentario Recalibracao', 'Selecionado', 'Timestamp Avaliacao'];

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
      now, grupo, projeto, file.getId(), file.getUrl(), 'pendente', '', '', '', '', '', ''
    ]]);
  } else {
    sheet.appendRow([now, grupo, projeto, file.getId(), file.getUrl(), 'pendente', '', '', '', '', '', '']);
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

    // Etapa 1: nota individual por projeto, cada PDF avaliado isoladamente.
    var avaliados = [];
    var comErro = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var grupo = row[1], projeto = row[2], fileId = row[3];
      if (!fileId) continue;

      var rowNumber = i + 1;
      try {
        var blob = DriveApp.getFileById(fileId).getBlob();
        var base64Pdf = Utilities.base64Encode(blob.getBytes());
        var evalResult = callOpenRouter_(base64Pdf, criteriaText, projeto, grupo);
        avaliados.push({
          grupo: grupo, projeto: projeto,
          scoreIndividual: evalResult.score, justificativaIndividual: evalResult.justificativa,
          _row: rowNumber
        });
      } catch (err) {
        comErro.push({ grupo: grupo, projeto: projeto, erro: String(err), _row: rowNumber });
      }
      Utilities.sleep(1000);
    }

    // Etapa 2: recalibração vendo todos os grupos avaliados de uma vez (só texto, sem PDF),
    // para corrigir distorções de comparar notas que foram dadas em chamadas isoladas.
    var ranqueados;
    if (avaliados.length > 0) {
      try {
        ranqueados = recalibrar_(criteriaText, avaliados, topX);
      } catch (err) {
        ranqueados = avaliados.slice().sort(function (a, b) { return b.scoreIndividual - a.scoreIndividual; });
        ranqueados.forEach(function (r, idx) {
          r.scoreFinal = r.scoreIndividual;
          r.comentarioRecalibracao = 'Recalibração indisponível, mantida nota individual: ' + String(err);
          r.selecionado = idx < topX;
        });
      }
    } else {
      ranqueados = [];
    }

    ranqueados.forEach(function (r) {
      sheet.getRange(r._row, 6, 1, 7).setValues([[
        'avaliado', r.scoreIndividual, r.justificativaIndividual, r.scoreFinal,
        r.comentarioRecalibracao, r.selecionado, new Date()
      ]]);
    });
    comErro.forEach(function (r) {
      sheet.getRange(r._row, 6, 1, 7).setValues([[
        'erro', '', r.erro, '', '', false, new Date()
      ]]);
    });

    var resultados = ranqueados.map(function (r) {
      return {
        grupo: r.grupo, projeto: r.projeto, status: 'avaliado',
        score_individual: r.scoreIndividual, justificativa_individual: r.justificativaIndividual,
        score_final: r.scoreFinal, comentario_recalibracao: r.comentarioRecalibracao,
        selecionado: r.selecionado, erro: false
      };
    }).concat(comErro.map(function (r) {
      return {
        grupo: r.grupo, projeto: r.projeto, status: 'erro',
        score_individual: null, justificativa_individual: r.erro,
        score_final: null, comentario_recalibracao: null,
        selecionado: false, erro: true
      };
    }));

    resultados.sort(function (a, b) { return (b.score_final || -1) - (a.score_final || -1); });

    return { ok: true, resultados: resultados };
  } finally {
    lock.releaseLock();
  }
}

// Pede a uma segunda chamada (só texto) para comparar todas as notas individuais de uma vez
// e devolver um ranking final calibrado. Trata grupos omitidos/inventados pela IA e garante
// que no máximo topX fiquem marcados como selecionados, mesmo se a resposta exagerar.
function recalibrar_(criteriaText, avaliados, topX) {
  var ranking = callRecalibracao_(criteriaText, avaliados, topX);

  var porGrupo = {};
  avaliados.forEach(function (a) { porGrupo[normalize_(a.grupo)] = a; });

  var vistos = {};
  var resultado = [];

  ranking.forEach(function (item) {
    var original = porGrupo[normalize_(item.grupo)];
    if (!original) return; // grupo inventado pela IA, ignora
    vistos[normalize_(item.grupo)] = true;
    resultado.push({
      grupo: original.grupo, projeto: original.projeto,
      scoreIndividual: original.scoreIndividual, justificativaIndividual: original.justificativaIndividual,
      scoreFinal: (typeof item.score_final === 'number' && !isNaN(item.score_final)) ? item.score_final : original.scoreIndividual,
      comentarioRecalibracao: item.comentario || '',
      selecionado: !!item.selecionado,
      _row: original._row
    });
  });

  avaliados.forEach(function (a) {
    if (!vistos[normalize_(a.grupo)]) {
      resultado.push({
        grupo: a.grupo, projeto: a.projeto,
        scoreIndividual: a.scoreIndividual, justificativaIndividual: a.justificativaIndividual,
        scoreFinal: a.scoreIndividual,
        comentarioRecalibracao: '(grupo omitido na recalibração; mantida nota individual)',
        selecionado: false,
        _row: a._row
      });
    }
  });

  resultado.sort(function (a, b) { return (b.scoreFinal || -1) - (a.scoreFinal || -1); });

  var selecionadosCount = 0;
  resultado.forEach(function (r) {
    if (r.selecionado) {
      selecionadosCount++;
      if (selecionadosCount > topX) r.selecionado = false;
    }
  });

  return resultado;
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
      score_individual: row[6] === '' ? null : Number(row[6]),
      justificativa_individual: row[7],
      score_final: row[8] === '' ? null : Number(row[8]),
      comentario_recalibracao: row[9],
      selecionado: row[10] === true,
      erro: row[5] === 'erro'
    });
  }
  results.sort(function (a, b) { return (b.score_final || -1) - (a.score_final || -1); });
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

// Segunda chamada à IA, só com texto (sem reenviar os PDFs): compara as notas individuais
// de todos os grupos ao mesmo tempo e devolve um ranking final calibrado + seleção dos top X.
function callRecalibracao_(criteriaText, avaliados, topX) {
  var apiKey = getProp_('OPENROUTER_API_KEY');
  if (!apiKey) throw new Error('OPENROUTER_API_KEY não configurada nas Propriedades do Script');
  var model = getProp_('OPENROUTER_MODEL', 'google/gemini-2.5-flash');
  var url = 'https://openrouter.ai/api/v1/chat/completions';

  var listaTexto = avaliados.map(function (a, idx) {
    return (idx + 1) + '. Grupo: ' + a.grupo + ' | Projeto: ' + (a.projeto || '(não informado)') +
      ' | Nota individual: ' + a.scoreIndividual + ' | Justificativa: ' + a.justificativaIndividual;
  }).join('\n');

  var prompt = [
    'Você é um avaliador sênior de um hackathon corporativo, revisando notas dadas individualmente',
    'a cada projeto por outro avaliador que analisou cada PDF isoladamente, sem comparar com os demais.',
    '',
    '=== CRITÉRIOS DE AVALIAÇÃO ===',
    criteriaText,
    '=== FIM DOS CRITÉRIOS ===',
    '',
    'Notas individuais (isoladas) de cada grupo:',
    listaTexto,
    '',
    'Agora que você vê TODOS os grupos ao mesmo tempo, produza um ranking final calibrado,',
    'ajustando a ordem quando a comparação direta entre os projetos sugerir que uma nota',
    'isolada ficou alta ou baixa demais em relação às demais. Selecione exatamente os ' + topX,
    'melhores como selecionado=true (menos apenas se algum projeto claramente não atender',
    'aos critérios mínimos).',
    '',
    'Responda APENAS com um JSON no formato exato, sem markdown, incluindo TODOS os ' + avaliados.length + ' grupos:',
    '{"ranking": [{"grupo": "<nome exato do grupo>", "score_final": <numero 0-100>, "selecionado": <true|false>, "comentario": "<até 2 frases sobre a posição no ranking>"}]}'
  ].join('\n');

  var body = {
    model: model,
    messages: [{ role: 'user', content: prompt }]
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
      if (!parsed.ranking || !Array.isArray(parsed.ranking)) {
        throw new Error('Resposta de recalibração sem "ranking" válido');
      }
      return parsed.ranking;
    }

    if (code === 429 || code >= 500) {
      lastError = 'HTTP ' + code + ': ' + response.getContentText();
      Utilities.sleep(2000 * (attempt + 1));
      continue;
    }

    throw new Error('OpenRouter HTTP ' + code + ': ' + response.getContentText());
  }
  throw new Error(lastError || 'Falha ao chamar o OpenRouter (recalibração) após 3 tentativas');
}

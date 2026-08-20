/**
 * Backend do Avaliador de Hackathon.
 *
 * Script Properties necessárias (Configurações do projeto > Propriedades do script):
 *   OPENROUTER_API_KEY - chave da API do OpenRouter (openrouter.ai)
 *   OPENROUTER_MODEL   - ex. "google/gemini-2.5-flash" (opcional, default abaixo)
 *   TOP_X              - quantos grupos recomendar (opcional, default 10)
 *   CRITERIA_DOC_ID    - ID do Google Doc com os critérios de avaliação
 *   DRIVE_FOLDER_ID    - ID da pasta do Drive onde os PDFs são salvos
 *   SHEET_ID           - ID da Planilha usada como banco de dados
 *
 * Fluxo:
 *   submit   -> salva o PDF, gera FEEDBACK QUALITATIVO (sem nota) e devolve ao grupo
 *   evaluate -> nota individual secreta por projeto + recalibração comparativa (top X)
 *   results  -> lê o estado atual da planilha
 */

var HEADERS = ['Timestamp Envio', 'Grupo', 'Projeto', 'Drive File Id', 'Drive File Url',
  'Status', 'Feedback Qualitativo', 'Score Individual', 'Justificativa Individual',
  'Score Final', 'Comentario Recalibracao', 'Selecionado', 'Timestamp Avaliacao'];

// Índices de coluna (1-based), para não espalhar números mágicos pelo código.
var COL = {
  TIMESTAMP_ENVIO: 1, GRUPO: 2, PROJETO: 3, FILE_ID: 4, FILE_URL: 5,
  STATUS: 6, FEEDBACK: 7, SCORE_INDIVIDUAL: 8, JUSTIFICATIVA: 9,
  SCORE_FINAL: 10, COMENTARIO: 11, SELECIONADO: 12, TIMESTAMP_AVALIACAO: 13
};

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

function getCriteriaText_() {
  var text = DocumentApp.openById(getProp_('CRITERIA_DOC_ID')).getBody().getText();
  if (!text.trim()) throw new Error('O documento de critérios está vazio');
  return text;
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(getProp_('SHEET_ID'));
  var sheet = ss.getSheetByName('Submissoes');
  if (!sheet) {
    sheet = ss.insertSheet('Submissoes');
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Mantém o cabeçalho em dia quando o esquema muda entre versões do script.
  var current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (current.join('|') !== HEADERS.join('|')) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRowByGrupo_(sheet, grupoNormalizado) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalize_(data[i][COL.GRUPO - 1]) === grupoNormalizado) return i + 1;
  }
  return -1;
}

/**
 * Recebe a submissão do grupo: salva o PDF, gera o feedback qualitativo
 * (sem nenhuma nota) e grava/substitui a linha do grupo na planilha.
 * Uma nova submissão do mesmo nome de grupo substitui a anterior.
 */
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

  // O feedback é gerado antes de responder ao grupo, mas uma falha aqui não pode
  // fazer a submissão ser perdida: o PDF já está salvo e a linha é gravada de todo jeito.
  var base64Pdf = Utilities.base64Encode(bytes);
  var feedback = null;
  var feedbackErro = null;
  try {
    feedback = gerarFeedbackQualitativo_(base64Pdf, getCriteriaText_(), projeto, grupo);
  } catch (err) {
    feedbackErro = String(err);
  }

  var feedbackTexto = feedback ? formatFeedback_(feedback) : ('(feedback não gerado: ' + feedbackErro + ')');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();
    var rowIndex = findRowByGrupo_(sheet, normalize_(grupo));
    var linha = [new Date(), grupo, projeto, file.getId(), file.getUrl(),
      'aguardando avaliacao', feedbackTexto, '', '', '', '', '', ''];

    if (rowIndex > 0) {
      var oldFileId = sheet.getRange(rowIndex, COL.FILE_ID).getValue();
      if (oldFileId && oldFileId !== file.getId()) {
        try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (err) { /* já removido */ }
      }
      sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([linha]);
    } else {
      sheet.appendRow(linha);
    }
  } finally {
    lock.releaseLock();
  }

  if (!feedback) {
    return {
      ok: true,
      message: 'Projeto recebido! Não foi possível gerar o feedback automático agora, ' +
        'mas a submissão do seu grupo foi registrada com sucesso.',
      feedback: null
    };
  }

  return { ok: true, message: 'Projeto recebido com sucesso!', feedback: feedback };
}

/**
 * Feedback devolvido ao grupo na hora do envio. Estritamente qualitativo:
 * o prompt proíbe notas, pontuações, percentuais e comparações com outros grupos.
 */
function gerarFeedbackQualitativo_(base64Pdf, criteriaText, projectName, groupName) {
  var prompt = [
    'Você é um mentor de um hackathon corporativo dando retorno construtivo a um grupo',
    'que acabou de enviar o descritivo do projeto. Analise o PDF anexado à luz dos critérios abaixo.',
    '',
    '=== CRITÉRIOS DE AVALIAÇÃO DO EVENTO ===',
    criteriaText,
    '=== FIM DOS CRITÉRIOS ===',
    '',
    'Grupo: ' + groupName,
    'Projeto: ' + (projectName || '(não informado)'),
    '',
    'REGRAS OBRIGATÓRIAS DO FEEDBACK:',
    '- NUNCA inclua nota, pontuação, score, percentual, estrelas, conceito (A/B/C) ou',
    '  qualquer forma de medida numérica de qualidade.',
    '- NUNCA diga se o projeto é bom o suficiente, se vai ser selecionado, nem o compare',
    '  com outros grupos ou com uma média.',
    '- Escreva apenas análise qualitativa: o que está claro e bem construído, o que ficou',
    '  vago ou ausente, e o que o grupo pode desenvolver melhor.',
    '- Seja específico e acionável, citando trechos ou aspectos concretos do projeto.',
    '- Tom construtivo e encorajador, em português do Brasil.',
    '',
    'Responda APENAS com um JSON no formato exato, sem markdown:',
    '{"resumo": "<2 a 3 frases descrevendo o que a ferramenta entendeu do projeto>",',
    ' "pontos_fortes": ["<aspecto bem desenvolvido>", "..."],',
    ' "pontos_de_atencao": ["<aspecto vago, ausente ou frágil>", "..."],',
    ' "sugestoes": ["<recomendação prática para fortalecer o projeto>", "..."]}',
    '',
    'Use de 2 a 4 itens em cada lista.'
  ].join('\n');

  var parsed = chamarIA_(prompt, base64Pdf);
  return {
    resumo: String(parsed.resumo || ''),
    pontos_fortes: asStringArray_(parsed.pontos_fortes),
    pontos_de_atencao: asStringArray_(parsed.pontos_de_atencao),
    sugestoes: asStringArray_(parsed.sugestoes)
  };
}

function asStringArray_(value) {
  if (!value) return [];
  if (!Array.isArray(value)) return [String(value)];
  return value.map(function (v) { return String(v); });
}

function formatFeedback_(fb) {
  var partes = ['RESUMO: ' + fb.resumo];
  if (fb.pontos_fortes.length) partes.push('PONTOS FORTES:\n- ' + fb.pontos_fortes.join('\n- '));
  if (fb.pontos_de_atencao.length) partes.push('PONTOS DE ATENÇÃO:\n- ' + fb.pontos_de_atencao.join('\n- '));
  if (fb.sugestoes.length) partes.push('SUGESTÕES:\n- ' + fb.sugestoes.join('\n- '));
  return partes.join('\n\n');
}

function handleEvaluate_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error('Já existe uma avaliação em andamento. Aguarde terminar e tente de novo.');
  }

  try {
    var criteriaText = getCriteriaText_();
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var topX = Number(getProp_('TOP_X', '10'));

    // Etapa 1: nota individual secreta, cada PDF avaliado isoladamente,
    // levando em conta também o feedback qualitativo gerado no envio.
    var avaliados = [];
    var comErro = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var grupo = row[COL.GRUPO - 1];
      var projeto = row[COL.PROJETO - 1];
      var fileId = row[COL.FILE_ID - 1];
      var feedbackTexto = row[COL.FEEDBACK - 1];
      if (!fileId) continue;

      var rowNumber = i + 1;
      try {
        var blob = DriveApp.getFileById(fileId).getBlob();
        var base64Pdf = Utilities.base64Encode(blob.getBytes());
        var evalResult = avaliarProjeto_(base64Pdf, criteriaText, projeto, grupo, feedbackTexto);
        avaliados.push({
          grupo: grupo, projeto: projeto, feedback: String(feedbackTexto || ''),
          scoreIndividual: evalResult.score, justificativaIndividual: evalResult.justificativa,
          _row: rowNumber
        });
      } catch (err) {
        comErro.push({ grupo: grupo, projeto: projeto, erro: String(err), _row: rowNumber });
      }
      Utilities.sleep(1000);
    }

    // Etapa 2: recalibração vendo todos os grupos de uma vez (só texto, sem reenviar
    // os PDFs), combinando nota individual e análise qualitativa para recomendar o top X
    // por potencial de impacto.
    var ranqueados = [];
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
    }

    // Escreve status e o bloco de avaliação sem tocar na coluna de feedback qualitativo.
    ranqueados.forEach(function (r) {
      sheet.getRange(r._row, COL.STATUS).setValue('avaliado');
      sheet.getRange(r._row, COL.SCORE_INDIVIDUAL, 1, 6).setValues([[
        r.scoreIndividual, r.justificativaIndividual, r.scoreFinal,
        r.comentarioRecalibracao, r.selecionado, new Date()
      ]]);
    });
    comErro.forEach(function (r) {
      sheet.getRange(r._row, COL.STATUS).setValue('erro');
      sheet.getRange(r._row, COL.JUSTIFICATIVA).setValue(r.erro);
      sheet.getRange(r._row, COL.SELECIONADO).setValue(false);
      sheet.getRange(r._row, COL.TIMESTAMP_AVALIACAO).setValue(new Date());
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
    return { ok: true, resultados: resultados, top_x: topX };
  } finally {
    lock.releaseLock();
  }
}

/** Nota individual secreta de um projeto (nunca exposta ao grupo). */
function avaliarProjeto_(base64Pdf, criteriaText, projectName, groupName, feedbackTexto) {
  var prompt = [
    'Você é um avaliador de um hackathon corporativo. Avalie o projeto descrito no PDF anexado',
    'seguindo ESTRITAMENTE os critérios abaixo. Esta nota é interna e nunca será mostrada ao grupo.',
    '',
    '=== CRITÉRIOS DE AVALIAÇÃO ===',
    criteriaText,
    '=== FIM DOS CRITÉRIOS ===',
    '',
    'Grupo: ' + groupName,
    'Projeto: ' + (projectName || '(não informado)'),
    '',
    '=== ANÁLISE QUALITATIVA JÁ FEITA SOBRE ESTE PROJETO ===',
    String(feedbackTexto || '(sem análise qualitativa registrada)'),
    '=== FIM DA ANÁLISE QUALITATIVA ===',
    '',
    'Considere tanto o conteúdo do PDF quanto a análise qualitativa acima. Dê peso especial ao',
    'POTENCIAL DE IMPACTO do projeto segundo os critérios do evento.',
    '',
    'Responda APENAS com um JSON no formato exato, sem markdown:',
    '{"score": <numero de 0 a 100>, "justificativa": "<explicação em até 3 frases>"}'
  ].join('\n');

  var parsed = chamarIA_(prompt, base64Pdf);
  return { score: Number(parsed.score), justificativa: String(parsed.justificativa || '') };
}

/**
 * Compara todos os grupos avaliados de uma vez e devolve o ranking final calibrado.
 * Protege contra grupos omitidos/inventados na resposta e limita a seleção a topX.
 */
function recalibrar_(criteriaText, avaliados, topX) {
  var ranking = chamarRecalibracao_(criteriaText, avaliados, topX);

  var porGrupo = {};
  avaliados.forEach(function (a) { porGrupo[normalize_(a.grupo)] = a; });

  var vistos = {};
  var resultado = [];

  ranking.forEach(function (item) {
    var original = porGrupo[normalize_(item.grupo)];
    if (!original) return; // grupo inventado pela IA, ignora
    if (vistos[normalize_(item.grupo)]) return; // duplicado na resposta
    vistos[normalize_(item.grupo)] = true;
    resultado.push({
      grupo: original.grupo, projeto: original.projeto,
      scoreIndividual: original.scoreIndividual,
      justificativaIndividual: original.justificativaIndividual,
      scoreFinal: (typeof item.score_final === 'number' && !isNaN(item.score_final))
        ? item.score_final : original.scoreIndividual,
      comentarioRecalibracao: String(item.comentario || ''),
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

  var selecionados = 0;
  resultado.forEach(function (r) {
    if (r.selecionado) {
      selecionados++;
      if (selecionados > topX) r.selecionado = false;
    }
  });

  return resultado;
}

function chamarRecalibracao_(criteriaText, avaliados, topX) {
  var lista = avaliados.map(function (a, idx) {
    return [
      '--- Grupo ' + (idx + 1) + ' ---',
      'Nome do grupo: ' + a.grupo,
      'Projeto: ' + (a.projeto || '(não informado)'),
      'Nota individual (interna): ' + a.scoreIndividual,
      'Justificativa da nota: ' + a.justificativaIndividual,
      'Análise qualitativa do projeto: ' + (a.feedback || '(sem análise registrada)')
    ].join('\n');
  }).join('\n\n');

  var prompt = [
    'Você é o avaliador sênior de um hackathon corporativo. Outro avaliador analisou cada',
    'projeto isoladamente, sem comparar com os demais: você recebe agora a nota individual',
    'e a análise qualitativa de TODOS os grupos ao mesmo tempo.',
    '',
    '=== CRITÉRIOS DE AVALIAÇÃO ===',
    criteriaText,
    '=== FIM DOS CRITÉRIOS ===',
    '',
    '=== GRUPOS AVALIADOS (' + avaliados.length + ') ===',
    lista,
    '=== FIM DOS GRUPOS ===',
    '',
    'Sua tarefa: recomendar os ' + topX + ' grupos cujos projetos têm o MAIOR POTENCIAL DE IMPACTO',
    'segundo os critérios do evento. Combine as duas fontes de informação — a nota individual',
    'e a análise qualitativa — e ajuste a ordem quando a comparação direta entre os projetos',
    'mostrar que uma nota isolada ficou alta ou baixa demais em relação às demais.',
    'Marque exatamente ' + topX + ' grupos com selecionado=true (menos apenas se houver menos',
    'grupos que isso, ou se algum projeto claramente não atender aos critérios mínimos).',
    '',
    'Responda APENAS com um JSON no formato exato, sem markdown, incluindo TODOS os ' +
      avaliados.length + ' grupos, do mais forte para o mais fraco:',
    '{"ranking": [{"grupo": "<nome exato do grupo>", "score_final": <numero 0-100>,',
    ' "selecionado": <true|false>, "comentario": "<até 2 frases justificando a posição e o potencial de impacto>"}]}'
  ].join('\n');

  var parsed = chamarIA_(prompt, null);
  if (!parsed.ranking || !Array.isArray(parsed.ranking)) {
    throw new Error('Resposta de recalibração sem "ranking" válido');
  }
  return parsed.ranking;
}

/**
 * Chamada única ao OpenRouter. Se base64Pdf vier preenchido, anexa o PDF à mensagem.
 * Retorna o JSON já parseado da resposta do modelo. Faz retry em 429 e erros 5xx.
 */
function chamarIA_(prompt, base64Pdf) {
  var apiKey = getProp_('OPENROUTER_API_KEY');
  if (!apiKey) throw new Error('OPENROUTER_API_KEY não configurada nas Propriedades do Script');
  var model = getProp_('OPENROUTER_MODEL', 'google/gemini-2.5-flash');
  var url = 'https://openrouter.ai/api/v1/chat/completions';

  var content = [{ type: 'text', text: prompt }];
  if (base64Pdf) {
    content.push({
      type: 'file',
      file: { filename: 'projeto.pdf', file_data: 'data:application/pdf;base64,' + base64Pdf }
    });
  }

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': 'Avaliador Hackathon'
    },
    payload: JSON.stringify({
      model: model,
      messages: [{ role: 'user', content: content }]
    }),
    muteHttpExceptions: true
  };

  var lastError = null;
  for (var attempt = 0; attempt < 3; attempt++) {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code === 200) {
      var json = JSON.parse(response.getContentText());
      if (!json.choices || !json.choices.length) {
        throw new Error('Resposta da IA sem conteúdo: ' + response.getContentText().substring(0, 300));
      }
      var text = String(json.choices[0].message.content || '');
      var cleaned = text.replace(/```json/gi, '').replace(/```/g, '').trim();
      try {
        return JSON.parse(cleaned);
      } catch (err) {
        throw new Error('A IA não devolveu JSON válido: ' + cleaned.substring(0, 300));
      }
    }

    if (code === 429 || code >= 500) {
      lastError = 'HTTP ' + code + ': ' + response.getContentText().substring(0, 300);
      Utilities.sleep(2000 * (attempt + 1));
      continue;
    }

    throw new Error('OpenRouter HTTP ' + code + ': ' + response.getContentText().substring(0, 300));
  }
  throw new Error(lastError || 'Falha ao chamar o OpenRouter após 3 tentativas');
}

function handleResults_() {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[COL.GRUPO - 1]) continue;
    results.push({
      grupo: row[COL.GRUPO - 1],
      projeto: row[COL.PROJETO - 1],
      status: row[COL.STATUS - 1],
      feedback: row[COL.FEEDBACK - 1],
      score_individual: row[COL.SCORE_INDIVIDUAL - 1] === '' ? null : Number(row[COL.SCORE_INDIVIDUAL - 1]),
      justificativa_individual: row[COL.JUSTIFICATIVA - 1],
      score_final: row[COL.SCORE_FINAL - 1] === '' ? null : Number(row[COL.SCORE_FINAL - 1]),
      comentario_recalibracao: row[COL.COMENTARIO - 1],
      selecionado: row[COL.SELECIONADO - 1] === true,
      erro: row[COL.STATUS - 1] === 'erro'
    });
  }
  results.sort(function (a, b) { return (b.score_final || -1) - (a.score_final || -1); });
  return { ok: true, resultados: results, top_x: Number(getProp_('TOP_X', '10')) };
}

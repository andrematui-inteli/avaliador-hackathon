/**
 * Backend do Avaliador de Hackathon.
 *
 * Script Properties necessárias (Configurações do projeto > Propriedades do script):
 *   OPENROUTER_API_KEY - chave da API do OpenRouter (openrouter.ai)
 *   OPENROUTER_MODEL   - ex. "nvidia/nemotron-3-ultra-550b-a55b:free" (opcional, default abaixo)
 *   TOP_X              - quantos grupos recomendar (opcional, default 10)
 *   CRITERIA_DOC_ID    - ID do Google Doc com os critérios de avaliação
 *   DRIVE_FOLDER_ID    - ID da pasta do Drive onde os PDFs são salvos
 *   SHEET_ID           - ID da Planilha usada como banco de dados
 *
 * Fluxo:
 *   submit   -> salva o PDF, extrai o texto, gera FEEDBACK QUALITATIVO (sem nota)
 *               e devolve ao grupo
 *   evaluate -> nota individual secreta por projeto + recalibração comparativa (top X)
 *   results  -> lê o estado atual da planilha
 *
 * O PDF é convertido em texto pelo próprio Drive e só o texto vai para a IA.
 * Imagens, diagramas e o layout de tabelas do PDF não entram na avaliação.
 */

var HEADERS = ['Timestamp Envio', 'Grupo', 'Projeto', 'Drive File Id', 'Drive File Url',
  'Status', 'Feedback Qualitativo', 'Score Individual', 'Justificativa Individual',
  'Score Final', 'Comentario Recalibracao', 'Selecionado', 'Timestamp Avaliacao',
  'Texto Extraido'];

// Índices de coluna (1-based), para não espalhar números mágicos pelo código.
var COL = {
  TIMESTAMP_ENVIO: 1, GRUPO: 2, PROJETO: 3, FILE_ID: 4, FILE_URL: 5,
  STATUS: 6, FEEDBACK: 7, SCORE_INDIVIDUAL: 8, JUSTIFICATIVA: 9,
  SCORE_FINAL: 10, COMENTARIO: 11, SELECIONADO: 12, TIMESTAMP_AVALIACAO: 13,
  TEXTO: 14
};

// Célula de planilha aceita ~50 mil caracteres; deixamos margem.
var MAX_TEXTO_CELULA = 45000;

// Quantas chamadas à IA disparar em paralelo por lote. Cada chamada leva ~9s, então
// avaliar 40 grupos em sequência estouraria o limite de execução do Apps Script.
var TAMANHO_LOTE_IA = 10;

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
 * Serializa apenas as escritas na planilha. O lock é sempre de curta duração:
 * nenhuma chamada à IA acontece dentro dele, para não travar as submissões dos
 * grupos enquanto uma avaliação longa está rodando.
 */
function withSheetLock_(fn) {
  var lock = LockService.getScriptLock();
  var acquired = false;
  for (var attempt = 0; attempt < 3 && !acquired; attempt++) {
    acquired = lock.tryLock(10000);
  }
  if (!acquired) throw new Error('A planilha está ocupada no momento. Tente novamente em alguns segundos.');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

var EVAL_FLAG = 'EVAL_RUNNING_AT';
var EVAL_FLAG_TTL_MS = 30 * 60 * 1000; // avaliação presumida travada depois disso

/**
 * Impede duas avaliações simultâneas sem usar o lock do script, que ficaria
 * retido por minutos e derrubaria as submissões que chegassem nesse meio-tempo.
 */
function acquireEvalFlag_() {
  var props = PropertiesService.getScriptProperties();
  withSheetLock_(function () {
    var startedAt = props.getProperty(EVAL_FLAG);
    if (startedAt && (new Date().getTime() - Number(startedAt)) < EVAL_FLAG_TTL_MS) {
      throw new Error('Já existe uma avaliação em andamento. Aguarde terminar e tente de novo.');
    }
    props.setProperty(EVAL_FLAG, String(new Date().getTime()));
  });
}

function releaseEvalFlag_() {
  PropertiesService.getScriptProperties().deleteProperty(EVAL_FLAG);
}

/**
 * Extrai o texto de um PDF já salvo no Drive, usando a conversão nativa do Drive
 * (que aplica OCR quando o PDF é escaneado). Copia o arquivo como Google Doc,
 * lê o texto e descarta o Doc temporário.
 *
 * Só o texto é enviado à IA — imagens, diagramas e o layout das tabelas do PDF
 * não são considerados na avaliação.
 */
function extrairTextoPdf_(fileId) {
  var url = 'https://www.googleapis.com/drive/v3/files/' + fileId +
    '/copy?ocrLanguage=pt&supportsAllDrives=true';
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({
      name: 'tmp-ocr-' + fileId,
      mimeType: 'application/vnd.google-apps.document'
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Falha ao converter o PDF em texto (HTTP ' + response.getResponseCode() +
      '): ' + response.getContentText().substring(0, 300));
  }

  var docId = JSON.parse(response.getContentText()).id;
  try {
    var texto = DocumentApp.openById(docId).getBody().getText();
    if (!texto || !texto.trim()) {
      throw new Error('O PDF não contém texto extraível. Se ele for composto apenas de ' +
        'imagens, reenvie um PDF com texto selecionável.');
    }
    return texto.trim();
  } finally {
    try { DriveApp.getFileById(docId).setTrashed(true); } catch (err) { /* ignora limpeza */ }
  }
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

  // A linha é gravada ANTES de gerar o feedback: se a chamada à IA falhar ou demorar,
  // a submissão do grupo já está registrada e nunca é perdida.
  var rowNumber = withSheetLock_(function () {
    var sheet = getSheet_();
    var rowIndex = findRowByGrupo_(sheet, normalize_(grupo));
    var linha = [new Date(), grupo, projeto, file.getId(), file.getUrl(),
      'gerando feedback', '', '', '', '', '', '', '', ''];

    if (rowIndex > 0) {
      var oldFileId = sheet.getRange(rowIndex, COL.FILE_ID).getValue();
      if (oldFileId && oldFileId !== file.getId()) {
        try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (err) { /* já removido */ }
      }
      sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([linha]);
      return rowIndex;
    }
    sheet.appendRow(linha);
    return sheet.getLastRow();
  });

  // O texto é extraído uma única vez aqui e reaproveitado na avaliação,
  // para não converter o mesmo PDF de novo mais tarde.
  var feedback = null;
  var erroEtapa = null;
  var textoPdf = '';
  try {
    textoPdf = extrairTextoPdf_(file.getId());
    feedback = gerarFeedbackQualitativo_(textoPdf, getCriteriaText_(), projeto, grupo);
  } catch (err) {
    erroEtapa = String(err);
  }

  var feedbackTexto = feedback ? formatFeedback_(feedback) : ('(feedback não gerado: ' + erroEtapa + ')');
  withSheetLock_(function () {
    var sheet = getSheet_();
    sheet.getRange(rowNumber, COL.STATUS).setValue('aguardando avaliacao');
    sheet.getRange(rowNumber, COL.FEEDBACK).setValue(feedbackTexto);
    sheet.getRange(rowNumber, COL.TEXTO).setValue(textoPdf.substring(0, MAX_TEXTO_CELULA));
  });

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
function gerarFeedbackQualitativo_(textoProjeto, criteriaText, projectName, groupName) {
  var prompt = [
    'Você é um mentor de um hackathon corporativo dando retorno construtivo a um grupo',
    'que acabou de enviar o descritivo do projeto. Analise o descritivo à luz dos critérios abaixo.',
    '',
    '=== CRITÉRIOS DE AVALIAÇÃO DO EVENTO ===',
    criteriaText,
    '=== FIM DOS CRITÉRIOS ===',
    '',
    'Grupo: ' + groupName,
    'Projeto: ' + (projectName || '(não informado)'),
    '',
    '=== DESCRITIVO DO PROJETO ===',
    textoProjeto,
    '=== FIM DO DESCRITIVO ===',
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

  var parsed = chamarIA_(prompt);
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
  acquireEvalFlag_();

  try {
    var criteriaText = getCriteriaText_();
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var topX = Number(getProp_('TOP_X', '10'));

    // Etapa 1: nota individual secreta, cada projeto avaliado isoladamente,
    // levando em conta também o feedback qualitativo gerado no envio.
    var candidatos = [];
    var comErro = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[COL.FILE_ID - 1]) continue;
      candidatos.push({
        grupo: row[COL.GRUPO - 1],
        projeto: row[COL.PROJETO - 1],
        fileId: row[COL.FILE_ID - 1],
        feedback: String(row[COL.FEEDBACK - 1] || ''),
        texto: String(row[COL.TEXTO - 1] || '').trim(),
        _row: i + 1
      });
    }

    // Reaproveita o texto extraído no envio; só reconverte se estiver faltando
    // (por exemplo, submissões feitas antes desta versão do script).
    var prontos = [];
    candidatos.forEach(function (c) {
      if (c.texto) { prontos.push(c); return; }
      try {
        c.texto = extrairTextoPdf_(c.fileId);
        sheet.getRange(c._row, COL.TEXTO).setValue(c.texto.substring(0, MAX_TEXTO_CELULA));
        prontos.push(c);
      } catch (err) {
        comErro.push({ grupo: c.grupo, projeto: c.projeto, erro: String(err), _row: c._row });
      }
    });

    // As chamadas vão em lotes paralelos: em sequência, 40 grupos passariam de
    // 6 minutos só de espera de rede e estourariam o limite de execução.
    var respostas = chamarIALote_(prontos.map(function (c) {
      return promptAvaliacao_(c.texto, criteriaText, c.projeto, c.grupo, c.feedback);
    }));

    var avaliados = [];
    prontos.forEach(function (c, idx) {
      var resposta = respostas[idx];
      if (!resposta || !resposta.ok) {
        comErro.push({
          grupo: c.grupo, projeto: c.projeto,
          erro: (resposta && resposta.error) || 'Sem resposta da IA', _row: c._row
        });
        return;
      }
      try {
        var avaliacao = normalizarAvaliacao_(resposta.data);
        avaliados.push({
          grupo: c.grupo, projeto: c.projeto, feedback: c.feedback,
          scoreIndividual: avaliacao.score, justificativaIndividual: avaliacao.justificativa,
          _row: c._row
        });
      } catch (err) {
        comErro.push({ grupo: c.grupo, projeto: c.projeto, erro: String(err), _row: c._row });
      }
    });

    // Grava as notas individuais antes da recalibração: se a recalibração falhar ou
    // a execução for interrompida, o trabalho já feito não é perdido.
    withSheetLock_(function () {
      avaliados.forEach(function (a) {
        sheet.getRange(a._row, COL.SCORE_INDIVIDUAL, 1, 2).setValues([[
          a.scoreIndividual, a.justificativaIndividual
        ]]);
      });
    });

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

    // Completa o bloco de avaliação. As notas individuais já foram gravadas acima,
    // então aqui só entram o resultado da recalibração e o status.
    withSheetLock_(function () {
      ranqueados.forEach(function (r) {
        sheet.getRange(r._row, COL.STATUS).setValue('avaliado');
        sheet.getRange(r._row, COL.SCORE_FINAL, 1, 4).setValues([[
          r.scoreFinal, r.comentarioRecalibracao, r.selecionado, new Date()
        ]]);
      });
      comErro.forEach(function (r) {
        sheet.getRange(r._row, COL.STATUS).setValue('erro');
        sheet.getRange(r._row, COL.JUSTIFICATIVA).setValue(r.erro);
        sheet.getRange(r._row, COL.SELECIONADO).setValue(false);
        sheet.getRange(r._row, COL.TIMESTAMP_AVALIACAO).setValue(new Date());
      });
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
    releaseEvalFlag_();
  }
}

/** Prompt da nota individual secreta de um projeto (nunca exposta ao grupo). */
function promptAvaliacao_(textoProjeto, criteriaText, projectName, groupName, feedbackTexto) {
  return [
    'Você é um avaliador de um hackathon corporativo. Avalie o projeto descrito abaixo',
    'seguindo ESTRITAMENTE os critérios. Esta nota é interna e nunca será mostrada ao grupo.',
    '',
    '=== CRITÉRIOS DE AVALIAÇÃO ===',
    criteriaText,
    '=== FIM DOS CRITÉRIOS ===',
    '',
    'Grupo: ' + groupName,
    'Projeto: ' + (projectName || '(não informado)'),
    '',
    '=== DESCRITIVO DO PROJETO ===',
    textoProjeto,
    '=== FIM DO DESCRITIVO ===',
    '',
    '=== ANÁLISE QUALITATIVA JÁ FEITA SOBRE ESTE PROJETO ===',
    String(feedbackTexto || '(sem análise qualitativa registrada)'),
    '=== FIM DA ANÁLISE QUALITATIVA ===',
    '',
    'Considere tanto o descritivo quanto a análise qualitativa acima. Dê peso especial ao',
    'POTENCIAL DE IMPACTO do projeto segundo os critérios do evento.',
    '',
    'Responda APENAS com um JSON no formato exato, sem markdown:',
    '{"score": <numero de 0 a 100>, "justificativa": "<explicação em até 3 frases>"}'
  ].join('\n');
}

function normalizarAvaliacao_(parsed) {
  var score = Number(parsed.score);
  if (isNaN(score)) throw new Error('A IA não devolveu uma nota numérica');
  return { score: score, justificativa: String(parsed.justificativa || '') };
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

  var parsed = chamarIA_(prompt);
  if (!parsed.ranking || !Array.isArray(parsed.ranking)) {
    throw new Error('Resposta de recalibração sem "ranking" válido');
  }
  return parsed.ranking;
}

var URL_OPENROUTER = 'https://openrouter.ai/api/v1/chat/completions';

/** Monta a requisição de texto ao OpenRouter, no formato aceito por fetch e fetchAll. */
function montarRequestIA_(prompt) {
  var apiKey = getProp_('OPENROUTER_API_KEY');
  if (!apiKey) throw new Error('OPENROUTER_API_KEY não configurada nas Propriedades do Script');
  var model = getProp_('OPENROUTER_MODEL', 'google/gemini-3.7-flash');

  return {
    url: URL_OPENROUTER,
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': 'Avaliador Hackathon'
    },
    payload: JSON.stringify({
      model: model,
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.2
    }),
    muteHttpExceptions: true
  };
}

/** Interpreta uma resposta HTTP do OpenRouter e devolve o JSON do modelo. */
function lerRespostaIA_(response) {
  var json = JSON.parse(response.getContentText());
  if (!json.choices || !json.choices.length) {
    throw new Error('Resposta da IA sem conteúdo: ' + response.getContentText().substring(0, 300));
  }
  // Usa apenas 'content': em modelos com raciocínio, 'reasoning' costuma conter
  // rascunhos de JSON que não são a resposta final.
  return extrairJson_(String(json.choices[0].message.content || ''));
}

/**
 * Chamada única ao OpenRouter, com retry em 429 e erros 5xx.
 * Retorna o JSON já parseado da resposta do modelo.
 */
function chamarIA_(prompt) {
  var request = montarRequestIA_(prompt);
  var lastError = null;

  for (var attempt = 0; attempt < 3; attempt++) {
    var response = UrlFetchApp.fetch(request.url, request);
    var code = response.getResponseCode();

    if (code === 200) return lerRespostaIA_(response);

    if (code === 429 || code >= 500) {
      lastError = 'HTTP ' + code + ': ' + response.getContentText().substring(0, 300);
      Utilities.sleep(2000 * (attempt + 1));
      continue;
    }

    throw new Error('OpenRouter HTTP ' + code + ': ' + response.getContentText().substring(0, 300));
  }
  throw new Error(lastError || 'Falha ao chamar o OpenRouter após 3 tentativas');
}

/**
 * Executa vários prompts em lotes paralelos via fetchAll, reenviando apenas os que
 * falharam com 429 ou erro 5xx. Devolve um array na mesma ordem dos prompts, com
 * {ok: true, data} ou {ok: false, error} em cada posição — uma falha isolada nunca
 * interrompe as demais.
 */
function chamarIALote_(prompts) {
  var resultados = new Array(prompts.length);
  var pendentes = prompts.map(function (_, idx) { return idx; });

  for (var tentativa = 0; tentativa < 3 && pendentes.length > 0; tentativa++) {
    var reenviar = [];

    for (var inicio = 0; inicio < pendentes.length; inicio += TAMANHO_LOTE_IA) {
      var fatia = pendentes.slice(inicio, inicio + TAMANHO_LOTE_IA);
      var requests = fatia.map(function (idx) { return montarRequestIA_(prompts[idx]); });
      var responses = UrlFetchApp.fetchAll(requests);

      responses.forEach(function (response, k) {
        var idx = fatia[k];
        var code = response.getResponseCode();

        if (code === 200) {
          try {
            resultados[idx] = { ok: true, data: lerRespostaIA_(response) };
          } catch (err) {
            resultados[idx] = { ok: false, error: String(err) };
          }
          return;
        }

        resultados[idx] = {
          ok: false,
          error: 'OpenRouter HTTP ' + code + ': ' + response.getContentText().substring(0, 300)
        };
        if (code === 429 || code >= 500) reenviar.push(idx);
      });
    }

    pendentes = reenviar;
    if (pendentes.length > 0) Utilities.sleep(3000 * (tentativa + 1));
  }

  return resultados;
}

/**
 * Extrai o objeto JSON da resposta do modelo. Modelos menores costumam embrulhar o
 * JSON em cercas de markdown ou em texto explicativo, então isolamos o primeiro
 * objeto balanceado em vez de confiar na resposta vir limpa.
 */
function extrairJson_(texto) {
  var limpo = texto.replace(/```json/gi, '').replace(/```/g, '').trim();
  try {
    return JSON.parse(limpo);
  } catch (err) { /* tenta isolar o objeto abaixo */ }

  var inicio = limpo.indexOf('{');
  if (inicio === -1) {
    throw new Error('A IA não devolveu JSON: ' + limpo.substring(0, 300));
  }

  var profundidade = 0;
  var emString = false;
  var escapando = false;
  for (var i = inicio; i < limpo.length; i++) {
    var ch = limpo.charAt(i);
    if (escapando) { escapando = false; continue; }
    if (ch === '\\') { escapando = true; continue; }
    if (ch === '"') { emString = !emString; continue; }
    if (emString) continue;
    if (ch === '{') profundidade++;
    else if (ch === '}') {
      profundidade--;
      if (profundidade === 0) {
        var candidato = limpo.substring(inicio, i + 1);
        try {
          return JSON.parse(candidato);
        } catch (err2) {
          throw new Error('A IA não devolveu JSON válido: ' + candidato.substring(0, 300));
        }
      }
    }
  }
  throw new Error('A IA devolveu JSON incompleto: ' + limpo.substring(0, 300));
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

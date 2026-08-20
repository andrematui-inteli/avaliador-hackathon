(function () {
  const btnAvaliar = document.getElementById('btn-avaliar');
  const btnAtualizar = document.getElementById('btn-atualizar');
  const btnDiagnostico = document.getElementById('btn-diagnostico');
  const statusEl = document.getElementById('status');
  const tabela = document.getElementById('tabela-resultados');
  const corpoTabela = document.getElementById('corpo-tabela');
  const msgVazio = document.getElementById('msg-vazio');

  function setStatus(kind, message) {
    if (!message) {
      statusEl.className = 'status-msg';
      return;
    }
    statusEl.className = 'status-msg visible ' + kind;
    statusEl.innerHTML = message;
  }

  function badge(row) {
    if (row.status === 'erro' || row.erro) return '<span class="badge erro">erro</span>';
    if (row.status === 'avaliado') return '<span class="badge ok">avaliado</span>';
    return '<span class="badge pendente">pendente</span>';
  }

  function renderResultados(resultados) {
    if (!resultados || resultados.length === 0) {
      tabela.style.display = 'none';
      msgVazio.style.display = 'block';
      return;
    }
    msgVazio.style.display = 'none';
    tabela.style.display = 'table';
    corpoTabela.innerHTML = resultados.map((row, i) => {
      const linhaClasse = row.selecionado ? 'selecionado' : (row.status === 'erro' || row.erro ? 'erro' : '');
      const fmt = (v) => (v === null || v === undefined || v === '' ? '-' : v);
      return `
        <tr class="${linhaClasse}">
          <td>${i + 1}${row.selecionado ? ' 🏆' : ''}</td>
          <td>${escapeHtml(row.grupo)}</td>
          <td>${escapeHtml(row.projeto || '-')}</td>
          <td>${badge(row)}</td>
          <td>${fmt(row.score_individual)}</td>
          <td><strong>${fmt(row.score_final)}</strong></td>
          <td>${escapeHtml(row.justificativa_individual || '-')}</td>
          <td>${escapeHtml(row.comentario_recalibracao || '-')}</td>
        </tr>`;
    }).join('');
  }

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = String(str);
    return div.innerHTML;
  }

  async function callApi(action) {
    const response = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action })
    });
    const result = await response.json();
    if (!result.ok) throw new Error(result.error || 'Falha desconhecida');
    return result;
  }

  async function carregarResultados() {
    setStatus('info', '<span class="spinner"></span>Carregando resultados...');
    try {
      const { resultados } = await callApi('results');
      renderResultados(resultados);
      setStatus('', '');
    } catch (err) {
      setStatus('error', 'Não foi possível carregar resultados: ' + err.message);
    }
  }

  btnAvaliar.addEventListener('click', async () => {
    const confirmado = confirm(
      'Isso vai reavaliar TODOS os grupos que enviaram projeto, usando os critérios ' +
      'atuais do documento. Pode levar alguns minutos. Continuar?'
    );
    if (!confirmado) return;

    btnAvaliar.disabled = true;
    btnAtualizar.disabled = true;
    setStatus('info', '<span class="spinner"></span>Avaliando grupos, isso pode levar alguns minutos...');
    try {
      const { resultados, top_x } = await callApi('evaluate');
      renderResultados(resultados);
      const recomendados = resultados.filter((r) => r.selecionado).length;
      setStatus('success', `Avaliação concluída para ${resultados.length} grupo(s). ` +
        `${recomendados} grupo(s) recomendado(s) para apresentar${top_x ? ` (meta: top ${top_x})` : ''}.`);
    } catch (err) {
      setStatus('error', 'Erro ao avaliar: ' + err.message);
    } finally {
      btnAvaliar.disabled = false;
      btnAtualizar.disabled = false;
    }
  });

  btnAtualizar.addEventListener('click', carregarResultados);

  btnDiagnostico.addEventListener('click', async () => {
    btnDiagnostico.disabled = true;
    setStatus('info', '<span class="spinner"></span>Verificando configuração...');
    try {
      const response = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'diagnostico' })
      });
      const result = await response.json();
      const linhas = (result.checagens || []).map((c) =>
        `${c.ok ? '✅' : '❌'} <strong>${escapeHtml(c.item)}</strong>: ${escapeHtml(c.detalhe)}`
      ).join('<br>');
      setStatus(result.ok ? 'success' : 'error',
        (result.ok ? 'Configuração completa.' : `${result.falhas} problema(s) encontrado(s).`) +
        '<br><br>' + linhas);
    } catch (err) {
      setStatus('error', 'Não foi possível verificar: ' + err.message);
    } finally {
      btnDiagnostico.disabled = false;
    }
  });

  if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL.includes('COLE_AQUI')) {
    setStatus('error', 'A aplicação ainda não foi configurada (config.js).');
  } else {
    carregarResultados();
  }
})();

(function () {
  const btnAvaliar = document.getElementById('btn-avaliar');
  const btnAtualizar = document.getElementById('btn-atualizar');
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
      return `
        <tr class="${linhaClasse}">
          <td>${i + 1}${row.selecionado ? ' 🏆' : ''}</td>
          <td>${escapeHtml(row.grupo)}</td>
          <td>${escapeHtml(row.projeto || '-')}</td>
          <td>${badge(row)}</td>
          <td>${row.score === null || row.score === undefined ? '-' : row.score}</td>
          <td>${escapeHtml(row.justificativa || '-')}</td>
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
    return result.resultados || [];
  }

  async function carregarResultados() {
    setStatus('info', '<span class="spinner"></span>Carregando resultados...');
    try {
      const resultados = await callApi('results');
      renderResultados(resultados);
      setStatus('', '');
    } catch (err) {
      setStatus('error', 'Não foi possível carregar resultados: ' + err.message);
    }
  }

  btnAvaliar.addEventListener('click', async () => {
    const confirmado = confirm(
      'Isso vai avaliar TODOS os grupos com submissão pendente ou já avaliada, ' +
      'usando os critérios atuais do documento e a chave do Gemini configurada. ' +
      'Pode levar alguns minutos. Continuar?'
    );
    if (!confirmado) return;

    btnAvaliar.disabled = true;
    btnAtualizar.disabled = true;
    setStatus('info', '<span class="spinner"></span>Avaliando grupos, isso pode levar alguns minutos...');
    try {
      const resultados = await callApi('evaluate');
      renderResultados(resultados);
      setStatus('success', `Avaliação concluída para ${resultados.length} grupo(s).`);
    } catch (err) {
      setStatus('error', 'Erro ao avaliar: ' + err.message);
    } finally {
      btnAvaliar.disabled = false;
      btnAtualizar.disabled = false;
    }
  });

  btnAtualizar.addEventListener('click', carregarResultados);

  if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL.includes('COLE_AQUI')) {
    setStatus('error', 'A aplicação ainda não foi configurada (config.js).');
  } else {
    carregarResultados();
  }
})();

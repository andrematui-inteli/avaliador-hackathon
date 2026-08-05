(function () {
  const MAX_SIZE_MB = 15;

  const form = document.getElementById('form-submissao');
  const btn = document.getElementById('btn-enviar');
  const statusEl = document.getElementById('status');

  function setStatus(kind, message) {
    statusEl.className = 'status-msg visible ' + kind;
    statusEl.innerHTML = message;
  }

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result;
        resolve(result.substring(result.indexOf(',') + 1));
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  form.addEventListener('submit', async (event) => {
    event.preventDefault();

    const grupo = document.getElementById('grupo').value.trim();
    const projeto = document.getElementById('projeto').value.trim();
    const arquivoInput = document.getElementById('arquivo');
    const file = arquivoInput.files[0];

    if (!grupo || !file) {
      setStatus('error', 'Preencha o nome do grupo e selecione o arquivo PDF.');
      return;
    }
    if (file.type !== 'application/pdf') {
      setStatus('error', 'O arquivo precisa estar no formato PDF.');
      return;
    }
    if (file.size > MAX_SIZE_MB * 1024 * 1024) {
      setStatus('error', `O arquivo excede o limite de ${MAX_SIZE_MB}MB.`);
      return;
    }
    if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL.includes('COLE_AQUI')) {
      setStatus('error', 'A aplicação ainda não foi configurada (config.js). Avise a organização.');
      return;
    }

    btn.disabled = true;
    setStatus('info', '<span class="spinner"></span>Enviando projeto...');

    try {
      const base64Data = await fileToBase64(file);
      const response = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'submit',
          grupo,
          projeto,
          filename: file.name,
          data: base64Data
        })
      });

      const result = await response.json();
      if (!result.ok) throw new Error(result.error || 'Falha desconhecida');

      setStatus('success', `Projeto do grupo <strong>${grupo}</strong> enviado com sucesso!`);
      form.reset();
    } catch (err) {
      setStatus('error', 'Não foi possível enviar o projeto: ' + err.message);
    } finally {
      btn.disabled = false;
    }
  });
})();

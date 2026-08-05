# avaliador-hackathon

Aplicação para avaliar por IA os descritivos de projeto (PDF) de uma hackathon e selecionar os melhores grupos.

- Frontend estático (GitHub Pages): [`index.html`](index.html) (envio de projetos pelos grupos) e [`admin.html`](admin.html) (painel do organizador, link não divulgado).
- Backend: [`apps-script/Code.gs`](apps-script/Code.gs), publicado como Google Apps Script Web App. Guarda a chave do Gemini, os critérios de avaliação e os arquivos enviados — nada disso fica exposto no site estático.

Configuração completa (Sheet, Drive, Doc de critérios, deploy do Apps Script e GitHub Pages) descrita no plano de implementação.
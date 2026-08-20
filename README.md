# avaliador-hackathon

Aplicação para avaliar por IA os descritivos de projeto de uma hackathon, dar feedback
imediato aos grupos e recomendar os melhores para apresentar à banca.

## Links

- **Envio de projetos** (divulgar aos grupos): https://andrematui-inteli.github.io/avaliador-hackathon/
- **Painel do organizador** (não divulgar): https://andrematui-inteli.github.io/avaliador-hackathon/admin.html

## Como funciona

```
Grupos                            Organização
  │                                   │
  ▼                                   ▼
index.html                       admin.html
(envia PDF, recebe feedback)     (dispara avaliação, vê ranking)
  │                                   │
  └───────────────┬───────────────────┘
                  ▼
        Apps Script Web App
        ├── Drive: PDFs enviados
        ├── Sheet: submissões, feedback, notas e ranking
        ├── Doc:  critérios de avaliação (editável a qualquer momento)
        └── OpenRouter: chamadas à IA
```

O GitHub Pages serve apenas arquivos estáticos, então a chave de API nunca pode estar
no site — qualquer visitante a leria no DevTools. Por isso toda lógica sensível roda no
Apps Script, que guarda a chave em Propriedades do Script (no servidor do Google) e
pode ter esses valores trocados sem mexer no código nem reimplantar.

### Fluxo de envio

1. O grupo envia o PDF em `index.html`
2. O Apps Script salva o arquivo no Drive e registra a linha do grupo na planilha
3. O Drive converte o PDF em texto (com OCR quando o PDF é escaneado)
4. A IA gera um **feedback qualitativo** — sem nota, sem percentual, sem comparação com
   outros grupos — e o grupo vê na hora
5. Reenvio do mesmo nome de grupo substitui a submissão anterior

### Fluxo de avaliação

1. A organização clica em "Avaliar todos os grupos" em `admin.html`
2. Cada projeto recebe uma **nota interna de 0 a 100**, considerando o descritivo e o
   feedback qualitativo já gerado. Essa nota nunca é mostrada aos grupos
3. Uma segunda chamada compara **todos os grupos ao mesmo tempo** e recalibra o ranking,
   corrigindo distorções de notas dadas em análises isoladas
4. Os `TOP_X` grupos com maior potencial de impacto são marcados como recomendados

## Configuração

### Propriedades do Script

Em `script.google.com` → ⚙️ Configurações do projeto → Propriedades do script:

| Propriedade | Valor | Obrigatória |
|---|---|---|
| `OPENROUTER_API_KEY` | chave de openrouter.ai | sim |
| `OPENROUTER_MODEL` | `google/gemini-3.7-flash` | não (esse é o default) |
| `TOP_X` | quantos grupos recomendar, ex. `10` | não (default 10) |
| `CRITERIA_DOC_ID` | ID do Doc de critérios | sim |
| `DRIVE_FOLDER_ID` | ID da pasta dos PDFs | sim |
| `SHEET_ID` | ID da planilha | sim |

O ID de cada recurso do Google é o trecho da URL entre `/d/` (ou `/folders/`) e a barra
seguinte.

### Implantação

"Implantar" → "Gerenciar implantações" → ✏️ → em **Versão** escolher **"Nova versão"** →
"Implantar". Usar "Gerenciar implantações" mantém a mesma URL; criar uma implantação nova
gera outra URL e exige atualizar o `config.js`.

Atenção: se em "Versão" for escolhido um número já existente em vez de "Nova versão", o
Apps Script continua servindo o código antigo mesmo com o editor mostrando o novo.

### Critérios de avaliação

O texto do Doc apontado por `CRITERIA_DOC_ID` é lido a cada avaliação, então pode ser
editado livremente até o dia do evento. Há um rascunho de partida em
[`CRITERIOS-EXEMPLO.md`](CRITERIOS-EXEMPLO.md).

## Números medidos

Testes reais com `google/gemini-3.7-flash` em agosto de 2026:

| Medida | Valor |
|---|---|
| Latência de uma chamada | ~9 s (o modelo raciocina antes de responder e isso não pode ser desligado) |
| Lote de 10 chamadas em paralelo | ~20 s, 10/10 sucesso, sem rate limit |
| Custo de uma avaliação | ~US$ 0,0013 |
| Custo de um feedback | ~US$ 0,0026 |
| **Custo estimado do evento (40 grupos)** | **~US$ 0,20** |

A avaliação usa `UrlFetchApp.fetchAll` em lotes de 10, então 40 grupos levam cerca de
1 minuto em vez dos mais de 6 minutos que levariam em sequência.

## Limites conhecidos

- **Apenas texto é avaliado:** imagens, diagramas e o layout de tabelas do PDF não são
  considerados, porque o PDF é convertido em texto antes de ir para a IA.
- **PDF sem texto extraível** (composto só de imagens sem OCR possível) falha na geração
  do feedback. A submissão ainda é registrada e o erro aparece na planilha.
- **Tempo de execução do Apps Script:** 6 min em contas Google pessoais, 30 min em contas
  Workspace. Com os lotes paralelos a avaliação fica em ~1 min, bem dentro do limite.
- **Cota de modelos gratuitos do OpenRouter** (só relevante se usar um modelo `:free`):
  50 requisições/dia, ou 1.000/dia se a conta já comprou US$ 10 em créditos. Com 40 grupos
  o consumo passa de 81 requisições, então o limite de 50/dia não é suficiente.
- **Notas individuais ficam agrupadas:** em testes, projetos distintos receberam notas
  entre 70 e 81. É por isso que existe a etapa de recalibração — sem ela, a ordem dos
  primeiros colocados seria pouco confiável.

## Arquivos

| Arquivo | Papel |
|---|---|
| [`index.html`](index.html) / [`submit.js`](submit.js) | página de envio e exibição do feedback |
| [`admin.html`](admin.html) / [`admin.js`](admin.js) | painel do organizador |
| [`apps-script/Code.gs`](apps-script/Code.gs) | backend (copiar para o editor do Apps Script) |
| [`config.js`](config.js) | URL pública do Web App |
| [`style.css`](style.css) | estilos compartilhados |

O `Code.gs` deste repositório é a fonte da verdade; o editor do Apps Script é uma cópia
que precisa ser atualizada manualmente a cada mudança.

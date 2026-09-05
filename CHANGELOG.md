# Changelog — cep-ufv

## [1.7.1] — 2026-09-05
### Corrigido
- DOCX grava o total real no campo NUMPAGES: prévia do Google Docs mostra 'Página X de Y' correto (o Word recalcula ao abrir).

## [1.7.0] — 2026-09-05
### Adicionado
- Necessidade por caso: cada documento avaliado (TCLE↔dispensa excludentes; TALE/responsável só com menores; anuência SÓ com entrevistas/similares dentro de instituição; sigilo só com dado sigiloso/prontuário/secundário); N/A justificado no checklist, MANIFESTO e --validar-apenas.

## [1.6.1] — 2026-09-05
### Corrigido
- Paginação total em todos os documentos: docx com campos PAGE+NUMPAGES válidos e pdf com 'Página X de Y' (duas passagens, contagem via total real).

## [1.6.0] — 2026-09-05
### Adicionado
- Coleta online: doc 13-Coleta-Online-GoogleForms + criar-form-google.gs (Apps Script monta o Forms sozinho, consentimento obrigatório na Seção 1).
- Carta-Resposta só com --com-carta (sob demanda, quando há pendência).
- Documentos datados automaticamente (Cidade, DD de mês de AAAA); previsão de parecer pela data atual + calendário de reuniões (previsao.txt).
- Blocos de assinatura ABNT refeitos (linha + identificação + nome legível).

## [1.4.0] — 2026-09-05
### Corrigido
- Documentos finais limpos: títulos oficiais sem parênteses de modelo, sem menções à skill, sem OBS de template, sem `[PREENCHER]`; varredura fail-closed ampliada.
- Sigilo situacional (prontuário vs. bases); dispensa com texto limpo.
### Formatação
- ABNT/UFV via skill ufv-abnt (NBR 14724/2025): Arial 12, 1,5, margens 3/2, justificado, paginação — docx e pdf espelhados.

## [1.3.0] — 2026-09-05
### Adicionado
- Liberdade situacional dentro do checklist: cláusulas condicionais (vulneráveis, acompanhamento, saúde mental, SUS, biológico, voz/imagem) nos TCLEs.
- Docs novos: 00-Projeto-de-Pesquisa (8 seções) e 08b-Declaracao-Financiamento (próprio/externo).
- Modo FINAL padrão: sem `[PREENCHER]`, sem avisos internos; incompleto aborta (fail-closed); `--rascunho` só p/ revisão.

## [1.1.0] — 2026-09-05
### Adicionado
- `12-Roteiro-Submissao` (ETAPA 0–8) verificado em projeto-novo, pendência, cronograma e normativas.
- Pasta `modelos/` com os 9 oficiais (PDFs+ODTs de 05/09/2026).
- Estrutura GitHub completa (README, INSTALL, docs, evals, tests, CI).
### Mantido
- Trava de template: verbatim legal, `[PREENCHER]` + `pendencias.txt`, par docx+pdf fail-closed, fingerprints, 27 itens TCLE.

## [1.0.0] — 2026-09-05
- Versão inicial template-locked: 11 documentos x2 formatos (01–11) + MANIFESTO + pendências.

---
name: cep-ufv
description: >
  Gera o pacote documental CEP/UFV travado aos templates oficiais de
  https://cep.ufv.br/modelos/ a partir de um projeto.json. Preenche SOMENTE
  os campos variáveis dentro do texto verbatim dos modelos (TCLE adulto,
  TCLE responsável, TALE assentimento, anuência institucional, sigilo,
  cronograma, orçamento, dispensa, carta-resposta, checklist, roteiro passo a passo).
  NUNCA cria
  seção, parágrafo ou documento fora dos templates. Gera sempre o par
  .docx editável + .pdf idêntico por documento, com validação fail-closed.
  Os 9 arquivos-modelo oficiais ficam embutidos em `modelos/` (PDFs+ODTs
  baixados de cep.ufv.br em 05/09/2026). Use quando o usuário pedir documentos do CEP, TCLE, TALE, anuência,
  sigilo, cronograma CEP, submissão Plataforma Brasil, roteiro de submissão.
---

# CEP-UFV — Gerador travado aos templates oficiais

> **TRAVA ABSOLUTA DE TEMPLATE (inegociável):**
> 1. Esta skill **SÓ preenche campos** dentro do texto verbatim dos modelos oficiais
>    de https://cep.ufv.br/modelos/ (versões 2026 + checklist 2021 + ODTs 2016/2019).
> 2. É **PROIBIDO** criar, reescrever, resumir, modernizar ou acrescentar qualquer
>    parágrafo, seção ou documento fora desses modelos.
> 3. Campo sem dado = marcador `[PREENCHER: <campo>]` + entrada em `pendencias.txt`.
>    **NUNCA inventar** dado, frase legal ou número.
> 4. Todo documento sai **sempre em par idêntico**: `.docx` (editável) + `.pdf`.
>    Se um dos dois falhar, a geração inteira falha (fail-closed).
> 5. Base legal fixa (não editável pelo agente): Lei 14.874/2024, Decreto 12.651/2025,
>    Res. CNS 466/2012, Res. CNS 510/2016, Norma Operacional 01/2013, LGPD 13.709/2018,
>    + específicas quando aplicável (441/2011 biobanco, 304/2000 indígenas, 340/2004
>    genética, 674/2022 tipificação). Fonte: https://cep.ufv.br/normativas/.

## 0. Fontes oficiais congeladas nesta skill

| Documento | URL oficial | Versão na skill |
|---|---|---|
| Modelo de TCLE | https://cep.ufv.br/wp-content/uploads/2026/05/TCLE.pdf | 2026, 3p |
| Modelo de TCLE Responsável | https://cep.ufv.br/wp-content/uploads/2026/05/TCLE-Responsavel.pdf | 2026, 3p |
| Modelo de Termo de Assentimento (TALE) | https://cep.ufv.br/wp-content/uploads/2026/05/Termo-de-Assentimento.pdf | 2026, 3p |
| Modelo de Anuência Institucional | https://cep.ufv.br/wp-content/uploads/2026/05/Modelo-de-Autorizacao-Institucional.pdf | 2026, 1p |
| Modelo de Cronograma | https://cep.ufv.br/wp-content/uploads/2026/05/Modelo-de-cronograma.pdf | 2026/2027, 1p |
| Modelo de Relatório Final | https://cep.ufv.br/wp-content/uploads/2016/06/Modelo-Relatório-Final.pdf | 2016, 2p |
| Modelo de Carta Resposta | https://cep.ufv.br/wp-content/uploads/2016/05/Modelo-Carta-Resposta.odt | 2016 ODT |
| Modelo de Sigilo e Confidencialidade | https://cep.ufv.br/wp-content/uploads/2019/07/Modelo-de-termo-de-sigilo-e-confidencialidade.odt | 2019 ODT |
| Check List | https://cep.ufv.br/wp-content/uploads/2021/02/Ckeck-List-CEP-UFV.pdf | 2021 |
| Rito de submissão | https://cep.ufv.br/projeto-novo/ | 10 docs obrigatórios |
| Legislação | https://cep.ufv.br/normativas/ | Lei 14.874/24 + Decreto 12.651/25 + Res. 466/510 |

Se o site atualizar um modelo, **não adivinhe o novo texto**: avise o usuário,
mantenha a versão congelada e marque `[VALIDAÇÃO PENDENTE — template pode estar desatualizado]`.

> **Modelos embutidos:** a pasta `modelos/` dentro desta skill contém cópias fiéis dos
> 9 arquivos oficiais (TCLE.pdf, TCLE-Responsavel.pdf, Termo-de-Assentimento.pdf,
> Modelo-de-Autorizacao-Institucional.pdf, Modelo-de-cronograma.pdf,
> Modelo-Relatório-Final.pdf, Modelo-Carta-Resposta.odt,
> Modelo-de-termo-de-sigilo-e-confidencialidade.odt, Ckeck-List-CEP-UFV.pdf).
> O `run.py` tem liberdade situacional DENTRO do checklist: ativa/desativa cláusulas
> conforme o projeto (menores → TALE+TCLE-resp; biológico → Res. 441/2011; voz/imagem;
> vulneráveis; acompanhamento; saúde mental → não-diagnóstico; SUS → não confunde com
> assistência). Gera TODOS os documentos do checklist: 00-Projeto, 01, 02, 03, 04, 05,
> 06, 07, 08-Dispensa, 08b-Financiamento, 09-Carta, 10-Checklist, 11-Orientações e
> 12-Roteiro. Modo FINAL (padrão): sem `[PREENCHER]`, sem avisos internos — incompleto
> aborta; use `--rascunho` só para revisão interna (NÃO submetível).
> O documento `12-Roteiro-Submissao` é guia adicional permitido (além do
> `11-Orientacoes`), rotulado `NÃO SUBMETER`, com o passo a passo verificado em
> https://cep.ufv.br/projeto-novo/, https://cep.ufv.br/pendencia/ e
> https://cep.ufv.br/cronograma/ (reuniões 2026 + regra dia 20 + 30 dias de pendência).

## 1. Quando acionar

- "gerar documentos do CEP", "fazer TCLE", "TALE", "anuência", "sigilo",
  "cronograma CEP", "checklist CEP", "pacote Plataforma Brasil", "submeter ao CEP/UFV".
- Sempre que o usuário citar `cep-ufv`, `documentos CEP`, `TCLE UFV`.

## 2. Fluxo obrigatório do agente (o JSON é tarefa DO AGENTE, nunca do usuário)

```
1. ENTREVISTAR -> o AGENTE conduz a coleta (perguntas objetivas ou a partir de
              texto livre/projeto existente) e monta sozinho o projeto.json completo
              (schema em projeto-exemplo.json: dados do pesquisador, introdução,
              objetivo, justificativa, procedimentos, instrumentos, referências,
              local, tempo, riscos, medidas, benefícios, ressarcimento, retorno,
              contato de retirada + flags: menores, biologico, virtual, gravacao,
              vulneraveis, acompanhamento, saude_mental, sus, financiamento).
              Faltou dado? O AGENTE pergunta ao usuário. A skill NÃO inventa e o
              modo FINAL aborta se incompleto.
2. GERAR    -> o AGENTE roda: python3 run.py projeto.json --out outputs-cep-<slug>/
              (modo FINAL, sem marcações). --rascunho só p/ revisão interna;
              --validar-apenas p/ listar faltas.
3. VALIDAR  -> run.py já valida: 14 pares docx+pdf, checklist 27 itens TCLE,
              zero marcações, manifesto. Se exit != 0, o AGENTE coleta o que falta
              e regenera. NÃO entrega pacote com erro.
4. ENTREGAR -> listar arquivos gerados + pendencias.txt + MANIFESTO.txt.
              Avisar: folha de rosto sai da Plataforma Brasil (não geramos),
              coleta só após aprovação, pesquisador responsável = orientador.
```

O usuário NUNCA precisa abrir ou editar JSON: basta responder às perguntas do agente
(ou colar o projeto/texto base) e o agente devolve o pacote pronto.

Nunca edite `run.py` para "melhorar" texto legal. Só preencha via `projeto.json`.

## 3. Entrada: projeto.json (schema mínimo)

```json
{
  "titulo": "string OBRIGATÓRIA",
  "pesquisador": {"nome": "", "departamento": "", "endereco": "", "telefone": "", "email": "", "instituicao": "Universidade Federal de Viçosa"},
  "orientador_e_responsavel": "nome do orientador (= pesquisador responsável p/ CEP/UFV)",
  "objetivo": "", "justificativa": "", "procedimentos": "",
  "local": "", "tempo_estimado": "", "riscos": "", "medidas_risco": "",
  "beneficios": "", "ressarcimento": "texto ou 'Não haverá custos...'",
  "retorno": "como resultados serão devolvidos",
  "contato_retirada": "email/telefone para retirar consentimento",
  "participante_nome": "[PREENCHER no ato da assinatura se individual]",
  "usa_menores": false, "usa_biologico": false,
  "biologico_lab": "", "biologico_futuro": "",
  "ambiente_virtual": false, "gravacao": false,
  "instituicao_local": {"nome": "", "representante": "", "cargo": ""},
  "cronograma": [{"etapa": "", "periodo": ""}],
  "orcamento": [{"item": "", "valor": "", "fonte": ""}],
  "financiamento_proprio": true,
  "dispensa_tcle": {"solicitar": false, "justificativa": ""},
  "instrumentos": ["questionário anexo X"],
  "data_local": "Viçosa, ___ de ______ de 20___"
}
```

Campos ausentes → modo FINAL aborta e lista o que o AGENTE deve coletar
(`--validar-apenas`); `--rascunho` gera rascunho interno NÃO submetível.
A skill nunca inventa dado.

## 4. Saída: pacote fechado (par docx+pdf por documento)

```
outputs-cep-<slug>/            (pasta citada ao pesquisador — é lá que está tudo)
  00-Projeto-de-Pesquisa.docx + .pdf
  01-TCLE-Adulto.docx + .pdf
  02-TCLE-Responsavel.docx + .pdf
  03-TALE-Assentimento.docx + .pdf
  04-Anuencia-Institucional.docx + .pdf
  05-Sigilo-Confidencialidade.docx + .pdf
  06-Cronograma.docx + .pdf
  07-Orcamento.docx + .pdf
  08-Dispensa-TCLE.docx + .pdf        (sem solicitação, é declaração limpa de não-dispensa)
  08b-Declaracao-Financiamento.docx + .pdf
  10-Checklist-Conformidade.docx + .pdf   (USO INTERNO, não anexar)
  11-Orientacoes-Plataforma-Brasil.docx + .pdf  (guia, não anexar)
  12-Roteiro-Submissao.docx + .pdf    (passo a passo + previsão, guia, não anexar)
  [+ 13-Coleta-Online-GoogleForms.docx + .pdf + criar-form-google.gs — SÓ se coleta online]
  [+ 09-Carta-Resposta.docx + .pdf — SÓ com --com-carta, quando houver pendência; o agente gera na hora]
  MANIFESTO.txt + pendencias.txt (zerada) + previsao.txt
```

Regras:
- Nenhum arquivo fora dessa lista. Proibido "criar um TCLE simplificado/criativo".
- `10-Checklist`, `11-Orientacoes` e `12-Roteiro` são guias de USO INTERNO
  (não anexar na PB).
- Documentos finais 00–09 saem LIMPOS: títulos oficiais sem parênteses de modelo,
  sem menções à skill, sem OBS de template, sem `[PREENCHER]` — prontos para submissão.
- Folha de rosto: **não geramos** (sai da Plataforma Brasil). O guia explica assinaturas:
  responsável = orientador; proponente = chefia dept. (graduação/IC/extensão) ou
  coordenação PPG (pós); superior hierárquico assina se acúmulo; patrocinador
  externo preenche+assina, próprio suprime campo.
- Formatação ABNT/UFV (skill `ufv-abnt`, NBR 14724/2025): A4, Arial 12 preto,
  justificado, margens sup/esq 3 cm e inf/dir 2 cm, espaçamento 1,5, paginação.

## 5. Necessidade por caso (o run.py decide documento a documento)

| Documento | Gerado quando | Se desnecessário |
|---|---|---|
| 00-Projeto, 06-Cronograma, 07-Orçamento, 08b-Financiamento, 10-Checklist, 11, 12 | Sempre | — |
| 01-TCLE adulto | Sem dispensa total | N/A: TCLE dispensado |
| 02-TCLE responsável, 03-TALE | Com menores/incapazes E sem dispensa | N/A: motivo justificado |
| 04-Anuência | SÓ com entrevistas ou similares DENTRO de instituição (`entrevistas_instituicao`/`coleta_institucional`) | N/A: coleta online, via pública ou documental sem atividade presencial institucional |
| 05-Sigilo | Com `usa_prontuario`, `dados_sigilosos` ou `dados_secundarios` | N/A: sem dado sigiloso declarado |
| 08-Dispensa | Com `dispensa_tcle.solicitar: true` | N/A: há TCLE — dispensa desnecessária |
| 09-Carta | Só com `--com-carta` (há pendência; o agente gera na hora) | Fora do pacote |
| 13 + `criar-form-google.gs` | Com `coleta_online`/`ambiente_virtual` + `questionario` | Fora do pacote |

N/A aparece justificado no `10-Checklist`, no `MANIFESTO.txt` e no `--validar-apenas`.
Regra de ouro: **com TCLE não há dispensa, e com dispensa total não há TCLE.**

Conteúdo preenchível por documento: 01 (título, objetivo/justificativa, procedimentos,
tempo, local, riscos, medidas, benefícios, ressarcimento, retorno, contato de retirada,
pesquisador, cláusulas situacionais: biológico/voz-imagem/vulneráveis/acompanhamento/
saúde mental/SUS); 02 idem + representado; 03 versão simples; 04 (instituição,
representante, cargo, vínculo); 05 (pesquisador + fonte dos dados); 06 (etapas/períodos);
07 (itens/valores/fontes); 08 (justificativa); 13 (perguntas por tipo). Todo o resto =
verbatim oficial. Proibido alterar a ordem dos parágrafos, trocar "participante da pesquisa" por
"entrevistado/sujeito/respondente", dizer "pesquisa sem risco", ou inserir logo.

## 6. Validação fail-closed (run.py)

1. Completude: qualquer campo ausente aborta o modo FINAL listando o que o AGENTE
   deve coletar (`--validar-apenas`); nada com marcação é escrito.
2. Pós-geração: para cada documento do pacote, exige `.docx` e `.pdf` existentes
   e >0 bytes, senão exit 1. Contagem dinâmica conforme a necessidade do caso.
3. Checklist TCLE 27 itens: cada item verificado textualmente no 01-TCLE gerado
   (convite, sem logo, "participante da pesquisa", riscos≠"sem risco", ressarcimento,
   guarda 5 anos, retorno, paginação "Página X de Y", contatos pesquisador+CEP, etc.).
   Falha → exit 1 se item crítico ausente por bug de template.
4. Limpeza: varredura proíbe `[PREENCHER:`, `AVISO:`, `INFORMAÇÕES IMPORTANTES`,
   menções à skill e OBS de template nos documentos submetíveis (00–08b, 13).
5. Proibição de drift: `TEMPLATE_FINGERPRINTS` (trechos-âncora). Se âncora ausente, falha.
6. `MANIFESTO.txt` + `pendencias.txt` + `previsao.txt` sempre escritos.

## 7. Comandos (uso DO AGENTE — nunca pedir ao pesquisador para rodar)

```bash
python3 run.py projeto.json --out outputs-cep-meu-estudo/
python3 run.py projeto.json --out outputs-cep-meu-estudo/ --validar-apenas
python3 run.py projeto.json --out outputs-cep-pend/ --com-carta   # só com pendência
python3 run.py --exemplo  # escreve projeto-exemplo.json no cwd
```

O 12-Roteiro entregue ao pesquisador NÃO contém comandos: cita apenas a pasta
do pacote. Datas: documentos datados automaticamente (Cidade, DD de mês de AAAA);
o agente confere a data atual e o calendário de reuniões e informa a previsão
de parecer (ver `previsao.txt`).

Dependências: `pip install python-docx reportlab pymupdf pdfplumber`
Sem rede, sem LibreOffice, sem npm. Funciona offline.

## 8. Limites e avisos obrigatórios ao usuário

- Skill não substitui parecer do CEP, não garante aprovação, não emite juízo ético.
- Coleta iniciada antes da aprovação **não pode ser apreciada** — cronograma deve prever
  2–3 meses pós-submissão.
- Pesquisas CHS (Res. 510/2016) também passam pelo CEP quando há dado direto/identificável
  com risco acima do cotidiano.
- Sugerir Declaração de Uso de IA nas entregas finais.
- Em cada entrega, informar: pasta do pacote, previsão de parecer (reunião provável)
  e lembrar "assine nos blocos de assinatura; folha de rosto na Plataforma Brasil".

## 9. Localização

- Skill (runtime): `/root/.agents/skills/cep-ufv/SKILL.md` + `run.py` + `projeto-exemplo.json`
- Cópia persistente: `skills-cep-ufv/` na pasta PesquisAI (Drive no Colab · `~/PesquisAI` no offline)
- Entregáveis de cada estudo: `outputs-cep-<slug>/` (nunca dentro do vault sem cópia em outputs-)

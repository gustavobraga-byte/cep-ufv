# cep-ufv — Skill de Documentação CEP/UFV (travada aos templates oficiais)

Skill que gera o **pacote documental para submissão ao Comitê de Ética em Pesquisa com Seres Humanos da UFV (CEP/UFV)** — travada aos modelos oficiais de https://cep.ufv.br/modelos/. Preenche **somente** os campos variáveis dentro do texto verbatim dos templates. **Nunca** cria seção, parágrafo ou documento fora do padrão. Todo documento sai em par idêntico **.docx editável + .pdf**.

## Visão Geral

| Aspecto | Detalhe |
|---------|---------|
| **Nome** | `cep-ufv` |
| **Versão** | 1.7.1-total-no-docs |
| **Categoria** | Ética em pesquisa / Documentação regulatória |
| **Fontes** | CEP/UFV (modelos, rito, pendências, reuniões, legislação) |
| **Formato** | SKILL.md + run.py + modelos/ + docs + tests + evals |
| **Dependências** | `python-docx`, `reportlab`, `pymupdf`, `pdfplumber` (sem rede, sem LibreOffice, sem npm) |

## Documentos gerados (14 × 2 formatos, prontos p/ submissão, sem marcações)

| # | Arquivo | Origem |
|---|---------|--------|
| 01 | TCLE Adulto | Modelo 2026 verbatim |
| 02 | TCLE Responsável Legal | Modelo 2026 verbatim |
| 03 | TALE Assentimento | Modelo 2026 verbatim |
| 04 | Anuência Institucional | Modelo 2026 verbatim |
| 05 | Sigilo e Confidencialidade | Modelo ODT 2019 verbatim |
| 06 | Cronograma (+ compromisso expresso) | Modelo 2026 verbatim |
| 07 | Orçamento detalhado | Checklist 2021 (item obrigatório) |
| 08 | Dispensa de TCLE / declaração de não-dispensa | Norma aplicável (só com justificativa do pesquisador) |
| 09 | Carta-Resposta (vazia) | Modelo ODT 2016 verbatim |
| 10 | Checklist de Conformidade (STATUS por item) | Checklist 2021 transcrito |
| 00 | Projeto de Pesquisa (estrutura 8 seções) | https://cep.ufv.br/projeto-novo/ |
| 08b | Declaração de Financiamento (próprio/externo) | checklist 2021 |
| 11 | Orientações Plataforma Brasil (guia, NÃO SUBMETER) | https://cep.ufv.br/projeto-novo/ |
| 12 | Roteiro passo a passo ETAPA 0–8 (guia, NÃO SUBMETER) | projeto-novo + pendência + cronograma + normativas |

Mais `MANIFESTO.txt` (pares verificados) e `pendencias.txt` (campos `[PREENCHER]` a zerar).

## Estrutura do repositório

```
cep-ufv/
├── SKILL.md                 # Instruções da skill (trava de template)
├── run.py                   # Gerador travado (docx+pdf, fail-closed)
├── projeto-exemplo.json     # Entrada de exemplo
├── modelos/                 # 9 modelos oficiais (PDFs+ODTs de cep.ufv.br, 05/09/2026)
├── README.md                # Este arquivo
├── INSTALL.md               # Instalação
├── CHANGELOG.md             # Histórico
├── CONTRIBUTING.md          # Contribuição (sem tocar em texto legal)
├── LICENSE                  # MIT
├── pyproject.toml
├── requirements.txt
├── .gitignore
├── docs/
│   ├── guia-rapido.md
│   ├── fontes.md
│   ├── legislacao.md
│   └── roteiro-submissao.md
├── examples/
│   ├── projeto-exemplo.json
│   └── gerar_pacote.py
├── evals/
│   └── evals.json
├── tests/
│   ├── test_skill_structure.py
│   └── test_run.py
└── .github/
    ├── workflows/validate.yml
    ├── ISSUE_TEMPLATE/bug_report.md
    ├── ISSUE_TEMPLATE/feature_request.md
    └── PULL_REQUEST_TEMPLATE.md
```

## Uso rápido

```bash
pip install -r requirements.txt
python3 run.py --exemplo
python3 run.py projeto-exemplo.json --out outputs-cep-meu-estudo/
python3 run.py projeto.json --out outputs-cep-meu-estudo/ --validar-apenas
```

Ver `docs/guia-rapido.md` e `examples/gerar_pacote.py`.

## Regras inegociáveis

1. Só preencher campos; resto verbatim. Sem dado = `[PREENCHER]` + `pendencias.txt`. Nunca inventar.
2. Folha de rosto **não** é gerada (sai da Plataforma Brasil). Responsável = orientador.
3. Coleta só após aprovação. Prever 2–3 meses. Reuniões 2026 + dia 20 (ver `docs/roteiro-submissao.md`).
4. Pendência: 30 dias ou arquiva. Documental = substitui; ética = `X modificado` + Carta-Resposta.
5. Base legal: Lei 14.874/2024, Decreto 12.651/2025, Res. 466/2012, 510/2016, NO 001/2013, LGPD.

## Licença

MIT — ver `LICENSE`. Modelos em `modelos/` © CEP/UFV (uso orientativo; verificar vigência em https://cep.ufv.br/modelos/).

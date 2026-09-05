# Contribuindo — cep-ufv

## Regra de ouro
**Nunca altere texto legal dos templates.** Contribuições só podem: corrigir preenchimento de campos, melhorar validação, docs, testes e CI. Mudança de template oficial = atualizar `modelos/` + `CHANGELOG.md` + fingerprints no `run.py`, com link e data da fonte.

## Fluxo
1. Abra uma issue (bug ou proposta).
2. Fork + branch (`fix/...`, `docs/...`, `feat/...`).
3. Rode `pytest -q` e gere o pacote de exemplo (12/12 pares).
4. PR com: o que mudou, fontes oficiais consultadas, evidência de teste.

## Proibido
- Reescrever TCLE/TALE/anuência/sigilo/cronograma/carta com "texto melhor".
- Inventar norma, prazo ou contato.
- Gerar documento fora da lista de 12.

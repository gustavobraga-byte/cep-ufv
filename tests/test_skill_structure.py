#!/usr/bin/env python3
"""Estrutura mínima da skill cep-ufv (espelho do padrão agrobr)."""
from pathlib import Path

def test_skill_directory_structure():
    d = Path(__file__).parent.parent
    req = ["SKILL.md", "README.md", "LICENSE", "CHANGELOG.md", "CONTRIBUTING.md",
           "INSTALL.md", "pyproject.toml", "requirements.txt", ".gitignore",
           "run.py", "projeto-exemplo.json", "evals/evals.json",
           "examples/gerar_pacote.py", "examples/projeto-exemplo.json",
           "docs/guia-rapido.md", "docs/fontes.md", "docs/legislacao.md", "docs/roteiro-submissao.md",
           "tests/test_skill_structure.py", "tests/test_run.py",
           ".github/workflows/validate.yml",
           ".github/ISSUE_TEMPLATE/bug_report.md",
           ".github/ISSUE_TEMPLATE/feature_request.md",
           ".github/PULL_REQUEST_TEMPLATE.md"]
    miss = [f for f in req if not (d / f).exists()]
    assert not miss, f"Arquivos ausentes: {miss}"

def test_skill_md_format():
    c = (Path(__file__).parent.parent / "SKILL.md").read_text(encoding="utf-8")
    assert c.startswith("---")
    end = c.find("---", 3)
    assert end > 0
    fm = c[3:end]
    assert "name:" in fm and "description:" in fm and "cep-ufv" in fm

def test_modelos_present():
    d = Path(__file__).parent.parent / "modelos"
    assert d.is_dir()
    nomes = {p.name for p in d.iterdir()}
    for esperado in ["TCLE.pdf", "TCLE-Responsavel.pdf", "Termo-Assentimento.pdf", "Anuencia.pdf",
                     "Cronograma.pdf", "Relatorio-Final.pdf", "Carta-Resposta.odt", "Sigilo.odt", "Checklist.pdf"]:
        assert esperado in nomes, f"Modelo ausente: {esperado}"

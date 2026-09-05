#!/usr/bin/env python3
"""Geração fail-closed a partir do exemplo (13 pares padrão; carta e online condicionais)."""
import json, subprocess, sys
from pathlib import Path

SLUGS_PADRAO = ["00-Projeto-de-Pesquisa", "01-TCLE-Adulto", "02-TCLE-Responsavel", "03-TALE-Assentimento",
                "04-Anuencia-Institucional", "05-Sigilo-Confidencialidade", "06-Cronograma", "07-Orcamento",
                "08b-Declaracao-Financiamento",
                "10-Checklist-Conformidade", "11-Orientacoes-Plataforma-Brasil", "12-Roteiro-Submissao"]

def test_run_generates_12_pairs(tmp_path):
    base = Path(__file__).parent.parent
    out = tmp_path / "outputs-cep-teste"
    r = subprocess.run([sys.executable, str(base / "run.py"), str(base / "examples" / "projeto-exemplo.json"),
                        "--out", str(out)], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    for s in SLUGS_PADRAO:
        for ext in (".docx", ".pdf"):
            f = out / f"{s}{ext}"
            assert f.exists() and f.stat().st_size > 0, f"Ausente/vazio: {s}{ext}"
    assert (out / "MANIFESTO.txt").exists()
    assert (out / "pendencias.txt").exists()
    assert (out / "previsao.txt").exists()
    assert not (out / "08-Dispensa-TCLE.docx").exists()  # há TCLE: dispensa desnecessária
    man = (out / "MANIFESTO.txt").read_text(encoding="utf-8")
    assert "dispensa desnecessária" in man


def test_paginacao_total(tmp_path):
    """Checklist exige 'Página X de Y': campos PAGE+NUMPAGES no docx e texto no pdf."""
    import subprocess, sys
    from pathlib import Path as _P
    base = _P(__file__).parent.parent
    out = tmp_path / "pag"
    r = subprocess.run([sys.executable, str(base / "run.py"), str(base / "examples" / "projeto-exemplo.json"),
                        "--out", str(out)], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    from docx import Document as _D
    import re as _re
    f = out / "01-TCLE-Adulto.docx"
    paras = _D(f).sections[0].footer.paragraphs
    xml = "".join(p._p.xml for p in paras)
    assert "PAGE" in xml and "NUMPAGES" in xml and "separate" in xml
    import fitz
    d = fitz.open(out / "01-TCLE-Adulto.pdf")
    n = d.page_count
    for i, pg in enumerate(d):
        assert f"Página {i + 1} de {n}" in pg.get_text()
    # prévia sem recálculo (Google Docs): total real como texto do campo
    m = _re.findall(r"NUMPAGES.*?<w:t[^>]*>([^<]*)</w:t>", xml, _re.S)
    assert m and m[0] == str(n)
    vis = "".join(p.text for p in paras)
    assert vis == f"Página 1 de {n}"



def test_carta_somente_sob_demanda(tmp_path):
    base = Path(__file__).parent.parent
    out = tmp_path / "sem-carta"
    r = subprocess.run([sys.executable, str(base / "run.py"), str(base / "examples" / "projeto-exemplo.json"),
                        "--out", str(out)], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert not (out / "09-Carta-Resposta.docx").exists()
    out2 = tmp_path / "com-carta"
    r = subprocess.run([sys.executable, str(base / "run.py"), str(base / "examples" / "projeto-exemplo.json"),
                        "--out", str(out2), "--com-carta"], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert (out2 / "09-Carta-Resposta.docx").exists()


def test_necessidade_dispensa_vs_tcle(tmp_path):
    """Com TCLE não há dispensa; com dispensa total não há TCLE/TALE."""
    base = Path(__file__).parent.parent
    p = json.loads((base / "examples" / "projeto-exemplo.json").read_text(encoding="utf-8"))
    pj = tmp_path / "disp.json"
    p["dispensa_tcle"] = {"solicitar": True, "justificativa": "Banco anonimizado, sem contato."}
    pj.write_text(json.dumps(p, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "disp"
    r = subprocess.run([sys.executable, str(base / "run.py"), str(pj), "--out", str(out)],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert (out / "08-Dispensa-TCLE.docx").exists()
    assert not (out / "01-TCLE-Adulto.docx").exists()
    assert not (out / "02-TCLE-Responsavel.docx").exists()
    assert not (out / "03-TALE-Assentimento.docx").exists()
    man = (out / "MANIFESTO.txt").read_text(encoding="utf-8")
    assert "DESNECESSÁRIOS" in man and "TCLE dispensado" in man


def test_necessidade_sem_menores_sem_sigilo(tmp_path):
    base = Path(__file__).parent.parent
    p = json.loads((base / "examples" / "projeto-exemplo.json").read_text(encoding="utf-8"))
    p["usa_menores"] = False
    p["dados_sigilosos"] = False
    p["coleta_institucional"] = False
    p["instituicao_local"] = {"nome": "", "representante": "", "cargo": ""}
    pj = tmp_path / "ad.json"
    pj.write_text(json.dumps(p, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "ad"
    r = subprocess.run([sys.executable, str(base / "run.py"), str(pj), "--out", str(out)],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert (out / "01-TCLE-Adulto.docx").exists()
    for ausente in ["02-TCLE-Responsavel.docx", "03-TALE-Assentimento.docx",
                    "04-Anuencia-Institucional.docx", "05-Sigilo-Confidencialidade.docx",
                    "08-Dispensa-TCLE.docx"]:
        assert not (out / ausente).exists(), ausente


def test_coleta_online_gera_forms(tmp_path):
    base = Path(__file__).parent.parent
    p = json.loads((base / "examples" / "projeto-exemplo.json").read_text(encoding="utf-8"))
    p["coleta_online"] = True
    pj = tmp_path / "online.json"
    pj.write_text(json.dumps(p, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "online"
    r = subprocess.run([sys.executable, str(base / "run.py"), str(pj), "--out", str(out)],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert (out / "13-Coleta-Online-GoogleForms.docx").exists()
    assert (out / "criar-form-google.gs").exists()

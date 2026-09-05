#!/usr/bin/env python3
"""Exemplo: gera o pacote CEP/UFV a partir do projeto-exemplo.json."""
import subprocess, sys
from pathlib import Path
base = Path(__file__).resolve().parent.parent
out = base / "outputs-cep-exemplo"
r = subprocess.run([sys.executable, str(base / "run.py"), str(base / "examples" / "projeto-exemplo.json"), "--out", str(out)])
sys.exit(r.returncode)

# Instalação — cep-ufv

## Requisitos
- Python 3.9+
- Sem LibreOffice, sem npm, sem rede (após instalar deps).

## Instalar
```bash
pip install -r requirements.txt
# ou
pip install python-docx reportlab pymupdf pdfplumber
```

## Verificar
```bash
python3 run.py --exemplo
python3 run.py projeto-exemplo.json --out outputs-cep-exemplo/
pytest -q
```

Saída esperada: `OK: 12 documentos x2 formatos` + `pendencias.txt` + `MANIFESTO.txt`.

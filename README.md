# pdf2word — PDF to DOCX Converter

Package Python pour convertir des PDF en fichiers Word (.docx) éditables.
Supporte **MS Word** (Windows), **LibreOffice** (Linux), et **PyMuPDF** (partout).

## Installation

```bash
pip install pdf2word
```

### Prérequis système

| Moteur | Plateforme | Installation |
|---|---|---|
| **MS Word** (meilleure qualité) | Windows | MS Office installé |
| **LibreOffice** (recommandé Linux) | Linux/macOS | `sudo apt install libreoffice-writer` |
| **Tesseract** (OCR, optionnel) | Toutes | [Guide](https://github.com/tesseract-ocr/tesseract) |

## Utilisation

### Ligne de commande

```bash
# Conversion automatique (détecte le meilleur moteur)
pdf2word convert document.pdf -o document.docx

# Mode MS Word (Windows uniquement, meilleure qualité)
pdf2word convert document.pdf -o output.docx --mode msword

# Mode LibreOffice (Linux/macOS, open source)
pdf2word convert document.pdf -o output.docx --mode libreoffice

# Mode smart (partout, sans dépendances système)
pdf2word convert document.pdf -o output.docx --mode smart

# OCR pour PDF scannés
pdf2word convert scan.pdf -o scan.docx --force-ocr

# Reconvertir un Word en PDF
pdf2word reconvert document.docx -o output_dir/
```

### API Python

```python
from pdf2word.converter import PDFToWordConverter

converter = PDFToWordConverter(mode="libreoffice")
result = converter.convert("document.pdf", "document.docx")
print(result)
# {'output_path': 'document.docx', 'method': 'libreoffice', ...}
```

## Modes de conversion

| Mode | Qualité | Plateforme | Description |
|---|---|---|---|
| `msword` | ⭐⭐⭐⭐⭐ | Windows | MS Word PDF Reflow via COM |
| `libreoffice` | ⭐⭐⭐⭐ | Linux/macOS/Windows | LibreOffice headless |
| `smart` | ⭐⭐⭐ | Toutes | PyMuPDF extraction |
| `text` | ⭐⭐⭐ | Toutes | pdf2docx |
| `hybrid` | ⭐⭐⭐ | Toutes | Pages en images |
| `combined` | ⭐⭐⭐ | Toutes | Image + texte |
| `ocr` | ⭐⭐ | Toutes | OCR (Tesseract) |
| `auto` | Variable | Toutes | Détecte texte vs scanné |

**Fallback automatique** : si `msword` est demandé sur Linux → utilise `libreoffice` → puis `smart`.

## Déploiement Docker (Linux)

```bash
docker build -t pdf2word .
docker run -v $(pwd):/data pdf2word convert /data/input.pdf -o /data/output.docx --mode libreoffice
```

## Architecture

```
pdf2word/
├── converter.py              # Orchestrateur + routing
├── msword_converter.py       # MS Word (pywin32 COM)
├── libreoffice_converter.py  # LibreOffice headless
├── smart_converter.py        # PyMuPDF extraction
├── text_converter.py         # pdf2docx
├── hybrid_converter.py       # Pages en images
├── combined_converter.py     # Image + texte overlay
├── ocr_converter.py          # OCR (Tesseract)
├── docx_enhancer.py          # Post-processing
├── analyzer.py               # Détection type PDF
├── docx_to_pdf.py            # Reconversion DOCX→PDF
└── cli.py                    # CLI
```

## Licence

MIT

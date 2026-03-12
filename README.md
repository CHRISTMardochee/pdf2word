# pdf2word — PDF to DOCX Converter

Package Python pour convertir des PDF en fichiers Word (.docx) éditables.
Supporte **MS Word** (Windows), **LibreOffice** (Linux), **PyMuPDF Smart** (partout), **Docling ML** (IBM), et **Cloud API**.

## Installation

```bash
pip install pdf2word
```

### Dépendances optionnelles

```bash
# Pour le mode Docling ML (IBM) — Document Intelligence
pip install pdf2word[ml]

# Pour le mode Cloud API (ConvertAPI)
pip install pdf2word[cloud]
```

### Prérequis système

| Moteur | Plateforme | Installation |
|---|---|---|
| **MS Word** (meilleure qualité) | Windows | MS Office installé |
| **LibreOffice** (recommandé Linux/macOS) | Linux/macOS/Windows | Voir ci-dessous |
| **Tesseract** (OCR, optionnel) | Toutes | [Guide](https://github.com/tesseract-ocr/tesseract) |

**Linux (Debian/Ubuntu)** :
```bash
sudo apt install libreoffice-writer fonts-liberation fonts-dejavu
# Polices Microsoft (optionnel, améliore la fidélité) :
sudo add-apt-repository multiverse && sudo apt update && sudo apt install ttf-mscorefonts-installer
```

**macOS** :
```bash
brew install --cask libreoffice
# OCR (optionnel)
brew install tesseract
```

## Utilisation

### Ligne de commande

```bash
# Conversion automatique (détecte le meilleur moteur)
pdf2word convert document.pdf -o document.docx

# Mode Smart V5 (PyMuPDF + tables + couleurs, 100% open source)
pdf2word convert document.pdf -o output.docx --mode smart

# Mode Docling ML (IBM Document Intelligence, 100% open source)
pdf2word convert document.pdf -o output.docx --mode docling

# Mode Cloud API (ConvertAPI, qualité maximale)
pdf2word convert document.pdf -o output.docx --mode cloud

# Mode MS Word (Windows uniquement)
pdf2word convert document.pdf -o output.docx --mode msword

# Mode LibreOffice (Linux/macOS, open source)
pdf2word convert document.pdf -o output.docx --mode libreoffice

# OCR pour PDF scannés
pdf2word convert scan.pdf -o scan.docx --force-ocr
```

### Gestion de la clé API (mode cloud)

```bash
# Sauvegarder la clé ConvertAPI une fois pour toutes
pdf2word set-key VOTRE_CLE_SECRETE

# Supprimer la clé
pdf2word remove-key
```

### API Python

```python
from pdf2word.converter import PDFToWordConverter

# Mode Smart (recommandé, 100% gratuit)
converter = PDFToWordConverter(mode="smart")
result = converter.convert("document.pdf", "document.docx")

# Mode Docling ML (IBM, gratuit mais plus lent)
converter = PDFToWordConverter(mode="docling")
result = converter.convert("document.pdf", "document.docx")

# Mode Cloud (ConvertAPI, payant mais parfait)
converter = PDFToWordConverter(mode="cloud", api_key="votre_clé")
result = converter.convert("document.pdf", "document.docx")
```

## Modes de conversion

| Mode | Qualité | Coût | Plateforme | Description |
|---|---|---|---|---|
| `smart` | ⭐⭐⭐⭐ | Gratuit | Toutes | PyMuPDF + détection tables + couleurs vectorielles |
| `docling` | ⭐⭐⭐⭐ | Gratuit | Toutes | IBM Docling ML (Document Intelligence) |
| `cloud` | ⭐⭐⭐⭐⭐ | Payant | Toutes | ConvertAPI (qualité maximale garantie) |
| `msword` | ⭐⭐⭐⭐⭐ | Gratuit | Windows | MS Word PDF Reflow via COM |
| `libreoffice` | ⭐⭐⭐⭐ | Gratuit | Linux/macOS/Windows | LibreOffice headless |
| `text` | ⭐⭐⭐ | Gratuit | Toutes | pdf2docx |
| `hybrid` | ⭐⭐⭐ | Gratuit | Toutes | Pages en images |
| `combined` | ⭐⭐⭐ | Gratuit | Toutes | Image + texte |
| `ocr` | ⭐⭐ | Gratuit | Toutes | OCR (Tesseract) |
| `auto` | Variable | Gratuit | Toutes | Détecte texte vs scanné |

**Fallback automatique** : si `msword` est demandé sur Linux → utilise `libreoffice` → puis `smart`.

## Architecture

```
pdf2word/
├── converter.py              # Orchestrateur + routing
├── smart_converter.py        # PyMuPDF Smart V5 (tables, couleurs, colonnes)
├── docling_converter.py      # IBM Docling ML + post-traitement PyMuPDF
├── cloud_converter.py        # ConvertAPI (cloud)
├── config.py                 # Gestion clé API globale
├── msword_converter.py       # MS Word (pywin32 COM)
├── libreoffice_converter.py  # LibreOffice headless
├── text_converter.py         # pdf2docx
├── hybrid_converter.py       # Pages en images
├── combined_converter.py     # Image + texte overlay
├── ocr_converter.py          # OCR (Tesseract)
├── docx_enhancer.py          # Post-processing qualité
├── analyzer.py               # Détection type PDF
├── docx_to_pdf.py            # Reconversion DOCX→PDF
└── cli.py                    # CLI
```

## Déploiement Docker (Linux)

```bash
docker build -t pdf2word .
docker run -v $(pwd):/data pdf2word convert /data/input.pdf -o /data/output.docx --mode smart
```

## Licence

MIT

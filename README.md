# Traducere Document Word (RO → EN) cu Azure Translator

Aplicație desktop simplă (Tkinter) care încarcă un fișier `.docx`, îl trimite la
Azure Translator și salvează versiunea tradusă în engleză păstrând formatarea
(stiluri, fonturi, tabele, headere/footere).

## Instalare

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Configurare Azure

1. Creează o resursă **Translator** în portalul Azure.
2. Copiază `Key` și `Region`.
3. Copiază `.env.example` în `.env` și completează valorile (opțional — pot fi
   introduse și direct în interfață).

```bash
cp .env.example .env
```

## Utilizare

```bash
python translator_app.py
```

- Apasă **Răsfoiește…** și alege un fișier `.docx`.
- Verifică / completează cheia și regiunea.
- Apasă **Traduce**. Documentul tradus este salvat lângă cel original cu
  sufixul `.en.docx`.

## Unde se salvează rapoartele

Toate documentele traduse ajung în folderul `rapoarte_traduse/` creat automat
lângă aplicație (sau lângă `.exe` dacă rulezi varianta compilată). După
traducere, fișierul se deschide automat în Word.

## Build executabil Windows (.exe)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name TraduceRaport translator_app.py
```

`.exe`-ul rezultat e în `dist/TraduceRaport.exe`. Pentru a-l folosi, copiază
lângă el fișierul `.env` cu cheia Azure — la prima rulare se va crea automat
folderul `rapoarte_traduse/` în același loc.

## Cum păstrează formatarea

Traducerea se face la nivel de **run** din `python-docx` — fiecare segment de
text își păstrează stilul, fontul, culoarea și marcajul bold/italic. Sunt
parcurse paragrafele din corp, celulele tabelelor (recursiv) și paragrafele din
headere/footere.

"""
Aplicație desktop pentru traducerea documentelor Word din română în engleză
folosind Azure Translator. Păstrează formatarea (stiluri, fonturi, tabele, headers/footers).
"""
import os
import uuid
import copy
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import certifi
import requests
from docx import Document
from dotenv import load_dotenv

# Asigură că requests folosește bundle-ul certifi (rezolvă SSL: CERTIFICATE_VERIFY_FAILED)
os.environ.setdefault("SSL_CERT_FILE", certifi.where())
os.environ.setdefault("REQUESTS_CA_BUNDLE", certifi.where())

load_dotenv()

AZURE_KEY = os.getenv("AZURE_TRANSLATOR_KEY", "")
AZURE_REGION = os.getenv("AZURE_TRANSLATOR_REGION", "")
AZURE_ENDPOINT = os.getenv(
    "AZURE_TRANSLATOR_ENDPOINT", "https://api.cognitive.microsofttranslator.com"
)


class AzureTranslator:
    def __init__(self, key: str, region: str, endpoint: str):
        self.key = key
        self.region = region
        self.url = endpoint.rstrip("/") + "/translate"

    def translate_batch(self, texts, from_lang="ro", to_lang="en"):
        """Traduce o listă de texte. Returnează listă de aceeași lungime."""
        if not texts:
            return []
        # Azure permite max 1000 elemente / 50.000 caractere per request.
        results = [""] * len(texts)
        batch, batch_idx, batch_chars = [], [], 0
        MAX_ITEMS, MAX_CHARS = 100, 45000

        def flush():
            nonlocal batch, batch_idx, batch_chars
            if not batch:
                return
            translated = self._call(batch, from_lang, to_lang)
            for i, t in zip(batch_idx, translated):
                results[i] = t
            batch, batch_idx, batch_chars = [], [], 0

        for i, t in enumerate(texts):
            if not t.strip():
                results[i] = t
                continue
            if len(batch) >= MAX_ITEMS or batch_chars + len(t) > MAX_CHARS:
                flush()
            batch.append(t)
            batch_idx.append(i)
            batch_chars += len(t)
        flush()
        return results

    def _call(self, texts, from_lang, to_lang):
        params = {"api-version": "3.0", "from": from_lang, "to": [to_lang]}
        headers = {
            "Ocp-Apim-Subscription-Key": self.key,
            "Ocp-Apim-Subscription-Region": self.region,
            "Content-Type": "application/json",
            "X-ClientTraceId": str(uuid.uuid4()),
        }
        body = [{"text": t} for t in texts]
        r = requests.post(self.url, params=params, headers=headers, json=body,
                          timeout=60, verify=certifi.where())
        r.raise_for_status()
        data = r.json()
        return [item["translations"][0]["text"] for item in data]


def iter_paragraphs(doc):
    """Iterează paragrafele din corp, tabele (recursiv), headere și footere."""
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        yield from _iter_table(table)
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.first_page_header, section.first_page_footer,
                   section.even_page_header, section.even_page_footer):
            if hf is None:
                continue
            for p in hf.paragraphs:
                yield p
            for table in hf.tables:
                yield from _iter_table(table)


def _iter_table(table):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                yield p
            for t in cell.tables:
                yield from _iter_table(t)


def translate_document(input_path: str, output_path: str, translator: AzureTranslator,
                       progress_cb=None):
    """Traduce documentul păstrând formatarea la nivel de run."""
    doc = Document(input_path)

    runs = []
    for p in iter_paragraphs(doc):
        for run in p.runs:
            if run.text:
                runs.append(run)

    if progress_cb:
        progress_cb(0, len(runs))

    texts = [r.text for r in runs]
    translated = translator.translate_batch(texts, "ro", "en")
    for run, new_text in zip(runs, translated):
        run.text = new_text

    if progress_cb:
        progress_cb(len(runs), len(runs))

    doc.save(output_path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Traducere document Word — RO → EN (Azure)")
        self.geometry("560x320")
        self.resizable(False, False)

        pad = {"padx": 12, "pady": 6}

        ttk.Label(self, text="Document Word (.docx):").pack(anchor="w", **pad)
        row = ttk.Frame(self); row.pack(fill="x", **pad)
        self.path_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.path_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Răsfoiește…", command=self.browse).pack(side="left", padx=(6, 0))

        creds = ttk.LabelFrame(self, text="Credențiale Azure Translator")
        creds.pack(fill="x", **pad)
        self.key_var = tk.StringVar(value=AZURE_KEY)
        self.region_var = tk.StringVar(value=AZURE_REGION)
        ttk.Label(creds, text="Key:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(creds, textvariable=self.key_var, show="•", width=50).grid(row=0, column=1, padx=6, pady=4)
        ttk.Label(creds, text="Region:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(creds, textvariable=self.region_var, width=50).grid(row=1, column=1, padx=6, pady=4)

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", **pad)
        self.status = ttk.Label(self, text="Gata.")
        self.status.pack(anchor="w", **pad)

        ttk.Button(self, text="Traduce", command=self.start).pack(pady=10)

    def browse(self):
        path = filedialog.askopenfilename(
            title="Alege document Word",
            filetypes=[("Documente Word", "*.docx")],
        )
        if path:
            self.path_var.set(path)

    def start(self):
        path = self.path_var.get().strip()
        key = self.key_var.get().strip()
        region = self.region_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Eroare", "Selectează un fișier .docx valid.")
            return
        if not key or not region:
            messagebox.showerror("Eroare", "Completează cheia și regiunea Azure.")
            return

        base, ext = os.path.splitext(path)
        out_path = f"{base}.en{ext}"
        translator = AzureTranslator(key, region, AZURE_ENDPOINT)

        def worker():
            try:
                self.set_status("Se traduce…")
                translate_document(path, out_path, translator, self.set_progress)
                self.set_status(f"Gata: {out_path}")
                messagebox.showinfo("Succes", f"Document salvat:\n{out_path}")
            except requests.HTTPError as e:
                messagebox.showerror("Eroare Azure", f"{e}\n{e.response.text if e.response else ''}")
                self.set_status("Eroare.")
            except Exception as e:
                messagebox.showerror("Eroare", str(e))
                self.set_status("Eroare.")

        threading.Thread(target=worker, daemon=True).start()

    def set_progress(self, current, total):
        self.progress["maximum"] = max(total, 1)
        self.progress["value"] = current
        self.update_idletasks()

    def set_status(self, text):
        self.status.config(text=text)
        self.update_idletasks()


if __name__ == "__main__":
    App().mainloop()

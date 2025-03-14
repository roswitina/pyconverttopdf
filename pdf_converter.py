import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, IntVar
from tkinter.ttk import Progressbar
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
import subprocess
import os
import tempfile
import logging
import shutil
import sys
import threading
import win32com.client
from queue import Queue
from PyPDF2 import PdfWriter, PdfReader
from langdetect import detect, DetectorFactory
from langdetect.lang_detect_exception import LangDetectException

# Überprüfen, ob erforderliche Bibliotheken installiert sind
try:
    from docx2pdf import convert as docx2pdf
    import pandas as pd
    from pptx import Presentation
    from reportlab.pdfgen import canvas
    from PIL import Image
    import pytesseract
    from pdf2image import convert_from_path
except ImportError as e:
    messagebox.showerror("Fehler", f"Erforderliche Bibliothek nicht installiert: {e.name}")
    sys.exit(1)

# Ghostscript-Pfad automatisch suchen
def find_ghostscript():
    possible_paths = [
        r"C:\Program Files\gs\gs10.04.0\bin\gswin64c.exe",
        r"C:\Program Files (x86)\gs\gs10.04.0\bin\gswin32c.exe",
        shutil.which("gs"),
        shutil.which("gswin64c"),
        shutil.which("gswin32c")
    ]
    for path in possible_paths:
        if path and os.path.exists(path):
            return path
    return None

# Tesseract-Pfad finden
def find_tesseract():
    possible_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        shutil.which("tesseract"),
        "/usr/bin/tesseract",
        "/usr/local/bin/tesseract"
    ]
    for path in possible_paths:
        if path and os.path.exists(path):
            return path
    return None

GHOSTSCRIPT_PATH = find_ghostscript()
TESSERACT_PATH = find_tesseract()

if not GHOSTSCRIPT_PATH:
    messagebox.showerror("Fehler", "Ghostscript ist nicht installiert oder der Pfad ist falsch.")
    sys.exit(1)

if not TESSERACT_PATH:
    messagebox.showerror("Fehler", "Tesseract OCR ist nicht installiert oder der Pfad ist falsch.")
    sys.exit(1)

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# Logging konfigurieren
def setup_logging():
    logging.basicConfig(filename="converter.log", level=logging.DEBUG,
                        format="%(asctime)s - %(levelname)s - %(message)s")

setup_logging()

def handle_error(e, file_path):
    logging.exception(f"Fehler bei {file_path}: {e}")
    messagebox.showerror("Fehler", f"Fehler bei {os.path.basename(file_path)}: {e}")

def sanitize_filename(filename):
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '-')
    return filename

def detect_language(text):
    try:
        language = detect(text)
        return language
    except LangDetectException:
        return 'eng'  # Fallback-Sprache

def apply_ocr_with_language(input_path, output_pdf, language='deu+eng', dpi=300):
    try:
        img = Image.open(input_path)
        text = pytesseract.image_to_pdf_or_hocr(img, lang=language, extension='pdf')
        with open(output_pdf, 'wb') as f:
            f.write(text)
        logging.info(f"OCR erfolgreich auf {input_path} angewendet")
    except Exception as e:
        handle_error(e, input_path)

def apply_ocr_to_pdf(input_pdf, output_pdf, language='deu+eng', dpi=300, page_numbers=None):
    try:
        images = convert_from_path(input_pdf, dpi=dpi)
        with tempfile.TemporaryDirectory() as tmpdir:
            image_paths = []
            for i, img in enumerate(images):
                if page_numbers and i not in page_numbers:
                    continue
                img_path = os.path.join(tmpdir, f"page_{i}.jpg")
                img.save(img_path, "JPEG")
                image_paths.append(img_path)

            pdf_pages = []
            for img_path in image_paths:
                pdf_bytes = pytesseract.image_to_pdf_or_hocr(img_path, lang=language, extension='pdf')
                pdf_pages.append(pdf_bytes)

            with open(output_pdf, 'wb') as out_file:
                for page in pdf_pages:
                    out_file.write(page)
        logging.info(f"OCR erfolgreich auf PDF {input_pdf} angewendet")
    except Exception as e:
        handle_error(e, input_pdf)

def convert_docx_to_pdf(input_path, output_path):
    try:
        docx2pdf(input_path, output_path)
        logging.info(f"Erfolgreich konvertiert: {input_path} zu {output_path}")
    except Exception as e:
        handle_error(e, input_path)

def convert_excel_to_pdf(input_path, output_path):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(input_path)
        ws = wb.ActiveSheet
        ws.ExportAsFixedFormat(0, output_path)
        wb.Close(False)
        excel.Quit()
        logging.info(f"Erfolgreich konvertiert: {input_path} zu {output_path}")
    except Exception as e:
        handle_error(e, input_path)

def convert_pptx_to_pdf(input_path, output_path):
    try:
        prs = Presentation(input_path)
        image_paths = []
        with tempfile.TemporaryDirectory() as tmpdirname:
            for i, slide in enumerate(prs.slides):
                image_path = os.path.join(tmpdirname, f"slide_{i}.png")
                slide.export(image_path)
                image_paths.append(image_path)

            pdf_canvas = canvas.Canvas(output_path, pagesize=letter)
            for image_path in image_paths:
                img = Image.open(image_path)
                width, height = letter
                pdf_canvas.drawImage(image_path, 0, 0, width, height)
                pdf_canvas.showPage()
            pdf_canvas.save()
        logging.info(f"Erfolgreich konvertiert: {input_path} zu {output_path}")
    except Exception as e:
        handle_error(e, input_path)

def convert_csv_to_pdf(input_path, output_path):
    try:
        df = pd.read_csv(input_path)
        pdf_canvas = canvas.Canvas(output_path, pagesize=letter)
        text = df.to_string(index=False)
        y = 750
        for line in text.split("\n"):
            pdf_canvas.drawString(50, y, line)
            y -= 15
            if y < 50:
                pdf_canvas.showPage()
                y = 750
        pdf_canvas.save()
        logging.info(f"Erfolgreich konvertiert: {input_path} zu {output_path}")
    except Exception as e:
        handle_error(e, input_path)

def convert_txt_to_pdf(input_path, output_path):
    try:
        pdf_canvas = canvas.Canvas(output_path, pagesize=letter)
        y = 750
        with open(input_path, 'r', encoding='utf-8') as file:
            for line in file:
                pdf_canvas.drawString(50, y, line.strip())
                y -= 15
                if y < 50:
                    pdf_canvas.showPage()
                    y = 750
        pdf_canvas.save()
        logging.info(f"Erfolgreich konvertiert: {input_path} zu {output_path}")
    except Exception as e:
        handle_error(e, input_path)

def convert_image_to_pdf(input_path, output_path, use_ocr=False, language='deu+eng', dpi=300):
    try:
        if use_ocr:
            apply_ocr_with_language(input_path, output_path, language, dpi)
        else:
            img = Image.open(input_path)
            img.save(output_pdf, "PDF", resolution=dpi)
        logging.info(f"Erfolgreich konvertiert: {input_path} zu {output_path}")
    except Exception as e:
        handle_error(e, input_path)

def convert_to_pdfa(input_path, output_path, pdfa_type="PDF/A-1b"):
    pdfa_version = {
        "PDF/A-1b": 1,
        "PDF/A-2b": 2,
        "PDF/A-3b": 3,
    }.get(pdfa_type, 1)

    command = [
        GHOSTSCRIPT_PATH,
        "-dBATCH",
        "-dNOPAUSE",
        "-dNOOUTERSAVE",
        f"-dPDFA={pdfa_version}",
        "-sDEVICE=pdfwrite",
        "-dPDFACompatibilityPolicy=1",
        f"-sOutputFile={output_path}",
        input_path
    ]

    if pdfa_type in ["PDF/A-2b", "PDF/A-3b"]:
        command.extend(["-sColorConversionStrategy=RGB", "-sProcessColorModel=DeviceRGB"])
    else:
        command.extend(["-sColorConversionStrategy=CMYK", "-sProcessColorModel=DeviceCMYK"])

    try:
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if result.returncode != 0:
            logging.error(f"PDF/A-Konvertierung fehlgeschlagen: {result.stderr}")
            raise RuntimeError("PDF/A-Konvertierung fehlgeschlagen. Details siehe converter.log.")
        logging.info(f"Erfolgreich in PDF/A konvertiert: {input_path} zu {output_path}")
    except subprocess.CalledProcessError as e:
        handle_error(e, input_path)

def set_pdf_properties(output_path, title, author):
    try:
        reader = PdfReader(output_path)
        writer = PdfWriter()
        writer.append_pages_from_reader(reader)
        metadata = {
            '/Title': title,
            '/Author': author
        }
        writer.add_metadata(metadata)
        with open(output_path, 'wb') as f:  # Use output_path instead of output_pdf
            writer.write(f)
    except Exception as e:
        handle_error(e, output_path)


def get_unique_filename(output_path):
    base, ext = os.path.splitext(output_path)
    counter = 1
    while os.path.exists(output_path):
        output_path = f"{base}_V{counter}{ext}"
        counter += 1
    return output_path

def update_progress(file_name, value, total):
    """Aktualisiert die Fortschrittsanzeige in der GUI."""
    progress_bar['value'] = value
    progress_label.config(text=f"Konvertiere: {file_name} ({value}/{total})")
    root.update_idletasks()

def convert_files(file_paths, output_dir, pdfa_type, title="", author="", dpi=300, page_numbers=None):
    temp_files = []
    use_ocr = ocr_var.get()
    language = ocr_language_var.get()
    total_files = len(file_paths)

    try:
        for index, file_path in enumerate(file_paths):
            try:
                file_name = os.path.basename(file_path)
                base_name = sanitize_filename(os.path.splitext(file_name)[0])

                if pdfa_type != "Standard":
                    clean_pdfa_type = sanitize_filename(pdfa_type)
                    base_name += f"_{clean_pdfa_type}"

                output_path = os.path.join(output_dir, f"{base_name}.pdf")
                output_path = get_unique_filename(output_path)

                temp_pdf_path = os.path.join(tempfile.gettempdir(), f"temp_output_{file_name}.pdf")
                temp_files.append(temp_pdf_path)

                if file_path.endswith('.docx'):
                    convert_docx_to_pdf(file_path, temp_pdf_path)
                elif file_path.endswith('.xlsx'):
                    convert_excel_to_pdf(file_path, temp_pdf_path)
                elif file_path.endswith('.pptx'):
                    convert_pptx_to_pdf(file_path, temp_pdf_path)
                elif file_path.endswith('.csv'):
                    convert_csv_to_pdf(file_path, temp_pdf_path)
                elif file_path.endswith('.txt'):
                    convert_txt_to_pdf(file_path, temp_pdf_path)
                elif file_path.endswith(('.jpg', '.jpeg', '.png')):
                    convert_image_to_pdf(file_path, temp_pdf_path, use_ocr=use_ocr, language=language, dpi=dpi)
                elif file_path.endswith('.pdf'):
                    if use_ocr:
                        apply_ocr_to_pdf(file_path, temp_pdf_path, language=language, dpi=dpi, page_numbers=page_numbers)
                    else:
                        shutil.copy(file_path, temp_pdf_path)

                if os.path.getsize(temp_pdf_path) == 0:
                    raise ValueError("Temporäre PDF-Datei ist leer.")

                if pdfa_type != "Standard":
                    convert_to_pdfa(temp_pdf_path, output_path, pdfa_type)
                else:
                    os.rename(temp_pdf_path, output_path)

                if title or author:
                    set_pdf_properties(output_path, title, author)

                root.after(0, update_progress, file_name, index + 1, total_files)
            except Exception as e:
                handle_error(e, file_path)

    finally:
        for temp_file in temp_files:
            secure_delete(temp_file)
        root.after(0, reset_gui)

def secure_delete(file_path):
    try:
        os.remove(file_path)
    except Exception as e:
        logging.error(f"Fehler beim Löschen der Datei {file_path}: {e}")

def reset_gui():
    progress_bar['value'] = 0
    progress_label.config(text="Bereit")
    root.update_idletasks()

def start_conversion():
    file_paths = filedialog.askopenfilenames(
        filetypes=[
            ("Alle unterstützten Dateien", "*.docx *.xlsx *.pptx *.csv *.txt *.jpg *.jpeg *.png *.pdf"),
            ("Word-Dokumente", "*.docx"),
            ("Excel-Tabellen", "*.xlsx"),
            ("PowerPoint-Präsentationen", "*.pptx"),
            ("CSV-Dateien", "*.csv"),
            ("Textdateien", "*.txt"),
            ("Bilder", "*.jpg *.jpeg *.png"),
            ("PDF-Dateien", "*.pdf"),
            ("Alle Dateien", "*.*")
        ],
        title="Dokumente zur Konvertierung auswählen"
    )
    if not file_paths:
        return

    output_dir = filedialog.askdirectory(title="Ausgabeordner wählen")
    if not output_dir:
        return

    pdfa_type = pdf_version_var.get().replace("Standard-PDF", "Standard")
    title = title_var.get()
    author = author_var.get()
    dpi = dpi_var.get()
    page_numbers = page_numbers_var.get()
    if page_numbers:
        page_numbers = list(map(int, page_numbers.split(',')))

    progress_bar['maximum'] = len(file_paths)
    progress_bar['value'] = 0
    root.update_idletasks()

    threading.Thread(target=convert_files, args=(file_paths, output_dir, pdfa_type, title, author, dpi, page_numbers), daemon=True).start()

def show_help():
    help_text = """
Hier ist eine detaillierte Installationsanleitung für das erweiterte PDF-Konverter-Skript. Diese Anleitung führt Sie durch die erforderlichen Schritte, um das Skript auf Ihrem System einzurichten und auszuführen.

Installationsanleitung

1. Voraussetzungen

Stellen Sie sicher, dass Python auf Ihrem System installiert ist. Sie können Python von der offiziellen [Python-Website](https://www.python.org/) herunterladen und installieren.

2. Abhängigkeiten installieren

Das Skript benötigt mehrere Python-Bibliotheken. Sie können diese mit `pip` installieren. Öffnen Sie ein Terminal oder eine Eingabeaufforderung und führen Sie die folgenden Befehle aus:

pip install docx2pdf pandas python-pptx reportlab Pillow pytesseract pdf2image langdetect pywin32

3. Ghostscript installieren

Ghostscript wird benötigt, um PDFs in das PDF/A-Format zu konvertieren. Laden Sie Ghostscript von der offiziellen [Ghostscript-Website](https://www.ghostscript.com/) herunter und installieren Sie es. Stellen Sie sicher, dass der Pfad zu `gswin64c.exe` oder `gswin32c.exe` in Ihrer PATH-Umgebungsvariable enthalten ist.

4. Tesseract OCR installieren

Tesseract wird für die OCR-Funktionalität benötigt. Laden Sie Tesseract von der offiziellen [Tesseract-Website](https://github.com/tesseract-ocr/tesseract) herunter und installieren Sie es. Stellen Sie sicher, dass der Pfad zu `tesseract.exe` in Ihrer PATH-Umgebungsvariable enthalten ist.

5. Skript herunterladen

Laden Sie das Skript herunter oder kopieren Sie es in eine Datei namens `pdf_converter.py`.

6. Skript ausführen

Navigieren Sie im Terminal oder in der Eingabeaufforderung zu dem Verzeichnis, in dem sich `pdf_converter.py` befindet, und führen Sie das Skript aus:

python pdf_converter.py

7. GUI verwenden

Nach dem Starten des Skripts wird ein GUI-Fenster angezeigt. Folgen Sie diesen Schritten, um das Skript zu verwenden:

- Wählen Sie die Dateien aus, die Sie konvertieren möchten, indem Sie auf die Schaltfläche "Datei(en) auswählen / Konvertieren" klicken.
- Wählen Sie den Ausgabeordner, in dem die konvertierten PDFs gespeichert werden sollen.
- Passen Sie die Einstellungen nach Bedarf an, z.B. PDF-Version, OCR-Einstellungen, Titel und Autor des PDFs.
- Klicken Sie auf "Datei(en) auswählen / Konvertieren", um den Konvertierungsprozess zu starten.

8. Protokolldatei

Das Skript erstellt eine Protokolldatei namens `converter.log`, in der alle Aktionen und Fehler protokolliert werden. Sie können diese Datei öffnen, um Details zu den Konvertierungsprozessen zu überprüfen.

Hinweise

- Stellen Sie sicher, dass alle Pfade zu den ausführbaren Dateien von Ghostscript und Tesseract korrekt sind.
- Wenn Sie auf Probleme stoßen, überprüfen Sie die Protokolldatei auf Fehlerhinweise.
- Das Skript unterstützt mehrere Dateitypen, einschließlich DOCX, XLSX, PPTX, CSV, TXT, JPG, JPEG, PNG und PDF.

    """
    messagebox.showinfo("Hilfe", help_text)

def create_gui():
    global root, progress_bar, progress_label, pdf_version_var, ocr_var, ocr_language_var, title_var, author_var, dpi_var, page_numbers_var

    root = tk.Tk()
    root.title("PDF Konverter")
    root.geometry("500x700")

    frame = tk.Frame(root)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    tk.Label(frame, text="Dokument(e) auswählen und konvertieren:", font=("Arial", 12, "bold")).pack(pady=5)
    tk.Label(frame, text="*.docx *.xlsx *.pptx *.csv *.txt *.jpg *.jpeg *.png *.pdf").pack(pady=5)

    pdf_version_var = StringVar(value="Standard-PDF")
    tk.Label(frame, text="Konvertiere zu PDF-Version:").pack(pady=5)
    tk.OptionMenu(frame, pdf_version_var, "Standard-PDF", "PDF/A-1b", "PDF/A-2b", "PDF/A-3b").pack(pady=5)

    ocr_var = tk.BooleanVar(value=False)
    tk.Checkbutton(frame, text="OCR auf Bilder/PDFs anwenden", variable=ocr_var).pack(pady=5)

    ocr_language_var = StringVar(value='deu+eng')
    tk.Label(frame, text="OCR-Sprache (optional):").pack(pady=5)
    tk.Entry(frame, textvariable=ocr_language_var).pack(pady=5)

    title_var = StringVar()
    tk.Label(frame, text="PDF-Titel:").pack(pady=5)
    tk.Entry(frame, textvariable=title_var).pack(pady=5)

    author_var = StringVar()
    tk.Label(frame, text="PDF-Autor:").pack(pady=5)
    tk.Entry(frame, textvariable=author_var).pack(pady=5)

    dpi_var = IntVar(value=300)
    tk.Label(frame, text="OCR-DPI:").pack(pady=5)
    tk.Entry(frame, textvariable=dpi_var).pack(pady=5)

    page_numbers_var = StringVar()
    tk.Label(frame, text="Seiten für OCR (kommagetrennt, optional):").pack(pady=5)
    tk.Entry(frame, textvariable=page_numbers_var).pack(pady=5)

    tk.Button(frame, text="Datei(en) auswählen / Konvertieren", command=start_conversion).pack(pady=10)

    progress_bar = Progressbar(frame, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=10)

    progress_label = tk.Label(frame, text="")
    progress_label.pack(pady=5)

    tk.Button(frame, text="Protokoll öffnen", command=lambda: os.startfile("converter.log")).pack(pady=5)
    tk.Button(frame, text="Hilfe", command=show_help).pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    create_gui()

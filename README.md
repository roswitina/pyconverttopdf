## Installationsanleitung für das PDF-Konverter-Skript

# Systemanforderungen
Betriebssystem: Windows (aufgrund der Verwendung von "win32com.client"), der Rest sollte auch unter Linux und MacOS funktionieren

Speicherplatz: Mindestens 500 MB freier Speicherplatz

Python: Python 3.7 oder höher

# 1. Voraussetzungen
Stellen Sie sicher, dass Python auf Ihrem System installiert ist. Sie können Python von der offiziellen Python-Website herunterladen und installieren.

# 2. Abhängigkeiten installieren
Das Skript benötigt mehrere Python-Bibliotheken. Öffnen Sie ein Terminal oder eine Eingabeaufforderung und führen Sie die folgenden Befehle aus:

``pip install docx2pdf pandas python-pptx reportlab Pillow pytesseract pdf2image langdetect pywin32``

# 3. Ghostscript installieren
Ghostscript wird benötigt, um PDFs in das PDF/A-Format zu konvertieren. Laden Sie Ghostscript von der offiziellen Ghostscript-Website herunter und installieren Sie es. Stellen Sie sicher, dass der Pfad zu gswin64c.exe oder gswin32c.exe in Ihrer PATH-Umgebungsvariable enthalten ist.

# 4. Tesseract OCR installieren
Tesseract wird für die OCR-Funktionalität benötigt. Laden Sie Tesseract von der offiziellen Tesseract-Website herunter und installieren Sie es. Stellen Sie sicher, dass der Pfad zu tesseract.exe in Ihrer PATH-Umgebungsvariable enthalten ist.

# 5. Skript herunterladen
Laden Sie das Skript herunter oder kopieren Sie es in eine Datei namens pdf_converter.py.

# 6. Skript ausführen
Navigieren Sie im Terminal oder in der Eingabeaufforderung zu dem Verzeichnis, in dem sich pdf_converter.py befindet, und führen Sie das Skript aus:

``python pdf_converter.py``

# 7. GUI verwenden
Nach dem Starten des Skripts wird ein GUI-Fenster angezeigt. Folgen Sie diesen Schritten, um das Skript zu verwenden:

Dateien auswählen: Klicken Sie auf "Datei(en) auswählen / Konvertieren", um die zu konvertierenden Dateien auszuwählen.

Ausgabeordner wählen: Wählen Sie den Ordner, in dem die konvertierten PDFs gespeichert werden sollen.

Einstellungen anpassen: Passen Sie die Einstellungen nach Bedarf an, z.B. PDF-Version, OCR-Einstellungen, Titel und Autor des PDFs.

Konvertierung starten: Klicken Sie erneut auf "Datei(en) auswählen / Konvertieren", um den Konvertierungsprozess zu starten.

# 8. Protokolldatei
Das Skript erstellt eine Protokolldatei namens converter.log, in der alle Aktionen und Fehler protokolliert werden. Sie können diese Datei öffnen, um Details zu den Konvertierungsprozessen zu überprüfen.

# Fehlerbehebung
Fehlende Abhängigkeiten: Stellen Sie sicher, dass alle erforderlichen Bibliotheken installiert sind.

Pfadprobleme: Überprüfen Sie, ob die Pfade zu Ghostscript und Tesseract korrekt in der PATH-Umgebungsvariable eingetragen sind.

Berechtigungen: Stellen Sie sicher, dass Sie die erforderlichen Berechtigungen haben, um Dateien zu lesen und zu schreiben.

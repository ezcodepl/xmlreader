import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkhtmlview import HTMLLabel
import xml.etree.ElementTree as ET
import tempfile
import webbrowser
import base64
import os
import io
import zipfile

# === Pomocnicze ===

def strip_ns(tag):
    return tag.split('}', 1)[1] if '}' in tag else tag

def is_base64_string(s):
    try:
        return len(s) > 100 and base64.b64encode(base64.b64decode(s)).decode()[:100] in s[:110]
    except Exception:
        return False

def guess_office_extension(file_bytes):
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            namelist = z.namelist()
            if any(name.startswith('word/') for name in namelist): return '.docx'
            elif any(name.startswith('xl/') for name in namelist): return '.xlsx'
            elif any(name.startswith('ppt/') for name in namelist): return '.pptx'
            else: return '.zip'
    except: pass

    ole_magic = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
    if file_bytes.startswith(ole_magic):
        header_sample = file_bytes[512:2048].decode('latin1', errors='ignore').lower()
        if 'worddocument' in header_sample: return '.doc'
        elif 'workbook' in header_sample: return '.xls'
        elif 'powerpoint document' in header_sample: return '.ppt'
        else: return '.ole'
    try:
        text_sample = file_bytes[:2048].decode('utf-8')
        if not text_sample.lstrip().startswith('<?xml'): return '.txt'
    except: pass
    return 'nieznany'

def guess_extension_from_bytes(data_bytes):
    header = data_bytes[:100].lstrip()
    if header.startswith(b'%PDF-'): return '.pdf'
    elif header.startswith(b'\xFF\xD8\xFF'): return '.jpg'
    elif header.startswith(b'\x89PNG\r\n\x1a\n'): return '.png'
    elif header.startswith(b'PK\x03\x04'): return guess_office_extension(data_bytes)
    elif header.startswith(b'<?xml') or header.startswith(b'<'): return '.xml'
    elif header.startswith(b'From:') or b'\r\nFrom:' in data_bytes[:200] or b'\nFrom:' in data_bytes[:200]: return '.eml'
    elif data_bytes.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'):
        sample = data_bytes[:2048].lower()
        if b'outlook message' in sample or b'microsoft outlook' in sample: return '.msg'
        return guess_office_extension(data_bytes)
    else: return '.nieznany'

# === Kluczowa funkcja ===

def extract_all_text_elements(root, skip_signature_blocks=False):
    attachments = []
    lines = []

    def recurse(elem, depth=0, parent_filename=None):
        tag = strip_ns(elem.tag)
        if skip_signature_blocks and tag in ("SignatureValue", "X509Certificate"):
            return  # pomijamy cały ten blok

        text = (elem.text or "").strip()

        # Sprawdź atrybut nazwaPliku w elemencie (np. str:Zalacznik)
        filename = elem.attrib.get("nazwaPliku") or elem.attrib.get("Nazwa") or parent_filename

        if text and is_base64_string(text):
            current_filename = filename or f"zalacznik_{len(attachments) + 1}"
            b64_clean = "".join(text.split())
            try:
                decoded_bytes = base64.b64decode(b64_clean)
                ext = os.path.splitext(current_filename)[1].lower()
                if not ext:
                    ext = guess_extension_from_bytes(decoded_bytes)
                    current_filename += ext
            except Exception:
                if not os.path.splitext(current_filename)[1]:
                    current_filename += '.bin'

            attachments.append((current_filename, text))
            indent = "  " * depth
            lines.append(f"{indent}{tag}:")
            lines.append(f"{indent}  Nazwa załącznika: {current_filename}")

        elif text:
            indent = "  " * depth
            if tag == "Informacja":
                lines.append(f"{indent}## {text}")
            elif ':' in text:
                lines.append(f"{indent}{text}")
            else:
                lines.append(f"{indent}{tag}: {text}")

        for child in elem:
            recurse(child, depth + 1, filename)

    recurse(root)
    return lines, attachments

# === HTML GENERATOR ===

def generate_html_from_text_lines(lines, filename=None, font="Arial", font_size="14"):
    filtered_lines = [line for line in lines if line.strip()]
    content_blocks = "\n".join(
        f"""<div class="mainTxtContainer"><pre class="mainTxt">{line}</pre></div>"""
        for line in filtered_lines
    )
    filename_display = f"<p><b>Nazwa pliku:</b> {filename}</p>" if filename else ""
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{
                margin: 0;
                padding: 10px;
                font-family: {font};
                background: #f0f0f0;
                overflow-y: scroll;
            }}
            .mainTxt {{
                font-size: {font_size}px;
                line-height: 1.2;
                margin: 2px 0;
                white-space: pre-wrap;
            }}
        </style>
    </head>
    <body>
        <h3>Dokument XML</h3>
        <h4>Zawartość pliku:</h4>
        {filename_display}
        {content_blocks}
    </body>
    </html>
    """
    return html

# === APLIKACJA ===

class XMLViewerApp:
    def __init__(self, root):
        self.root = root
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.root.geometry("1500x900")
        self.root.title("🧾 Uniwersalny podgląd dokumentu XML")

        self.frame = ctk.CTkFrame(root)
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.controls_frame = ctk.CTkFrame(self.frame)
        self.controls_frame.pack(fill="x", padx=10, pady=(5, 5))

        self.open_button = ctk.CTkButton(self.controls_frame, text="📂 Wczytaj plik XML", command=self.load_xml)
        self.open_button.pack(side="left", padx=5)

        self.print_button = ctk.CTkButton(self.controls_frame, text="🖨 Podgląd do wydruku", command=self.print_html)
        self.print_button.pack(side="left", padx=5)

        # Dodatkowy przycisk do zapisu załączników
        self.save_attachments_button = ctk.CTkButton(self.controls_frame, text="💾 Zapisz załączniki", command=self.save_attachments)
        self.save_attachments_button.pack(side="left", padx=5)

        self.font_family = ctk.StringVar(value="Arial")
        self.font_size = ctk.StringVar(value="14")
        self.skip_signature = ctk.BooleanVar(value=True)

        self.font_menu = ctk.CTkOptionMenu(self.controls_frame, values=["Arial", "Courier New", "Times New Roman"], variable=self.font_family, command=self.refresh_html)
        self.font_menu.pack(side="left", padx=5)

        self.size_menu = ctk.CTkOptionMenu(self.controls_frame, values=["12", "14", "16", "18"], variable=self.font_size, command=self.refresh_html)
        self.size_menu.pack(side="left", padx=5)

        self.skip_checkbox = ctk.CTkCheckBox(self.controls_frame, text="Pomiń SignatureValue i X509Certificate", variable=self.skip_signature, command=self.refresh_html)
        self.skip_checkbox.pack(side="left", padx=5)

        self.attachments_info_label = ctk.CTkLabel(self.frame, text="", font=("Arial", 14), anchor="w", justify="left", wraplength=1400)
        self.attachments_info_label.pack(fill="x", padx=10, pady=(5, 5))

        self.html_frame = ctk.CTkFrame(self.frame, fg_color="white")
        self.html_frame.pack(fill="both", expand=True)

        self.html_view = HTMLLabel(self.html_frame, html="<p>Tu pojawi się podgląd dokumentu</p>",
                                   background="white", foreground="black",
                                   font=("Arial", 14), padx=20, pady=10)
        self.html_view.pack(fill="both", expand=True, padx=5, pady=5)

        self.current_html = ""
        self.attachments = []
        self.current_xml_root = None
        self.current_filename = ""

    def load_xml(self):
        file_path = filedialog.askopenfilename(filetypes=[("Pliki XML", "*.xml")])
        if not file_path:
            return
        try:
            tree = ET.parse(file_path)
            self.current_xml_root = tree.getroot()
            self.current_filename = os.path.basename(file_path)
            self.refresh_html()

        except Exception as e:
            messagebox.showerror("Błąd", f"Błąd przetwarzania XML:\n{e}")

    def refresh_html(self, *_):
        if self.current_xml_root is None:
            return

        lines, attachments = extract_all_text_elements(self.current_xml_root, skip_signature_blocks=self.skip_signature.get())
        html = generate_html_from_text_lines(lines, filename=self.current_filename,
                                             font=self.font_family.get(),
                                             font_size=self.font_size.get())

        self.current_html = html
        self.attachments = attachments
        self.html_view.set_html(html)

        if attachments:
            self.show_attachments_info()
        else:
            self.attachments_info_label.configure(text="")

    def show_attachments_info(self):
        info_lines = ["📎 Dokument zawiera załączniki:\n"]
        for i, (filename, b64data) in enumerate(self.attachments, start=1):
            try:
                data = base64.b64decode("".join(b64data.split()))
                ext = guess_extension_from_bytes(data)
                info_lines.append(f"{i}. {filename} ({ext})")
            except:
                info_lines.append(f"{i}. {filename} (nieznany typ)")
        self.attachments_info_label.configure(text="\n".join(info_lines))

    def print_html(self):
        if not self.current_html:
            messagebox.showwarning("Brak danych", "Najpierw wczytaj plik XML.")
            return
        with tempfile.NamedTemporaryFile('w', delete=False, suffix='.html', encoding='utf-8') as f:
            f.write(self.current_html)
            webbrowser.open(f.name)

    def save_attachments(self):
        if not self.attachments:
            messagebox.showinfo("Informacja", "Brak załączników do zapisania.")
            return

        folder = filedialog.askdirectory(title="Wybierz folder do zapisu załączników")
        if not folder:
            return  # użytkownik anulował wybór folderu

        saved_files = []
        for filename, b64data in self.attachments:
            try:
                data = base64.b64decode("".join(b64data.split()))
                path = os.path.join(folder, filename)
                with open(path, 'wb') as f:
                    f.write(data)
                saved_files.append(filename)
            except Exception as e:
                messagebox.showerror("Błąd", f"Nie udało się zapisać {filename}:\n{e}")
                return

        messagebox.showinfo("Zapisano", f"Zapisano {len(saved_files)} załączników w:\n{folder}")

# === Start ===
if __name__ == "__main__":
    root = ctk.CTk()
    app = XMLViewerApp(root)
    root.mainloop()

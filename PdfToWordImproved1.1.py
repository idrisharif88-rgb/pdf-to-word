# === OCR GUI v3.5 Jet-Black + Neon-Green Edition ===
# Part 1 / 10  –  Imports, dark theme, helpers, GUI skeleton
from google.oauth2 import service_account
import os
import re
import gc
import shutil
import logging
import platform
import threading
import multiprocessing
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pypdf import PdfReader
from concurrent.futures import ThreadPoolExecutor, as_completed
from google.api_core.client_options import ClientOptions
from google.cloud import documentai  # type: ignore

IS_WINDOWS = platform.system() == "Windows"
if IS_WINDOWS:
    import winsound  # type: ignore

# --------------------------------------------------------------------
# Configuration: read from the environment, with a .env file if present.
def _load_dotenv() -> None:
    """Minimal .env loader so no extra dependency is required."""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.isfile(env_path):
        return
    with open(env_path, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip("'\""))


_load_dotenv()

GCP_PROJECT_ID = os.environ.get("GCP_PROJECT_ID", "")
GCP_LOCATION = os.environ.get("GCP_LOCATION", "eu")
GCP_PROCESSOR_ID = os.environ.get("GCP_PROCESSOR_ID", "")
# Optional: path to a service-account JSON key. When unset, Application
# Default Credentials are used (gcloud auth application-default login).
GCP_CREDENTIALS_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "")

# --------------------------------------------------------------------
# Logging
logging.basicConfig(
    filename=os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "ocr_gui.log"
    ),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


# --------------------------------------------------------------------
# Cross-platform notification / file-reveal helpers
def notify_sound() -> None:
    """Play a short completion sound, silently doing nothing if unavailable."""
    try:
        if IS_WINDOWS:
            winsound.MessageBeep()
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
        elif shutil.which("canberra-gtk-play"):
            subprocess.Popen(["canberra-gtk-play", "-i", "complete"])
        elif shutil.which("paplay"):
            subprocess.Popen(
                ["paplay", "/usr/share/sounds/freedesktop/stereo/complete.oga"]
            )
        else:
            print("\a", end="", flush=True)
    except Exception as e:
        logging.warning(f"Could not play notification sound: {e}")


def reveal_in_file_manager(path: str) -> None:
    """Open the folder containing `path` in the desktop file manager."""
    if IS_WINDOWS:
        subprocess.Popen(f'explorer /select,"{path}"')
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", "-R", path])
    else:
        subprocess.Popen(["xdg-open", os.path.dirname(path)])

# --------------------------------------------------------------------
# Arabic-aware text cleaner (regex based, offline)
def clean_text_smart(text: str) -> str:
    text = re.sub(r"[\\(\\)\\[\\]\\{\\}\\<\\>]", "", text)
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[\u0000-\u001F\u202A-\u202E]", "", text)
    text = text.replace(" ,", ",").replace(" .", ".")
    text = text.replace(" ?", "؟").replace(";", "؛")
    text = text.replace("|", "").replace("¬", "")
    return text.strip()

# --------------------------------------------------------------------
# Arabic paragraph styling
def apply_arabic_paragraph_format(paragraph):
    try:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p = paragraph._element
        p.set(qn("w:bidi"), "true")
        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = Pt(24)
            run.bold = True
            rPr = run._element.get_or_add_rPr()
            rtl_tag = OxmlElement("w:rtl")
            rtl_tag.set(qn("w:val"), "true")
            rPr.append(rtl_tag)
    except Exception as e:
        logging.warning(f"Could not apply Arabic paragraph format: {e}")

# --------------------------------------------------------------------
# Page-range parser
def parse_page_range(page_input: str, total_pages: int) -> List[int]:
    page_input = page_input.strip().lower()
    if page_input == "all":
        return list(range(1, total_pages + 1))
    pages = set()
    for part in page_input.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-")
            pages.update(range(int(start), int(end) + 1))
        else:
            pages.add(int(part))
    return sorted(p for p in pages if 1 <= p <= total_pages)

# --------------------------------------------------------------------
# Jet-black + neon-green theme
def apply_dark_theme(root):
    root.configure(bg="#000000")
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(
        "TButton",
        background="#00FF00",
        foreground="#000000",
        font=("Consolas", 10, "bold"),
        padding=5,
    )
    style.map("TButton",
              foreground=[("active", "#000000")],
              background=[("active", "#00DD00")])
    style.configure("TProgressbar", troughcolor="#111111", background="#00FF00")

# --------------------------------------------------------------------
# Main GUI class header
class OCRApp:
    def __init__(self, root):
        apply_dark_theme(root)
        self.root = root
        self.root.title("Enterprise OCR v3.5 — Jet-Black Edition")
        self.root.geometry("820x640")
        self.root.resizable(False, False)

        if IS_WINDOWS:
            try:
                import ctypes
                ctypes.windll.user32.ShowWindow(
                    ctypes.windll.kernel32.GetConsoleWindow(), 0
                )
            except Exception:
                pass

        # Cloud configuration (see .env / .env.example)
        self.project_id = GCP_PROJECT_ID
        self.location = GCP_LOCATION
        self.processor_id = GCP_PROCESSOR_ID

        self.cancel_requested = False

        # --- GUI layout ---
        tk.Label(root, text="Input File:", fg="#00FF00", bg="#000000",
                 font=("Consolas", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        self.file_entry = tk.Entry(root, width=80, bg="#111111", fg="#00FF00",
                                   insertbackground="#00FF00")
        self.file_entry.pack(padx=10, pady=2)
        ttk.Button(root, text="Browse", command=self.browse_file).pack(pady=2)

        tk.Label(root, text="Output DOCX File:", fg="#00FF00", bg="#000000",
                 font=("Consolas", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        self.output_entry = tk.Entry(root, width=80, bg="#111111", fg="#00FF00",
                                     insertbackground="#00FF00")
        self.output_entry.pack(padx=10, pady=2)

        tk.Label(root, text="Page Range (e.g., all, 1-5, 2,5,7):",
                 fg="#00FF00", bg="#000000", font=("Consolas", 10, "bold"))\
            .pack(anchor="w", padx=10, pady=5)
        self.page_entry = tk.Entry(root, width=40, bg="#111111",
                                   fg="#00FF00", insertbackground="#00FF00")
        self.page_entry.insert(0, "all")
        self.page_entry.pack(padx=10, pady=2)

        self.quality_var = tk.BooleanVar(value=True)
        tk.Checkbutton(root, text="Enable Quality Scoring",
                       variable=self.quality_var, bg="#000000",
                       fg="#00FF00", selectcolor="#111111",
                       font=("Consolas", 10)).pack(anchor="w", padx=20, pady=5)

        self.cleanup_var = tk.BooleanVar(value=True)
        tk.Checkbutton(root, text="Enable Smart Text Cleanup",
                       variable=self.cleanup_var, bg="#000000",
                       fg="#00FF00", selectcolor="#111111",
                       font=("Consolas", 10)).pack(anchor="w", padx=20, pady=2)

        # Buttons
        btn_frame = tk.Frame(root, bg="#000000")
        btn_frame.pack(pady=10)
        self.run_btn = ttk.Button(btn_frame, text="Run OCR", command=self.run_ocr)
        self.run_btn.grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel_ocr).grid(row=0, column=1, padx=5)

        # Status + progress bar + log console
        self.status_label = tk.Label(root, text="Ready.", anchor="w",
                                     fg="#00FF00", bg="#000000",
                                     font=("Consolas", 10, "bold"))
        self.status_label.pack(fill="x", padx=10, pady=8)
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(root, variable=self.progress_var, maximum=100)
        self.progress.pack(fill="x", padx=15, pady=10)
        self.log_box = tk.Text(root, height=8, bg="#000000", fg="#00FF00",
                               font=("Consolas", 9), state="disabled")
        self.log_box.pack(fill="x", padx=10, pady=10)

    # Simple logger for the GUI
    def gui_log(self, msg):
        self.root.after(0, lambda: self._append_log(msg))

    def _append_log(self, msg):
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n")
        total_lines = int(self.log_box.index('end-1c').split('.')[0])
        if total_lines > 200:
            self.log_box.delete("1.0", f"{total_lines - 180}.0")
        self.log_box.see("end")
        self.log_box.config(state="disabled")


    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF or Image", "*.pdf *.jpg *.jpeg")])
        if filename:
            # 1. Update the input field with the chosen file
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filename)
            
            # 2. Extract the file's directory and its name
            input_dir = os.path.dirname(filename)  # Gets the folder containing the file
            base_name = os.path.splitext(os.path.basename(filename))[0]
            
            # 3. Build the path using the same folder as the input
            output_path = os.path.join(input_dir, f"{base_name}.docx")
            
            # 4. Update the GUI output field
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_path)
            
            # Log the destination for the user
            self.gui_log(f"Output path set to: {input_dir}")
    def run_ocr(self):
        """Start OCR process when 'Run OCR' is clicked."""
        if self.cancel_requested:
            self.cancel_requested = False  # reset cancel flag

        file_path = self.file_entry.get().strip()
        output_path = self.output_entry.get().strip()
        page_input = self.page_entry.get().strip()
        enable_quality = self.quality_var.get()
        cleanup = self.cleanup_var.get()

        if not file_path or not output_path:
            messagebox.showerror("Error", "Please select input and output files.")
            return

        try:
            # Determine MIME type
            mime_type = "application/pdf" if file_path.lower().endswith(".pdf") else "image/jpeg"
            total_pages = len(PdfReader(file_path).pages) if mime_type == "application/pdf" else 1
            selected_pages = parse_page_range(page_input, total_pages)

            if not selected_pages:
                messagebox.showwarning("Page Range", "No valid pages selected.")
                return

            self.gui_log(f"Starting OCR for pages: {selected_pages}")
            self.safe_status_update("Starting OCR process...")
            self.safe_progress_update(0)
            self.run_btn.config(state="disabled")

            # Start threaded OCR
            t = threading.Thread(
                target=self.threaded_ocr,
                args=(
                    self.project_id,
                    self.location,
                    self.processor_id,
                    file_path,
                    mime_type,
                    output_path,
                    selected_pages,
                    enable_quality,
                    cleanup,
                ),
                daemon=True,
            )
            t.start()

        except Exception as e:
            logging.error(f"Error during OCR start: {e}")
            messagebox.showerror("Error", str(e))
            
    
    # --- Cancel OCR ---
    def cancel_ocr(self):
        self.cancel_requested = True
        self.safe_status_update("Canceling process...")
        self.gui_log("User requested cancel.")
    
    def safe_status_update(self, msg):
        self.root.after(0, lambda: self.status_label.config(text=msg))
        
    def safe_progress_update(self, value):
        self.root.after(0, lambda: self.progress_var.set(value))

    # Tk is not thread-safe: dialogs raised from worker threads must be
    # marshalled back onto the main loop.
    def safe_info(self, title, msg):
        self.root.after(0, lambda: messagebox.showinfo(title, msg))

    def safe_warning(self, title, msg):
        self.root.after(0, lambda: messagebox.showwarning(title, msg))

    def safe_error(self, title, msg):
        self.root.after(0, lambda: messagebox.showerror(title, msg))
        
    
    def threaded_ocr(self, project_id, location, processor_id,
                     file_path, mime_type, output_path, selected_pages,
                     enable_quality, cleanup):
        """Run OCR in background threads, update progress bar, and handle resume."""
        try:
            if file_path.lower().endswith(".pdf"):
                total_pages = len(PdfReader(file_path).pages)
            else:
                total_pages = 1

            
            output_dir = os.path.dirname(os.path.abspath(output_path))
            os.makedirs(output_dir, exist_ok=True)
            base_name = os.path.splitext(os.path.basename(output_path))[0]
            batch_size = 15
            total_parts = (len(selected_pages) + batch_size - 1) // batch_size
            completed = 0

            # Smart resume detection
            existing_parts = [f for f in os.listdir(output_dir)
                              if f.startswith(base_name + "_part_") and f.endswith(".docx")]
            start_index = len(existing_parts) * batch_size
            if existing_parts:
                self.gui_log(f"Resuming from part {len(existing_parts)+1}")
            else:
                self.gui_log("Starting new OCR job...")

            # Auto thread scaling
            max_workers = max(2, min(6, multiprocessing.cpu_count() - 1))
            self.gui_log(f"Using {max_workers} threads")

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = []
                for i in range(start_index, len(selected_pages), batch_size):
                    if self.cancel_requested:
                        break
                    part_num = (i // batch_size) + 1
                    batch_pages = selected_pages[i:i + batch_size]
                    part_output = os.path.join(output_dir, f"{base_name}_part_{part_num}.docx")
                    futures.append(executor.submit(
                        self.process_part,
                        project_id, location, processor_id,
                        file_path, mime_type, batch_pages,
                        enable_quality, part_output, part_num, cleanup
                    ))

                for future in as_completed(futures):
                    if self.cancel_requested:
                        break
                    future.result()
                    completed += 1
                    progress = (completed / total_parts) * 100
                    self.safe_progress_update(progress)
                    self.safe_status_update(f"Processed {completed}/{total_parts} parts...")
                    self.gui_log(f"Completed part {completed}/{total_parts}")

            if not self.cancel_requested:
                self.merge_and_finish(output_dir, base_name)
            else:
                self.safe_status_update("OCR canceled by user.")
                self.gui_log("Processing canceled.")
        except Exception as e:
            logging.error(f"OCR error: {e}")
            self.safe_error("Error", str(e))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state="normal"))
            
     
    def process_part(self, project_id, location, processor_id,
                     file_path, mime_type, pages, enable_quality,
                     part_output, part_num, cleanup):
        """Handle one OCR batch (part)."""
        if self.cancel_requested:
            return
        self.safe_status_update(f"Processing Part {part_num} (Pages {pages[0]}–{pages[-1]})...")
        self.gui_log(f"Starting part {part_num} for pages {pages}")
        doc = extract_batch(project_id, location, processor_id, file_path,
                            mime_type, pages, enable_quality, cleanup)
        doc.save(part_output)
        self.gui_log(f"Saved {part_output}")
        del doc
        gc.collect()      
        
        
        
    def merge_and_finish(self, output_dir, base_name):
        """Merge all temporary part files, delete them, and open the merged document."""
        try:
            merged_path = os.path.join(output_dir, f"{base_name}_merged.docx")
            merged_doc = Document()
            part_files = []

            # Collect parts and order them numerically. Lexical sorting would
            # give part_1, part_10, part_2, ... and scramble long documents.
            part_pattern = re.compile(
                rf"^{re.escape(base_name)}_part_(\d+)\.docx$"
            )
            numbered_parts = []
            for filename in os.listdir(output_dir):
                match = part_pattern.match(filename)
                if match:
                    numbered_parts.append((int(match.group(1)), filename))
            numbered_parts.sort()

            for _, filename in numbered_parts:
                part_path = os.path.join(output_dir, filename)
                part_files.append(part_path)
                self.gui_log(f"Merging {filename}")
                try:
                    part_doc = Document(part_path)
                    for element in part_doc.element.body:
                        merged_doc.element.body.append(element)
                except Exception as e:
                    self.gui_log(f"Error reading {filename}: {e}")

            if not part_files:
                self.safe_warning("Merge", "No part files found to merge.")
                return

            merged_doc.save(merged_path)
            self.gui_log(f"Merged file saved: {merged_path}")

            # Delete temp part files
            for f in part_files:
                try:
                    os.remove(f)
                    self.gui_log(f"Deleted temporary file: {os.path.basename(f)}")
                except Exception as e:
                    self.gui_log(f"Could not delete {os.path.basename(f)}: {e}")

            self.safe_progress_update(100)
            self.safe_status_update("All parts merged successfully!")
            notify_sound()
            self.safe_info("Success", f"OCR completed!\nFinal file:\n{merged_path}")
            self.root.after(3000, lambda: self.safe_status_update("Ready."))
            self.root.after(3000, lambda: self.run_btn.config(state="normal"))

            # Reveal the finished file in the desktop file manager
            try:
                reveal_in_file_manager(merged_path)
                self.gui_log("Opened final file in the file manager.")
            except Exception as e:
                self.gui_log(f"Auto-open failed: {e}")

        except Exception as e:
            logging.error(f"Merge error: {e}")
            self.safe_error("Error", f"Error during merging: {e}")  
    
    # === Part 2 / 10 — OCR extraction logic ===

def extract_batch(project_id, location, processor_id,
                  file_path, mime_type, pages,
                  enable_quality, cleanup_enabled=True) -> Document:
    """Run Google Document AI OCR on a batch of pages and return a Word Document."""
    opts = ClientOptions(api_endpoint=f"{location}-documentai.googleapis.com")

    # Prefer an explicit service-account key when one is configured, otherwise
    # fall back to Application Default Credentials.
    if GCP_CREDENTIALS_FILE:
        if not os.path.isfile(GCP_CREDENTIALS_FILE):
            raise FileNotFoundError(
                f"Service-account key not found: {GCP_CREDENTIALS_FILE}\n"
                "Fix GOOGLE_APPLICATION_CREDENTIALS in .env, or unset it to use "
                "'gcloud auth application-default login'."
            )
        creds = service_account.Credentials.from_service_account_file(
            GCP_CREDENTIALS_FILE
        )
        client = documentai.DocumentProcessorServiceClient(
            credentials=creds, client_options=opts
        )
    else:
        client = documentai.DocumentProcessorServiceClient(client_options=opts)

    name = client.processor_path(project_id, location, processor_id)


    with open(file_path, "rb") as f:
        file_content = f.read()

    # Configure OCR
    ocr_config = documentai.OcrConfig(
        enable_native_pdf_parsing=False,
        enable_image_quality_scores=enable_quality,
        enable_symbol=True,
        hints=documentai.OcrConfig.Hints(language_hints=["ar", "en"]),
    )
    process_options = documentai.ProcessOptions(
        ocr_config=ocr_config,
        individual_page_selector=documentai.ProcessOptions.IndividualPageSelector(pages=pages),
    )
    request = documentai.ProcessRequest(
        name=name,
        raw_document=documentai.RawDocument(content=file_content, mime_type=mime_type),
        process_options=process_options,
    )

    result = client.process_document(request=request)
    document = result.document
    doc = Document()

    # Default font
    try:
        style = doc.styles["Normal"]
        style.font.name = "Arial"
        style.font.size = Pt(24)
        style.font.bold = True
    except Exception as e:
        logging.warning(f"Normal style font not set: {e}")

    # Extract text per page
    for i, page in enumerate(document.pages, start=1):
        page_num = pages[i - 1]
        p_header = doc.add_paragraph(f"Page {page_num}")
        p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p_header.runs:
            run.bold = True
            run.font.size = Pt(18)
            run.font.name = "Arial"

        for block in page.blocks:
            block_text = "".join(
                document.text[int(seg.start_index):int(seg.end_index)]
                for seg in block.layout.text_anchor.text_segments
            )
            refined = clean_text_smart(block_text) if cleanup_enabled else block_text
            if refined.strip():
                paragraph = doc.add_paragraph(refined)
                apply_arabic_paragraph_format(paragraph)

        if i < len(document.pages):
            doc.add_page_break()

    return doc


# === Part 3 / 10 — File browsing, cancel logic, and small helpers ===

def safe_thread(func):
    """Decorator to run a function in a background thread."""
    def wrapper(self, *args, **kwargs):
        t = threading.Thread(target=func, args=(self, *args), kwargs=kwargs, daemon=True)
        t.start()
    return wrapper



    

    
    # --- Safe GUI update helpers ---
    

    
        # === Part 4 / 10 — Threaded OCR controller and batch handling ===

    


    


    # --- GUI log wrapper (safe thread) ---
    
        
        
    # === Part 5 / 10 — Merge .docx parts and auto-open final file ===

    



    # === Part 6 / 10 — Run OCR button logic ===

    








# === Part 7 / 10 — Application entry point ===
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = OCRApp(root)
        # bring window to front
        root.after(200, lambda: root.focus_force())
        root.mainloop()
    except Exception as e:
        logging.error(f"Fatal error: {e}")
        messagebox.showerror("Startup Error", f"Fatal error: {e}")



import os
import sys
import time
import threading
import subprocess
import win32com.client
import pythoncom
import webbrowser
from tkinter import *
from tkinter import filedialog, scrolledtext
from tkinterdnd2 import DND_FILES, TkinterDnD  # Requires `pip install tkinterdnd2`

def resource_path(relative_path):
    """Get absolute path to resource (for dev and PyInstaller)."""
    try:
        base_path = sys._MEIPASS  # PyInstaller temp folder
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# --- Document Conversion Logic ---

def force_kill_word():
    try:
        subprocess.run([
            "taskkill", "/F", "/IM", "WINWORD.EXE"
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        log(f"⚠️ Unable to force close Word: {e}")

def docx2pdf(input_path, output_path):
    try:
        word = win32com.client.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(os.path.abspath(input_path), ReadOnly=True)
        doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
        doc.Close(SaveChanges=False)
        word.Quit()

        log(f"✅ Converted: {output_path}")
    except Exception as e:
        log(f"❌ Error converting {input_path}: {e}")
    finally:
        force_kill_word()

# --- GUI Logic ---

selected_files = []

def log(message):
    status_text.insert(END, message + "\n")
    status_text.see(END)

def browse_files():
    global selected_files
    file_paths = filedialog.askopenfilenames(filetypes=[("Word Documents", "*.docx")])
    selected_files = list(file_paths)
    update_file_list()

def handle_drop(event):
    global selected_files
    try:
        paths = root.tk.splitlist(event.data)
        docx_files = [os.path.abspath(p.strip('{}')) for p in paths if p.lower().endswith(".docx")]
        selected_files = docx_files
        update_file_list()
        log(f"📂 Dropped {len(docx_files)} file(s). Ready to convert.")
    except Exception as e:
        log(f"❌ Error parsing dropped files: {e}")

def update_file_list():
    file_list.delete(0, END)
    for f in selected_files:
        file_list.insert(END, os.path.basename(f))

def start_conversion():
    if not selected_files:
        log("⚠️ No files selected.")
        return
    threading.Thread(target=convert_files, daemon=True).start()

def convert_files():
    pythoncom.CoInitialize()
    output_folder = os.path.join(os.getcwd(), "pdf")
    os.makedirs(output_folder, exist_ok=True)

    for input_path in selected_files:
        filename = os.path.basename(input_path)
        output_path = os.path.join(output_folder, filename.replace(".docx", ".pdf"))
        log(f"🔄 Converting: {filename} → {output_path}")
        docx2pdf(input_path, output_path)
        time.sleep(0.1)

    pythoncom.CoUninitialize()
    log("✅ All files converted. Finished!")

    if open_folder_after.get():
        webbrowser.open(output_folder)

# --- UI Setup ---

root = TkinterDnD.Tk()
root.title("Docx2PDF")
root.geometry("600x540")
root.configure(bg="white")

open_folder_after = BooleanVar(value=True)

custom_font = ("Calibri", 20, "bold")
title_label = Label(root, text="Docx", font=custom_font, fg="#0077ff", bg="white")
title_label.pack(pady=(10, 0))
title_2_label = Label(root, text="2PDF", font=custom_font, fg="red", bg="white")
title_2_label.pack(pady=(0, 10))

upload_frame = Frame(root, bg="#e9f3fc", bd=2, relief=SOLID)
upload_frame.pack(pady=5, padx=10, fill=BOTH, expand=False)

upload_label = Label(upload_frame, text="Drag & Drop or click to upload files", fg="#333", bg="#e9f3fc", font=("Calibri", 12))
upload_label.pack(pady=10)
upload_label.bind("<Button-1>", lambda e: browse_files())
upload_label.drop_target_register(DND_FILES)
upload_label.dnd_bind('<<Drop>>', handle_drop)

file_list = Listbox(upload_frame, font=("Calibri", 11), bg="white", activestyle='none', height=6)
file_list.pack(padx=10, pady=5, fill=X)

convert_button = Button(root, text="START", font=("Calibri", 14, "bold"), fg="white", bg="#0077ff", bd=0, padx=20, pady=10, command=start_conversion)
convert_button.pack(pady=10)

check_open = Checkbutton(root, text="Open folder when done", variable=open_folder_after, bg="white", font=("Calibri", 11))
check_open.pack()

status_text = scrolledtext.ScrolledText(root, height=10, font=("Courier", 10))
status_text.pack(fill=BOTH, padx=10, pady=(0,10), expand=True)

root.mainloop()
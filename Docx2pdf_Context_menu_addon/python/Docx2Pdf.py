import os
import sys
import subprocess
import threading
import comtypes.client
import pythoncom
import tkinter as tk
from tkinter import messagebox

class ConverterApp:
    def __init__(self, files):
        self.files = files
        self.total = len(files)
        self.current = 0
        self.final_log = []

        self.root = tk.Tk()
        self.root.title("Docx2Pdf - Converting")
        self.root.resizable(False, False)

        self.label = tk.Label(self.root, text="Starting conversion...", font=("Calibri", 12), padx=20, pady=20)
        self.label.pack()

        # Start conversion in a background thread
        threading.Thread(target=self.start_conversion, daemon=True).start()

        self.root.mainloop()

    def update_progress(self):
        self.label.config(text=f"Converting {self.current}/{self.total}")
        self.root.update_idletasks()

    def force_kill_word(self):
        try:
            subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except:
            pass

    def convert(self, input_path):
        pythoncom.CoInitialize()  # Initialize COM for this thread
        try:
            output_path = os.path.splitext(input_path)[0] + ".pdf"

            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            word.DisplayAlerts = False

            doc = word.Documents.Open(os.path.abspath(input_path), ReadOnly=True)
            doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
            doc.Close(SaveChanges=False)
            word.Quit()

            self.final_log.append(f"✅ Success: {output_path}")
        except Exception as e:
            self.final_log.append(f"❌ Error: {input_path} → {e}")
        finally:
            self.force_kill_word()
            pythoncom.CoUninitialize()  # Always uninitialize COM at end

    def start_conversion(self):
        for input_file in self.files:
            self.current += 1
            self.update_progress()

            if os.path.isfile(input_file) and input_file.lower().endswith((".doc", ".docx")):
                self.convert(input_file)
            else:
                self.final_log.append(f"⚠️ Skipped (not doc/docx): {input_file}")

        # Final done message
        self.label.config(text="✅ Conversion Complete!")
        self.root.update_idletasks()

        # Show result popup
        messagebox.showinfo("Conversion Completed", "\n".join(self.final_log))
        self.root.destroy()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        tk.Tk().withdraw()
        messagebox.showerror("Error", "Please provide at least one .doc or .docx file.")
        sys.exit(1)

    ConverterApp(sys.argv[1:])

import os
import time
import subprocess
import comtypes.client

def force_kill_word():
    """Forcefully terminates Microsoft Word process to release file locks."""
    try:
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        print(f"⚠️ Unable to force close Word: {e}")

def docx2pdf(input_path, output_path):
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False  # Run Word in the background
        word.DisplayAlerts = False  # Suppress any popups

        # Open the document in Read-Only mode
        doc = word.Documents.Open(os.path.abspath(input_path), ReadOnly=True)
        doc.SaveAs(os.path.abspath(output_path), FileFormat=17)  # 17 = wdFormatPDF
        doc.Close(SaveChanges=False)  # Ensure no changes are saved
        word.Quit()

        print(f"✅ Converted: {output_path}")
    except Exception as e:
        print(f"❌ Error converting {input_path}: {e}")
    finally:
        force_kill_word()  # Ensure Word is fully closed

def batch_convert():
    input_folder = os.path.join(os.getcwd(), "docx")
    output_folder = os.path.join(os.getcwd(), "pdf")

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    docx_files = [f for f in os.listdir(input_folder) if f.endswith(".docx") and not f.startswith("~$")]

    if not docx_files:
        print("⚠️ No valid .docx files found in the 'docx' folder.")
        return

    for filename in docx_files:
        input_path = os.path.join(input_folder, filename)
        output_path = os.path.join(output_folder, filename.replace(".docx", ".pdf"))

        print(f"📄 Converting: {filename} → {output_path}")
        docx2pdf(input_path, output_path)
        time.sleep(0)  # Small delay to prevent crashes

if __name__ == "__main__":
    batch_convert()

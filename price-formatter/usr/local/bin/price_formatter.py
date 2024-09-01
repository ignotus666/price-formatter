import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import zipfile
import os
import tempfile
import shutil
import re
from lxml import etree

def show_error(message):
    messagebox.showerror("Error", message)

def show_success(message):
    messagebox.showinfo("Success", message)

def replace_prices_in_docx(input_file):
    # Create a backup of the original file
    input_dir = os.path.dirname(input_file)
    input_filename = os.path.basename(input_file)
    backup_file = os.path.join(input_dir, f"{os.path.splitext(input_filename)[0]}_backup.docx")
    shutil.copy2(input_file, backup_file)

    # Create a temporary directory
    temp_dir = tempfile.mkdtemp()

    try:
        # Unzip the .docx file
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        document_xml = os.path.join(temp_dir, 'word', 'document.xml')

        if not os.path.exists(document_xml):
            show_error("document.xml not found in the .docx file.")
            return

        # Parse the XML content
        parser = etree.XMLParser(ns_clean=True, recover=True)
        tree = etree.parse(document_xml, parser)
        root = tree.getroot()
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        # Traverse paragraphs and process each run of text
        for para in root.findall('.//w:p', namespaces=ns):
            texts = para.findall('.//w:t', namespaces=ns)

            if texts:
                combined_text = ''.join([t.text for t in texts if t.text])

                # Replace valid prices while ignoring unwanted matches like "x2"
                def replace_price(match):
                    # Only replace if it's a standalone price
                    if not re.search(r'x\d', match.group(0)):  # Ensure no "x2" pattern
                        return f'€{match.group(1)}.{match.group(2)}'
                    return match.group(0)

                # Update the regex to carefully target prices
                new_combined_text = re.sub(r'(\d+),(\d{2})\s*€', replace_price, combined_text)

                # Update the text nodes back with the corrected price
                offset = 0
                for t in texts:
                    if t.text:
                        length = len(t.text)
                        t.text = new_combined_text[offset:offset+length]
                        offset += length

        # Write the modified XML back to the document.xml
        tree.write(document_xml, encoding='utf-8', xml_declaration=True)

        # Repackage the .docx file
        with zipfile.ZipFile(input_file, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_ref.write(file_path, arcname)

        show_success(f"Price format replacement completed successfully in {input_file}")

    except Exception as e:
        show_error(f"An error occurred: {e}")

    finally:
        # Clean up the temporary directory
        shutil.rmtree(temp_dir)

def on_file_drop(event):
    file_path = event.data.strip('{}')
    replace_prices_in_docx(file_path)

def main():
    root = TkinterDnD.Tk()
    root.title("Format those prices!")
    root.geometry("400x200")

    # Keep the window always on top
    root.attributes("-topmost", True)

    label = tk.Label(root, text="Drag and drop a DOCX file here", width=40, height=10, bg="black")
    label.pack(pady=20)

    # Bind the drag-and-drop event
    label.drop_target_register(DND_FILES)
    label.dnd_bind('<<Drop>>', on_file_drop)

    root.mainloop()

if __name__ == "__main__":
    main()

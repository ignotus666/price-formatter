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


def replace_prices_in_docx(input_file, direction='es_to_en'):
    input_dir = os.path.dirname(input_file)
    input_filename = os.path.basename(input_file)
    output_filename = f"{os.path.splitext(input_filename)[0]}_conv.docx"
    output_file = os.path.join(input_dir, output_filename)
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        document_xml = os.path.join(temp_dir, 'word', 'document.xml')
        if not os.path.exists(document_xml):
            show_error("document.xml not found in the .docx file.")
            return
        parser = etree.XMLParser(ns_clean=True, recover=True)
        tree = etree.parse(document_xml, parser)
        root = tree.getroot()
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        for para in root.findall('.//w:p', namespaces=ns):
            texts = para.findall('.//w:t', namespaces=ns)
            if texts:
                combined_text = ''.join([t.text for t in texts if t.text])
                if direction == 'es_to_en':
                    # ES to EN: 12,34€ -> €12.34, 12,1€ -> €12.1, 12€ -> €12
                    def replace_price_decimal2(match):
                        euros = match.group(1)
                        cents = match.group(2)
                        return f'€{euros}.{cents}'
                    def replace_price_decimal1(match):
                        euros = match.group(1)
                        cents = match.group(2)
                        return f'€{euros}.{cents}'
                    def replace_price_integer(match):
                        euros = match.group(1)
                        return f'€{euros}'
                    # Decimal prices (2 digits)
                    price_regex_decimal2 = r'(?<![\d.,])(\d{1,3}(?:\.\d{3})*),(\d{2})\s*€(?![\w.])'
                    # Decimal prices (1 digit)
                    price_regex_decimal1 = r'(?<![\d.,])(\d{1,3}(?:\.\d{3})*),(\d{1})\s*€(?![\w.])'
                    # Integer prices
                    price_regex_integer = r'(?<![\d.,])(\d{1,3}(?:\.\d{3})*)\s*€(?![\w.,\d])'
                    new_combined_text = re.sub(price_regex_decimal2, replace_price_decimal2, combined_text)
                    new_combined_text = re.sub(price_regex_decimal1, replace_price_decimal1, new_combined_text)
                    new_combined_text = re.sub(price_regex_integer, replace_price_integer, new_combined_text)
                else:
                    # EN to ES: €12.34 -> 12,34€, €12.1 -> 12,1€, €12 -> 12€
                    def replace_price_decimal2(match):
                        euros = match.group(1)
                        cents = match.group(2)
                        return f'{euros},{cents}€'
                    def replace_price_decimal1(match):
                        euros = match.group(1)
                        cents = match.group(2)
                        return f'{euros},{cents}€'
                    def replace_price_integer(match):
                        euros = match.group(1)
                        return f'{euros}€'
                    # Decimal prices (2 digits)
                    price_regex_decimal2 = r'€\s*(\d{1,3}(?:\.\d{3})*)\.(\d{2})(?![\w])'
                    # Decimal prices (1 digit)
                    price_regex_decimal1 = r'€\s*(\d{1,3}(?:\.\d{3})*)\.(\d{1})(?![\w])'
                    # Integer prices
                    price_regex_integer = r'€\s*(\d{1,3}(?:\.\d{3})*)(?![\w.,\d])'
                    new_combined_text = re.sub(price_regex_decimal2, replace_price_decimal2, combined_text)
                    new_combined_text = re.sub(price_regex_decimal1, replace_price_decimal1, new_combined_text)
                    new_combined_text = re.sub(price_regex_integer, replace_price_integer, new_combined_text)
                offset = 0
                for t in texts:
                    if t.text:
                        length = len(t.text)
                        t.text = new_combined_text[offset:offset+length]
                        offset += length
        tree.write(document_xml, encoding='utf-8', xml_declaration=True)
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_ref.write(file_path, arcname)
        show_success(f"Price format replacement completed successfully in {output_file}")
    except Exception as e:
        show_error(f"An error occurred: {e}")
    finally:
        # Clean up the temporary directory
        shutil.rmtree(temp_dir)


# Global state for conversion direction
conversion_direction = ['es_to_en']  # mutable for closure

def on_file_drop(event):
    file_path = event.data.strip('{}')
    replace_prices_in_docx(file_path, direction=conversion_direction[0])


def main():
    root = TkinterDnD.Tk()
    root.title("Format those prices!")
    root.geometry("400x300")
    root.attributes("-topmost", True)

    label = tk.Label(root, text="Drag and drop a DOCX file here", width=40, height=10, bg="black")
    label.pack(pady=20)

    # Conversion direction toggle button
    def toggle_direction():
        if conversion_direction[0] == 'es_to_en':
            conversion_direction[0] = 'en_to_es'
            btn.config(text="Switch to ES > EN")
            dir_label.config(text="Current: EN > ES")
        else:
            conversion_direction[0] = 'es_to_en'
            btn.config(text="Switch to EN > ES")
            dir_label.config(text="Current: ES > EN")

    dir_label = tk.Label(root, text="Current: ES > EN", fg="white", bg="black", width=40)
    dir_label.pack(pady=5)
    btn = tk.Button(root, text="Switch to EN > ES", command=toggle_direction)
    btn.pack(pady=5)

    label.drop_target_register(DND_FILES)
    label.dnd_bind('<<Drop>>', on_file_drop)
    root.mainloop()

if __name__ == "__main__":
    main()

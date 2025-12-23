# price-formatter

<img width="403" height="325" alt="price_formatter" src="https://github.com/user-attachments/assets/8250e382-848d-4023-86db-b5a0ed0a1902" />



A simple GUI application for formatting prices (ES<>EN) in DOCX files. When translating files in English/Spanish that have prices in them (in particular menus), it's a PITA to manually change all the prices from the format 10,00€ to €10.00 or vice-versa.

## Usage:
- After finishing (or before starting) the translation, launch price-formatter.
- It will indicate the translation direction (default is ES>EN). There is a button to choose the desired direction.
- Drag and drop the docx file into the window. It will (hopefully!) change all the prices to the desired format and will create a new file with '_conv' added to its name. The original file is left as-is in case anything went wrong.
- It supports prices with 1, 2 or no decimals.

Files with weird formatting (e.g. those produced with OCR) might not work as expected. A .deb file is provided [here](https://github.com/ignotus666/price-formatter/releases/tag/v1.0.0), otherwise you can run the `price-formatter/usr/local/bin/price_formatter.py` script.

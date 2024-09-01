# price-formatter

![price-formatter_gui](https://github.com/user-attachments/assets/3e2392bd-0f4f-4760-86bf-dfcf6c5198ce)


A simple GUI application for formatting prices (ES>EN) in DOCX files. When translating files in Spanish that have prices in them (in particular menus), it's a PITA to manually change all the prices from the format 10,00€ to €10.00. Just leave them in the Spanish format, and when finished, open this application, drag and drop the docx file into the window, and it will change all the prices to the correct format, It will also create a backup file in the same directory in case anything goes wrong. Files with weird formatting (e.g. those produced with OCR) might not work as expected. A .deb file is provided; otherwise package the price-formatter folder with your packager of choice.

# E-Certificate_generator
📜 E-Certificate Generator
This project helps you generate personalized certificates from an Excel sheet.
It automatically creates both PNG images and PDFs for each participant, based on a template.

🚀 Features
- Automatically installs required libraries (first run only).

- Pop-up file picker for selecting your Excel sheet.

- Generates both Image (.png) and PDF (.pdf) certificates.

- Organized output into All_Certificates/Images and All_Certificates/PDFs.

📂 Setup
> Download the Code

> Clone this repo or download the ZIP and extract.

> Prepare the Files

> Place your Excel sheet with participant details in the same folder.

> Ensure your certificate template image (e.g., data_alchemy.png) is in the same folder.

> Ensure your font file (e.g., LucidaUnicodeCalligraphy.ttf) is available in the folder.

▶️ Usage
Open the folder in VS Code (or any IDE).

**Run the script:**

python certificate_generator.py

A file picker pop-up will appear.
Select your Excel file (.xlsx / .xlsm).

Certificates will be generated in:
All_Certificates/Images/ → .png files
All_Certificates/PDFs/ → .pdf files

📝 Notes
On the first run, the script will install required libraries (img2pdf, openpyxl, Pillow) automatically.
After the first successful run, you can comment out these lines in the script to speed up execution:

install('img2pdf')
install('openpyxl')
install('Pillow')

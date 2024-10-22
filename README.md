#  PDF Data Extraction to Excel with Embedded Images
Project Overview
This project is a Python-based solution for extracting both text and images from PDF files and saving them into an Excel sheet. The text is structured in a readable format, and images are embedded directly into the corresponding cells in Excel. It is ideal for scenarios where you need to convert PDFs into an editable Excel format while maintaining the visual elements.

Features
Text Extraction: Extracts all the text content from each page of the PDF.
Image Extraction: Detects and extracts images from PDF pages and embeds them directly into Excel cells.
Formatted Output: Generates a well-organized Excel file where text and images are aligned per page.
Image Resizing: Resizes the images to fit neatly within Excel cells.
Installation
Clone the repository:

bash
Copy code
git clone https://github.com/yourusername/pdf-to-excel.git
Install required libraries:

The project relies on the following Python libraries:

pandas: For managing Excel output.
PyMuPDF (fitz): For PDF parsing.
Pillow: For image manipulation.
openpyxl: For writing to Excel.
To install the dependencies, run:

bash
Copy code
pip install pandas PyMuPDF Pillow openpyxl
Usage
Place your PDF file in the project directory.

Edit the Python script to specify the path of your PDF:

Open the script and modify the following line to point to your PDF file:

python
Copy code
pdf_path = "path_to_your_pdf_file.pdf"
Run the script:

bash
Copy code
python extract_pdf_to_excel.py
Output: The extracted data (text and images) will be saved as an Excel file named output.xlsx in the working directory.

Example
Given a PDF with multiple pages of text and images, the script will generate an Excel file where:

Column A contains the page numbers.
Column B contains the extracted text from each page.
Column C contains the corresponding images (embedded in the cells).
Directory Structure
bash
Copy code
pdf-to-excel/
│
├── extract_pdf_to_excel.py   # Main script to extract text and images
├── sample.pdf                # Sample PDF (replace with your own file)
├── output.xlsx               # Generated Excel file with extracted data
└── README.md                 # Project documentation
Requirements
Python 3.x
The following libraries:
pandas
PyMuPDF
Pillow
openpyxl
Future Enhancements
Selective Page Extraction: Option to extract data from specific pages of the PDF.
More Image Formats: Support for embedding other image types such as JPEG and GIF.
Advanced Formatting: Enhance text formatting in Excel to match the original PDF layout.
License
This project is licensed under the MIT License. See the LICENSE file for details.

Contributing
Contributions are welcome! Feel free to fork the repository, submit pull requests, or open issues for any bugs or feature requests.

Contact
For questions or support, feel free to reach out through kundan.k2205@gmail.com.


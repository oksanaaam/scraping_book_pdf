# Scraping Book PDF

The Abstract Book PDF Scraper is a Python-based tool designed to extract article information from PDF files.  
It utilizes the PyMuPDF library to parse PDF documents,  
extract relevant text blocks, and retrieve article details such as session names, titles, authors, affiliations, locations, and presentation abstracts.

The extracted article data can be conveniently stored in an Excel file using the openpyxl library, allowing for easy access and further analysis.  
This tool aims to automate the process of extracting article information from PDFs, saving time and effort for researchers, conference organizers, or anyone dealing with large collections of articles.

## Features
1. Extracts article information from PDF files.
2. Retrieves session names, article titles, authors, affiliations, locations, and presentation abstracts.
3. Stores extracted data in an Excel file for easy management and analysis.


### Installation

1. Clone the repo
`git clone https://github.com/oksanaaam/scraping_book_pdf.git`
2. Open the project folder in your IDE
3. Open a terminal in the project folder
4. If you are using PyCharm - it may propose you to automatically create venv for your project and install requirements in it, but if not:
```
python -m venv venv
venv\Scripts\activate (on Windows)
source venv/bin/activate (on macOS)
pip install -r requirements.txt
```

## Usage

Make sure that your PDF file placed in the project directory.

Modify the pdf_file and xl_file variables in the main.py file to specify the input PDF file path and the desired output Excel file path.
```
pdf_file = "your path to Abstract Book.pdf here"
xl_file = "your path to task.xlsx  here"
```

Run the script:
``` 
python main.py
```

The extracted article information will be displayed in the console and saved in the Excel file specified.  
You can open the Excel file to verify the extracted data.

Data before:
![data_before.png](img%20for%20README.md%2Fdata_before.png)

Running script:
![run_script.png](img%20for%20README.md%2Frun_script.png)

Result:
![result_data.png](img%20for%20README.md%2Fresult_data.png)

# IDEA Policy Guidance Scraper

## Project Overview
This project is designed to scrape and extract structured text data from the **IDEA (Individuals with Disabilities Education Act) policy guidance website**. The extracted content includes policy document metadata, descriptions, and full text from downloadable PDF and Word files. The output is stored in both **JSON** and **Excel** formats for further processing and analysis.

## Features
- Scrapes policy guidance documents from [IDEA Policy Guidance](https://sites.ed.gov/idea/policy-guidance/).
- Extracts text from **PDF, DOCX, and DOC** files.
- Uses **multithreading** to speed up processing.
- Saves the structured data in **JSON** and **Excel** formats.
- Supports **Windows-only DOC file extraction** using `win32com.client`.

## Requirements
This project requires **Python 3.8+** and the following dependencies:

```bash
pip install requests beautifulsoup4 tqdm PyPDF2 docx2txt pandas openpyxl
```

**Note:** If running on Windows, ensure `win32com.client` is available for DOC file extraction.

## Installation
Clone this repository and install the required packages:

```bash
git clone https://github.com/your-repo/idea-scraper.git
cd idea-scraper
pip install -r requirements.txt
```

## How It Works
1. **Web Scraping**: The script fetches the main policy guidance webpage and extracts document metadata.
2. **Document Processing**:
   - **PDFs**: Extracts text using `PyPDF2`.
   - **DOCX**: Extracts text using `docx2txt`.
   - **DOC (Windows only)**: Uses `win32com.client` for text extraction.
3. **Data Structuring**: Organizes extracted data into a structured JSON format.
4. **Saving Output**:
   - **JSON**: Stores structured data with extracted text.
   - **Excel**: Saves metadata (excluding full text) for easy reference.

## Usage
Run the script:

```bash
python scrape_idea.py
```

Upon completion, the following files will be generated:
- `data_cleaned.json`: JSON file containing structured data.
- `IDEAbot_reference.xlsx`: Excel file with metadata for easy browsing.

## File Structure
```plaintext
idea-scraper/
│── scrape_idea.py          # Main script
│── requirements.txt        # Dependencies
│── data_cleaned.json       # Extracted data (JSON output)
│── IDEAbot_reference.xlsx  # Extracted metadata (Excel output)
```

## Notes
- The script **prioritizes PDF documents** if both PDF and Word versions are available.
- Some files may fail to extract text; the script logs errors in JSON.
- Ensure your Python environment is configured correctly before running.

## License
This project is licensed under the MIT License.

## Author
[Sydney Leiher] - [sydney.leiher@gmail.com]

---
For any issues or feature requests, please open an issue in the repository.


import requests
from bs4 import BeautifulSoup
import re
import io
import os
import uuid
import json
import concurrent.futures
import platform
import pandas as pd
from datetime import datetime
from tqdm import tqdm
from PyPDF2 import PdfReader
import docx2txt

# Check if running on Windows for win32com compatibility
IS_WINDOWS = platform.system() == "Windows"
if IS_WINDOWS:
    import win32com.client

# Initialize a session for requests reuse
session = requests.Session()
MAIN_PAGE = "https://sites.ed.gov/idea/policy-guidance/"

# Fetch the main page
response = session.get(MAIN_PAGE)
response.raise_for_status()
soup = BeautifulSoup(response.text, "html.parser")

# Extract "idea-file-item" elements
file_items = soup.find_all("div", class_="idea-file-item")
if not file_items:
    raise ValueError("No policy guidance documents found on the page.")

# Extract text from PDFs
def get_pdf_text(pdf_link):
    try:
        response = session.get(pdf_link, stream=True, timeout=10)
        response.raise_for_status()
        pdf_reader = PdfReader(io.BytesIO(response.content))
        text = "\n".join(page.extract_text() or "" for page in pdf_reader.pages).strip()
        return text if text else "Could not extract text from PDF"
    except Exception as e:
        return f"PDF Error ({pdf_link}): {e}"

# Extract text from DOCX and DOC files
def get_doc_text(doc_link):
    try:
        response = session.get(doc_link, stream=True, timeout=10)
        response.raise_for_status()
        file_extension = doc_link.lower().split(".")[-1]

        if file_extension == "docx":
            return docx2txt.process(io.BytesIO(response.content)).strip()

        elif file_extension == "doc" and IS_WINDOWS:
            return extract_text_win32com(doc_link, response.content)

    except Exception as e:
        return f"DOC Error ({doc_link}): {e}"

# Extract text from DOC using win32com (Windows only)
def extract_text_win32com(doc_link, file_content):
    try:
        temp_filename = f"{uuid.uuid4()}.doc"
        with open(temp_filename, "wb") as f:
            f.write(file_content)

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(temp_filename))
        text = doc.Content.Text.strip()
        doc.Close(False)
        word.Quit()
        os.remove(temp_filename)
        return text
    except Exception as e:
        return f"Win32com Error ({doc_link}): {e}"

# Extract a 4-digit year from the title
def extract_date_from_title(title):
    match = re.search(r"\b(\d{4})\b", title)
    return match.group(1) if match else "Unknown"

# Process each file item
def process_file_item(file_item):
    try:
        title_element = file_item.find("h3").find("a")
        title = title_element.text.strip()
        link = title_element["href"]

        topic_area = file_item.find("div", class_="topic-area-list").get_text(strip=True).replace("Topic Areas: ", "")
        description = file_item.find("div", class_="description").get_text(strip=True).replace("Read More", "")

        document_date = extract_date_from_title(title)

        all_links = {a["href"]: a.text.strip() for a in file_item.find_all("a")}
        pdf_links = {k: v for k, v in all_links.items() if k.lower().endswith(".pdf")}
        doc_links = {k: v for k, v in all_links.items() if k.lower().endswith(('.doc', '.docx'))}

        # Prioritize PDFs, otherwise take the first Word doc
        selected_doc = next(iter(pdf_links.items()), next(iter(doc_links.items()), None))

        docs = []
        if selected_doc:
            doc_link, doc_title = selected_doc
            doc_text = get_pdf_text(doc_link) if doc_link.lower().endswith(".pdf") else get_doc_text(doc_link)
            docs.append({"link": doc_link, "title": doc_title, "text": doc_text})

        return {"title": title, "date": document_date, "link": link, "topic_area": topic_area, "description": description, "docs": docs}

    except Exception as e:
        return {"error": f"Processing Error: {e}"}

# Process all file items concurrently
with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
    data = list(tqdm(executor.map(process_file_item, file_items), total=len(file_items)))

# Save JSON file
with open("data_cleaned.json", "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2, ensure_ascii=False)

# Convert to DataFrame and Save to Excel
def save_to_excel(data, filename):
    excel_data = [
        {"Title": entry["title"], "Date": entry["date"], "Link": entry["link"], "Topic Area": entry["topic_area"], "Description": entry["description"]}
        for entry in data if "error" not in entry
    ]
    pd.DataFrame(excel_data).to_excel(filename, index=False)

save_to_excel(data, "IDEAbot_reference.xlsx")
print("Processing complete!")

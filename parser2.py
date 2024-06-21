import pdfplumber
import spacy
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# SpaCy modelini yükleyin
nlp = spacy.load("en_core_web_sm")

# PDF'den metin çıkarma fonksiyonu
def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        all_text = ""
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                all_text += page_text + "\n"
    return all_text

# Metinden veri çıkarma fonksiyonu
def extract_info(text):
    doc = nlp(text)
    info = {
        "name": None,
        "education": [],
        "skills": {
            "languages": [],
            "frameworks": [],
            "developer_tools": [],
            "libraries": []
        }
    }

    # Adı çıkarma (örnek)
    for ent in doc.ents:
        if ent.label_ == "PERSON" and info["name"] is None:
            info["name"] = ent.text

    education_keywords = ["university", "college", "bachelor", "master", "phd"]
    technical_skills_sections = {
        "languages": ["languages:"],
        "frameworks": ["frameworks:"],
        "developer_tools": ["developer tools:"],
        "libraries": ["libraries:"]
    }

    for line in text.split('\n'):
        # Eğitim bilgilerini çıkarma
        for keyword in education_keywords:
            if keyword in line.lower():
                info["education"].append(line.strip())

        # Teknik becerileri çıkarma
        for section, keywords in technical_skills_sections.items():
            for keyword in keywords:
                if keyword in line.lower():
                    skills = line.split(':')[1].strip().split(',')
                    skills = [skill.strip() for skill in skills]
                    info["skills"][section].extend(skills)

    return info

# JSON'a dönüştürme fonksiyonu (opsiyonel olarak kullanılabilir)
def convert_to_json(info, json_path):
    with open(json_path, 'w') as json_file:
        json.dump(info, json_file, indent=4)

# Excel'e dönüştürme fonksiyonu
def convert_to_excel(info, excel_path):
    # Veriyi DataFrame'e dönüştürme
    data = {
        "Name": [info["name"]],
        "Education": ["\n".join(info["education"])],
        "Languages": [", ".join(info["skills"]["languages"])],
        "Frameworks": [", ".join(info["skills"]["frameworks"])],
        "Developer Tools": [", ".join(info["skills"]["developer_tools"])],
        "Libraries": [", ".join(info["skills"]["libraries"])]
    }

    df = pd.DataFrame(data)

    # Excel dosyasına yazma
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Resume')

        # Başlık biçimlendirme
        workbook = writer.book
        worksheet = writer.sheets['Resume']
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top'})
        worksheet.set_row(0, None, header_format)

        # Hücre boyutlandırma
        worksheet.set_column('A:F', 20)

# Ana akış
def main():
    def browse_pdf():
        filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filename:
            pdf_entry.delete(0, tk.END)
            pdf_entry.insert(tk.END, filename)

    def start_parsing():
        pdf_path = pdf_entry.get()
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file.")
            return
        
        text = extract_text_from_pdf(pdf_path)
        info = extract_info(text)
        
        # JSON olarak kaydetme (opsiyonel)
        convert_to_json(info, 'output.json')
        
        # Excel olarak kaydetme
        excel_path = 'output.xlsx'
        convert_to_excel(info, excel_path)
        
        messagebox.showinfo("Success", f"Data parsed and saved to {excel_path}")

    window = tk.Tk()
    window.title("Resume Parser")

    tk.Label(window, text="Select PDF file:").pack(pady=10)
    pdf_entry = tk.Entry(window, width=50)
    pdf_entry.pack(pady=10)

    browse_button = tk.Button(window, text="Browse", command=browse_pdf)
    browse_button.pack(pady=10)

    start_button = tk.Button(window, text="Start", command=start_parsing)
    start_button.pack(pady=10)

    window.mainloop()

if __name__ == "__main__":
    main()

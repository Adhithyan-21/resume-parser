import os
import docx
from PyPDF2 import PdfReader
import re
from tabulate import tabulate
import pandas as pd

folder_path = "C:/Users/SADASIVAM/Documents"
keywords=["education","name","objectivr","skills","resume"]

def extracted_data(full_text):
    name = re.findall(r"name\s*[:\-]\s*([A-Za-z]+)",full_text,re.I)
    mailid = re.findall(r"[\w\.-]+@[\w\.-]+\.\w+", full_text, re.I)
    dob = re.findall(r"date of birth\s*[:\-]\s*(\d{1,2}[-/.\s]\d{1,2}[-/.\s]\d{2,4})", full_text, re.I)
    return{
        "name":name[0]if name else None,
        "dob":dob[0]if dob else None,
        "mailid":mailid[0] if mailid else None,
    }
all_data=[]
for file in os.listdir(folder_path):      
            file_path = os.path.join(folder_path,file)
            if file.lower().endswith(".pdf"):
                try:
                    reader = PdfReader(file_path)
                    full_text=" "
                    for page in reader.pages:
                        full_text += page.extract_text() or " "
                    if any(keyword in full_text.lower() for keyword in keywords):
                        print(f"matched:{file}")
                        data=extracted_data(full_text)
                        data["filename"] = file
                        all_data.append(data)
                except Exception as e:
                    print(f"error in {file}:{e}")
            elif file.lower().endswith(".docx"):
                try:
                    doc =docx.Document(file_path)
                    full_text = "\n".join([para.text for para in doc.paragraphs])
                    if any(keyword in full_text.lower() for keyword in keywords):
                        print(f"matched:{file}")
                        data = extracted_data(full_text)
                        data["filename"] = file
                        all_data.append(data)
                except Exception as e:
                    print(f"error in {file}:{e}")
full_data=pd.DataFrame(all_data)
data_exel="resume_extract.xlsx"
full_data.to_excel(data_exel,index=False)
os.startfile(data_exel)
print(tabulate(all_data,headers="keys",tablefmt="grid"))

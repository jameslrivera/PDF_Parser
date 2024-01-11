import requests
import pdfplumber
import re
import pandas as pd
from collections import namedtuple
import openpyxl
import xlsxwriter

doc_name = "Traction_Power_Facilities_Installation_Requirements"
TPFI_requirements = namedtuple(doc_name, 'section' 'title')
doc = '/Users/jamesrivera/Downloads/34 21 70 Traction Power Facilities Installation Requirements (1).pdf'

with pdfplumber.open(doc) as pdf:
    page1 = pdf.pages[0]  
    page1_text = page1.extract_text()
    print(page1_text)
    print("END OF PAGE ONE")

pattern8 = re.compile(r'([A-Z]+\s\d+\s+\d+\s+\d+)([A-Z\s]{52})')
found_first_match = False
section_data = []

for match in pattern8.finditer(page1_text):
    if not found_first_match:
        section = match.group(1)
        title = match.group(2).strip()
        section_data.append((section, title))
        found_first_match = True


df1 = pd.DataFrame(section_data, columns=["Doc Section Number", "Doc Title"])

#print(df)

#df.to_csv("output.csv", index=False, sep='\t', encoding='utf-8-sig')

#excel_file = 'output.xlsx'

# Export the DataFrame to Excel
#df.to_excel(excel_file, index=False)


pattern = r'(\d+\.\d+)\s+([A-Z\s]+)\s*(.*?)\s*(?=\d+\.\d+|\Z)'

# Initialize a list to store extracted sections from all pages
all_sections = []

with pdfplumber.open(doc) as pdf:
    for page_number, page in enumerate(pdf.pages[2:], start=3):  # Start from page 3
        text = page.extract_text()

        # Find all matches in the text for the current page
        matches = re.findall(pattern, text, re.DOTALL)

        sections = []

        for match in matches:
            section_number = match[0]
            section_title = match[1]
            section_text = match[2].strip()

            # Skip sections with specific keywords in section_text
            if "BART FACILITIES STANDARDS" in section_text or "ISSUED: APRIL 2018 PAGE" in section_text:
                continue

            sections.append({
                "Section Number": section_number,
                "Section Title": section_title.strip(),
                "Section Text": section_text
            })

        # Append sections from the current page to the list
        all_sections.extend(sections)

# Create a DataFrame from the list of sections
df2 = pd.DataFrame(all_sections)

# Export the DataFrame to CSV with formatting options
#csv_file = 'output2.csv'
#df.to_csv(csv_file, index=False, encoding='utf-8-sig', sep='\t')
excel_file = 'parsed.xlsx'
#df.to_excel(excel_file, index=False)



with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
    
    df1.to_excel(writer, sheet_name='Sheet1', startrow=0, startcol=8, index=False) 
    df2.to_excel(writer, sheet_name='Sheet1', startrow=5, startcol=2, index=False)
    worksheet = writer.sheets['Sheet1']
    worksheet.set_column('C:E', 15)
    worksheet.set_column('I:J', 15)







'''
pattern = r'(\d+\.\d+)\s+([A-Z\s]+)\s*(.*?)\s*(?=\d+\.\d+|\Z)'

# Initialize a list to store extracted sections from all pages
all_sections = []

with pdfplumber.open(doc) as pdf:
    for page_number, page in enumerate(pdf.pages[2:], start=3):  # Start from page 3
        text = page.extract_text()

        # Find all matches in the text for the current page
        matches = re.findall(pattern, text, re.DOTALL)

        sections = []

        for match in matches:
            section_number = match[0]
            section_title = match[1]
            section_text = match[2].strip()

            # Skip sections with specific keywords in section_text
            if "BART FACILITIES STANDARDS" in section_text or "ISSUED: APRIL 2018 PAGE" in section_text:
                continue

            sections.append({
                "section_number": section_number,
                "section_title": section_title.strip(),
                "section_text": section_text
            })

        # Append sections from the current page to the list
        all_sections.extend(sections)

        # Print the extracted sections for the current page
        print(f"Page {page_number} Sections:")
        for section in sections:
            print("Section Number:", section["section_number"])
            print("Section Title:", section["section_title"])
            print("Section Text:", section["section_text"])
            print("\n")

# Print the extracted sections for all pages
print("All Extracted Sections:")
for section in all_sections:
    print("Section Number:", section["section_number"])
    print("Section Title:", section["section_title"])
    print("Section Text:", section["section_text"])
    print("\n")

'''
'''
pattern = r'(\d+\.\d+)\s+([A-Z\s]+)\s*(.*?)\s*(?=\d+\.\d+|\Z)'

# Initialize a list to store extracted sections from all pages
all_sections = []

with pdfplumber.open(doc) as pdf:
    for page_number, page in enumerate(pdf.pages[2:], start=3):  # Start from page 3
        text = page.extract_text()

        # Find all matches in the text for the current page
        matches = re.findall(pattern, text, re.DOTALL)

        sections = []

        for match in matches:
            section_number = match[0]
            section_title = match[1]
            section_text = match[2].strip()

            sections.append({
                "section_number": section_number,
                "section_title": section_title.strip(),
                "section_text": section_text
            })

        # Append sections from the current page to the list
        all_sections.extend(sections)

        # Print the extracted sections for the current page
        print(f"Page {page_number} Sections:")
        for section in sections:
            print("Section Number:", section["section_number"])
            print("Section Title:", section["section_title"])
            print("Section Text:", section["section_text"])
            print("\n")

# Print the extracted sections for all pages
print("All Extracted Sections:")
for section in all_sections:
    print("Section Number:", section["section_number"])
    print("Section Title:", section["section_title"])
    print("Section Text:", section["section_text"])
    print("\n")
'''
'''
with pdfplumber.open(doc) as pdf:
    for page in pdf.pages[2:3]:  
        text = page.extract_text()
        #print(text)

pattern = r'(\d+\.\d+)\s+([A-Z\s]+)\s*(.*?)\s*(?=\d+\.\d+|\Z)'

# Find all matches in the text
matches = re.findall(pattern, text, re.DOTALL)

sections = []

for match in matches:
    section_number = match[0]
    section_title = match[1]
    section_text = match[2].strip()
    
    sections.append({
        "section_number": section_number,
        "section_title": section_title.strip(),
        "section_text": section_text
    })

# Print the extracted sections
for section in sections:
    print("Section Number:", section["section_number"])
    print("Section Title:", section["section_title"])
    print("Section Text:", section["section_text"])
    print("\n")
    '''
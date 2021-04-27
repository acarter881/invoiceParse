#! python3
import pdfplumber
import re
import pandas as pd

# Create an empty list that stores all of the invoice data 
records = []

# Loop through each page of the PDF file
with pdfplumber.open('C:\\Users\\Alex\\Desktop\\INVOICES.pdf') as pdf:
    for i in range(len(pdf.pages)):
        print(f'Searching on page: {i+1}')
        page = pdf.pages[i]
        text = page.extract_text()
        invoice_number = re.search(r'[4]\d{7}', text).group(0)
        ship_to = re.findall(r'(\d{8})', text)[1]
        sold_to = re.findall(r'(\d{8})', text)[2]
        try:
            state_tax = re.search(r'STATE SALES TAX \$(\d+\.\d{2})', text).group(1)
        except AttributeError:
            state_tax = 0
        try:
            county_tax = re.search(r'COUNTY SALES TAX \$(\d+\.\d{2})', text).group(1)
        except AttributeError:
            county_tax = 0
        try:
            city_tax = re.search(r'CITY SALES TAX \$(\d+\.\d{2})', text).group(1)
        except AttributeError:
            city_tax = 0
        try:
            local_tax = re.search(r'LOCAL TAX \$(\d+\.\d{2})', text).group(1)
        except AttributeError:
            local_tax = 0
        if state_tax:
            records.append((invoice_number, ship_to, sold_to, float(state_tax), float(county_tax), float(city_tax), float(local_tax)))

# Create pandas data frame    
df = pd.DataFrame(records, columns=['Invoice Number', 'Ship To', 'Sold To', 'State Tax', 'County Tax', 'City Tax', 'Local Tax'])

# Write relevant data to an Excel file
df.to_excel('Invoice_Data.xlsx', sheet_name='Invoice_Data', index=False, freeze_panes=(1,0))

import PyPDF2
import os
import re
#import pdfplumber

def extract_invoice_info(pdf_file_path):

    with open(pdf_file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ''
        
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
          

        invoice_number_pattern = r'INVOICE\s*\s*(\d+)'
        bill_to_pattern = r'Bill\s*To\s*:\s*(.*)'
        items_pattern = r'(.*?)\s*(\d+)\s*(\d+\.\d{2})\s*(\d+\.\d{2})'
        notes_terms_pattern = r'Notes\s*:\s*(.*?)\s*Terms\s*:\s*(.*)'
       
        #Extracted information
        '''invoice_number = invoice_number_match.group (1) if invoice_number_match else None 
        bill_to = bill_to_match.group(1) if bill_to_match else None
        items = items_matches
        notes, terms = notes_terms_match.groups() if notes_terms_match else (None, None) 
        match = re.search(discount_tax_pattern, text)
        discount_percentage = match.group(1)
        tax_percentagematch.group(2)

        print(f'Invoice Number: {invoice_number}') 
        print('Bill To: {bill_to}') 
        print('Items: ')
        for item in items:
            print(f' - (item[0]): (item[1]) units x (item[2]) = {item[3]}')
        print (f'Notes: {notes}')
        print('Terms: {terms}')
        print('Discount Percentage: {discount_percentage}') 
        print (f'Tax Percentage: {tax_percentage}')'''
        
        print(text)
extract_invoice_info("invoices/invoice3.pdf")
        
        
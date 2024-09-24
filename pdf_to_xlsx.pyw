# pip install PyPDF2
# pip install openpyxl
#     Copy a shortcut of script to this location to have a SendTo shortcut:
#       %AppData%\Microsoft\Windows\SendTo

import argparse
import os.path
import sys
import PyPDF2
import openpyxl

def extract_from_pdf_to_xlsx(input_pdf, output_xlsx, sheet_name):
    with open(input_pdf, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        account_list = []
        payment_list = []
        transaction_count = 0
        
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
                
            lines = text.split('\n')
            for i, line in enumerate(lines):
                if 'Plata domestica' in line:                
                    fields = line.split()
                    next_line = lines[i+1]
                    subtext = next_line[-29:]
                    account = subtext[:24]
                    payment = fields[-2]
                    payment = payment.replace('.','')
                    payment = payment.replace(',','.')
                    account_list.append(account)
                    payment_list.append(float(payment))
                    transaction_count = transaction_count + 1

        workbook = openpyxl.load_workbook(output_xlsx)
        sheet = workbook[sheet_name] 

        for a in sheet['A1':'B999']: 
            for cell in a:
                cell.value = None
        sheet[f'A1'] = 'Account'        
        sheet[f'B1'] = 'Payment'
        sheet.column_dimensions['A'].width = 30
        sheet.column_dimensions['B'].width = 10
        for i, (account_list, payment_list) in enumerate(zip(account_list, payment_list), start=1):
            sheet[f'A{i+1}'] = account_list
            sheet[f'B{i+1}'] = payment_list

        workbook.save(output_xlsx)
        print(f'S-au extras {transaction_count} tranzactii dintr-un fisier cu {len(pdf_reader.pages)} pagini')

if __name__ == "__main__":
    param_count = len(sys.argv)
    script_path = os.path.abspath(__file__)
    script_name = os.path.basename(script_path)
    input_file1, _ = os.path.splitext(script_name)
    input_file = os.path.join(os.path.dirname(script_path), input_file1 + '.pdf')
    input_file2, _ = os.path.splitext(script_name)
    output_xlsx = os.path.join(os.path.dirname(script_path), input_file2 + '.xlsx')

    if param_count-1 == 0:
        print('Nu s-a specificat nici un fisier, se foloseste: ', input_file)
        if os.path.isfile(f'{input_file}'):
            input_pdf = input_file
            sheet_name = 'Sheet1'
        else:
            input_pdf = ''
            print('Nu a fost gasit fisierul: ', input_file)
    elif param_count-1 == 1:
        parser = argparse.ArgumentParser(description='Extrage dintr-un fisier PDF lista tranzactiilor.')    
        parser.add_argument('input_file', help='Calea catre fisierul .pdf de procesat')
        args = parser.parse_args()
        input_file = args.input_file
        input_file2, _ = os.path.splitext(input_file)
        output_xlsx = os.path.join(os.path.dirname(input_file2), input_file2 + '.xlsx')
    elif param_count-1 == 2:
        parser = argparse.ArgumentParser(description='Extrage dintr-un fisier PDF lista tranzactiilor.')
        parser.add_argument('input_file', help='Calea catre fisierul .pdf de procesat')
        parser.add_argument('output_xlsx', help='Calea catre fisierul .xlsx rezultat')
        args = parser.parse_args()
        input_file = args.input_file
        output_xlsx = args.output_xlsx        

    if input_file == "":        
        print('Nu a fost gasit fisierul: ', input_file)
    else:
        print('Procesez fisierul: ', input_file)
        print('Salvez in fisierul: ', output_xlsx)
        if os.path.isfile(f'{output_xlsx}'):
            os.remove(output_xlsx)
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        workbook.save(output_xlsx)
        extract_from_pdf_to_xlsx(input_file, output_xlsx, 'Sheet1')
                


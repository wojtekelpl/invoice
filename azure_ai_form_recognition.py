import argparse
import datetime
import json
import os
import pandas as pd
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from dateutil.parser import parse
import re
import dateparser
from dateutil.relativedelta import relativedelta
import requests
from config import month_dict
from myconfig import endpoint, api_key, nip_check_url, faktury_koszty_path
# Konfiguracja 
client = DocumentAnalysisClient(endpoint, AzureKeyCredential(api_key))
 

def append_warnings(warnings, warning):
    if warning:
        warnings += warning + "; "
    return warnings

def convert_date(date_string, miesiac):
    try:
        # Try to parse the date
        for month in month_dict:
            date_string = date_string.replace(month, month_dict[month])
        date = dateparser.parse(date_string)
        uwagi = ""
        if date.strftime('%m') != miesiac:
            uwagi = "Data niezgodna z miesiącem faktury" 
            if date.strftime('%d') == miesiac:
                uwagi = "Data niezgodna z miesiącem faktury, ale poprawiona, prosze sprawdzić"
                return (date.strftime('%Y-%d-%m'), uwagi)

        # Convert the date to the common format (YYYY-MM-DD)
        return (date.strftime('%Y-%m-%d'), uwagi)
    except ValueError:
        uwagi = "Nie można przetworzyć daty"
        # If the date cannot be parsed, return the input string
        return (date_string, uwagi)
    
# Funkcja do przetwarzania dokumentów
def analyze_document(file_path):
    with open(file_path, "rb") as f:
        poller = client.begin_analyze_document("prebuilt-invoice", f)
        result = poller.result()
    return result

def convert_to_number(text):
    try:
        # Remove currency codes
        text = re.sub(r'PLN|zl|\s', '', text)
        # Replace comma with dot
        text = text.replace(',', '.')
        # Remove all non-numeric characters except dot
        text = re.sub(r'[^\d.]', '', text)
        # Convert to float
        number = float(text)
        return number
    except ValueError as e:
        print(f"Error converting text to number: {text}")
        return 0

def check_nip(nip):    
    current_date = datetime.date.today().strftime('%Y-%m-%d')    
    url =  f"{nip_check_url}{nip.replace('-', '')}?date={current_date}"
    print(f"Sprawdzanie NIP: {nip} url {url}")
    response = requests.get(url)    
    data = json.loads(response.content.decode('utf-8'))

    status_vat = data.get('result', {}).get('subject', {}).get('statusVat', None)
    print(f"Status NIP: {status_vat}")
    
    return status_vat

# Funkcja do wyodrębniania danych i zapisywania ich w pliku CSV
def extract_and_save_to_csv(files, output_csv, month, year):
    all_invoices = []
    for file_path in files:
        result = analyze_document(file_path)
        
        for idx, invoice in enumerate(result.documents):
            invoice_data = {}
            uwagi = ""
            invoiceId = invoice.fields.get("InvoiceId").content
            invoice_data['Numer Faktury'] = invoiceId
            nazwa_sprzedajacego = invoice.fields.get("VendorName").content
            invoice_data['Nazwa Sprzedającego'] = nazwa_sprzedajacego
            invoice_data['Adres Sprzedającego'] = invoice.fields.get("VendorAddress").content
            invoice_data['NIP'] = invoice.fields.get("VendorTaxId").content
            dataFaktury, uwagiData = convert_date(invoice.fields.get("InvoiceDate").content, month)
            invoice_data['DataFaktury'] = dataFaktury
            uwagi = append_warnings(uwagi, uwagiData)
            subTotalAttr = getattr(invoice.fields.get("SubTotal"), 'content', None)
            SubTotal = 0
            if subTotalAttr is None:
                uwagi = append_warnings(uwagi, "Brak wartości netto")
                 
            else:
                SubTotal = convert_to_number(invoice.fields.get("SubTotal").content)
               
                
            invoice_data['Netto'] = SubTotal
            totalTax = getattr(invoice.fields.get("TotalTax"), 'content', None)
            InvoiceTotal= convert_to_number(invoice.fields.get("InvoiceTotal").content)
            
            if totalTax is None:
                invoice_data['VAT'] = InvoiceTotal - SubTotal
                uwagi = append_warnings(uwagi, "Brak wartości VAT")
            else:
                invoice_data['VAT'] = convert_to_number(invoice.fields.get("TotalTax").content)
            invoice_data['Brutto'] = InvoiceTotal
            invoice_data['Uwagi'] = uwagi
                
            if InvoiceTotal != SubTotal + invoice_data['VAT']:
                uwagi = append_warnings(uwagi, "Wartość brutto niezgodna z wartością netto i VAT")
            
            if check_nip(invoice_data['NIP']) != 'Czynny':
                uwagi = append_warnings(uwagi, "NIP nieaktywny")
                
            print("Wczytano fakturę: ", invoiceId, "z pliku: ", file_path)
        invoice_data["Nazwa pliku"] = rename_file(file_path, invoiceId, nazwa_sprzedajacego, year, month)
        all_invoices.append(invoice_data)
        
        
        
    df = pd.DataFrame(all_invoices)
    df.to_excel(output_csv, index=False)
    
        
def rename_file(file_path, invoiceId, nazwa_sprzedajacego, year, month):
    directory_path = os.path.dirname(file_path)
    new_file_name = (f"{year}-{month}-{replace_polish_and_special_chars(nazwa_sprzedajacego)}-{replace_polish_and_special_chars(invoiceId)}.pdf")
    new_file_path = os.path.join(directory_path, new_file_name)
    os.rename(file_path, new_file_path)
    print(f"Zmieniono nazwę pliku: {file_path} na {new_file_path}")
    return new_file_name
def replace_polish_and_special_chars(text):
    translation_table = str.maketrans({
        '.':'',' ':'','\\':'','/': '', 'ą': 'a', 'ć': 'c', 'ę': 'e', 'ł': 'l','\n':'',
        'ń': 'n', 'ó': 'o', 'ś': 's', 'ż': 'z', 'ź': 'z',
        'Ą': 'A', 'Ć': 'C', 'Ę': 'E', 'Ł': 'L',
        'Ń': 'N', 'Ó': 'O', 'Ś': 'S', 'Ż': 'Z', 'Ź': 'Z'
    })
    # Use the translate method with the translation table
    new_text = str.upper(text).translate(translation_table)
    return new_text

# Lista plików do przetworzenia
 
def get_all_file_paths(directory):
    
    return [os.path.join(directory, f) for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))and (f.endswith('.jpg') or f.endswith('.pdf'))]

previous_month_date = datetime.date.today() - relativedelta(months=1)
parser = argparse.ArgumentParser(description='Przetwarzanie faktur.')
parser.add_argument('month', type=str, nargs='?', default=previous_month_date.strftime('%m'), help='Miesiąc faktur (opcjonalnie, domyślnie bieżący miesiąc)')

# Parsowanie argumentów
year = previous_month_date.strftime('%Y')
args = parser.parse_args()

month = args.month
    
print(f"Przetwarzanie faktur z miesiąca: {month}") 

# Ścieżka do pliku wynikowego CSV
output_csv = f"{faktury_koszty_path}{year}\\{month}\\{year}-{month}-faktury.xlsx"

# Przetwarzanie i zapisanie wyników
listfiles = get_all_file_paths(f"{faktury_koszty_path}{year}\\{month}")
print (listfiles)
extract_and_save_to_csv(listfiles, output_csv, month, year)
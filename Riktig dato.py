#!/usr/bin/env python
# coding: utf-8

# In[7]:


import os
import pandas as pd
from glob import glob
import re
import datetime

# Katalogen der filene ligger
directory = '/Users/excel'

# Finn alle Excel-filer i katalogen
excel_files = glob(os.path.join(directory, '*.xlsx'))

# Finn den nyeste filen basert på modifikasjonstid
latest_file = max(excel_files, key=os.path.getmtime)

# Les inn den nyeste Excel-filen, og hopp over de første 7 radene
df = pd.read_excel(latest_file, skiprows=7)

# Legg til en ny kolonne til venstre for "Navn"
df.insert(0, 'Avd', '')

# Funksjon for å trekke ut tallverdiene fra "Navn"
def extract_number(name):
    match = re.match(r'(\d+)', str(name))
    return int(match.group(1)) if match else None

# Fyll ut kolonnen "Avd" med tallverdiene fra "Navn"
current_number = None
for index, row in df.iterrows():
    number = extract_number(row['Navn'])
    if number is not None:
        current_number = number
    df.at[index, 'Avd'] = current_number

# Fyll eventuelle gjenværende NaN-verdier med siste gyldige verdier
df['Avd'] = df['Avd'].fillna(method='ffill')
df['Avd'] = df['Avd'].fillna(0).astype(int)

# Funksjon for å filtrere bort rader basert på kriterier for "Konto"
def filter_konto(konto):
    if pd.isna(konto) or konto == '' or konto == ' ':
        return False
    if isinstance(konto, str) and len(konto) == 1 and konto.isalpha():
        return False
    if isinstance(konto, str) and konto.strip() == '1529':
        return False
    return True

# Filtrering av DataFrame basert på "Konto"-kolonnen
df = df[df['Konto'].apply(filter_konto)]

# Konverter celler under og til høyre for D1 til heltall der det er mulig
for col in df.columns[3:]:  # Kolonne D er den fjerde kolonnen (indeks 3)
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

# Konverter kolonnenavnene fra D1 og utover til yyyy/mm/dd format
new_columns = {}
for col in df.columns[3:]:
    try:
        # Bruk formatet 'dd.mm.yyyy' og konverter til 'yyyy/mm/dd'
        date_val = pd.to_datetime(col, format='%d.%m.%Y')
        # Lagre i riktig format
        new_columns[col] = date_val.strftime("%Y/%m/%d")
    except ValueError:
        pass  # Hopp over kolonner som ikke kan konverteres

# Oppdater kolonnenavnene i DataFrame
df.rename(columns=new_columns, inplace=True)

# Filtrer bort rader med nummer 10 og 90 i kolonnen "Avd"
df = df[~df['Avd'].isin([10, 90])]

# Sti til den nye Excel-filen
output_file_path = '/Users/modified_file.xlsx'

# Lagre DataFrame til en ny Excel-fil med formler
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Legg til DATEVALUE formler i kolonneoverskriftene
    for col_num, value in enumerate(df.columns.values):
        if isinstance(value, str) and re.match(r'^\d{4}/\d{2}/\d{2}$', value):
            worksheet.write_formula(0, col_num, f'=DATEVALUE("{value}")')

print(f"Done")


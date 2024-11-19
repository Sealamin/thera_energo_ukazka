import pandas as pd

import numpy as np

# Poznámka k programu
# Program načítá údaje z Excel tabulky => mohl by i stahovat z přímo interní firemní databáze

file_path = 'data.xlsx'  # zde se vloží odkaz na konkrétní sešit .xlsx

sheet_name = 'Sheet1'  # zde se vloží název konkrétního listu

data = pd.read_excel(file_path, sheet_name=sheet_name)


# Kontrola prázdných buněk v prvním sloupci listu

empty_cells = data['Column1'].isnull().sum()

print(f"Počet prázdných buněk činí: {empty_cells}")


# Ukázka automatizace procesu hledání pomocí kombinace INDEX a MATCH

# Indexujeme konkrétní položku z druhého sloupce

value_to_find = 'SomeValue'  # 

index = data[data['Column2'] == value_to_find].index.tolist()

print(f"Index of '{value_to_find}' in 'Column2': {index}")



lookup_data = pd.DataFrame({

    'Key': ['A', 'B', 'C'],

    'Value': [10, 20, 30]

})


# Spojení dvou sloupců

data = pd.merge(data, lookup_data, left_on='Column2', right_on='Key', how='left')


# Přejmenování sloupce pro větší přehlednost

data.rename(columns={'Value': 'LookupValue'}, inplace=True)


# Apliakce podmíněného formátování na hodnoty

data['Condition'] = np.where(data['LookupValue'] > 15, 'High', 'Low')


# Uložení nové dataframu do souboru

output_file_path = 'output_data.xlsx'

with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:

    data.to_excel(writer, index=False, sheet_name='ProcessedData')

    workbook = writer.book

    worksheet = writer.sheets['ProcessedData']

    

    # Definování formátu podmíněného formátování

    high_format = workbook.add_format({'bg_color': 'yellow', 'font_color': 'black'})

    

    # Aplikace podmíněného formátování na konkrétní sloupec tabulky

    worksheet.conditional_format('F2:F{}'.format(len(data) + 1), 

                                  {'type': 'text',

                                   'criteria': 'containing',

                                   'value': 'High',

                                   'format': high_format})


print("Zpracovaná data byla uložena do souboru 'output_data.xlsx'.")

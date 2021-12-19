import io
import math

import requests
import zipfile
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

Tk().withdraw()
filename = askopenfilename()


def get_table_from_ruchess(table_name):
    r = requests.get(f'https://ratings.ruchess.ru/api/{table_name}.csv.zip')
    z = zipfile.ZipFile(io.BytesIO(r.content))
    table = pd.read_csv(z.open(f'{table_name}.csv'))
    return table


standard_rating = get_table_from_ruchess('smanager_standard')
rapid_rating = get_table_from_ruchess('smanager_rapid')
blitz_rating = get_table_from_ruchess('smanager_blitz')

table_to_update = pd.read_excel(filename)

standard_rating_merge = pd.merge(table_to_update, standard_rating[['ID_No', 'Rtg_Nat']], how='left', left_on='ФШР ID', right_on='ID_No')
rapid_rating_merge = pd.merge(table_to_update, rapid_rating[['ID_No', 'Rtg_Nat']], how='left', left_on='ФШР ID', right_on='ID_No')
blitz_rating_merge = pd.merge(table_to_update, blitz_rating[['ID_No', 'Rtg_Nat']], how='left', left_on='ФШР ID', right_on='ID_No')

table_to_update['Classic'] = standard_rating_merge['Rtg_Nat']
table_to_update['Rapid'] = rapid_rating_merge['Rtg_Nat']
table_to_update['Blitz'] = blitz_rating_merge['Rtg_Nat']


def make_fcr_hyperlink(value):
    if math.isnan(value):
        return value
    url = "https://ratings.ruchess.ru/people/{}"
    return '=HYPERLINK("%s", "%s")' % (url.format(int(value)), int(value))


def make_fide_hyperlink(value):
    if math.isnan(value):
        return value
    url = "https://ratings.fide.com/profile/{}"
    return '=HYPERLINK("%s", "%s")' % (url.format(int(value)), int(value))


table_to_update['ФШР ID'] = table_to_update['ФШР ID'].apply(lambda x: make_fcr_hyperlink(x))
table_to_update['Fide ID'] = table_to_update['Fide ID'].apply(lambda x: make_fide_hyperlink(x))


def write_to_excel(table, output_file):
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter', date_format="YYYY-MM-DD", datetime_format="YYYY-MM-DD")
    table.to_excel(writer, index=False, sheet_name='Rating')

    worksheet = writer.sheets['Rating']

    worksheet.set_column('A:C', 20)
    worksheet.set_column('E:E', 20)
    worksheet.set_column('F:N', 10)
    worksheet.set_column('O:O', 25)
    worksheet.set_column('P:P', 20)

    writer.save()


write_to_excel(table_to_update, 'output.xlsx')


import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter
from urllib.parse import quote_plus
from sqlalchemy import create_engine
import win32com.client as win32
from datetime import datetime, timedelta

def count_images_today(directory):
    today = datetime.today().date()
    image_count = 0
    files_counted = set()
    extensions_image = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']
    for filename in os.listdir(directory):
        if any(filename.lower().endswith(ext) for ext in extensions_image):
            file_path = os.path.join(directory, filename)
            creation_time = datetime.fromtimestamp(os.path.getctime(file_path)).date()
            if creation_time == today:
                prefix = filename[:5]
                if prefix not in files_counted:
                    files_counted.add(prefix)
                    image_count += 1
    return image_count
directory_paths = {
    'Transcamino': 'C:\\Users\\lmsantos\\OneDrive - Ecourbis Ambiental SA\\Chamado 191667 - QR-Code Containeres\\01 - Transcamino',
    'MC Lopes': 'C:\\Users\\lmsantos\\OneDrive - Ecourbis Ambiental SA\\Chamado 191667 - QR-Code Containeres\\02 - MC Lopes\\FOTOS',
    'Rodomarca': 'C:\\Users\\lmsantos\\OneDrive - Ecourbis Ambiental SA\\Chamado 191667 - QR-Code Containeres\\03 - Rodomarca',
    'Dois Aranha': 'C:\\Users\\lmsantos\\OneDrive - Ecourbis Ambiental SA\\Chamado 191667 - QR-Code Containeres\\04 - Dois Aranha'}
today = datetime.today().strftime('%d/%m/%Y')
results = {'DATA': today}
conn_str = (
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=SEVSUL-22\\WEB;'
    'DATABASE=CONTEINERESSGC;'
    'Trusted_Connection=yes;')
conn_str = quote_plus(conn_str)
engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')
consulta_sql = """
SELECT NomeEmpresa, COUNT(*) AS Contagem
FROM tbConteineres
WHERE DataQrCode >= CAST(GETDATE() AS DATE)  AND NumeroQrCode > '0'
GROUP BY NomeEmpresa
"""
results_df = pd.read_sql_query(consulta_sql, engine)
resultado_transcamino = 0
resultado_mc_lopes = 0
resultado_rodomarca = 0
resultado_dois_aranha = 0
resultado_ecourbis = 0
for index, row in results_df.iterrows():
    if row['NomeEmpresa'] == 'TRANSCAMINO':
        resultado_transcamino = row['Contagem']
    elif row['NomeEmpresa'] == 'MC LOPES':
        resultado_mc_lopes = row['Contagem']
    elif row['NomeEmpresa'] == 'RODOMARCA':
        resultado_rodomarca = row['Contagem']
    elif row['NomeEmpresa'] == 'DOIS ARANHA':
        resultado_dois_aranha = row['Contagem']
    elif row['NomeEmpresa'] == 'ECOURBIS AMBIENTAL SA':
        resultado_ecourbis = row['Contagem']
consulta_sql_cadastradas = """
SELECT COUNT(*)
FROM tbConteineres
WHERE DataQrCode >= CAST(GETDATE() AS DATE)  AND NumeroQrCode > '0'
"""
results_df_cadastradas = pd.read_sql_query(consulta_sql_cadastradas, engine)
resultado_cadastradas = results_df_cadastradas.iloc[0, 0]
consulta_sql_sul = """
SELECT COUNT(*)
FROM tbConteineres
WHERE DataQrCode >= CAST(GETDATE() AS DATE)  AND NumeroQrCode > '0' and UnidadeId = '1'
"""
results_df_sul = pd.read_sql_query(consulta_sql_sul, engine)
resultado_sul = results_df_sul.iloc[0, 0]
consulta_sql_leste = """
SELECT COUNT(*)
FROM tbConteineres
WHERE DataQrCode >= CAST(GETDATE() AS DATE) AND NumeroQrCode > '0' and UnidadeId = '2'
"""
results_df_leste = pd.read_sql_query(consulta_sql_leste, engine)
resultado_leste = results_df_leste.iloc[0, 0]
df_results = pd.DataFrame([results])
for folder_name, directory_path in directory_paths.items():
    quantity_images = count_images_today(directory_path)
    results[folder_name] = quantity_images
results['Geral'] = sum(results[folder_name] for folder_name in directory_paths.keys())
results['Cadastradas_Ecourbis'] = resultado_ecourbis
results['Cadastradas_Transcamino'] = resultado_transcamino
results['Cadastradas_MC_Lopes'] = resultado_mc_lopes
results['Cadastradas_Rodomarca'] = resultado_rodomarca
results['Cadastradas_Dois_Aranha'] = resultado_dois_aranha
results['Cadastradas_Geral'] = resultado_cadastradas
results[''] = ''
results['Cadastrados_Sul'] = resultado_sul
results['Cadastrados_Leste'] = resultado_leste
df_results = pd.DataFrame([results])
excel_name = 'C:\\Users\\lmsantos\\Documents\\imagens_adicionadas.xlsx'
try:
    df_existent = pd.read_excel(excel_name, engine='openpyxl', header=1)  # Ler o arquivo com o cabeçalho na segunda linha
    df_existent = df_existent[df_existent['DATA'] != 'Total']
    df_final = pd.concat([df_existent, df_results], ignore_index=True)
except FileNotFoundError:
    df_final = df_results
total_line = {'DATA': 'Total'}
total_line['Transcamino'] = df_final['Transcamino'].sum()
total_line['MC Lopes'] = df_final['MC Lopes'].sum()
total_line['Rodomarca'] = df_final['Rodomarca'].sum()
total_line['Dois Aranha'] = df_final['Dois Aranha'].sum()
total_line['Cadastradas_Ecourbis'] = df_final['Cadastradas_Ecourbis'].sum()
total_line['Cadastradas_Transcamino'] = df_final['Cadastradas_Transcamino'].sum()
total_line['Cadastradas_MC_Lopes'] = df_final['Cadastradas_MC_Lopes'].sum()
total_line['Cadastradas_Rodomarca'] = df_final['Cadastradas_Rodomarca'].sum()
total_line['Cadastradas_Dois_Aranha'] = df_final['Cadastradas_Dois_Aranha'].sum()
total_line['Geral'] = df_final['Geral'].sum()
total_line['Cadastradas_Geral'] = df_final['Cadastradas_Geral'].sum()
total_line[''] = ''
total_line['Cadastrados_Sul'] = df_final['Cadastrados_Sul'].sum()
total_line['Cadastrados_Leste'] = df_final['Cadastrados_Leste'].sum()
df_final.loc[len(df_final)] = total_line
if 'Cadastradas_Ecourbis' not in df_final.columns:
    df_final.insert(df_final.columns.get_loc('Geral') + 1, 'Cadastradas_Ecourbis', '')
if 'Cadastrados_Sul' not in df_final.columns:
    df_final.insert(df_final.columns.get_loc('Cadastradas_Geral') + 2, 'Cadastrados_Sul', '')
if 'Cadastrados_Leste' not in df_final.columns:
    df_final.insert(df_final.columns.get_loc('Cadastrados_Sul') + 1, 'Cadastrados_Leste', '')

df_final_shifted = pd.DataFrame(columns=df_final.columns)
df_final_shifted.loc[0] = df_final.columns
df_final_shifted = pd.concat([df_final_shifted, df_final], ignore_index=True)

with pd.ExcelWriter(excel_name, engine='openpyxl') as writer:
    df_final_shifted.to_excel(writer, index=False, header=False, startrow=1)
wb = load_workbook(excel_name)
ws = wb.active

if not any(['B1:F1' in str(range) for range in ws.merged_cells.ranges]):
    ws.merge_cells('B1:F1')
    blue_fill = PatternFill(start_color='90D5AC', end_color='90D5AC', fill_type='solid')
    green_fill = PatternFill(start_color="60BB47", end_color="60BB47", fill_type="solid")
    for row in ws.iter_rows(min_row=1, max_row=2, min_col=2, max_col=6):
        for cell in row:
            cell.fill = blue_fill
            cell.font = Font(bold=True)
    for cell in ws['A']:
        cell.font = Font(bold=True)
    for col_num, key in enumerate(total_line.keys(), start=1):
        cell = ws.cell(row=ws.max_row, column=6)
        cell.fill = green_fill
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))
        cell.border = thin_border

if not any(['G1:L1' in str(range) for range in ws.merged_cells.ranges]):
    ws.merge_cells('G1:L1')
    blue_fill = PatternFill(start_color='1995a8', end_color='1995a8', fill_type='solid')
    for row in ws.iter_rows(min_row=1, max_row=2, min_col=7, max_col=12):
        for cell in row:
            cell.fill = blue_fill
            cell.font = Font(bold=True)
    for col_num, key in enumerate(total_line.keys(), start=1):
        if col_num <= 12:
            cell = ws.cell(row=ws.max_row, column=col_num)
            cell.fill = green_fill
            cell.border = thin_border
            cell.font = Font(bold=True)

if not any(['N1:O1' in str(range) for range in ws.merged_cells.ranges]):
    ws.merge_cells('N1:O1')
    blue_fill = PatternFill(start_color='90D5AC', end_color='90D5AC', fill_type='solid')
    for row in ws.iter_rows(min_row=1, max_row=2, min_col=14, max_col=15):
        for cell in row:
            cell.fill = blue_fill
            cell.font = Font(bold=True)
            cell.border = thin_border
    for col_num, key in enumerate(total_line.keys(), start=1):
        if col_num in [14, 15]:
            cell = ws.cell(row=ws.max_row, column=col_num)
            cell.fill = green_fill
            cell.border = thin_border
            cell.font = Font(bold=True)

ws['B1'] = "Inclusão das fotos das Placas QR Code"
ws['G1'] = "Cadatros das Placas QR Code no Sistema"
ws['N1'] = "Cadastro Por Unidade"
ws['M2'] = ""
ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
ws['B1'].alignment = Alignment(horizontal='center', vertical='center')
ws['G1'].alignment = Alignment(horizontal='center', vertical='center')
ws['N1'].alignment = Alignment(horizontal='center', vertical='center')
columns_to_resize = ['A', 'B', 'C', 'D', 'E', 'F']
for col in columns_to_resize:
    ws.column_dimensions[col].width = 100 / 7.5
columns_to_resize = ['H', 'I', 'J', 'K']
for col in columns_to_resize:
    ws.column_dimensions[col].width = 175 / 7.5
columns_to_resize = ['G', 'L']
for col in columns_to_resize:
    ws.column_dimensions[col].width = 135 / 7.5
columns_to_resize = ['N', 'O']
for col in columns_to_resize:
    ws.column_dimensions[col].width = 140 / 7.5
for row in ws.iter_rows(min_row=1, max_row=2, min_col=2, max_col=12):
    for cell in row:
        cell.border = thin_border

results['Cadastrados_Sul'] = resultado_sul
results['Cadastrados_Leste'] = resultado_leste
df_results = pd.DataFrame([results])
try:
    df_existent = pd.read_excel(excel_name, engine='openpyxl', header=1)
    df_existent = df_existent[df_existent['DATA'] != 'Total']
    df_final = pd.concat([df_existent, df_results], ignore_index=True)
except FileNotFoundError:
    df_final = df_results
total_line['Cadastrados_Sul'] = df_final['Cadastrados_Sul'].sum()
total_line['Cadastrados_Leste'] = df_final['Cadastrados_Leste'].sum()
df_final.loc[len(df_final)] = total_line
wb.save(excel_name)
print(f"Os resultados foram salvos no arquivo {excel_name}.")
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
signature_path = os.path.join(os.environ['APPDATA'], r"Microsoft\Signatures\Leandro.txt")
with open(signature_path, 'r', encoding='UTF-16 LE') as file:
    signature = file.read()
signature = signature.replace('\n', '<br>')
email.To = ""#
email.Subject = f"QR-Codes contêineres - {today}"
image_path = r"C:\Users\lmsantos\AppData\Roaming\Microsoft\Signatures\Leandro (lmsantos@ecourbis.com.br)_arquivos\image001.png"
if not os.path.exists(image_path):
    raise FileNotFoundError(f"O arquivo de imagem não foi encontrado: {image_path}")
image_attachment = email.Attachments.Add(image_path)
image_attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image001.png")
excel_attachment = email.Attachments.Add(excel_name)
email.HTMLBody = f"""
<p style="margin: 0;">Prezados,</p>
<p style="margin: 0;">Boa Tarde!</p>
<br style="line-height: 0;">
<p style="margin: 0;">Segue em Anexo o relatório do comparativo entre as fotos inseridas e os QR-Codes cadastrados no sistema.</p>
<br style="line-height: 0;">
<p style="margin: 10;"><i><b>HOJE {today}</b></i></p>
<p style="margin: 0;">Fotos incluidas: {sum(results[folder_name] for folder_name in directory_paths.keys())} | Informações Incluida no sistema: {resultado_cadastradas}</p>
<br style="line-height: 0;">
<br style="line-height: 0;">
<p style="margin: 10;"><i><b>Geral</b></i></p>
<p style="margin: 0;">Fotos incluidas: {total_line['Geral']} | Informações Incluida no sistema: {total_line['Cadastradas_Geral']}</p>
<br style="line-height: 0;">
<br style="line-height: 0;">
<p style="margin: 0;">Atenciosamente,</p>
<table>
    <tr>
        <td><img src="cid:image001.png" width="150" height="100"></td> 
        <td>{signature}</td>
    </tr>
</table>
"""
email.Send()
print("Email Enviado")

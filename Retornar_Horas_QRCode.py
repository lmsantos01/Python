import os
from datetime import datetime
import pandas as pd
from urllib.parse import quote_plus
from sqlalchemy import create_engine
import win32com.client as win32

def count_images_today(directory):
    today = datetime.today().date()
    image_count = 0
    files_counted = set()
    extensions_image = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']
    for filename in os.listdir(directory):
        if any(filename.lower().endswith(ext) for ext in extensions_image):
            file_path = os.path.join(directory, filename)
            creation_time = datetime.fromtimestamp(os.path.getctime(file_path)).date()
            #Confere os 5 primeiros Caracteres.
            if creation_time == today:
                prefix = filename[:5]
                if prefix not in files_counted:
                    files_counted.add(prefix)
                    image_count += 1
    return image_count
directory_paths = {
    'Transcamino': 'C:\\Users\\lmsantos\\OneDrive - Ecourbis Ambiental SA\\Chamado 191667 - QR-Code Containeres\\01 - Transcamino',
    'MC Lopes': 'C:\\Users\\lmsantos\\OneDrive - Ecourbis Ambiental SA\\Chamado 191667 - QR-Code Containeres\\02 - MC Lopes',
    'Rodomarca': 'C:\\Users\\lmsantos\\OneDrive - Ecourbis Ambiental SA\\Chamado 191667 - QR-Code Containeres\\03 - Rodomarca',
    'Metal Serv': 'C:\\Users\\lmsantos\\OneDrive - Ecourbis Ambiental SA\\Chamado 191667 - QR-Code Containeres\\04 - Dois Aranha'}
today = datetime.today().strftime('%d/%m/%Y')
results = {'DATA': today}

#Conexão com o Banco para extrair os Cadastrados.
conn_str = (
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=SEVSUL-22\\WEB;'
    'DATABASE=CONTEINERESSGC;'
    'Trusted_Connection=yes;')
conn_str = quote_plus(conn_str)
engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')
consulta_sql = """
SELECT COUNT(*) FROM tbConteineres WHERE DataQrCode >= CAST(GETDATE() AS DATE)
"""
results_df = pd.read_sql_query(consulta_sql, engine)
resultado_final = results_df.iloc[0, 0]
df_results = pd.DataFrame([results])
for folder_name, directory_path in directory_paths.items():
    quantity_images = count_images_today(directory_path)
    results[folder_name] = quantity_images
results['Geral'] = sum(results[folder_name] for folder_name in directory_paths.keys())
results['Cadastradas'] = resultado_final
df_results = pd.DataFrame([results])
excel_name = 'C:\\Users\\lmsantos\\Documents\\imagens_adicionadas.xlsx'
try:
    df_existent = pd.read_excel(excel_name, engine='openpyxl')
    df_existent = df_existent[df_existent['DATA'] != 'Total']
    df_final = pd.concat([df_existent, df_results], ignore_index=True)
except FileNotFoundError:
    df_final = df_results
total_line = {'DATA': 'Total'}
for column in directory_paths.keys():
    total_line[column] = df_final[column].sum()
total_line['Geral'] = df_final['Geral'].sum()
total_line['Cadastradas'] = df_final['Cadastradas'].sum()
df_final.loc[len(df_final)] = total_line
df_final.to_excel(excel_name, index=False)
print(f"Os resultados foram salvos no arquivo {excel_name}.")

#Envio do E-mail.
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
signature_path = os.path.join(os.environ['APPDATA'], r"Microsoft\Signatures\Leandro.txt")
with open(signature_path, 'r', encoding='UTF-16 LE') as file:
    signature = file.read()
signature = signature.replace('\n', '<br>')
email.To = "lmsantos@ecourbis.com.br"
#email.BCC = "dnsilva@ecourbis.com.br"
email.Subject = f"QR-Codes - {today}"
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
<p style="margin: 0;">Fotos incluidas no dia: {sum(results[folder_name] for folder_name in directory_paths.keys())}</p>
<p style="margin: 0;">Total Incluida no sistema no dia: {resultado_final}</p>
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

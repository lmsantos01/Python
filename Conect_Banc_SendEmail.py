import pandas as pd
import pyodbc
from sqlalchemy import create_engine
from datetime import datetime
from urllib.parse import quote_plus
import win32com.client as win32
import os
conn_str = (
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=SEVSUL-26\\TOTVS;'
    'DATABASE=DADOSADV;'
    'Trusted_Connection=yes;'
)
conn_str = quote_plus(conn_str)
engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')
now = datetime.now()
mes = (now.month - 1) if now.month > 1 else 12
ano = now.year if now.month > 1 else now.year - 1
mes = f"{mes:02d}"

def mes_por_extenso(mes):
    meses = {
        '01': 'janeiro',
        '02': 'fevereiro',
        '03': 'março',
        '04': 'abril',
        '05': 'maio',
        '06': 'junho',
        '07': 'julho',
        '08': 'agosto',
        '09': 'setembro',
        '10': 'outubro',
        '11': 'novembro',
        '12': 'dezembro'
    }
    return meses.get(mes, 'Mês inválido')

mes_extenso = mes_por_extenso(mes)
consulta_sql = """
SELECT 
    C.cd_veiculo, 
    V.placa, 
    C.dh_abastec, 
    C.cd_filial, 
    SUBSTRING(CONVERT(VARCHAR, C.dh_abastec, 120), 1, 4) AS ano,
    SUBSTRING(CONVERT(VARCHAR, C.dh_abastec, 120), 6, 2) AS mes,
    SUBSTRING(CONVERT(VARCHAR, C.dh_abastec, 120), 9, 2) AS dia,
    (CASE 
        WHEN C.cd_ccusto = 1 OR C.cd_ccusto = 11 THEN '102010000' 
        WHEN C.cd_ccusto = 2 OR C.cd_ccusto = 21 THEN '101010000' 
        WHEN C.cd_ccusto = 12 THEN '102020000' 
        WHEN C.cd_ccusto = 13 THEN '102040000' 
        WHEN C.cd_ccusto = 14 THEN '102030000' 
        WHEN C.cd_ccusto = 22 THEN '101020000' 
        WHEN C.cd_ccusto = 25 THEN '101030000' 
        WHEN C.cd_ccusto = 23 THEN '101040000' 
        WHEN C.cd_ccusto = 24 THEN '104080000' 
        WHEN C.cd_ccusto = 1121 THEN '102400000' 
        ELSE '999999999' 
    END) AS cd_ccusto, 
    C.qt_litros 
FROM 
    [LinkedGuberman].[GUBERMAN].[dbo].[CONSUMO] C 
INNER JOIN 
    [LinkedGuberman].[GUBERMAN].[dbo].[Veiculo] V 
ON 
    C.cd_veiculo = V.cd_veiculo
WHERE 
    SUBSTRING(CONVERT(VARCHAR, C.dh_abastec, 120), 6, 2) = ? AND SUBSTRING(CONVERT(VARCHAR, C.dh_abastec, 120), 1, 4) = ?
"""
df = pd.read_sql_query(consulta_sql, engine, params=(mes, ano))
excel_path = f'C:\\Users\\lmsantos\\Desktop\\Curso\\Relatorio_Abastecimento_{mes}_{ano}.xlsx'
df.to_excel(excel_path, index=False)
# Envio do e-mail
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
signature_path = os.path.join(os.environ['APPDATA'], r"Microsoft\Signatures\Leandro.txt")
with open(signature_path, 'r', encoding='UTF-16 LE') as file:
    signature = file.read()
signature = signature.replace('\n', '<br>')
email.To = "nshiobara@ecourbis.com.br;jmatheus@ecourbis.com.br"
email.BCC = "lmsantos@ecourbis.com.br"
email.Subject = f"Relatório de Abastecimento para área Técnica - {mes_extenso}"
image_path = r"C:\Users\lmsantos\AppData\Roaming\Microsoft\Signatures\Leandro (lmsantos@ecourbis.com.br)_arquivos\image001.png"
if not os.path.exists(image_path):
    raise FileNotFoundError(f"O arquivo de imagem não foi encontrado: {image_path}")
image_attachment = email.Attachments.Add(image_path)
image_attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image001.png")
excel_attachment = email.Attachments.Add(excel_path)

email.HTMLBody = f"""
<p style="margin: 0;">Prezados,</p>
<p style="margin: 0;">Boa Tarde!</p>
<br style="line-height: 0;">
<p style="margin: 0;">Segue relatório de abastecimento para área técnica referente ao mês de {mes_extenso}.</p>
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

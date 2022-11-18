import os
import pandas as pd
import fdb


nome = os.getlogin()

dst_path = r'MTK:C:/Microsys/MsysIndustrial/Dados/MSYSDADOS.FDB'
excel_path = r'C:/Users/'+nome+'/Desktop/morpheus.xlsx'

TABLE_NAME = 'CLIENTES'

SELECT = 'select CLI_NOME, CLI_FONE, CLI_FAX, CLI_CELULAR, CLI_ENT_CIDADE, CLI_ENT_ESTADO' \
         ' from %s' % (TABLE_NAME)

con = fdb.connect(dsn=dst_path, user='SYSDBA', password='masterkey', charset='UTF8')

cur = con.cursor()
cur.execute(SELECT)

table_rows = cur.fetchall()

df = pd.DataFrame(table_rows)
df = df.applymap(lambda x: x.encode('unicode_escape').
                 decode('utf-8') if isinstance(x, str) else x)

print(df.head(10))

df.to_excel(excel_path, index=False)





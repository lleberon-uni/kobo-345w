from pathlib import Path
from openpyxl import load_workbook
p=Path(__file__).resolve().parent.parent/'345w form fields v9.4.xlsx'
w=load_workbook(p,read_only=True,data_only=True)
for sheet in ('Logframe Actividades EQ','Logframe Actividades','Logframe Indicadores EQ','Logframe Indicadores'):
 print('\n',sheet)
 for r in w[sheet].iter_rows(values_only=True):
  if any(str(x or '').startswith(('GCLCO/CCA4','GCLCO/CCA1','GCLCO/CCA3','GCLCO/SCA13')) for x in r):
   print(repr(r[:9]).encode('ascii','backslashreplace').decode())

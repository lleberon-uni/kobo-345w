from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
import json
root=Path(__file__).resolve().parent.parent
w=load_workbook(root/'345w form fields v9.4.xlsx',data_only=True)
for s in w:
 print('HOJA',s.title,'TABLAS',[(t.name,t.ref) for t in s.tables.values()])
 for t in s.tables.values():
  if t.name.lower()=='eat':
   a,b,c,d=range_boundaries(t.ref)
   for row in s.iter_rows(min_row=b,max_row=min(d,b+5),min_col=a,max_col=c,values_only=True): print('EAT',json.dumps(row,ensure_ascii=True,default=str))
s=w['Vocab']
for row in s.iter_rows(min_col=8,max_col=8,values_only=True):
 if row[0]: print('TIPO',repr(row[0]))
for p in sorted((root/'upload templates').glob('*.xlsx')):
 b=load_workbook(p,read_only=True,data_only=True);s=b.worksheets[0];print('TEMPLATE',p.name)
 for row in s.iter_rows(min_row=1,max_row=3,values_only=True):print(json.dumps(row,ensure_ascii=True,default=str))
 b.close()
w.close()

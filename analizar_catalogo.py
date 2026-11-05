from openpyxl import load_workbook
from pathlib import Path
from collections import Counter
p=Path(__file__).resolve().parent.parent/'345w form fields v9.4.xlsx'
w=load_workbook(p,read_only=True,data_only=True)
for r in w['Organizaciones'].iter_rows(values_only=True):
 if any(x in str(r).lower() for x in ['aseinc','unicef','cecodap','fundainil','accion contra','acció']): print('ORG',r[:12])
raise SystemExit
for s in w:
 print('\nHOJA',s.title,s.max_row,s.max_column)
 for i,r in enumerate(s.iter_rows(values_only=True),1):
  tags=[(j,str(v)) for j,v in enumerate(r) if isinstance(v,str) and v.startswith('#')]
  if tags:
   print('FILA',i,'TAGS',tags)
   for row in s.iter_rows(min_row=i+1,max_row=min(i+3,s.max_row),values_only=True):
    print('EJEMPLO',[(j,str(row[j])[:120]) for j,v in tags if row[j] is not None])

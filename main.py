# -*- coding: utf-8 -*-

import sys, os
import pandas as pd

len_args = len(sys.argv)

def index_argv_bytext(texto):
   i   = 0
   enc = False
   while (not enc) and (i < len_args):
      enc = sys.argv[i] == texto
      if not enc:
         i = i + 1
         
   if enc:
      return i
   else:
      return -999
      
def text_argv_byindex(index):
   if (index >= 0) and (index < len_args):
      return sys.argv[index]
   else: 
      return ''

if (len_args < 2) or (index_argv_bytext('--help') > -1):
   print('ExcelToTXT --file [excel_file] --sheet [sheet_name] --out [txt_file] --log [S/N]')
   sys.exit(1)
else:
   try:
      filename  = text_argv_byindex(index_argv_bytext('--file') + 1)
      sheetname = text_argv_byindex(index_argv_bytext('--sheet') + 1) 
      outname   = text_argv_byindex(index_argv_bytext('--out') + 1)
      log       = text_argv_byindex(index_argv_bytext('--log') + 1).lower() == 's'

      if sheetname == '':
         x_sheet = 0
      elif sheetname.isnumeric():
         x_sheet = int(eval(sheetname))
      else:
         x_sheet = sheetname.strip()

      if outname == '':
         outname = os.path.splitext(filename)[0] + '.txt'
  
      if log:
         print('Excel File : ', filename)
         print('Nombre Hoja: ', x_sheet)
         print('Out: ', outname)
       
      if (filename != '') and os.path.exists(filename):
         
         df = pd.read_excel(filename, sheet_name=x_sheet)
         
         #print(df)
         if log:
            print('Exportando ', len(df), ' registros en ', outname)
         df.to_csv(outname, sep='\t', doublequote=True, index=False, header=False, date_format='dd/mm/yyyy')

         sys.exit(0)
      else:
         sys.exit(2)
   except:
      sys.exit(3)

# -*- coding: utf-8 -*-

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import colors, Border, Color, Side, PatternFill, Font, GradientFill, Alignment, NamedStyle
from datetime import datetime

def create_workbook(title):
  wb = Workbook()
  ws = wb.active
  ws.title = title
  return [wb, ws]

def header(ws, enc, h1, h2):
  ws['A1'] = enc
  title = ws['A1']
  title = format_text(ws['A1'], "center", "center")
  ws.merge_cells('A1:K1')

  ws['A3'] = h1
  ws['B3'] = h2
  title = ws['A3']
  title = format_text(ws['A3'], "left", "center")

  ws['A4'] = 'FECHA'
  ws['B4'] = 'Desde  Al '
  title = ws['A4']
  title = format_text(ws['A4'], "left", "center")

  ws['A5'] = 'RUTA PREVENTA'
  ws['B5'] = ''
  title = ws['A5']
  title = format_text(ws['A5'], "left", "center")

  ws['A6'] = 'RUTA ENTREGA'
  ws['B6'] = ''
  title = ws['A6']
  title = format_text(ws['A6'], "left", "center")

  ws['A7'] = 'PRODUCTO'
  ws['B7'] = ''
  title = ws['A7']
  title = format_text(ws['A7'], "left", "center")

  ws['A8'] = 'DESCRIPCION CANAL'
  ws['B8'] = ''
  title = ws['A8']
  title = format_text(ws['A8'], "left", "center")

  ws['A9'] = 'GRUPO'
  ws['B9'] = ''
  title = ws['A9']
  title = format_text(ws['A9'], "left", "center")
  return ws

def format_text(celda, alingH, alingV):
  title = celda
  title.font = Font(size=12,bold=True)
  title.alignment = Alignment(horizontal=alingH, vertical=alingV)

def resize_cells(ws, size):
  dims = {}
  for row in ws.rows:
    for cell in row:
      if cell.value:
        dims[cell.column] = size
  for col, value in dims.items():
    ws.column_dimensions[col].width = value
  return ws

def load_rows(ws, data):
  for row in data:
    row_list = list(row)
    ws.append(row_list)
  return ws

def paint_par(ws, cell_header, data, num_col):
  rowPar = NamedStyle(name="rowPar")
  rowPar.fill = PatternFill("solid", fgColor="E0ECF8")

  for column in range(1,len(cell_header)+1):
    column_letter = get_column_letter(column)
    for rowD in range(11,len(data)+11):
      if(rowD % 2 == 0):
        ws[column_letter + str(rowD)].style = rowPar
      if(column > num_col):
        ws[column_letter + str(rowD)].number_format = '#,#0.0'
  return ws

def load_filters(ws, init_vector):
  FullRange = init_vector +':' + get_column_letter(ws.max_column) \
  + str(ws.max_row)
  ws.auto_filter.ref = FullRange
  return ws

def adds_title_format(ws, table_header, font_color="FFFFFF", fill_color="afbcd7"):
  headOpe = NamedStyle(name="headOpe")
  headOpe.alignment = Alignment(horizontal='center')
  headOpe.fill = PatternFill("solid", fgColor=fill_color)
  headOpe.font = Font(color=font_color, size=12, bold=True)
  ws.insert_rows(10)
  for row in ws.iter_rows('A11:'+get_column_letter(len(table_header))+'11'):
    for cell in row:
      cell.style = headOpe
  return ws

def format_date(_date, time):
  date_time = _date + " " + time
  return date_time

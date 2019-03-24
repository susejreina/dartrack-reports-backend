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

def header(ws, enc, h1, h2, filters):
  ws['A1'] = enc
  title = ws['A1']
  title = format_text(ws['A1'], "center", "center")
  ws.merge_cells('A1:K1')

  ws['A3'] = h1
  ws['B3'] = h2
  title = ws['A3']
  title = format_text(ws['A3'], "left", "center")

  init_date  = filters.get('init_date', False)
  last_date  = filters.get('last_date', False)
  canal_est = filters.get('canal_est', False)
  presale_route  = filters.get('presale_route', False)
  delivery_route  = filters.get('delivery_route', False)
  product = filters.get('product', False)
  group_product = filters.get('group_product',False)

  nro = 3
  if init_date and last_date:
    nro += 1
    ws['A'+str(nro)] = 'FECHA'
    ws['B'+str(nro)] = 'Desde '+init_date+' Al '+last_date
    title = ws['A'+str(nro)]
    title = format_text(ws['A'+str(nro)], "left", "center")

  if presale_route:
    if presale_route["presale_route_id"] != 0:
      nro += 1
      ws['A'+str(nro)] = 'RUTA PREVENTA'
      ws['B'+str(nro)] = str(presale_route["presale_route"])
      title = ws['A'+str(nro)]
      title = format_text(ws['A'+str(nro)], "left", "center")

  if delivery_route:
    if delivery_route["delivery_route_id"] != 0:
      nro += 1
      ws['A'+str(nro)] = 'RUTA ENTREGA'
      ws['B'+str(nro)] = str(delivery_route["delivery_route"])
      title = ws['A'+str(nro)]
      title = format_text(ws['A'+str(nro)], "left", "center")

  if product:
    if product["product_id"] != 0:
      nro += 1
      ws['A'+str(nro)] = 'PRODUCTO'
      ws['B'+str(nro)] = str(product["product"])
      title = ws['A'+str(nro)]
      title = format_text(ws['A'+str(nro)], "left", "center")

  if canal_est:
    if canal_est["canal_est_id"] != 0:
      nro += 1
      ws['A'+str(nro)] = 'DESCRIPCION CANAL'
      ws['B'+str(nro)] = str(canal_est["canal_est"])
      title = ws['A'+str(nro)]
      title = format_text(ws['A'+str(nro)], "left", "center")

  if group_product:
    if group_product["group_product_id"] != 0:
      nro += 1
      ws['A'+str(nro)] = 'GRUPO'
      ws['B'+str(nro)] = str(group_product["group_product"])
      title = ws['A'+str(nro)]
      title = format_text(ws['A'+str(nro)], "left", "center")
  nro += 1
  return [ws,nro]

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

def paint_par(ws, cell_header, data, num_col,row=11, col_money=[],col_porc=[]):
  rowPar = NamedStyle(name="rowPar")
  rowPar.fill = PatternFill("solid", fgColor="E0ECF8")

  for column in range(1,len(cell_header)+1):
    column_letter = get_column_letter(column)
    for rowD in range(row,len(data)+row+1):
      if(rowD % 2 == 0):
        ws[column_letter + str(rowD)].style = rowPar
      if(column > num_col):
        ws[column_letter + str(rowD)].number_format = '#,##0.00'
      if(column in col_money):
        ws[column_letter + str(rowD)].number_format = '#,##0.00 $'
      if(column in col_porc):
        ws[column_letter + str(rowD)].number_format = '#,##0.00 %'
  return ws

def load_filters(ws, init_vector):
  FullRange = init_vector +':' + get_column_letter(ws.max_column) \
  + str(ws.max_row)
  ws.auto_filter.ref = FullRange
  return ws

def total_summary(ws, listTotal, numberRow, lonTableHeader, font_color="FFFFFF", fill_color="afbcd7",formatPercent=[],formatMoney=[],formatNumber=[]):
  
  ws.append(listTotal)
  
  totalOpe = NamedStyle(name="totalOpe")
  totalOpe.alignment = Alignment(horizontal='center')
  totalOpe.fill = PatternFill("solid", fgColor=fill_color)
  totalOpe.font = Font(color=font_color, size=12, bold=True)
  for row in ws.iter_rows('A'+str(numberRow)+':'+get_column_letter(lonTableHeader)+str(numberRow)):
    for cell in row:
      cell.style = totalOpe

  for col_letter in formatPercent:
    ws[col_letter+str(numberRow)].number_format = '#,##0.00 %'
  for col_letter in formatMoney:
    ws[col_letter+str(numberRow)].number_format = '#,##0.00 $'
  for col_letter in formatNumber:
    ws[col_letter+str(numberRow)].number_format = '#,##0.00'

  return ws

def adds_title_format(ws, lonTableHeader, font_color="FFFFFF", fill_color="afbcd7",rows=10):
  headOpe = NamedStyle(name="headOpe")
  headOpe.alignment = Alignment(horizontal='center')
  headOpe.fill = PatternFill("solid", fgColor=fill_color)
  headOpe.font = Font(color=font_color, size=12, bold=True)
  ws.insert_rows(rows)
  rows = rows + 1
  for row in ws.iter_rows('A'+str(rows)+':'+get_column_letter(lonTableHeader)+str(rows)):
    for cell in row:
      cell.style = headOpe
  return ws

def format_date(_date, time):
  date_time = _date + " " + time
  return date_time

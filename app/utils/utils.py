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

  for row in ws.iter_rows('A1:'+get_column_letter(len(table_header))+'1'):
    for cell in row:
      cell.style = headOpe
  return ws

def format_date(_date, time):
  date_time = _date + " " + time
  return date_time
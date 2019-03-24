# -*- coding: utf-8 -*-

import os
import sys
import psycopg2
from datetime import datetime, time, date

from flask import Flask, send_file, request, jsonify, json
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from app import app
from app.models import models

from app.utils import utils

@app.route('/api/v2/detail_client_sell', methods=['POST'])
def detail_client_sell(charset='utf-8'):
  jsonResponse = json.loads(request.data)
  filters = jsonResponse.get('filters', False)
  data = models.detail_client_sell(filters)
  wb, ws = utils.create_workbook("VTAS X CTE-13.1")
  print(jsonResponse.get('table', False))
  # table_exists = data.get('table', False)
  # headers_exists = data['table'].get('table', False)
  # if table_exists and headers_exists:
  table_header = jsonResponse['table']['headers']
  # else:
  #   response = { 'response': 'El header de la tabla es requerido'}
  #   return jsonify(response)  
  if len(data)>0:
    ws.append(list(table_header))
    # Method for load rows
    ws = utils.load_rows(ws, data)
    # Method for rezise cells
    # This method recives workbook active instance and
    ws = utils.resize_cells(ws, 15)
    # Method for load filters
    # This method recives workbook active instance and 
    # init vector on init the header_table
    ws = utils.load_filters(ws, 'A1')
    # Method for adds color to titles, first color si for the font
    # second color is for fill cell. Colors is in format RGB
    ws = utils.adds_title_format(ws, len(table_header), "000000", "afbcd7")

    nombre_archivo ="RC-13-1-"+datetime.now().date().strftime('%Y%m%d')+".xlsx"
    wb.save(nombre_archivo)
    return send_file('../'+nombre_archivo, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', True, nombre_archivo)
  else:
    response = { 'response': 'No hay registros'}
    return jsonify(response)
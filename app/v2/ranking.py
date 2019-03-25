# -*- coding: utf-8 -*-

import os
import sys
import psycopg2
from datetime import datetime, time, date

from flask import Flask, send_file, request, jsonify, json
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from app import app
from app.models import ranking

from app.utils import utils

@app.route('/api/v2/ranking_client', methods=['POST'])
def ranking_client(charset='utf-8'):
  jsonResponse = json.loads(request.data)
  filters = jsonResponse.get('filters', False)
  data = ranking.ranking_client(filters)
  wb, ws = utils.create_workbook("Vtas Cltes Rankin Acum Detallada Filtros")
  ws,row = utils.header(ws, "Vtas Cltes Rankin Acum Detallada Filtros", 'CEDIS', 'LAGOS DE MORENO',filters)
  # print(jsonResponse.get('table', False))
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

    col_money = [10,11,12,13,14,15,16,17] # columnas que indican dinero
    col_porc = [9,20] # columnas que indican porcentaje
    ws = utils.paint_par(ws, table_header, data, 7,row, col_money, col_porc)
    # Method for rezise cells
    # This method recives workbook active instance and
    ws = utils.resize_cells(ws, 20)
    # Method for load filters
    # This method recives workbook active instance and
    # init vector on init the header_table
    # ws = utils.load_filters(ws, 'A11')
    # Method for adds color to titles, first color si for the font
    # second color is for fill cell. Colors is in format RGB
    ws = utils.adds_title_format(ws, len(table_header), "FFFFFF", "4F81BD",row)

    total_start = row + 2
    total_end = (len(data) + total_start) -1
    total_total = total_end + 1
    listTotal = ['Total','','','','','','','= SUM(H'+str(total_start)+':H'+str(total_end)+')','= SUM(I'+str(total_start)+':I'+str(total_end)+')','= SUM(J'+str(total_start)+':J'+str(total_end)+')','= SUM(K'+str(total_start)+':K'+str(total_end)+')','= SUM(L'+str(total_start)+':L'+str(total_end)+')','= SUM(M'+str(total_start)+':M'+str(total_end)+')','= SUM(N'+str(total_start)+':N'+str(total_end)+')','= SUM(O'+str(total_start)+':O'+str(total_end)+')','= SUM(P'+str(total_start)+':P'+str(total_end)+')','= SUM(Q'+str(total_start)+':Q'+str(total_end)+')','= SUM(R'+str(total_start)+':R'+str(total_end)+')','= SUM(S'+str(total_start)+':S'+str(total_end)+')','= ((S'+str(total_total)+')/R'+str(total_total)+')']
    formatPercent = ['I','T']
    formatMoney = ['J','K','L','M','N','O','P','Q']
    formatNumber = ['H','R','S']
    ws = utils.total_summary(ws, listTotal, total_total, len(table_header), "FFFFFF", "4F81BD", formatPercent,formatMoney,formatNumber)

    nombre_archivo = "16"+datetime.now().date().strftime('%Y%m%d')+"Vtas Cltes Rankin Acum Detallada Filtros.xlsx"
    wb.save(nombre_archivo)
    return send_file('../'+nombre_archivo, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', True, nombre_archivo)
  else:
    response = { 'response': 'No hay registros'}
    return jsonify(response)

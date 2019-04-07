# -*- coding: utf-8 -*-

import os
import sys
import psycopg2
from datetime import datetime, time, date

from flask import Flask, send_file, request, jsonify, json
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from app import app
from app.models import model_product_ranking_week

from app.utils import utils

@app.route('/api/v2/ranking_product_week', methods=['POST'])
def ranking_product_week(charset='utf-8'):
  jsonResponse = json.loads(request.data)
  filters = jsonResponse.get('filters', False)
  data, arrDates = model_product_ranking_week.ranking_week(filters)
  table_header = jsonResponse['table']['headers']
  week = -1

  for day in arrDates:
    if(day["week"]!=week):
      if week!=-1:
        table_header.append("Total "+str(week))
      week = day["week"]
    table_header.append(day["day"])
  table_header.append("Total "+str(week))

  wb, ws = utils.create_workbook("Vtas Marca Producto")
  ws,row = utils.header(ws, "Vtas Marca Producto Acum mes sem dia Filtros", 'CEDIS', 'LAGOS DE MORENO',filters)
  # ws = utils.week_header(ws, start, row, week_start, week_end)
  if len(data)>0:
    ws.append(list(table_header))
    nombre_archivo = datetime.now().date().strftime('%Y %m %d')+" 18Vtas Cltes rankin x Semana Detallada Filtros.xlsx"
    wb.save(nombre_archivo)
    return send_file('../'+nombre_archivo, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', True, nombre_archivo)
  else:
    response = { 'response': 'No hay registros'}
    return jsonify(response)


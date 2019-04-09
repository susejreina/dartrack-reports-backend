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
  start = len(table_header)

  week = -1 if len(arrDates)<=0 else arrDates[0]["week"]
  month = -1 if len(arrDates)<=0 else str(arrDates[0]["monthString"])+" "+str(arrDates[0]["year"])
  arrWeeks = []
  arrMonths = []
  colsWeek = 0
  colsMonth = 0
  for day in arrDates:
    colsWeek += 1
    colsMonth += 1
    new_month = str(day["monthString"])+" "+str(day["year"])
    if(day["week"]!=week):
      colsMonth +=1
      table_header.append("Total "+str(week))
      arrWeeks.append({'text':"Semana "+str(week),'cols':colsWeek})
      week = day["week"]
      colsWeek = 1
    if(new_month!=month):
      colsMonth -= 1
      arrMonths.append({'text':month,'cols':colsMonth})
      month=new_month
      colsMonth = 1
    table_header.append(day["day"])

  table_header.append("Total "+str(week))
  colsWeek += 1
  arrWeeks.append({'text':"Semana "+str(week),'cols':colsWeek})
  colsMonth += 1
  arrMonths.append({'text':new_month,'cols':colsMonth})
  wb, ws = utils.create_workbook("Vtas Marca Producto")
  ws,row = utils.header(ws, "Vtas Marca Producto Acum mes sem dia Filtros", 'CEDIS', 'LAGOS DE MORENO',filters)
  
  if len(data)>0:
    ws.append(list(table_header))
    ws.insert_rows(row,3)
    row=6
    ws = utils.month_header_array(ws, start, row, arrMonths)
    row=7
    ws = utils.week_header_array(ws, start, row, arrWeeks)
    total = 0
    
    ws = utils.load_rows(ws, data)

    nombre_archivo = datetime.now().date().strftime('%Y %m %d')+" 19Vtas Marca Producto Acum mes sem dia Filtros.xlsx"
    wb.save(nombre_archivo)
    return send_file('../'+nombre_archivo, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', True, nombre_archivo)
  else:
    response = { 'response': 'No hay registros'}
    return jsonify(response)


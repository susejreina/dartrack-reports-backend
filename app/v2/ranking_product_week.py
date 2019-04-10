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
  dataBefore, arrDates = model_product_ranking_week.ranking_week(filters)
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
  table_header.append("Total Acumulado")
  colsWeek += 1
  arrWeeks.append({'text':"Semana "+str(week),'cols':colsWeek})
  colsMonth += 1
  arrMonths.append({'text':new_month,'cols':colsMonth})
  wb, ws = utils.create_workbook("Vtas Marca Producto")
  ws,row = utils.header(ws, "Vtas Marca Producto Acum mes sem dia Filtros", 'CEDIS', 'LAGOS DE MORENO',filters)
  print(row)
  
  if len(dataBefore)>0:
    ws.append(list(table_header))
    longHeader = len(table_header)
    ws = utils.adds_title_format_new(ws, longHeader, "FFFFFF", "4F81BD",row)

    ws.insert_rows(row,3)
    row=6
    ws = utils.month_header_array(ws, start, row, arrMonths)
    row=7
    ws = utils.week_header_array(ws, start, row, arrWeeks)
    total = 0

    arrTotales = [start] if len(arrWeeks) <=0 else [start+int(arrWeeks[0]['cols'])]
    i=0
    for w in range(1,len(arrWeeks)):
      arrTotales.append(arrTotales[i]+int(arrWeeks[w]['cols']))
      i+=1
    lenTotales = len(arrTotales)
    data = []
    for d in dataBefore:
      n = []
      totales = 0
      total_week = 0
      init  = start
      hasta = arrTotales[totales] - (2 + totales)
      for h in range(0,len(d)):
        if h<start:
          n.append(str(d[h]))
        else:
          n.append(float(d[h]))
        if init <= h <= hasta:
          total_week += float(d[h])
          if h == hasta:
            n.append(total_week)
            total_week = 0
            totales+=1
            init = hasta + 1
            if totales<lenTotales:
              hasta = arrTotales[totales] - (2 + totales)
        data.append(n)
    longData = len(data)
    ws = utils.load_rows(ws, data)
    ws = utils.freeze_row(ws,'E',row)

    col_nro_dec = []
    for col in range(start+1, len(data[0])+1):
      col_nro_dec.append(col)

    ws = utils.paint_par(ws, longHeader, longData, start,row+2, [], [],col_nro_dec,[])
    ws = utils.paint_total(ws, longHeader, longData, arrTotales, row+2)

    ws = utils.resize_cells(ws, 20)

    listTotal = ['Total','','','']
    formatPercent = []
    formatMoney = []
    formatNumberInteger = []
    formatNumberDecimal = []

    total_start = row + 2
    total_end = (longData+ total_start) -1
    total_total = total_end + 2

    for col in range(start+1, longHeader+1):
      letter_col = get_column_letter(col)
      listTotal.append('= SUM('+letter_col+str(total_start)+':'+letter_col+str(total_end)+')')
      formatNumberDecimal.append(letter_col)

    ws = utils.total_summary(ws, listTotal, total_total, longHeader, "FFFFFF", "4F81BD", formatPercent,formatMoney,formatNumberInteger,formatNumberDecimal)

    nombre_archivo = datetime.now().date().strftime('%Y %m %d')+" 19Vtas Marca Producto Acum mes sem dia Filtros.xlsx"
    wb.save(nombre_archivo)
    return send_file('../'+nombre_archivo, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', True, nombre_archivo)
  else:
    response = { 'response': 'No hay registros'}
    return jsonify(response)


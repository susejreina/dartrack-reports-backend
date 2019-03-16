# -*- coding: utf-8 -*-

from app.models import models
from config import db
from flask import Flask, send_file, request
from datetime import datetime
from app.utils import utils

def build_aditional_filters(filters):
  init_date           = filters.get('init_date', False)
  last_date           = filters.get('last_date', False)
  delivered_route_id  = filters.get('delivered_route_id', False)
  presale_route_id    = filters.get('presale_route_id', False)
  strategic_channel   = filters.get('strategic_channel', False)
  gec                 = filters.get('gec', False)
  description_channel = filters.get('description_channel', False)

  filters = ""

  if (delivered_route_id):
    filters = filters + """
      AND R.id ="""+str(delivered_route_id)+"""
    """
  if (presale_route_id):
    filters = filters + """
      AND R.id ="""+str(presale_route_id)+"""
    """
  if (strategic_channel):
    filters = filters + """
      AND CH.id ="""+str(strategic_channel)+"""
    """
  if (init_date and last_date):
    date_start = utils.format_date(init_date, '01:00')
    date_end = utils.format_date(last_date, '23:59')
    filters = filters + """
      AND ordered_at BETWEEN ('"""+date_start+"""') AND ('"""+date_end+"""')
      """
  return filters

def detail_client_sell(filters):
  dynamic_filters = build_aditional_filters(filters)
  cur = db.conn.cursor()
  query_select = """SELECT 
                CT.address_id,
                O.id as order_id,
                C.company_code as id_contpaq,
                O.client_id as id, 
                C.business_name as negocio,
                AD.location as town,
                BR.name as branding,
                P.name as producto,
                P.pres_ccm as presccm,
                OD.hlts as hlts,
                (OD.quantity_delivered * OD.list_price) as venta_total,
                OD.discount_promo as discount_promo,
                OD.discount_product as discount_product,
                OD.discount_bonification as discount_bonification,
                OD.discount_payment as discount_payment,
                OD.total as total_neto,
                OD.discount_fba as discount_fba,
                OD.quantity as total_boxes,
                OD.quantity_delivered as fullfilment,
                null as delivered"""
  from_select = """
            FROM orders O 
            LEFT JOIN clients C
            ON O.client_id = C.id
            LEFT JOIN order_details OD
            ON OD.order_id = O.id
            LEFT JOIN products P
            ON P.id = OD.product_id
            LEFT JOIN centers CT
            ON CT.id = C.center_id
            LEFT JOIN addresses AD
            ON AD.id = CT.address_id
            LEFT JOIN brands BR
            ON BR.id = P.brand_id
            LEFT JOIN routes R
            ON R.id = O.route_id
            LEFT JOIN channels CH
            ON CH.company_id = C.company_id
            LEFT JOIN route_types RT
            ON RT.id = R.route_type_id"""
  filter_query = """
            WHERE O.id IN
            (
            SELECT MAX(id) as id
            FROM orders
            GROUP BY client_id
            ORDER BY client_id
            )""" + dynamic_filters
  others_commands = """
            ORDER BY O.client_id;"""
  query = query_select+from_select+filter_query+others_commands
  cur.execute(query)
  data = cur.fetchall()
  return data

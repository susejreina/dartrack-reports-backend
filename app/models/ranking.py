# -*- coding: utf-8 -*-


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
  ''''
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
      AND o.ordered_at BETWEEN ('"""+date_start+"""') AND ('"""+date_end+"""')
      """
      '''
  return filters

def ranking_client(filters):
  dynamic_filters = build_aditional_filters(filters)
  cur = db.conn.cursor()
  query_select = """SELECT clientes.id_g_suc, clientes.id_clte, clientes.negocio, clientes.poblacion,
                            clientes.desc_canal, clientes.desc_canal_est, ordenes.rank, ordenes.htls,
                            ordenes.htls_percentage, ordenes.total, ordenes.desc_promo, ordenes.desc_produc, ordenes.bonif, ordenes.venta_neta, ordenes.bonif_fba, ordenes.venta_final, ordenes.boxes_requested, ordenes.boxes_delivered, ordenes.ent_ped"""
  from_select = """
            FROM (
	SELECT  C.company_code as id_g_suc,C.id as id_clte, C.business_name as negocio, AD.location as poblacion, Cl.name as desc_canal,
			CH.name as desc_canal_est
	FROM clients C
	LEFT JOIN addresses AD ON AD.id = C.address_id
	LEFT JOIN channels CH ON C.channel_id = CH.id
	LEFT JOIN client_types CL ON C.client_type_id=Cl.id
	ORDER BY C.id
) as clientes
LEFT JOIN
(
	SELECT	o.client_id as client, rank() over (order by SUM((p.hlts * od.quantity_delivered))::DOUBLE PRECISION DESC) as rank,
			SUM((p.hlts * od.quantity_delivered))::DOUBLE PRECISION "htls",
		(SUM((p.hlts * od.quantity_delivered))::DOUBLE PRECISION * 100 /
		(SELECT	SUM((p.hlts * od.quantity_delivered))::DOUBLE PRECISION
		 FROM orders o
		 LEFT JOIN order_details od ON o.id = od.order_id
		 LEFT JOIN products  p ON p.id = od.product_id
		 WHERE od.is_devolution = false AND o.active = true
            AND o.ordered_at BETWEEN ('2018-10-08 00:00:00') AND ('2019-03-14 23:59:59')
		)) as htls_percentage,
		(SUM(od.total) + SUM(od.discount_promo) + SUM(od.discount_bonification))::DOUBLE PRECISION "total",
		SUM(od.discount_promo)::DOUBLE PRECISION "desc_promo",
		SUM(od.discount_product)::DOUBLE PRECISION "desc_produc",
		SUM(od.discount_bonification)::DOUBLE PRECISION "bonif",
		SUM(od.total)::DOUBLE PRECISION "venta_neta",
		SUM(od.discount_fba)::DOUBLE PRECISION "bonif_fba",
		(SUM(od.total) - SUM(od.discount_fba))::DOUBLE PRECISION "venta_final",
		SUM(od.quantity)::INTEGER "boxes_requested",
		SUM(od.quantity_delivered)::INTEGER "boxes_delivered",
		CASE
			WHEN SUM(od.quantity_delivered)::INTEGER != 0
			THEN ((SUM(od.quantity_delivered)*100)/SUM(od.quantity))
			ELSE 0
		END as "ent_ped"
		FROM orders o
		LEFT JOIN order_details  od ON o.id = od.order_id
		LEFT JOIN products  p ON p.id = od.product_id
		WHERE od.is_devolution = false AND o.active = true AND o.ordered_at
		BETWEEN ('2018-10-08 00:00:00') AND ('2019-03-14 23:59:59')
		GROUP BY O.client_id
		ORDER BY O.client_id
	) as ordenes
	ON clientes.id_clte = ordenes.client"""
  filter_query = """
             """ + dynamic_filters
  others_commands = """
            ORDER BY clientes.id_clte;"""
  query = query_select+from_select+filter_query+others_commands
  cur.execute(query)
  data = cur.fetchall()
  return data

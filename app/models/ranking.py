# -*- coding: utf-8 -*-


from config import db
from flask import Flask, send_file, request
from datetime import datetime
from app.utils import utils

def build_clients_filters(filters):
	canal_giro = filters.get('canal_giro', False)
	canal_est = filters.get('canal_est', False)
	filters = ""
	
	if (canal_giro):
		if canal_giro["canal_giro_id"] != 0:
			filters = filters + """
				WHERE Cl.id ="""+str(canal_giro["canal_giro_id"])+"""
			"""
	if (canal_est):
		if canal_est["canal_est_id"] != 0:
			if filters == "":
				filters += " WHERE "
			else:
				filters += " AND "
			filters += " C.channel_id ="+str(canal_est["canal_est_id"])

	return filters

def build_aditional_filters(filters):
	init_date  = filters.get('init_date', False)
	last_date  = filters.get('last_date', False)
	presale_route  = filters.get('presale_route', False)
	delivery_route  = filters.get('delivery_route', False)
	product = filters.get('product', False)
	group_product = filters.get('group_product',False)

	filters = ""
	if (presale_route):
		if presale_route["presale_route_id"] != 0:
			filters = filters + """
				AND o.route_id ="""+str(presale_route["presale_route_id"])+"""
			"""
	if (delivery_route):
		if delivery_route["delivery_route_id"] != 0:
			filters = filters + """
				AND o.route_delivery_id ="""+str(delivery_route["delivery_route_id"])+"""
			"""
	if (product):
		if product["product_id"] != 0:
			filters = filters + """
				AND p.id ="""+str(product["product_id"])+"""
			"""
	if (group_product):
		if group_product["group_product_id"] != 0:
			filters = filters + """
				AND p.product_group_id ="""+str(group_product["group_product_id"])+"""
			"""
	if (init_date and last_date):
		date_start = utils.format_date(init_date, '01:00')
		date_end = utils.format_date(last_date, '23:59')
		filters = filters + """
			AND o.ordered_at BETWEEN ('"""+date_start+"""') AND ('"""+date_end+"""')
			"""
	return filters

def ranking_client(filters):
	dynamic_filters = build_aditional_filters(filters)
	client_filters = build_clients_filters(filters)
	cur = db.conn.cursor()
	query_select = """SELECT clientes.id_g_suc, clientes.id_clte, 
								CASE WHEN ordenes.rank IS NULL THEN 0 ELSE ordenes.rank END, 
								clientes.negocio, clientes.poblacion, clientes.canal_giro, clientes.canal_est, 
								CASE WHEN ordenes.htls IS NULL THEN 0 ELSE ordenes.htls END,
								CASE WHEN ordenes.htls_percentage IS NULL THEN 0 ELSE ordenes.htls_percentage END,
								CASE WHEN ordenes.total IS NULL THEN 0 ELSE ordenes.total END,
								CASE WHEN ordenes.desc_promo IS NULL THEN 0 ELSE ordenes.desc_promo END,
								CASE WHEN ordenes.desc_produc IS NULL THEN 0 ELSE ordenes.desc_produc END,
								CASE WHEN ordenes.bonif IS NULL THEN 0 ELSE ordenes.bonif END,
								CASE WHEN ordenes.discount_payment IS NULL THEN 0 ELSE ordenes.discount_payment END,
								CASE WHEN ordenes.venta_neta IS NULL THEN 0 ELSE ordenes.venta_neta END,
								CASE WHEN ordenes.bonif_fba IS NULL THEN 0 ELSE ordenes.bonif_fba END,
								CASE WHEN ordenes.venta_final IS NULL THEN 0 ELSE ordenes.venta_final END,
								CASE WHEN ordenes.boxes_requested IS NULL THEN 0 ELSE ordenes.boxes_requested END,
								CASE WHEN ordenes.boxes_delivered IS NULL THEN 0 ELSE ordenes.boxes_delivered END,
								CASE WHEN ordenes.ent_ped IS NULL THEN 0 ELSE ordenes.ent_ped END"""
	from_select = """
						FROM (
	SELECT  C.company_code as id_g_suc,C.id as id_clte, C.business_name as negocio, AD.location as poblacion, Cl.name as canal_giro,
			CH.name as canal_est
	FROM clients C
	LEFT JOIN addresses AD ON AD.id = C.address_id
	LEFT JOIN channels CH ON C.channel_id = CH.id
	LEFT JOIN client_types CL ON C.client_type_id=Cl.id """+client_filters+"""
	ORDER BY C.id
) as clientes
LEFT JOIN
(
	SELECT	o.client_id as client, rank() over (order by SUM((p.hlts * od.quantity_delivered))::DOUBLE PRECISION DESC) as rank,
			SUM((p.hlts * od.quantity_delivered))::DOUBLE PRECISION "htls",
		(SUM((p.hlts * od.quantity_delivered))::DOUBLE PRECISION /
		(SELECT	SUM((p.hlts * od.quantity_delivered))::DOUBLE PRECISION
		 FROM orders o
		 LEFT JOIN order_details od ON o.id = od.order_id
		 LEFT JOIN products  p ON p.id = od.product_id
		 WHERE od.is_devolution = false AND o.active = true """+dynamic_filters+"""
		)) as htls_percentage,
		(SUM(od.total) + SUM(od.discount_promo) + SUM(od.discount_bonification))::DOUBLE PRECISION "total",
		SUM(od.discount_promo)::DOUBLE PRECISION "desc_promo",
		SUM(od.discount_product)::DOUBLE PRECISION "desc_produc",
		SUM(od.discount_bonification)::DOUBLE PRECISION "bonif",
		SUM(od.discount_payment)::DOUBLE PRECISION "discount_payment",
		SUM(od.total)::DOUBLE PRECISION "venta_neta",
		SUM(od.discount_fba)::DOUBLE PRECISION "bonif_fba",
		(SUM(od.total) - SUM(od.discount_fba) - SUM(od.discount_payment))::DOUBLE PRECISION "venta_final",
		SUM(od.quantity)::INTEGER "boxes_requested",
		SUM(od.quantity_delivered)::INTEGER "boxes_delivered",
		CASE
			WHEN SUM(od.quantity_delivered)::INTEGER != 0
			THEN ((SUM(od.quantity_delivered) ::DOUBLE PRECISION)/SUM(od.quantity ::DOUBLE PRECISION))
			ELSE 0
		END as "ent_ped"
		FROM orders o
		LEFT JOIN order_details  od ON o.id = od.order_id
		LEFT JOIN products  p ON p.id = od.product_id
		WHERE od.is_devolution = false AND o.active = true """+dynamic_filters+"""
		GROUP BY O.client_id
		ORDER BY O.client_id
	) as ordenes
	ON clientes.id_clte = ordenes.client"""
	others_commands = """
						ORDER BY clientes.id_clte;"""
	query = query_select+from_select+others_commands
	cur.execute(query)
	data = cur.fetchall()
	return data

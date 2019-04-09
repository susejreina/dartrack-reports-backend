# -*- coding: utf-8 -*-


from config import db
from flask import Flask, send_file, request
from datetime import date, timedelta, datetime
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

def build_ranking_filters(filters):
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

def build_orders_filters(filters,week):
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
	if week>0:
		filters = filters + """
			AND EXTRACT(WEEK FROM o.ordered_at) = """+str(week)+"""
			"""
	return filters

def ranking_week(filters):
	init_date  = filters.get('init_date', False)
	last_date  = filters.get('last_date', False)


	init_date  = datetime.strptime(init_date,'%d/%m/%Y')
	last_date  = datetime.strptime(last_date,'%d/%m/%Y')
	arrDates = utils.daysBetweenDates(init_date, last_date)

	'''
	client_filters = build_clients_filters(filters)
	ranking_filters = build_ranking_filters(filters)
	'''
	months = [{'month':arrDates[0]["monthString"],
						'dateStart':str(arrDates[0]["year"])+"-"+str(arrDates[0]["month"])+"-"+str(arrDates[0]["day"]),
						'dateEnd': '',
						'dayStart':int(arrDates[0]["day"]),
						'dayEnd':31,
						'dias': '',
						'key': str(arrDates[0]["month"])+"_"+str(arrDates[0]["year"])
						}]
	
	cur = db.conn.cursor()
	query_select = """SELECT productos.marca, productos.sku, productos.name,productos.pres_ccm, """
	query_weeks = ""
	m = 0
	dias = ""
	for i in range(0,len(arrDates)):
		d = arrDates[i] 
		if months[m]["month"]!=d["monthString"]:
			months[m]['dateEnd']=str(arrDates[i-1]["year"])+"-"+str(arrDates[i-1]["month"])+"-"+str(arrDates[i-1]["day"])
			months[m]['dayEnd']=int(arrDates[i-1]["day"])
			months[m]['dias'] = dias
			months.append({'month':d["monthString"],
										'dateStart':str(d["year"])+"-"+str(d["month"])+"-"+str(d["day"]),
										'dateEnd': '',
										'dayStart':int(d["day"]),
										'dayEnd':31,
										'dias': '',
										'key': str(d["month"])+"_"+str(d["year"])
									})
			dias = ""
			m += 1
		dias += "Dia"+str(d["day"])+" NUMERIC(10,6),"
		query_weeks += "CASE WHEN "+str(d["monthString"])+".Dia"+str(d["day"])+" IS NULL THEN 0 ELSE "+str(d["monthString"])+".Dia"+str(d["day"])+" END, "

	months[m]['dateEnd']=str(arrDates[i]["year"])+"-"+str(arrDates[i]["month"])+"-"+str(arrDates[i]["day"])
	months[m]['dayEnd']=int(arrDates[i]["day"])
	months[m]['dias'] = dias

	query_weeks += "productos.total_htls "
	query = query_select+query_weeks

	query_from_productos = """ FROM (
		SELECT P.id, B.name as marca, P.sku, P.name,P.pres_ccm, SUM((P.hlts * OD.quantity_delivered))::DOUBLE PRECISION total_htls
		FROM orders O
		LEFT JOIN order_details  OD ON O.id = OD.order_id
		LEFT JOIN products P ON P.id = OD.product_id
		LEFT JOIN brands B ON B.id=P.brand_id
		WHERE od.is_devolution = false AND o.active = true  AND o.ordered_at BETWEEN '2018-12-01 00:00:00' AND '2019-01-31 23:59:59'
		GROUP BY B.name, P.id,P.sku, P.name,P.pres_ccm
		ORDER BY B.name) as productos"""
	
	query_from_days = ""

	for d in months:
		query_from_days += """ LEFT JOIN 
			(SELECT *
			FROM crosstab(
			'SELECT P.id as key_"""+d["key"]+""", 
			EXTRACT(DAY FROM  o.ordered_at) as day,  SUM((P.hlts * OD.quantity_delivered))::DOUBLE PRECISION total_htls
			FROM orders O
			LEFT JOIN order_details  OD ON O.id = OD.order_id
			LEFT JOIN products P ON P.id = OD.product_id
			WHERE od.is_devolution = false AND o.active = true AND 
			o.ordered_at BETWEEN ''"""+d["dateStart"]+""" 00:00:00'' AND ''"""+d["dateEnd"]+""" 23:59:59''
			GROUP BY key_"""+d["key"]+""",day
			ORDER BY 1,2',
			'SELECT day FROM generate_series("""+str(d["dayStart"])+""","""+str(d["dayEnd"])+""") AS day'
			) AS (
			key_"""+d["key"]+""" integer,
			"""+d["dias"][:-1]+"""
			)) as """+d["month"]+"""
			ON productos.id = """+d["month"]+""".key_"""+d["key"]+"""
		"""
	
	query = query_select+query_weeks+query_from_productos+query_from_days
	cur.execute(query)
	data = cur.fetchall()

	return [data,arrDates]

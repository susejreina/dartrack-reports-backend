# -*- coding: utf-8 -*

import psycopg2
from config import config

conn = psycopg2.connect(database=config.DATABASE,user=config.USER,password=config.PASS, host=config.HOST)
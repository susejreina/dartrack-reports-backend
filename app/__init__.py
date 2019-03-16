# -*- coding: utf-8 -*-
import os
from flask import Flask

app = Flask(__name__)

from app import routes
from app.v2 import clients

port = int(os.environ.get("PORT", 3000))
app.run(host='0.0.0.0', port=port)
# Define the application directory
import os
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

DATABASE = os.environ.get('DATABASE', 'DATABASE_DEV')
USER = os.environ.get('USER', 'USER_DEV')
PASS = os.environ.get("PASS", "PASS_DEV")
HOST = os.environ.get("HOST", "HOST_DEV")
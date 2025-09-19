import os
import mysql.connector

def get_db_config():
    return {
        "host": os.getenv("DB_HOST", "localhost"),
        "port": int(os.getenv("DB_PORT", "10645")),
        "user": os.getenv("DB_USER", "deepinspector"),
        "password": os.getenv("DB_PASSWORD", "xoaud17!@"),
        "database": os.getenv("DB_NAME", "db_deepinspector"),
        "charset": "utf8mb4",
        "use_pure":True,
    }

def connect_db():
    return mysql.connector.connect(**get_db_config())

# 사용처:
# conn = connect_db()

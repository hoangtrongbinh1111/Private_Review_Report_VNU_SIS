from flask import Flask
from flask_sqlalchemy import SQLAlchemy
import os
import sqlite3

basedir = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
app.config['SECRET_KEY'] = 'super secret key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///mydatabase.db'
# app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'mydatabase.db')
db = SQLAlchemy(app)

app.config.from_object(__name__)
from app import views
# Path to your schema.sql file
SCHEMA = os.path.join(basedir, 'schema.sql')

def create_database():
    conn = sqlite3.connect(os.path.join(basedir, 'mydatabase.db'))
    cursor = conn.cursor()
    with open(SCHEMA, 'r') as f:
        schema_sql = f.read()
    cursor.executescript(schema_sql)
    conn.commit()
    conn.close()

@app.before_first_request
def initialize_database():
    create_database()
    print("Database created successfully from schema.sql.")

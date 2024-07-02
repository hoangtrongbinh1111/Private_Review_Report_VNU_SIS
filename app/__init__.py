from flask import Flask
from flask_sqlalchemy import SQLAlchemy
import os

basedir = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
app.config['SECRET_KEY'] = 'super secret key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'mydatabase.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# Path to your schema.sql file
SCHEMA = os.path.join(basedir, 'schema.sql')

def create_database():
    if not os.path.exists(os.path.join(basedir, 'mydatabase.db')):
        with app.app_context():
            conn = db.engine.raw_connection()
            cursor = conn.cursor()
            with open(SCHEMA, 'r') as f:
                schema_sql = f.read()
            cursor.executescript(schema_sql)
            conn.commit()
            conn.close()
            print("Database created successfully from schema.sql.")

@app.before_first_request
def initialize_database():
    create_database()

from app import views

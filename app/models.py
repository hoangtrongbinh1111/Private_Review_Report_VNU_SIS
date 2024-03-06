from app import db

class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key = True)
    name = db.Column(db.String(255))
    uuid = db.Column(db.String(255), unique=True)
    rowIndex = db.Column(db.Integer, unique=True)

    def __init__(self, name, uuid, rowIndex):
        self.name = name
        self.uuid = uuid
        self.rowIndex = rowIndex

    def __repr__(self):
        return '<User %r>' % self.name

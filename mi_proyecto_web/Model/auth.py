from flask_sqlalchemy import SQLAlchemy 
from werkzeug.security import generate_password_hash, check_password_hash
import re

db = SQLAlchemy()

class User(db.Model):
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=db.func.current_timestamp())
    last_login = db.Column(db.DateTime, nullable=True)

    def set_password(self, password):
        """Hash password before storing"""
        if not self.validate_password(password):
            raise ValueError("Invalid password")
        self.password_hash = generate_password_hash(password, method='pbkdf2:sha256')

    def check_password(self, password):
        """Verify password against hash"""
        return check_password_hash(self.password_hash, password)

    @staticmethod
    def validate_password(password):
        """Validate password complexity"""
        return (len(password) >= 8 and 
                re.search(r'[A-Z]', password) and 
                re.search(r'[a-z]', password) and 
                re.search(r'\d', password))

    @staticmethod
    def validate_username(username):
        """Validate username format"""
        return re.match(r'^[a-zA-Z0-9_]+$', username) is not None
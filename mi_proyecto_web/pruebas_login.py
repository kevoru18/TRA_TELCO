from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import re

# Configuración para la aplicación principal
app = Flask(__name__)

# Configuración para conectar con SQL Server
app.config['SQLALCHEMY_DATABASE_URI'] = 'mssql+pyodbc://sa:infinity@192.168.201.12/_RRHH?driver=ODBC+Driver+17+for+SQL+Server'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Definición del modelo User
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

# Blueprints
from flask import Blueprint, request, jsonify, render_template, redirect
from flask_login import login_user, login_required, logout_user
from datetime import datetime

auth = Blueprint('auth', __name__)

@auth.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')

        # Server-side validation
        if not User.validate_username(username):
            return jsonify({'success': False, 'message': 'Invalid username'})

        user = User.query.filter_by(username=username, is_active=True).first()
        
        if user and user.check_password(password):
            user.last_login = datetime.utcnow()
            db.session.commit()
            login_user(user)
            return jsonify({'success': True, 'redirect': '/dashboard'})
        
        return jsonify({'success': False, 'message': 'Invalid credentials'})
    
    return render_template('login.html')

@auth.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect('/login')

# Registro del Blueprint
app.register_blueprint(auth)

# Ejemplo de consulta personalizada
@app.route('/active_staff')
def active_staff():
    result = db.session.execute(
        """
        SELECT [ID], [Apellido1], [Apellido2], [Nombre], [categoria_telco], [estado_telco], [usuario_windows]
        FROM [_RRHH].[dbo].[Contactos]
        WHERE categoria_telco='STAFF' AND estado_telco='ACTIVO'
        """
    )
    contacts = [dict(row) for row in result]
    return {"contacts": contacts}

if __name__ == '__main__':
    with app.app_context():
        db.create_all()  # Crea las tablas en la base de datos si no existen
    app.run(debug=True)

# Importamos las dependencias necesarias
from flask import Flask  # Framework web principal para manejar solicitudes y rutas.
from flask_sqlalchemy import SQLAlchemy  # Extensión para interactuar con bases de datos usando ORM.
from werkzeug.security import generate_password_hash, check_password_hash  # Herramientas para el manejo seguro de contraseñas.
import re  # Módulo para trabajar con expresiones regulares.

# Configuración para la aplicación principal
app = Flask(__name__)  # Creamos una instancia de la aplicación Flask.

# Configuración para conectar con SQL Server
app.config['SQLALCHEMY_DATABASE_URI'] = (
    'mssql+pyodbc://sa:infinity@192.168.201.12/_RRHH?driver=ODBC+Driver+17+for+SQL+Server'
)  # Cadena de conexión para la base de datos. Recomendación: usar variables de entorno para evitar exponer credenciales.
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False  # Desactiva el seguimiento de modificaciones para mejorar el rendimiento.

# Inicializamos SQLAlchemy con la configuración de Flask
db = SQLAlchemy(app)  # Vincula la instancia de la base de datos a la aplicación Flask.

# Definimos el modelo de datos para los usuarios
class User(db.Model):  # La clase `User` hereda de `db.Model`, lo que la convierte en un modelo de SQLAlchemy.
    __tablename__ = 'Contactos'  # Nombre de la tabla en la base de datos.

  # Mapeo de columnas
    id = db.Column(db.Integer, primary_key=True)  # Corresponde al campo `ID`.
    apellido1 = db.Column(db.String(100), nullable=False)  # Corresponde al campo `Apellido1`.
    apellido2 = db.Column(db.String(100), nullable=True)  # Corresponde al campo `Apellido2`.
    nombre = db.Column(db.String(100), nullable=False)  # Corresponde al campo `Nombre`.
    estado_telco = db.Column(db.String(50), nullable=True)  # Corresponde al campo `estado_telco`.
    categoria_telco = db.Column(db.String(50), nullable=True)  # Corresponde al campo `categoria_telco`.
    usuario_windows = db.Column(db.String(100), unique=True, nullable=False)  # Corresponde al campo `usuario_windows`.
    actualizado = db.Column(db.DateTime, nullable=True)  # Corresponde al campo `ACTUALIZADO`.

    # Método para establecer la contraseña del usuario
    def set_password(self, password):
        """
        Genera un hash seguro para la contraseña proporcionada y lo almacena en `password_hash`.
        Valida primero la complejidad de la contraseña antes de encriptarla.
        """
        self.password_hash = generate_password_hash(password, method='pbkdf2:sha256')  # Hash seguro para la contraseña.
        # if not self.validate_password(password):  # Verificamos si la contraseña cumple los criterios de complejidad.
        #     raise ValueError("Invalid password")  # Si no cumple, se lanza un error.
        # self.password_hash = generate_password_hash(password, method='pbkdf2:sha256')  # Hash seguro para la contraseña.

    # Método para verificar una contraseña
    def check_password(self, password):
        """
        Compara la contraseña proporcionada con la almacenada encriptada.
        Devuelve `True` si coinciden; de lo contrario, `False`.
        """
        return check_password_hash(self.password_hash, password)

    # Método estático para validar contraseñas
    @staticmethod
    def validate_password(password):
        """
        Verifica si una contraseña cumple con los criterios de seguridad:
        - Al menos 8 caracteres.
        - Al menos una letra mayúscula.
        - Al menos una letra minúscula.
        - Al menos un número.
        """
        return (
            len(password) >= 8 and  # Longitud mínima.
            re.search(r'[A-Z]', password) and  # Al menos una letra mayúscula.
            re.search(r'[a-z]', password) and  # Al menos una letra minúscula.
            re.search(r'\d', password)  # Al menos un dígito.
        )

    # Método estático para validar nombres de usuario
    @staticmethod
    def validate_username(username):
        """
        Verifica que el nombre de usuario solo contenga letras, números y guiones bajos.
        """
        return re.match(r'^[a-zA-Z0-9_]+$', username) is not None

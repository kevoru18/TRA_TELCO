import pyodbc
from werkzeug.security import generate_password_hash, check_password_hash

def validate_password(password):
   return (len(password) >= 8 and 
           any(c.isupper() for c in password) and
           any(c.islower() for c in password) and 
           any(c.isdigit() for c in password))

def set_password(password):
   if not validate_password(password):
       raise ValueError("Password must be at least 8 chars with upper, lower and numbers")
   return generate_password_hash(password, method='pbkdf2:sha256')

# Connect to database
conn = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.201.12;DATABASE=_krubi_tests;UID=sa;PWD=infinity')
# (
#      'mssql+pyodbc://sa:infinity@192.168.201.12/_RRHH?driver=ODBC+Driver+17+for+SQL+Server'
#  )
cursor = conn.cursor()

# Get users
cursor.execute('SELECT username FROM USUARIOS_APP')
users = [row[0] for row in cursor.fetchall()]

# Show users and get selection
print("Available users:")
for i, user in enumerate(users):
   print(f"{i+1}. {user}")

selection = int(input("\nSelect user number: ")) - 1
selected_user = users[selection]

# Get new password
new_password = input("Enter new password: ")

# Hash and update password
password_hash = set_password(new_password)
cursor.execute("UPDATE USUARIOS_APP SET password_hash = ? WHERE username = ?", 
             (password_hash, selected_user))
conn.commit()

print(f"Password updated for user {selected_user}")
import re

def limpiar_telefono(telefono):
    # Eliminar espacios, puntos, guiones y barras
    telefono = re.sub(r'[ .\-\/]', '', telefono)
    # Eliminar prefijos internacionales (+34, 0034)
    telefono = re.sub(r'^(\+34|0034)', '', telefono)
    # Verificar que el teléfono tenga 9 dígitos después de la limpieza
    if len(telefono) == 9 and telefono.isdigit():
        return telefono
    else:
        return None  # Teléfono inválido

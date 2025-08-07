import re
import pandas as pd
from tkinter import Tk, filedialog
from datetime import datetime
import os

# Diccionario de clientes
clientes = {
    "milagro": {
        "nombre": "MILAGRO DORADEA DE DIAZ",
        "nit": "04190304681010",
        "nrc": "266354-0"
    },
    "lucrecia": {
        "nombre": "LUCRECIA DELGADO DE ROMERO",
        "nit": "07081612701011",
        "nrc": "2101313"
    },
    # Agrega más clientes si deseas
}

# Selección de archivos .txt
Tk().withdraw()
rutas = filedialog.askopenfilenames(title="Selecciona los archivos .txt de PORTILLO", filetypes=[("Archivos de texto", "*.txt")])

if not rutas:
    print("No se seleccionaron archivos.")
    exit()

# Selección de cliente
alias_cliente = input("Alias del cliente (ej: milagro, lucrecia): ").strip().lower()
if alias_cliente not in clientes:
    print("Alias de cliente no válido.")
    exit()

datos_cliente = clientes[alias_cliente]

# Datos del proveedor PORTILLO
nit_proveedor = "04071310871027"
nrc_proveedor = "244994-1"
nombre_proveedor = "AGROFERRETERÍA EL POTRILLO"

# Lista para almacenar resultados
filas = []

for ruta in rutas:
    with open(ruta, 'r', encoding='utf-8', errors='ignore') as f:
        contenido = f.read()

    # Extraer campos clave
    codigo = re.search(r'Código de Generación:\s*(DTE[^\n]+)', contenido)
    control = re.search(r'Número de Control:\s*([A-Z0-9\-]+)', contenido)
    sello = re.search(r'(?:Sello|Número de Control):\s*([A-Z0-9]{20,})', contenido)
    fecha = re.search(r'Fecha y Hora de Generación:\s*(\d{2}/\d{2}/\d{4})', contenido)
    subtotal = re.search(r'Sub-Total:\s*\$?\s*([\d.]+)', contenido)
    iva = re.search(r'Impuesto al Valor Agregado 13%:\s*\$?\s*([\d.]+)', contenido)
    total = re.search(r'Monto Total de la Operación:\s*\$?\s*([\d.]+)', contenido)

    # Limpieza y asignación
    codigo = codigo.group(1).strip() if codigo else ""
    control = control.group(1).strip() if control else ""
    sello = sello.group(1).strip() if sello else ""
    fecha = fecha.group(1).strip() if fecha else ""
    subtotal = float(subtotal.group(1)) if subtotal else 0.00
    iva = float(iva.group(1)) if iva else round(subtotal * 0.13, 2)
    total = float(total.group(1)) if total else round(subtotal + iva, 2)

    # FOVIAL y COTRANS no aplican
    fovial = 0.00
    cotrans = 0.00
    cantidad = ""  # no se extrae individualmente

    filas.append({
        "Archivo": os.path.basename(ruta),
        "Código generación": codigo,
        "Número de control": control,
        "Sello de recepción": sello,
        "Fecha": fecha,
        "NIT proveedor": nit_proveedor,
        "NRC proveedor": nrc_proveedor,
        "Nombre proveedor": nombre_proveedor,
        "Cantidad": cantidad,
        "Subtotal": subtotal,
        "IVA": iva,
        "FOVIAL": fovial,
        "COTRANS": cotrans,
        "Total": total,
        "Nombre cliente": datos_cliente["nombre"],
        "NIT cliente": datos_cliente["nit"],
        "NRC cliente": datos_cliente["nrc"]
    })

# Crear Excel
df = pd.DataFrame(filas)

# Ruta de salida
salida = "C:/Users/PC/OneDrive/Escritorio/DOCUMENTOS EXCEL/facturas_potrillo_txt.xlsx"
df.to_excel(salida, index=False)

print(f"\n✅ ¡Listo Karla! Excel generado:\n{salida}")

import os
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xml.etree.ElementTree as ET
import pandas as pd
from PIL import Image, ImageTk


def limpiar_entero(valor, relleno=0):
    """Convierte de forma segura cualquier dato de Excel a un entero limpio sin fallar por NaNs."""
    if pd.isna(valor):
        return relleno
    try:
        return int(float(str(valor).strip()))
    except ValueError:
        return relleno


def limpiar_texto_sri(texto):
    """Elimina tildes, eñes y caracteres especiales no permitidos por el SRI."""
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    replacements = {
        "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
        "Ñ": "N", "Ü": "U"
    }
    for orig, rep in replacements.items():
        texto = texto.replace(orig, rep)
    # Solo permite letras, números y espacios
    texto = re.sub(r"[^A-Z0-9 ]", "", texto)
    return texto


def generar_xml_ats(archivo_excel, ruta_salida):
    df = pd.read_excel(archivo_excel, sheet_name="ATS")

    # Eliminar filas vacías
    df = df.dropna(subset=["idProv", "secuencial"], how="all")

    if df.empty:
        raise Exception(
            "El archivo Excel está vacío o no contiene registros válidos."
        )

    primera_fila = df.iloc[0]

    # 1. Crear nodo raíz 'iva'
    iva = ET.Element("iva")

    # 2. Cabecera obligatoria
    id_informante = str(primera_fila["IdInformante"]).split(".")[0].strip()
    if len(id_informante) < 13:
        id_informante = id_informante.zfill(13)

    ET.SubElement(iva, "IdInformante").text = id_informante
    ET.SubElement(iva, "razonSocial").text = limpiar_texto_sri(primera_fila["razonSocial"])
    ET.SubElement(iva, "Anio").text = str(limpiar_entero(primera_fila["Anio"]))
    ET.SubElement(iva, "Mes").text = str(limpiar_entero(primera_fila["Mes"])).zfill(2)
    ET.SubElement(iva, "numEstabRuc").text = str(limpiar_entero(primera_fila["numEstabRuc"])).zfill(3)

    try:
        total_ventas = float(primera_fila["totalVentas"])
    except:
        total_ventas = 0.00
    ET.SubElement(iva, "totalVentas").text = f"{total_ventas:.2f}"
    ET.SubElement(iva, "codigoOperativo").text = "IVA"

    # 3. Bloque de Compras
    df_compras = df[df["idProv"].notna()]
    if not df_compras.empty:
        compras_nodo = ET.SubElement(iva, "compras")

        for _, fila in df_compras.iterrows():
            detalle = ET.SubElement(compras_nodo, "detalleCompras")

            ET.SubElement(detalle, "codSustento").text = str(limpiar_entero(fila["codSustento"])).zfill(2)
            ET.SubElement(detalle, "tpIdProv").text = str(limpiar_entero(fila["tpIdProv"])).zfill(2)
            ET.SubElement(detalle, "idProv").text = str(fila["idProv"]).split(".")[0]
            ET.SubElement(detalle, "tipoComprobante").text = str(limpiar_entero(fila["tipoComprobante"])).zfill(2)
            ET.SubElement(detalle, "fechaRegistro").text = str(fila["fechaRegistro"]).strip()
            ET.SubElement(detalle, "establecimiento").text = str(limpiar_entero(fila["establecimiento"])).zfill(3)
            ET.SubElement(detalle, "puntoEmision").text = str(limpiar_entero(fila["puntoEmision"])).zfill(3)
            ET.SubElement(detalle, "secuencial").text = str(limpiar_entero(fila["secuencial"])).zfill(9)
            ET.SubElement(detalle, "fechaEmision").text = str(fila["fechaEmision"]).strip()
            ET.SubElement(detalle, "autorizacion").text = str(fila["autorizacion"]).split(".")[0]

            def safe_float(val):
                return float(val) if not pd.isna(val) else 0.00

            ET.SubElement(detalle, "baseNoGraIva").text = f"{safe_float(fila['baseNoGraIva']):.2f}"
            ET.SubElement(detalle, "baseImponible").text = f"{safe_float(fila['baseImponible']):.2f}"
            ET.SubElement(detalle, "baseImpGrav").text = f"{safe_float(fila['baseImpGrav']):.2f}"
            ET.SubElement(detalle, "baseImpExe").text = f"{safe_float(fila['baseImpExe']):.2f}"
            ET.SubElement(detalle, "montoIce").text = f"{safe_float(fila['montoIce']):.2f}"
            ET.SubElement(detalle, "montoIva").text = f"{safe_float(fila['montoIva']):.2f}"
            ET.SubElement(detalle, "valorRetBienes").text = f"{safe_float(fila['valorRetBienes']):.2f}"
            ET.SubElement(detalle, "valorRetServicios").text = f"{safe_float(fila['valorRetServicios']):.2f}"
            ET.SubElement(detalle, "valRetServ100").text = f"{safe_float(fila.get('valRetServ100', 0.00)):.2f}"

            ET.SubElement(detalle, "pagoLocExt").text = str(limpiar_entero(fila["pagoLocExt"], 1))
            ET.SubElement(detalle, "formaPago").text = str(limpiar_entero(fila["formaPago"], 1))

    # SOLUCIÓN RADICAL AL ERROR DEL DIMM: 
    # Generamos el string XML del contenido sin ninguna declaración por defecto.
    xml_contenido = ET.tostring(iva, encoding="utf-8").decode("utf-8")
    
    # Concatenamos de forma manual y estricta la cabecera en la misma línea sin caracteres '\n'
    cabecera = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    xml_plano_final = cabecera + xml_contenido

    # Guardamos forzando codificación binaria directa libre de saltos de línea del sistema operativo
    with open(ruta_salida, "w", encoding="utf-8", newline="") as f:
        f.write(xml_plano_final)


def ejecutar_conversion():
    archivo_excel = filedialog.askopenfilename(
        title="Selecciona el Reporte Excel ATS",
        filetypes=[("Archivos de Excel", "*.xlsx")],
    )
    if not archivo_excel:
        return

    try:
        directorio, nombre_base = os.path.split(archivo_excel)
        nombre_sin_ext = os.path.splitext(nombre_base)[0]
        
        # Guardamos con la marca de tiempo requerida
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nuevo_nombre_xml = f"{nombre_sin_ext}_{timestamp}.xml"
        ruta_xml = os.path.join(directorio, nuevo_nombre_xml)
        
        generar_xml_ats(archivo_excel, ruta_xml)
        messagebox.showinfo(
            "¡Proceso Exitoso!", f"XML generado correctamente en:\n{ruta_xml}"
        )
    except Exception as e:
        messagebox.showerror(
            "Error de Conversión", f"Ocurrió un problema:\n{str(e)}"
        )


# --- Configuración Interfaz Gráfica ---
ventana = tk.Tk()
ventana.title("World Class Ecuador - Convertidor ATS")
ventana.geometry("460x340")
ventana.resizable(False, False)

estilo = ttk.Style()
estilo.theme_use("clam")

control_pestañas = ttk.Notebook(ventana)
pestaña_principal = ttk.Frame(control_pestañas)
pestaña_creditos = ttk.Frame(control_pestañas)

control_pestañas.add(pestaña_principal, text=" Converter ")
control_pestañas.add(pestaña_creditos, text=" Créditos e Info ")
control_pestañas.pack(expand=1, fill="both")

# --- PESTAÑA PRINCIPAL ---
lbl_titulo = tk.Label(
    pestaña_principal,
    text="Generador de XML ATS Express",
    font=("Arial", 12, "bold"),
    pady=20,
)
lbl_titulo.pack()

lbl_instruccion = tk.Label(
    pestaña_principal,
    text="Selecciona el reporte Excel exportado para procesar el XML",
    font=("Arial", 9),
    fg="gray",
    pady=5,
)
lbl_instruccion.pack()

btn_convertir = tk.Button(
    pestaña_principal,
    text="🚀 Cargar Excel y Convertir",
    command=ejecutar_conversion,
    bg="#1A5276",
    fg="white",
    font=("Arial", 11, "bold"),
    padx=15,
    pady=12,
    relief="raised",
    bd=2,
)
btn_convertir.pack(pady=20)

# --- PESTAÑA CRÉDITOS ---
try:
    img_logo = Image.open("worldclass-logo.png")
    img_logo = img_logo.resize((120, 120), Image.Resampling.LANCZOS)
    foto_logo = ImageTk.PhotoImage(img_logo)
    lbl_logo = tk.Label(pestaña_creditos, image=foto_logo)
    lbl_logo.image = foto_logo
    lbl_logo.pack(pady=10)
except:
    lbl_respaldo = tk.Label(
        pestaña_creditos,
        text="WORLD CLASS ECUADOR",
        font=("Arial", 12, "bold"),
        fg="#1A5276",
    )
    lbl_respaldo.pack(pady=10)

lbl_empresa = tk.Label(
    pestaña_creditos,
    text="World Class Ecuador",
    font=("Arial", 11, "bold"),
    fg="#2C3E50",
)
lbl_empresa.pack()

lbl_autor = tk.Label(
    pestaña_creditos,
    text="Desarrollado por: Douglas Rujana\nDepartamento de Sistemas",
    font=("Arial", 10, "italic"),
    fg="#566573",
    pady=5,
)
lbl_autor.pack()

lbl_soporte = tk.Label(
    pestaña_creditos,
    text="✉ Soporte: soporteworldclass@gmail.com",
    font=("Arial", 9, "bold"),
    fg="#C0392B",
    pady=5,
)
lbl_soporte.pack()

ventana.mainloop()
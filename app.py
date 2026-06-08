import os
import re
import sys
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xml.etree.ElementTree as ET
import pandas as pd
from PIL import Image, ImageTk


def obtener_ruta_recurso(nombre_archivo):
    """Obtiene la ruta absoluta para un recurso, funciona para desarrollo y para PyInstaller."""
    try:
        # PyInstaller crea una carpeta temporal en _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, nombre_archivo)


def verificar_archivo_bloqueado(ruta_archivo):
    """Comprueba si el archivo Excel está abierto por otro programa (como Excel)."""
    if not os.path.exists(ruta_archivo):
        return
    try:
        # Intentamos abrir el archivo en modo append exclusivo
        with open(ruta_archivo, 'a'):
            pass
    except IOError:
        raise Exception(
            "El archivo Excel seleccionado está abierto por otro programa (ej. Microsoft Excel).\n"
            "Por favor, ciérralo antes de continuar."
        )


def formatear_fecha(valor):
    """Convierte un valor de fecha (Timestamp o string) al formato DD/MM/AAAA requerido por el SRI."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.strftime("%d/%m/%Y")
    valor_str = str(valor).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", valor_str):
        try:
            dt = pd.to_datetime(valor_str)
            return dt.strftime("%d/%m/%Y")
        except:
            pass
    return valor_str


def validar_datos_excel(df):
    """Realiza validaciones completas sobre el DataFrame del Excel."""
    columnas_requeridas = [
        "IdInformante", "razonSocial", "Anio", "Mes", "numEstabRuc",
        "totalVentas", "idProv", "codSustento", "tpIdProv", "tipoComprobante",
        "fechaRegistro", "establecimiento", "puntoEmision", "secuencial",
        "fechaEmision", "autorizacion", "baseNoGraIva", "baseImponible",
        "baseImpGrav", "baseImpExe", "montoIce", "montoIva",
        "valorRetBienes", "valorRetServicios"
    ]
    
    # 1. Verificar columnas obligatorias
    columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
    if columnas_faltantes:
        raise Exception(
            f"Faltan las siguientes columnas obligatorias en la plantilla Excel:\n"
            f"{', '.join(columnas_faltantes)}"
        )
        
    # 2. Validar datos de cabecera (primera fila)
    primera_fila = df.iloc[0]
    
    id_informante = str(primera_fila["IdInformante"]).split(".")[0].strip()
    if not id_informante or len(id_informante) < 10 or len(id_informante) > 13:
        raise Exception(
            f"Error en 'IdInformante' (Cabecera): Debe ser un RUC o Cédula de 10 a 13 dígitos.\n"
            f"Valor encontrado: '{id_informante}'"
        )
        
    anio = limpiar_entero(primera_fila["Anio"])
    if anio < 2000 or anio > 2100:
        raise Exception(
            f"Error en 'Anio' (Cabecera): Año inválido.\n"
            f"Valor encontrado: '{anio}'"
        )
        
    mes = limpiar_entero(primera_fila["Mes"])
    if mes < 1 or mes > 12:
        raise Exception(
            f"Error en 'Mes' (Cabecera): Debe estar entre 1 y 12.\n"
            f"Valor encontrado: '{mes}'"
        )
        
    # 3. Validar registros de compras
    df_compras = df[df["idProv"].notna()]
    for idx, fila in df_compras.iterrows():
        n_fila = idx + 2  # Excel es 1-indexed y tiene fila de cabecera
        
        # Validar idProv (RUC del proveedor)
        tp_id_prov = str(limpiar_entero(fila["tpIdProv"])).zfill(2)
        id_prov = str(fila["idProv"]).split(".")[0].strip()
        if tp_id_prov == "01": # RUC
            id_prov = id_prov.zfill(13)
        elif tp_id_prov == "02": # Cédula
            id_prov = id_prov.zfill(10)

        if not id_prov or len(id_prov) < 10 or len(id_prov) > 13:
            raise Exception(
                f"Fila {n_fila}: RUC/Cédula del proveedor ('idProv') inválido o no corresponde al tipo.\n"
                f"Valor encontrado: '{id_prov}' (Tipo ID: {tp_id_prov})"
            )
            
        # Validar secuencial (debe ser numérico)
        sec = str(limpiar_entero(fila["secuencial"]))
        if not sec.isdigit():
            raise Exception(
                f"Fila {n_fila}: El número de comprobante ('secuencial') debe ser numérico.\n"
                f"Valor encontrado: '{sec}'"
            )
            
        # Validar fechas
        for campo_fecha in ["fechaRegistro", "fechaEmision"]:
            fecha_val = formatear_fecha(fila[campo_fecha])
            if not re.match(r"^\d{2}/\d{2}/\d{4}$", fecha_val):
                raise Exception(
                    f"Fila {n_fila}: La fecha en '{campo_fecha}' debe ser válida y tener el formato DD/MM/AAAA.\n"
                    f"Valor encontrado: '{fecha_val}'"
                )


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
    verificar_archivo_bloqueado(archivo_excel)
    
    try:
        # Intentamos abrir el Excel usando calamine (para soporte de Strict Open XML y velocidad)
        # o openpyxl como fallback
        try:
            xls = pd.ExcelFile(archivo_excel, engine="calamine")
        except Exception:
            xls = pd.ExcelFile(archivo_excel, engine="openpyxl")
            
        nombres_hojas = xls.sheet_names
        hoja_seleccionada = None
        
        # 1. Buscar coincidencia exacta ignorando mayúsculas/minúsculas y espacios
        for nombre in nombres_hojas:
            if nombre.strip().upper() == "ATS":
                hoja_seleccionada = nombre
                break
                
        # 2. Si no se encuentra, buscar cualquier hoja que contenga "ATS"
        if not hoja_seleccionada:
            for nombre in nombres_hojas:
                if "ATS" in nombre.upper():
                    hoja_seleccionada = nombre
                    break
                    
        # 3. Si sigue sin encontrarse, usar la primera hoja por defecto
        if not hoja_seleccionada and nombres_hojas:
            hoja_seleccionada = nombres_hojas[0]
            
        if not hoja_seleccionada:
            raise Exception("El archivo Excel seleccionado no contiene ninguna pestaña.")
            
        df = pd.read_excel(xls, sheet_name=hoja_seleccionada)
        
        # Limpiar espacios en blanco en los nombres de las columnas
        df.columns = [str(col).strip() for col in df.columns]

        # Comprobar si las columnas requeridas están en las columnas del dataframe.
        # Si no están, buscar en las primeras filas si alguna fila contiene los encabezados reales.
        columnas_requeridas_set = {"IdInformante", "razonSocial", "Anio", "Mes", "idProv", "secuencial"}
        columnas_actuales = set(df.columns)
        
        if not columnas_requeridas_set.issubset(columnas_actuales):
            for i in range(min(5, len(df))):
                fila_valores = [str(val).strip() for val in df.iloc[i].values]
                # Limpiar posibles decimales como '.0' de los nombres de columnas
                fila_valores_clean = [v.split('.')[0] if v.endswith('.0') else v for v in fila_valores]
                if columnas_requeridas_set.issubset(set(fila_valores_clean)):
                    # Promover esta fila como encabezado
                    df.columns = fila_valores_clean
                    df = df.iloc[i+1:].reset_index(drop=True)
                    break
    except Exception as e:
        if "Worksheet named" in str(e) or "not found" in str(e):
            raise Exception("No se encontró la pestaña llamada 'ATS' ni ninguna otra hoja en el archivo Excel.")
        raise Exception(f"Error al leer el archivo Excel: {str(e)}")

    # Eliminar filas vacías
    df = df.dropna(subset=["idProv", "secuencial"], how="all")

    if df.empty:
        raise Exception(
            "El archivo Excel está vacío o no contiene registros válidos."
        )

    validar_datos_excel(df)

    primera_fila = df.iloc[0]

    # 1. Crear nodo raíz 'iva'
    iva = ET.Element("iva")

    # 2. Cabecera obligatoria
    tipo_id_informante = str(primera_fila.get("TipoIDInformante", primera_fila.get("TipoIdInformante", "R"))).strip().upper()
    ET.SubElement(iva, "TipoIDInformante").text = tipo_id_informante

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

            tp_id_prov = str(limpiar_entero(fila["tpIdProv"])).zfill(2)
            id_prov = str(fila["idProv"]).split(".")[0].strip()
            if tp_id_prov == "01": # RUC
                id_prov = id_prov.zfill(13)
            elif tp_id_prov == "02": # Cédula
                id_prov = id_prov.zfill(10)

            ET.SubElement(detalle, "codSustento").text = str(limpiar_entero(fila["codSustento"])).zfill(2)
            ET.SubElement(detalle, "tpIdProv").text = tp_id_prov
            ET.SubElement(detalle, "idProv").text = id_prov
            ET.SubElement(detalle, "tipoComprobante").text = str(limpiar_entero(fila["tipoComprobante"])).zfill(2)
            ET.SubElement(detalle, "fechaRegistro").text = formatear_fecha(fila["fechaRegistro"])
            ET.SubElement(detalle, "establecimiento").text = str(limpiar_entero(fila["establecimiento"])).zfill(3)
            ET.SubElement(detalle, "puntoEmision").text = str(limpiar_entero(fila["puntoEmision"])).zfill(3)
            ET.SubElement(detalle, "secuencial").text = str(limpiar_entero(fila["secuencial"])).zfill(9)
            ET.SubElement(detalle, "fechaEmision").text = formatear_fecha(fila["fechaEmision"])
            ET.SubElement(detalle, "autorizacion").text = str(fila["autorizacion"]).split(".")[0]

            def safe_float(val):
                return float(val) if not pd.isna(val) else 0.00

            ET.SubElement(detalle, "baseNoGraIva").text = f"{safe_float(fila['baseNoGraIva']):.2f}"
            ET.SubElement(detalle, "baseImponible").text = f"{safe_float(fila['baseImponible']):.2f}"
            ET.SubElement(detalle, "baseImpGrav").text = f"{safe_float(fila['baseImpGrav']):.2f}"
            ET.SubElement(detalle, "baseImpExe").text = f"{safe_float(fila['baseImpExe']):.2f}"
            ET.SubElement(detalle, "montoIce").text = f"{safe_float(fila['montoIce']):.2f}"
            ET.SubElement(detalle, "montoIva").text = f"{safe_float(fila['montoIva']):.2f}"
            ET.SubElement(detalle, "valRetBien10").text = f"{safe_float(fila.get('valRetBien10', 0.00)):.2f}"
            ET.SubElement(detalle, "valRetServ20").text = f"{safe_float(fila.get('valRetServ20', 0.00)):.2f}"
            ET.SubElement(detalle, "valorRetBienes").text = f"{safe_float(fila.get('valorRetBienes', 0.00)):.2f}"
            ET.SubElement(detalle, "valRetServ50").text = f"{safe_float(fila.get('valRetServ50', 0.00)):.2f}"
            ET.SubElement(detalle, "valorRetServicios").text = f"{safe_float(fila.get('valorRetServicios', 0.00)):.2f}"
            ET.SubElement(detalle, "valRetServ100").text = f"{safe_float(fila.get('valRetServ100', 0.00)):.2f}"
            # Forzamos totbasesImpReemb a 0.00 ya que este convertidor no genera el bloque <reembolsos>.
            # El SRI exige que si no se reporta el bloque de reembolsos, este total sea obligatoriamente 0.00.
            ET.SubElement(detalle, "totbasesImpReemb").text = "0.00"

            # pagoExterior
            pago_exterior = ET.SubElement(detalle, "pagoExterior")
            pago_loc_ext = str(limpiar_entero(fila.get("pagoLocExt"), 1)).zfill(2)
            ET.SubElement(pago_exterior, "pagoLocExt").text = pago_loc_ext
            
            pais_efec_pago = str(fila.get("paisEfecPago")).strip().upper() if not pd.isna(fila.get("paisEfecPago")) else "NA"
            ET.SubElement(pago_exterior, "paisEfecPago").text = pais_efec_pago
            
            aplic_conv_dob_trib = str(fila.get("aplicConvDobTrib")).strip().upper() if not pd.isna(fila.get("aplicConvDobTrib")) else "NA"
            ET.SubElement(pago_exterior, "aplicConvDobTrib").text = aplic_conv_dob_trib
            
            pag_ext_suj_ret_nor_leg = str(fila.get("pagExtSujRetNorLeg")).strip().upper() if not pd.isna(fila.get("pagExtSujRetNorLeg")) else "NA"
            ET.SubElement(pago_exterior, "pagExtSujRetNorLeg").text = pag_ext_suj_ret_nor_leg

            # Calcular total de la transacción para evaluar si se reporta forma de pago (obligatorio si supera 1000 USD)
            total_transaccion = (
                safe_float(fila.get("baseNoGraIva", 0.00)) +
                safe_float(fila.get("baseImponible", 0.00)) +
                safe_float(fila.get("baseImpGrav", 0.00)) +
                safe_float(fila.get("baseImpExe", 0.00)) +
                safe_float(fila.get("montoIce", 0.00)) +
                safe_float(fila.get("montoIva", 0.00))
            )
            if total_transaccion >= 1000.00:
                formas_de_pago = ET.SubElement(detalle, "formasDePago")
                forma_pago = str(limpiar_entero(fila.get("formaPago"), 1)).zfill(2)
                ET.SubElement(formas_de_pago, "formaPago").text = forma_pago

            # Bloque de Impuesto a la Renta (AIR) - va al final de detalleCompras según esquema XSD
            cod_ret_air = str(fila.get("codRetAir", "")).split(".")[0].strip()
            if cod_ret_air and cod_ret_air != "nan" and cod_ret_air != "None" and cod_ret_air != "":
                air_nodo = ET.SubElement(detalle, "air")
                detalle_air = ET.SubElement(air_nodo, "detalleAir")
                ET.SubElement(detalle_air, "codRetAir").text = cod_ret_air
                
                base_imp_air = safe_float(fila.get("baseImpAir", 0.00))
                # Si baseImpAir es 0, usamos baseImpGrav + baseImponible como fallback
                if base_imp_air == 0.00:
                    base_imp_air = safe_float(fila.get("baseImpGrav", 0.00)) + safe_float(fila.get("baseImponible", 0.00))
                ET.SubElement(detalle_air, "baseImpAir").text = f"{base_imp_air:.2f}"
                
                porcentaje_air = safe_float(fila.get("porcentajeAir", 0.00))
                ET.SubElement(detalle_air, "porcentajeAir").text = f"{porcentaje_air:.2f}"
                
                val_ret_air = safe_float(fila.get("valRetAir", 0.00))
                # Fallback cálculo automático si valRetAir es 0
                if val_ret_air == 0.00 and porcentaje_air > 0.00:
                    val_ret_air = base_imp_air * (porcentaje_air / 100.00)
                ET.SubElement(detalle_air, "valRetAir").text = f"{val_ret_air:.2f}"

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
    # El usuario elige el directorio inicial estándar (su carpeta Home / de usuario)
    dir_inicial = os.path.expanduser("~")

    archivo_excel = filedialog.askopenfilename(
        title="Selecciona el Reporte Excel ATS",
        initialdir=dir_inicial,
        filetypes=[("Archivos de Excel", "*.xlsx")],
    )
    if not archivo_excel:
        return

    try:
        # 1. Comprobar si el archivo está bloqueado antes de abrir diálogo de guardado
        verificar_archivo_bloqueado(archivo_excel)
        
        # 2. Preparar el nombre sugerido para el archivo XML
        directorio, nombre_base = os.path.split(archivo_excel)
        nombre_sin_ext = os.path.splitext(nombre_base)[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_sugerido = f"{nombre_sin_ext}_{timestamp}.xml"
        
        # 3. Diálogo "Guardar como" para elegir ubicación y nombre de destino del XML
        ruta_xml = filedialog.asksaveasfilename(
            title="Guardar archivo XML ATS",
            initialdir=directorio,
            initialfile=nombre_sugerido,
            defaultextension=".xml",
            filetypes=[("Archivos XML", "*.xml")],
        )
        if not ruta_xml:
            return  # El usuario canceló la selección de guardado
        
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
    ruta_logo = obtener_ruta_recurso("worldclass-logo.png")
    img_logo = Image.open(ruta_logo)
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
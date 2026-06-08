import pandas as pd

# Creamos un diccionario con las columnas reales de tu ATS
columnas = [
    "TipoIDInformante",
    "IdInformante",
    "razonSocial",
    "Anio",
    "Mes",
    "numEstabRuc",
    "totalVentas",
    "codigoOperativo",
    "codSustento",
    "tpIdProv",
    "idProv",
    "parteRel",
    "tipoComprobante",
    "fechaRegistro",
    "establecimiento",
    "puntoEmision",
    "secuencial",
    "fechaEmision",
    "autorizacion",
    "baseNoGraIva",
    "baseImponible",
    "baseImpGrav",
    "baseImpExe",
    "montoIce",
    "montoIva",
    "valRetBien10",
    "valRetServ20",
    "valorRetBienes",
    "valRetServ50",
    "valorRetServicios",
    "valRetServ100",
    "totbasesImpReemb",
    "pagoLocExt",
    "paisEfecPago",
    "aplicConvDobTrib",
    "pagExtSujRetNorLeg",
    "pagoRegFis",
    "formaPago",
    "codRetAir",
    "baseImpAir",
    "porcentajeAir",
    "valRetAir",
    "numCajBan",
    "precCajBan",
    "estabRetencion1",
    "ptoEmiRetencion1",
    "secRetencion1",
    "autRetencion1",
    "fechaEmiRet1",
    "docModificado",
    "estabModificado",
    "ptoEmiModificado",
    "secModificado",
    "autModificado",
    "DenoProv",
]

# Fila de ejemplo basada en tus datos reales
datos_ejemplo = {
    "TipoIDInformante": ["R"],
    "IdInformante": [993383273001],
    "razonSocial": ["RAPIVISA S.A."],
    "Anio": [2026],
    "Mes": [5],
    "numEstabRuc": [1],
    "totalVentas": [0],
    "codigoOperativo": ["IVA"],
    "codSustento": [1],
    "tpIdProv": [1],
    "idProv": [992560754001],
    "parteRel": ["NO"],
    "tipoComprobante": [1],
    "fechaRegistro": ["01/05/2026"],
    "establecimiento": [1],
    "puntoEmision": [1],
    "secuencial": [942581],
    "fechaEmision": ["01/05/2026"],
    "autorizacion": ["0105202601099256075400120010010009425815629771814"],
    "baseNoGraIva": [0.0],
    "baseImponible": [0.0],
    "baseImpGrav": [36.75],
    "baseImpExe": [0.0],
    "montoIce": [0.0],
    "montoIva": [5.51],
    "valorRetBienes": [0.0],
    "valorRetServicios": [0.0],
    "codRetAir": [303],
    "baseImpAir": [36.75],
    "porcentajeAir": [1.0],
    "valRetAir": [0.37],
    "DenoProv": ["ZUKALO S.A."],
}

# Rellenar el resto de columnas vacías automáticamente
for col in columnas:
    if col not in datos_ejemplo:
        datos_ejemplo[col] = [None]

# Generar el Excel de inmediato
df = pd.DataFrame(datos_ejemplo)
with pd.ExcelWriter("ReporteATSM052026111531.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="ATS", index=False)

print("¡Excel de prueba generado exitosamente sin abrir programas pesados!")
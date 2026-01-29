import pandas as pd

nombre_archivo_input = 'movimientos2.xlsx' 
try:
    df = pd.read_excel(nombre_archivo_input)
    df.columns = df.columns.str.strip()
    print(f"--- Archivo '{nombre_archivo_input}' cargado correctamente ---")
except Exception as e:
    print(f"Error al cargar: {e}")
    exit()

# 1. LIMPIEZA DE MONTO
if df['Monto'].dtype == object:
    df['Monto'] = df['Monto'].replace(r'[\$,]', '', regex=True).astype(float)

# 2. NORMALIZACIÓN DE TIPOS (ESTO ES LO QUE TE FALTABA)
df['Tipo'] = df['Tipo'].str.upper()
df['Tipo'] = df['Tipo'].replace({'DEPOSITO': 'INGRESO'})
df['Tipo'] = df['Tipo'].str.capitalize() # Deja solo 'Ingreso' y 'Egreso'

# 3. LIMPIEZA DE FECHAS
df['Fecha'] = pd.to_datetime(df['Fecha']).dt.date
df['Fecha contable'] = pd.to_datetime(df['Fecha contable']).dt.date

# 4. MOTOR DE CATEGORIZACIÓN MEJORADO
def motor_categorizacion(row):
    detalle = str(row['Detalle']).lower()
    beneficiario = str(row['Beneficiario']).lower()
    tipo = str(row['Tipo']).lower()

    if tipo == 'ingreso':
        if 'reverso' in beneficiario or 'anul' in detalle: 
            return 'Ajustes/Devoluciones'
        if 'efectivo' in detalle or 'deposito' in detalle: 
            return 'Depósitos en Efectivo'
        return 'Otros Ingresos'

    # --- EGRESOS ---
    if 'claro' in detalle: return 'Servicios'
    if 'spotify' in beneficiario or 'multicin' in detalle: return 'Entretenimiento/Suscripciones'
    
    if any(x in detalle for x in ['comis.', 'costo tj', 'iva', 'cost-serv']):
        return 'Gastos Bancarios e Impuestos'
    
    if 'caj/auto.ret' in detalle or 'retiro cnb' in detalle: return 'Efectivo/Cajeros'
    
    if 'kairostex' in beneficiario or 'maestro' in detalle: return 'Compras y Retail'
    
    if 'transf' in detalle or 'pago directo' in detalle: return 'Transferencias a Terceros'

    return 'Varios/Por Clasificar'

# Aplicamos la categorización
df['Categoria_Analisis'] = df.apply(motor_categorizacion, axis=1)

# 5. EXPORTAR RESULTADOS
nombre_salida = 'DatosLimpiados.xlsx'
df.to_excel(nombre_salida, index=False)

print(f"--- PROYECTO COMPLETADO ---")
print(f"El archivo '{nombre_salida}' ahora incluye depósitos como ingresos.")
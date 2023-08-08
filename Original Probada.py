import sqlite3
import pandas as pd
import requests
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import os
from os.path import dirname

# Crear una ventana de diálogo para seleccionar el archivo Excel
root = tk.Tk()
root.withdraw()  # Para esconder la ventana de Tkinter
file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xls *.xlsx *.xlsm')]) 

# Si no se elige ningun archivo
if file_path == '':
    print("Elegir archivo")
    exit () #Salir del Programa

# Obtener la ruta de la carpeta del archivo seleccionado
folder_path_bajada = dirname(file_path)

# Leer el archivo de Excel
xls = pd.ExcelFile(file_path)

# Verificar si el archivo tiene las hojas requeridas
required_sheets = ['Datos', 'Publicaciones', 'Ingresos Brutos', 'Combos', 'Premium']
for sheet in required_sheets:
    if sheet not in xls.sheet_names:
        print(f"El archivo no contiene las hojas necesarias")
        exit()  # Salir del Programa


# Leer el csv desde el enlace proporcionado
url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ2tO5rbdyNMqMkiwQs_fMk9326xylPkDINmsiNKBNLAEaRiblF7BTO7CBp2HuHIUittOfDGrHgibw1/pub?output=csv'
data = pd.read_csv(url)

# Convertir las columnas relevantes a un diccionario
data_dict = dict(zip(data.Titulo, data.Monto))

# Convertir la columna de fecha en un formato de fecha y hora de Python
data['Fecha finalizacion'] = pd.to_datetime(data['Fecha finalizacion'], errors='coerce', dayfirst=True)

# Ahora puedes acceder a los valores en el diccionario usando los nombres de las variables
publiclasica = data_dict.get("publiclasica")
publipremium = data_dict.get("publipremium")
comiscuotas_3 = float(data_dict.get("comiscuotas_3"))
comiscuotas_6 = float(data_dict.get("comiscuotas_6"))
comisahora_3 = float(data_dict.get("comisahora_3"))
comisahora_6 = float(data_dict.get("comisahora_6"))
comisahora_12 = float(data_dict.get("comisahora_12"))
comisfija = int(data_dict.get("Costo Fijo"))
monenviogratis = int(data_dict.get("Monto Envio Gratis"))
comismax = int(data_dict.get("Comision Max"))




#Traemos los datos especificos del vendedor incluyendo el usuario

# Leer los datos específicos del archivo Excel
df_ingresos_brutos = pd.read_excel(file_path, 
                                   sheet_name='Ingresos Brutos', 
                                   header=None,  # No hay encabezado
                                   usecols="A:F",  # Solo las columnas A a F
                                   nrows=2)  # Solo las dos primeras filas



# Para seccion calculo IIBB A CAMBIAR
iibb = df_ingresos_brutos.iloc[1, 0]
enviominimo = df_ingresos_brutos.iloc[1, 1]
porcgananciaoriginal = df_ingresos_brutos.iloc[1, 2]
costoventaweb = df_ingresos_brutos.iloc[1, 3]
gananciaventaweb = df_ingresos_brutos.iloc[1, 4] 
usuario = df_ingresos_brutos.iloc[1, 5] 




# Verificar si el usuario está en la columna 'Titulo'
if usuario in data['Titulo'].values:
    # Buscar 'usuario' en la columna 'Titulo' y obtener los valores correspondientes en las columnas 'Monto' y 'Fecha finalizacion'
    idml = int(data.loc[data['Titulo'] == usuario, 'Monto'].values[0])
    fechafinalizacion = data.loc[data['Titulo'] == usuario, 'Fecha finalizacion'].values[0]
else:
    print("Usuario incorrecto")
    exit()  # Salir del programa

# Comprobar si la 'fechafinalizacion' es anterior a la fecha actual
if pd.to_datetime(fechafinalizacion) < datetime.now():
    print('Tu mes está vencido')
    exit()  # Salir del programa





# Leer los datos del archivo Excel
df_productos = pd.read_excel(file_path, sheet_name='Datos')
df_publicaciones = pd.read_excel(file_path, sheet_name='Publicaciones')



# Creación de la base de datos y la tabla de Productos
con_productos = sqlite3.connect('productos.db')
cursor_productos = con_productos.cursor()

cursor_productos.execute("DROP TABLE IF EXISTS productos")

cursor_productos.execute('''
    CREATE TABLE IF NOT EXISTS productos (
        Codigo TEXT,
        Costo REAL,
        IVA REAL,
        MinimoML REAL,
        MinimoPremium REAL,
        MinimoWEB REAL,
        GananciaML REAL,
        GananciaMLPREMIUM REAL,
        GananciaWEB REAL
    )
''')
con_productos.commit()


# Crear una lista de tuplas de los datos del DataFrame
data_to_insert = df_productos[['Codigo', 'Costo', 'IVA', 'Minimo ML', 'Minimo Premium', 'Minimo WEB', 'Ganancia ML', 'Ganancia ML Premium', 'Ganancia WEB']].values.tolist()

# Inserción de datos en la tabla productos
cursor_productos.executemany(
    'INSERT INTO productos VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)', 
    data_to_insert
)

con_productos.commit()



# Leer los datos de la hoja "Combos"
df_combos = pd.read_excel(file_path, sheet_name='Combos')

# Eliminar las filas donde 'Codigo' está vacío
df_combos = df_combos[pd.notnull(df_combos['Codigo'])]

# Crear una lista de tuplas de los datos del DataFrame
data_to_insert = df_combos[['Codigo', 'Costo', 'IVA', 'Minimo ML', 'Minimo Premium', 'Minimo WEB', 'Ganancia ML', 'Ganancia ML Premium', 'Ganancia WEB']].values.tolist()

# Inserción de datos en la tabla productos
cursor_productos.executemany(
    'INSERT INTO productos VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)', 
    data_to_insert
)

con_productos.commit()





# Creación de la base de datos y la tabla de Publicaciones
con_publicaciones = sqlite3.connect('publicaciones.db')
cursor_publicaciones = con_publicaciones.cursor()

cursor_publicaciones.execute("DROP TABLE IF EXISTS publicaciones")

cursor_publicaciones.execute('''
    CREATE TABLE IF NOT EXISTS publicaciones (
        Publicacion INT,
        Codigo TEXT,
        Category_ID TEXT,
        Listing_Type_ID TEXT,
        Base_Price REAL,
        ID TEXT,
        Estado TEXT,
        Catalogo BOOLEAN,
        Vendedor INT,
        Costo_Envio REAL,
        comisionml REAL
    )
''')
con_publicaciones.commit()





# Reemplazar 'MLA' por '' en la columna 'Publicacion'
df_publicaciones['Publicacion'] = df_publicaciones['Publicacion'].astype(str).str.replace('MLA', '')

# Filtrar filas con celdas vacías en 'Publicacion' o 'Codigo'
df_missing = df_publicaciones[(df_publicaciones['Publicacion'].isna() | (df_publicaciones['Publicacion'] == 'nan')) | 
                              (df_publicaciones['Codigo'].isna() | (df_publicaciones['Codigo'] == ''))]

# Cambiar los valores 'nan' en 'Publicacion' a None y eliminar .0 al final si existe
df_missing.loc[:, 'Publicacion'] = df_missing['Publicacion'].apply(lambda x: '' if x == 'nan' else x.rstrip('.0') if isinstance(x, str) and x.endswith('.0') else x)

# Agregar una columna 'Errores' que representa el índice original más 2 en formato "Fila X"
df_missing['Errores'] = 'Fila ' + (df_missing.index + 2).astype(str)

# Reordenar las columnas
df_missing = df_missing[['Errores', 'Publicacion', 'Codigo']]

# Verificar si el DataFrame df_missing tiene filas
if not df_missing.empty:
    try:
        # Obtener el nombre de usuario del sistema
        user_name = os.getlogin()

        # Definir la ruta de la carpeta de Descargas
        # En Windows
        if os.name == 'nt':
            folder_path = f"C:/Users/{user_name}/Downloads"
        # En macOS y Linux
        else:
            folder_path = f"/home/{user_name}/Downloads"

        # Verificar si la carpeta de Descargas existe
        if not os.path.exists(folder_path):
            raise FileNotFoundError

    except FileNotFoundError:
        # Crear una ventana de diálogo para seleccionar la carpeta
        root = tk.Tk()
        root.withdraw()  # Para esconder la ventana de Tkinter
        folder_path = filedialog.askdirectory()

    # Definir el nombre del archivo
    file_name = 'Correcciones.xlsx'

    # Combinar la ruta de la carpeta con el nombre del archivo
    file_path2 = os.path.join(folder_path, file_name)
    
    with pd.ExcelWriter(file_path2, engine='xlsxwriter') as writer:
        df_missing.to_excel(writer, sheet_name='Sheet1', index=False)
        
        worksheet = writer.sheets['Sheet1']
        # Crear un formato de celda para texto
        text_format = writer.book.add_format({'num_format': '@'})
        # Establecer el formato de la columna 'B' (Publicacion) como texto y ancho de 15
        worksheet.set_column('B:B', 15, text_format)
        # Establecer el ancho de las columnas 'A', 'C' (Errores, Codigo) a 15
        worksheet.set_column('A:A', 15)
        worksheet.set_column('C:C', 15)
        
        red_format = writer.book.add_format({'bg_color': 'red'})
        worksheet.conditional_format('B2:C'+str(len(df_missing)+1), {
            'type': 'blanks',
            'format': red_format
        })

    print('No se pudo procesar: Existen errores. Revise el archivo "Errores de Carga.xlsx".')
    exit()






# Crear una lista de tuplas con los datos a insertar
data_to_insert = [(pub, code, None, None, None, None, None, None, None, None, None) 
                  for pub, code in zip(df_publicaciones['Publicacion'], df_publicaciones['Codigo']) 
                  if pd.notna(pub) and pd.notna(code)]

# Inserción de datos en la tabla publicaciones
cursor_publicaciones.executemany(
    'INSERT INTO publicaciones VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', 
    data_to_insert
)

con_publicaciones.commit()





# Convertir los valores no numéricos de "Costo" a NaN
df_productos['Costo'] = pd.to_numeric(df_productos['Costo'], errors='coerce')

# Convertir los valores no numéricos de "Iva" a NaN
df_productos['IVA'] = pd.to_numeric(df_productos['IVA'], errors='coerce')

# Verificar las filas donde "Costo" es 0, "IVA" no está entre 0 y 1, o donde "IVA", "Costo" o "Publicacion" están vacíos
# Aquí suponemos que quieres verificar los datos en df_productos. Si quieres verificar los datos en df_publicaciones, reemplaza df_productos con df_publicaciones.
df_missing_or_zero = df_productos[
    (df_productos['Costo'].isnull()) | 
    (df_productos['IVA'].isnull()) | 
    (df_productos['Codigo'].isnull()) | 
    (df_productos['Costo'] == 0) | 
    (df_productos['IVA'] < 0) | 
    (df_productos['IVA'] > 1)
]

# Si hay filas con "Costo" vacío o con texto, "IVA" vacío o con texto, "Publicacion" vacío, "Costo" igual a 0, o "IVA" no está entre 0 y 1
if not df_missing_or_zero.empty:
    try:
        # Obtener el nombre de usuario del sistema
        user_name = os.getlogin()

        # Definir la ruta de la carpeta de Descargas
        # En Windows
        if os.name == 'nt':
            folder_path = f"C:/Users/{user_name}/Downloads"
        # En macOS y Linux
        else:
            folder_path = f"/home/{user_name}/Downloads"

        # Verificar si la carpeta de Descargas existe
        if not os.path.exists(folder_path):
            raise FileNotFoundError

    except FileNotFoundError:
        # Crear una ventana de diálogo para seleccionar la carpeta
        root = tk.Tk()
        root.withdraw()  # Para esconder la ventana de Tkinter
        folder_path = filedialog.askdirectory()

    # Definir el nombre del archivo
    file_name = 'Correcciones.xlsx'

    # Combinar la ruta de la carpeta con el nombre del archivo
    file_path2 = os.path.join(folder_path, file_name)

    # Guardar el DataFrame en un archivo de Excel
    with pd.ExcelWriter(file_path2, engine='xlsxwriter') as writer:
        df_missing_or_zero.to_excel(writer, sheet_name='Sheet1', index=False)

        # Obtener la hoja de cálculo de xlsxwriter
        worksheet = writer.sheets['Sheet1']

        # Crear un formato de celda con fondo rojo
        red_format = writer.book.add_format({'bg_color': 'red'})

        # Aplicar formato de fondo rojo a las celdas con valores faltantes en la columna 'Publicacion'
        worksheet.conditional_format('A2:A'+str(len(df_missing_or_zero)+1), {
            'type': 'blanks',
            'format': red_format
        })

        # Aplicar formato de fondo rojo a las celdas con valores faltantes o valor 0 en la columna 'Costo'
        worksheet.conditional_format('B2:B'+str(len(df_missing_or_zero)+1), {
            'type': 'cell',
            'criteria': '=',
            'value': 0,
            'format': red_format
        })
        worksheet.conditional_format('B2:B'+str(len(df_missing_or_zero)+1), {
            'type': 'blanks',
            'format': red_format
        })

        # Aplicar formato de fondo rojo a las celdas con valores faltantes o valores no entre 0 y 1 en la columna 'IVA'
        worksheet.conditional_format('C2:C'+str(len(df_missing_or_zero)+1), {
            'type': 'cell',
            'criteria': '>',
            'value': 1,
            'format': red_format
        })
        worksheet.conditional_format('C2:C'+str(len(df_missing_or_zero)+1), {
            'type': 'cell',
            'criteria': '<',
            'value': 0,
            'format': red_format
        })
        worksheet.conditional_format('C2:C'+str(len(df_missing_or_zero)+1), {
            'type': 'blanks',
            'format': red_format
        })

    print('No se pudo procesar: Existen errores. Revise el archivo "Errores de Carga.xlsx".')
    exit()




# Creación de la base de datos de prodcutos premium
con_premium = sqlite3.connect('premium.db')
cursor_premium = con_premium.cursor()

# Creación de la tabla 'cuotas_3'
cursor_premium.execute("DROP TABLE IF EXISTS cuotas_3")
cursor_premium.execute('''
    CREATE TABLE IF NOT EXISTS cuotas_3 (
        data INT
    )
''')

# Creación de la tabla 'cuotas_6'
cursor_premium.execute("DROP TABLE IF EXISTS cuotas_6")
cursor_premium.execute('''
    CREATE TABLE IF NOT EXISTS cuotas_6 (
        data INT
    )
''')

# Creación de la tabla 'ahora_3'
cursor_premium.execute("DROP TABLE IF EXISTS ahora_3")
cursor_premium.execute('''
    CREATE TABLE IF NOT EXISTS ahora_3 (
        data INT
    )
''')

# Creación de la tabla 'ahora_6'
cursor_premium.execute("DROP TABLE IF EXISTS ahora_6")
cursor_premium.execute('''
    CREATE TABLE IF NOT EXISTS ahora_6 (
        data INT
    )
''')

# Creación de la tabla 'ahora_12'
cursor_premium.execute("DROP TABLE IF EXISTS ahora_12")
cursor_premium.execute('''
    CREATE TABLE IF NOT EXISTS ahora_12 (
        data INT
    )
''')

con_premium.commit()

# Leemos los datos del archivo Excel
df = pd.read_excel(file_path, sheet_name='Premium')

# Inserción de datos en la tabla 'cuotas_3'
for index, row in df.iterrows():
    data = str(row['3 Cuotas']).replace('MLA', '')
    cursor_premium.execute(
        'INSERT INTO cuotas_3 VALUES (?)', 
        (data,)
    )

con_premium.commit()

# Inserción de datos en la tabla 'cuotas_6'
for index, row in df.iterrows():
    data = str(row['6 Cuotas']).replace('MLA', '')
    cursor_premium.execute(
        'INSERT INTO cuotas_6 VALUES (?)', 
        (data,)
    )

con_premium.commit()

# Inserción de datos en la tabla 'ahora_3'
for index, row in df.iterrows():
    data = str(row['3 Ahora']).replace('MLA', '')
    cursor_premium.execute(
        'INSERT INTO ahora_3 VALUES (?)', 
        (data,)
    )

con_premium.commit()

# Inserción de datos en la tabla 'ahora_6'
for index, row in df.iterrows():
    data = str(row['6 Ahora']).replace('MLA', '')
    cursor_premium.execute(
        'INSERT INTO ahora_6 VALUES (?)', 
        (data,)
    )

con_premium.commit()

# Inserción de datos en la tabla 'ahora_12'
for index, row in df.iterrows():
    data = str(row['12 Ahora']).replace('MLA', '')
    cursor_premium.execute(
        'INSERT INTO ahora_12 VALUES (?)', 
        (data,)
    )

con_premium.commit()



# Importar el módulo 'time' para ralentizar la solicitud de la API
import time


# Crear un diccionario para mapear el ID de la publicación con su código
publicacion_codigo_map = {str(row['Publicacion']).replace('MLA', ''): row['Codigo'] for index, row in df_publicaciones.iterrows()}

# Llamada a la API de MercadoLibre y recopilación de información
cursor_publicaciones.execute('SELECT * FROM publicaciones')
rows = cursor_publicaciones.fetchall()

batch_size = 20
for i in range(0, len(rows), batch_size):
    batch_rows = rows[i:i+batch_size]
    batch_publicaciones = [row[0] for row in batch_rows]  # publicacion IDs
    string_ids = ",".join(["MLA" + str(id) for id in batch_publicaciones])

    response = requests.get(f"https://api.mercadolibre.com/items?ids={string_ids}")
    infopublicaciones = response.json()

    # Llamada a la API para obtener el costo de envío
    response_shipping = requests.get(f"https://api.mercadolibre.com/items/shipping_options/free?ids={string_ids}")
    
    # Controlamos que la respuesta sea un JSON válido y un diccionario
    try:
        infoenvios = response_shipping.json()
        if not isinstance(infoenvios, dict):
            print("Respuesta de la API de envío no es un diccionario. Respuesta:", infoenvios)
            infoenvios = {}
    except ValueError:
        print("Respuesta de la API de envío no es un JSON válido. Respuesta:", response_shipping.text)
        infoenvios = {}

    for infopublicacion in infopublicaciones:
        id_mla = infopublicacion['body'].get('id').replace('MLA', '') # Getting the ID from API response
        codigo = publicacion_codigo_map[id_mla]  # Get the codigo for this publicacion from the map

        # En caso que la publicacion tire un error, no la calcula
        if infopublicacion.get('error') in ["resource not found", "not_found"]:
            continue

        # Extraigo la información necesaria de la variable infopublicacion
        category_id = infopublicacion['body'].get('category_id', 'Publicacion Inexistente')
        listing_type_id = infopublicacion['body'].get('listing_type_id', 'Publicacion Inexistente')
        base_price = infopublicacion['body'].get('base_price', 'Publicacion Inexistente')
        estado = infopublicacion['body'].get('status', 'Publicacion Inexistente')
        catalogo = infopublicacion['body'].get('catalog_listing', 'Publicacion Inexistente')
        vendedor = infopublicacion['body'].get('seller_id', 'Publicacion Inexistente')

        # Buscar la información de envío para la publicación actual
        envio_info = infoenvios.get('MLA' + id_mla, None)
        if envio_info is not None and 'coverage' in envio_info and 'all_country' in envio_info['coverage']:
            costoenvio = envio_info['coverage']['all_country'].get('list_cost', 0)
        else:
            costoenvio = 0

        # Actualización de la fila existente con los nuevos datos
        cursor_publicaciones.execute(
            '''
            UPDATE publicaciones 
            SET Category_ID = ?, Listing_Type_ID = ?, Base_Price = ?, ID = ?, Estado = ?, Catalogo = ?, Vendedor = ?, Costo_Envio = ? 
            WHERE Publicacion = ? AND Codigo = ?
            ''', 
            (category_id, listing_type_id, base_price, 'MLA' + id_mla, estado, catalogo, vendedor, costoenvio, id_mla, codigo)
        )


# Crear un diccionario para almacenar los category_ids y los respectivos meli_percentage_fee
category_fee_map = {}


# Leer la tabla 'publicaciones' en un DataFrame
df = pd.read_sql_query("SELECT * FROM publicaciones", con_publicaciones)

for i, row in df.iterrows():
    category_id = row['Category_ID']
    
    # Si el category_id ya fue procesado anteriormente, usamos el valor almacenado en el diccionario
    if category_id in category_fee_map:
        meli_percentage_fee = category_fee_map[category_id]
    else:
        # Si no, hacemos la llamada a la API
        url = f"https://api.mercadolibre.com/sites/MLA/listing_prices?price=10000&category_id={category_id}&listing_type_id=gold_special"
        response = requests.get(url)
        costoml = response.json()

        try:
            # Guardamos el meli_percentage_fee en el diccionario para futuras referencias
            meli_percentage_fee = costoml['sale_fee_details']['meli_percentage_fee']
            category_fee_map[category_id] = meli_percentage_fee
        except KeyError:
            meli_percentage_fee = None  # O asigna un valor predeterminado que te parezca razonable

    # Asignamos el meli_percentage_fee a la columna 'comisionml' para la fila actual
    df.at[i, 'comisionml'] = meli_percentage_fee

for i, row in df.iterrows():
    cursor_publicaciones.execute(
        '''
        UPDATE publicaciones 
        SET comisionml = ?
        WHERE Publicacion = ? AND Codigo = ?
        ''', 
        (row['comisionml'], row['Publicacion'], row['Codigo'])
    )

con_publicaciones.commit()


# Leer la tabla 'publicaciones' en un DataFrame
df = pd.read_sql_query("SELECT * FROM publicaciones", con_publicaciones)


con_productos = sqlite3.connect('productos.db')
cursor_productos = con_productos.cursor()

con_publicaciones = sqlite3.connect('publicaciones.db')
cursor_publicaciones = con_publicaciones.cursor()



#Traemos el valor de comisión maxima

precioalto = 1000000

url = f"https://api.mercadolibre.com/sites/MLA/listing_prices?price={precioalto}&category_id={category_id}&listing_type_id=gold_special"
response = requests.get(url)
costoml = response.json()

costomaximo_previo = costoml['sale_fee_amount']

while precioalto < 1000000000:
    precioalto += 1000000
    url = f"https://api.mercadolibre.com/sites/MLA/listing_prices?price={precioalto}&category_id={category_id}&listing_type_id=gold_special"
    response = requests.get(url)
    costoml = response.json()
    costomaximo_actual = costoml['sale_fee_amount']
    
    if costomaximo_actual == costomaximo_previo:
        break
    else:
        costomaximo_previo = costomaximo_actual

comismax = costomaximo_previo



df_output = pd.DataFrame(columns=['PUBLICACION', 'SKU', 'PRECIO ACTUAL', 'NUEVO PRECIO', 'TIPO DE PUBLICACION'])

df_general = pd.DataFrame(columns=['ID', 'SKU', 'COSTO', 'IVA', 'Tipo publicacion', 'Estado', 'Costo Envio','Costo IIBB','Comision ML %','Precio ML','Ganancia','Ganancia %','Precio WEB','Ganancia WEB', 'Ganancia WEB %'])

cursor_publicaciones.execute('SELECT * FROM publicaciones')
rows = cursor_publicaciones.fetchall()

for row in rows:
    publicacion = row[0]
    codigo = row[1]
    category_id = row[2]
    listing_type_id = row[3]
    base_price = row[4]
    id = row[5]
    estado = row[6]
    catalogo = row[7]
    vendedor = row[8]
    costoenvio = row[9]
    comisionml = row[10]

    if base_price == "Publicacion Inexistente":
            # Se ingresa el producto sin calcular
        df_general = pd.concat([df_general, pd.DataFrame({
            'ID': [id], 
            'SKU': [codigo], 
            'COSTO': ['ERROR EN NUMERO DE PUBLICACION'], 
            'IVA': ['-'],
            'Tipo publicacion': ['-'],
            'Estado': ['-'],
            'Costo Envio': ['-'],
            'Costo IIBB': ['-'],
            'Comision ML %': ['-'],
            'Precio ML': ['-'],
            'Ganancia': ['-'],
            'Ganancia %': ['-'],
            'Precio WEB': ['-'],
            'Ganancia WEB': ['-'],
            'Ganancia WEB %': ['-'],


        })], ignore_index=True)
        continue


    if idml != vendedor:
            # Se ingresa el producto sin calcular
        df_general = pd.concat([df_general, pd.DataFrame({
            'ID': [publicacion], 
            'SKU': [codigo], 
            'COSTO': ['PUBLICACION DE OTRO VENDEDOR'], 
            'IVA': ['-'],
            'Tipo publicacion': ['-'],
            'Estado': ['-'],
            'Costo Envio': ['-'],
            'Costo IIBB': ['-'],
            'Comision ML %': ['-'],
            'Precio ML': ['-'],
            'Ganancia': ['-'],
            'Ganancia %': ['-'],
            'Precio WEB': ['-'],
            'Ganancia WEB': ['-'],
            'Ganancia WEB %': ['-'],


        })], ignore_index=True)
        continue

# ANULADO POR AHORA -------------------------------------------
    # EN caso que la publicacion sea de catalogo, no la calcula (habría que tener un tilde)
    if catalogo == True:
        continue

    cursor_productos.execute(f'SELECT * FROM productos WHERE Codigo = "{codigo}"')
    result = cursor_productos.fetchone()
    if result:
        costo = result[1]
        iva = result[2]
        minimoml = result[3] or 0
        minimopremium = result[4] or 0
        minimoweb = result[5] or 0
        gananciapropuestaml = result[6] or porcgananciaoriginal
        gananciapropuestamlpremium = result[7] or porcgananciaoriginal
        gananciapropuestaweb = result[8] or gananciaventaweb
    else:
            df_general = pd.concat([df_general, pd.DataFrame({
            'ID': [id], 
            'SKU': [codigo], 
            'COSTO': ['CARGAR PRODUCTO'], 
            'IVA': ['-'],
            'Tipo publicacion': ['-'],
            'Estado': ['-'],
            'Costo Envio': ['-'],
            'Costo IIBB': ['-'],
            'Comision ML %': ['-'],
            'Precio ML': ['-'],
            'Ganancia': ['-'],
            'Ganancia %': ['-'],
            'Precio WEB': ['-'],
            'Ganancia WEB': ['-'],
            'Ganancia WEB %': ['-'],


            })], ignore_index=True)
            continue

    #Redondeo
    costo = round(costo,2)


    #Pone el costo minimo en caso que sea necesario
    if costoenvio != 0:
        if costoenvio < enviominimo:
            costoenvio = enviominimo/1.21
        else:
            costoenvio = costoenvio/1.21
#---------------------------------------------------------------------------------------------------
    # Arranca seccion calculo REAL

    # Comprueba si el costo o el IVA son None
    if costo is None or iva is None:
        continue


        # Verificar si el producto se encuentra en alguna tabla premium
    publicacion_int = int(publicacion)

    if publicacion_int in [row[0] for row in cursor_premium.execute("SELECT * FROM cuotas_3")]:
        comisionml += comiscuotas_3
    elif publicacion_int in [row[0] for row in cursor_premium.execute("SELECT * FROM cuotas_6")]:
        comisionml += comiscuotas_6
    elif publicacion_int in [row[0] for row in cursor_premium.execute("SELECT * FROM ahora_3")]:
        comisionml += comisahora_3
    elif publicacion_int in [row[0] for row in cursor_premium.execute("SELECT * FROM ahora_6")]:
        comisionml += comisahora_6
    elif publicacion_int in [row[0] for row in cursor_premium.execute("SELECT * FROM ahora_12")]:
        comisionml += comisahora_12
    elif listing_type_id == publipremium:
        comisionml += comiscuotas_6


    preciosiniva = round(base_price / (1+iva),2)

    costomlsiniva = round((base_price * comisionml/100) / 1.21,2)
    # Si el precio base es menor a monenviogratis, se le suma comisfija antes de dividir por 1.21
    if base_price < monenviogratis:
        costomlsiniva = round(((base_price * comisionml/100) + comisfija) / 1.21,2)

    # Si el valor de costomlsiniva es mayor a comismax, se define costomlsiniva como comismax / 1.21
    if (base_price * comisionml/100) > comismax:
        costomlsiniva = round(comismax / 1.21, 2)


    ingresosbrutos = round(base_price*iibb,2)
    ganancia = round(preciosiniva - costomlsiniva - ingresosbrutos - costoenvio - costo, 2)
    gananciaporcentaje = round(ganancia / base_price, 4)



#---------------------------------------------------------------------------------------------------
    # Arranca seccion calculo PROPUESTO
    # chequea si la publicacion es clasica o premium
    if listing_type_id == publiclasica:
        preciominimo = minimoml
        tipopublicacion = 'Clasica'
    elif listing_type_id == publipremium:
        preciominimo = minimopremium
        tipopublicacion = 'Premium'
        gananciapropuestaml = gananciapropuestamlpremium
    else:
        preciominimo = 0


    # Calculo la ganancia con el precio minimo de ML
    if preciominimo != 0:
        if preciominimo < monenviogratis:
            gananciapreciominimo = (preciominimo/(1+iva) - (preciominimo * (comisionml/100/1.21)) - (preciominimo * iibb) - costoenvio - costo - (comisfija/1.21))
        else: 
            gananciapreciominimo = (preciominimo/(1+iva) - (preciominimo * (comisionml/100/1.21)) - (preciominimo * iibb) - costoenvio - costo)
        porcgananciapreciominimo = gananciapreciominimo/preciominimo

    
    # Calculo la ganancia con la ganancia propuesta

    # calculo para productos con costo fijo
    preciogananciaquerida = ((1+iva)*( costoenvio + costo + (comisfija/1.21)))/((1+(1+iva)*(-gananciapropuestaml - (comisionml/1.21/100) - iibb)))
    if preciogananciaquerida > monenviogratis:
            preciogananciaquerida = ((1+iva)*( costoenvio + costo ))/((1+(1+iva)*(-gananciapropuestaml - (comisionml/1.21/100) - iibb)))

    preciogananciaquerida = round(preciogananciaquerida,2)

    # verifico si el precio minimo es mayor a la ganancia querida, me tiene que dar el precio minimo
    if preciominimo !=0:
        if porcgananciapreciominimo > gananciapropuestaml:
            preciofinalML = preciominimo
        else:
            preciofinalML = preciogananciaquerida
    else:
        preciofinalML = preciogananciaquerida

    preciofinalML = round(preciofinalML,2)




#---------------------------------------------------------------------------------------------------
    # Arranca seccion calculo PRECIO WEB

     # Calculo la ganancia con el precio minimo WEB
    if minimoweb != 0:
        gananciawebpreciominimo = (minimoweb/(1+iva) - (minimoweb * costoventaweb) - (minimoweb * iibb) - costo)
        porcgananciawebpreciominimo = gananciawebpreciominimo/minimoweb   
    

    # calculo con ganancia propuestaweb
    preciogananciawebquerida = ((1+iva)*(costo))/((1+(1+iva)*(-gananciapropuestaweb - iibb - costoventaweb)))


    if minimoweb != 0:
        if porcgananciawebpreciominimo > gananciapropuestaweb:
            preciofinalweb = minimoweb
            gananciaporcentajeweb = porcgananciawebpreciominimo
        else:
            preciofinalweb = preciogananciawebquerida
            gananciaporcentajeweb = gananciapropuestaweb
    else:
        preciofinalweb = preciogananciawebquerida
        gananciaporcentajeweb = gananciapropuestaweb

    gananciaweb = preciofinalweb * gananciaporcentajeweb
    preciofinalweb = round(preciofinalweb,2)
    gananciaweb = round(gananciaweb,2)

    
    # ANULADO POR AHORA -------------------------------------------

    #PONER UN TILDE PARA QUE EL PRECIO WEB NO SEA MAYOR AL DE ML
    if preciofinalweb > preciofinalML:
        preciofinalweb = preciofinalML

    # Redondeo Comision
    comisionml = round(comisionml,2)

    # EXCEL DE CORRECCIONES
    # En caso que el precio final sea distinto al original agregar la nueva fila al DataFrame de salida
    if base_price != preciofinalML:
        df_output = pd.concat([df_output, pd.DataFrame({
            'PUBLICACION': [id], 
            'SKU': [codigo], 
            'PRECIO ACTUAL': [base_price], 
            'NUEVO PRECIO': [preciofinalML],
            'TIPO DE PUBLICACION': [tipopublicacion]

        })], ignore_index=True)


    # EXCEL GENERAL DE PRODUCTOS
    # Se ingresan todos los productos calculados con sus características
    df_general = pd.concat([df_general, pd.DataFrame({
        'ID': [id], 
        'SKU': [codigo], 
        'COSTO': [costo], 
        'IVA': [iva],
        'Tipo publicacion': [tipopublicacion],
        'Estado': [estado],
        'Costo Envio': [costoenvio],
        'Costo IIBB': [ingresosbrutos],
        'Comision ML %': [comisionml/100],
        'Precio ML': [base_price],
        'Ganancia': [ganancia],
        'Ganancia %': [gananciaporcentaje],
        'Precio WEB': [preciofinalweb],
        'Ganancia WEB': [gananciaweb],
        'Ganancia WEB %': [gananciaporcentajeweb],


    })], ignore_index=True)


#Extraer Archivos en la carpeta desde dode subio el archivo original


# Definir el nombre del archivo
file_name = 'Extraccion.xlsx'

# Combinar la ruta de la carpeta con el nombre del archivo
file_path2 = folder_path_bajada + '/' + file_name

# Crear un ExcelWriter con el nombre del archivo
with pd.ExcelWriter(file_path2, engine='xlsxwriter') as writer:

    # Guardar el DataFrame de las Correcciones en la primera hoja
    df_output.to_excel(writer, sheet_name='Correcciones', index=False)

    # Obtener el libro de trabajo y la hoja de trabajo para cambiar el formato
    workbook  = writer.book
    worksheet = writer.sheets['Correcciones']

    # Ajustar el ancho de todas las columnas (200 pixels es aproximadamente 50 en Excel)
    worksheet.set_column('A:E', 20)

    # Crear un formato de moneda
    money_format = workbook.add_format({'num_format': '$#,##0.00'})

    # Aplicar el formato de moneda a las columnas C y D
    worksheet.set_column('C:C', 20, money_format)
    worksheet.set_column('D:D', 20, money_format)

    # Guardar el DataFrame de todos los productos calculados en la segunda hoja
    df_general.to_excel(writer, sheet_name='Productos Exportados', index=False)

    # Obtener la hoja de trabajo para cambiar el formato
    worksheet = writer.sheets['Productos Exportados']

    # Ajustar el ancho de las columnas
    worksheet.set_column('A:B', 20) # Ajustar ancho a 20 para las columnas A y B
    worksheet.set_column('C:O', 9.3) # Ajustar ancho a 9.3 para las columnas C a O

    # Aplicar el formato de moneda a las columnas especificadas
    for column in ['C', 'G', 'H', 'J', 'K', 'M', 'N']:
        worksheet.set_column(f'{column}:{column}', 9.3, money_format)

    # Crear un formato de porcentaje
    percent_format = workbook.add_format({'num_format': '0.00%'})

    # Aplicar el formato de porcentaje a las columnas especificadas
    for column in ['D', 'I', 'L', 'O']:
        worksheet.set_column(f'{column}:{column}', 9.3, percent_format)



con_productos.close()
con_publicaciones.close()
con_premium.close()

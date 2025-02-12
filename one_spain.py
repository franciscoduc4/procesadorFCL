# %%
import pandas as pd
import os
import re

def generate_template_dataframe():
    """Genera un dataframe con las columnas del template"""
    columns = ['ORIGIN LOCATION', 'ORIGIN PORT', 'DESTINATION PORT', 'DESTINATION LOCATION', 'CHARGE', 
               'RATE BASIS', 'CURRENCY', '20DRY', '40DRY', '40HDRY', '45HDRY', '40NOR', '20RF', '40RF', '40HCRF', '45RF', '20OT', '40OT', '40HCOT', 
               '20FR', '40FR', '40HCFR', 'CHARGE TYPE', 'PAYMENT TERM', 'PROVIDER', 'LIMITS', 'START DATE', 'EXPIRATION DATE', 'VIA', 
               'TRANSIT TIME', 'COMMODITY', 'SERVICE NAME', 'INCLUDED CHARGES', 'REMARKS', 'MODE OF TRANSPORT', 'EXCEPTIONS ORIGIN', 'EXCEPTIONS DESTINATION', 'NOTAS']
    return pd.DataFrame(columns=columns)


def save_dataframe_multi(df_tarifario, df_arbitraries, original_name, output_dir):
    """Guarda dos dataframes en un mismo archivo Excel, en dos hojas distintas."""
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(original_name)[0]
    counter = 1
    file_path = os.path.join(output_dir, f"{base_name}_procesado.xlsx")
    
    while os.path.exists(file_path):
        file_path = os.path.join(output_dir, f"{base_name}_procesado_{counter}.xlsx")
        counter += 1
        
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df_tarifario.to_excel(writer, sheet_name='Tarifario', index=False)
        df_arbitraries.to_excel(writer, sheet_name='Arbitraries', index=False)
    
    return file_path

def save_dataframe(df, original_name, output_dir):
    """Guarda el dataframe procesado"""
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(original_name)[0]
    counter = 1
    file_path = os.path.join(output_dir, f"{base_name}_procesado.xlsx")
    
    while os.path.exists(file_path):
        file_path = os.path.join(output_dir, f"{base_name}_procesado_{counter}.xlsx")
        counter += 1
        
    df.to_excel(file_path, index=False)
    return file_path


def obtener_df_archivo(archivo, hojas):
    for hoja in hojas[1:2]:
        print(f"LEYENDO HOJA {hoja}")
        df_completo = pd.read_excel(archivo, sheet_name=hoja, header=None)

        fila_encabezado_surcharges = None
        for i in range(20):
            fila = df_completo.iloc[i].astype(str)
            if 'SUBJECT SURCHARGES:' in ' '.join(fila.values):
                fila_encabezado_surcharges = i
                print("Fila encontrada:", fila_encabezado_surcharges)
                break

        if fila_encabezado_surcharges is None:
            raise ValueError("No se encontró el encabezado de surcharges en las primeras 20 filas.")

        df_surcharges = pd.read_excel(
            archivo,
            sheet_name=hoja,
            header=fila_encabezado_surcharges
        )

        first_empty_row = 0
        for index, row in df_surcharges.iterrows():
            if row.isna().all():
                first_empty_row = index
                break

        df_surcharges = df_surcharges.iloc[:first_empty_row]

        df_surcharges = df_surcharges.dropna(axis=1, how='all')
        df_surcharges.name = hoja
        #Cambiar el nombre de la columna 'Unnamed: 7' a Remarks
        df_surcharges.rename(columns={'Unnamed: 7': 'Remarks'}, inplace=True)
        print("[DF SURCHARGES] Columnas:", df_surcharges.columns)


        fila_encabezado_freights = None
        for i in range(fila_encabezado_surcharges + 20, fila_encabezado_surcharges + 60):
            fila = df_completo.iloc[i].astype(str)
            if 'EFFECTVE DATE' in ' '.join(fila.values):
                fila_encabezado_freights = i
                break

        if fila_encabezado_freights is None:
            raise ValueError("No se encontró el encabezado de freights en el rango esperado.")
        else:
            print("Fila encontrada:", fila_encabezado_freights)

        df_freights = pd.read_excel(
            archivo,
            sheet_name=hoja,
            header=fila_encabezado_freights
        )
        df_freights = df_freights.dropna(axis=1, how='all')
        df_freights.name = hoja
        # print(df_freights.head())
        df_freights.columns = df_freights.columns.astype(str)
        print("[DF FREIGHT] Columnas:", df_freights.columns)
        
        return df_freights, df_surcharges
        
def map_data_to_template(df_freights, df_surcharges):
    df_template = generate_template_dataframe()
    
    cargo_mappings = {
        'Dry General': {'20': ['20DRY', '20OT', '20FR'], '40': ['40DRY', '40OT', '40FR'], 'HC': ['40HDRY', '45HDRY']},
        'Reefer': {'20': '20RF', '40': '40RF', 'HC': '40HCRF'},
        'Reefer Dry': {'40': '40NOR'}
    }
    
    route_info = {}
    
    for index, row in df_freights.iterrows():
        origin = row['ORIGIN'] if pd.notna(row['ORIGIN']) else ''
        destination = row['DESTINATION'] if pd.notna(row['DESTINATION']) else ''
        key = (origin, destination)
        
        if key not in route_info:
            route_info[key] = {
                'START DATE': '',
                'EXPIRATION DATE': '',
                'SERVICE NAME': '',
                'ORIGIN PORT': origin,
                'DESTINATION PORT': destination,
                'VIA': '',
                'CURRENCY': '',
                'REMARKS': '',
                'INCLUDED CHARGES': '',
                'NOTAS': df_freights.name,
                **{col: '' for col in df_template.columns if col not in ['ORIGIN PORT', 'DESTINATION PORT', 'START DATE', 'EXPIRATION DATE', 'SERVICE NAME', 'VIA', 'CURRENCY', 'REMARKS', 'INCLUDED CHARGES', 'NOTAS']}
            }

        # Actualizar información solo si es la primera vez o si era nula antes
        field_mappings = {
            'START DATE': 'EFFECTVE DATE',
            'EXPIRATION DATE': 'EXPIRY DATE',
            'SERVICE NAME': 'SERVICE SCOPE',
            'VIA': 'T/S PORT',
            'CURRENCY': 'CUR.',
            'REMARKS': 'SERVICE',
            'INCLUDED CHARGES': 'INCLUDE SURCHARGES'
        }
        
        for field, source in field_mappings.items():
            if pd.notna(row.get(source)) and not route_info[key][field]:
                route_info[key][field] = row[source].strftime('%d-%m-%Y') if field in ['START DATE', 'EXPIRATION DATE'] else row[source]
        
        # Mapeo basado en 
        cargo_type = row['CARGO'] if pd.notna(row['CARGO']) else ''
        if cargo_type in cargo_mappings:
            if cargo_type == 'Dry General': 
                for size, fields in cargo_mappings['Dry General'].items():
                    if pd.notna(row[size]):
                        for field in fields:
                            if not route_info[key][field]:  # Asegurarse de que el campo existe
                                route_info[key][field] = row[size] if pd.notna(row[size]) else ''
            elif cargo_type == 'Reefer':
                for size, field in cargo_mappings['Reefer'].items():
                    if pd.notna(row[size]):
                        if isinstance(field, list):  # Si field es una lista, iterar sobre ella
                            for f in field:
                                if not route_info[key][f]:
                                    route_info[key][f] = row[size]
                        else:  # Si es un solo campo
                            if not route_info[key][field]:
                                route_info[key][field] = row[size]
            elif cargo_type == 'Reefer Dry':
                if pd.notna(row['40']) and not route_info[key]['40NOR']:
                    route_info[key]['40NOR'] = row['40']
        
        route_info[key]['RATE BASIS'] = 'PER CONTAINER'
        route_info[key]['CHARGE'] = 'FREIGHT'
        if route_info[key]['SERVICE NAME'][-2:] == 'MW' or route_info[key]['SERVICE NAME'][-2:] == 'EW':
            route_info[key]['SERVICE NAME'] = route_info[key]['SERVICE NAME'][-2:]
        else:
            route_info[key]['SERVICE NAME'] = ''
        
    # Convertir el diccionario en un DataFrame, manteniendo el orden del template
    result_df = pd.DataFrame.from_dict(route_info, orient='index')[df_template.columns]
    #Eliminar la ultima fila
    result_df = result_df[:-1]
    
    unique_routes = result_df[['ORIGIN PORT', 'DESTINATION PORT', '20DRY', '40DRY', '40HDRY', '45HDRY', '40NOR', '20RF', '40RF', '40HCRF', '45RF', '20OT', '40OT', '40HCOT', '20FR', '40FR', '40HCFR', 'START DATE', 'EXPIRATION DATE', 'VIA', 'SERVICE NAME']].drop_duplicates()
    surcharge_df = process_surcharges(df_surcharges, unique_routes)
    final_df = pd.concat([result_df, surcharge_df], ignore_index=True)
    
    return final_df

def process_surcharges(df_surcharges, unique_routes):
    surcharge_rows = []
    container_columns = ['20DRY', '40DRY', '40HDRY', '45HDRY', '40NOR', 
                         '20RF', '40RF', '40HCRF', '45RF', '20OT', '40OT', 
                         '40HCOT', '20FR', '40FR', '40HCFR']
    
    basis_mapping = {
        'PER TEU': 'PER_TEU',
        'PER BL': 'PER_BL',
        'PER CONTAINER': 'PER_CONTAINER',
        'UNIT': 'PER_CONTAINER',
        'TEU': 'PER_TEU',
        'PER BOOKING': 'PER_BOOKING'
    }

    # Lista de recargos permitidos
    valid_surcharges = ['OBS', 'ETS', 'SCT', 'HEA']

    # Determinar qué columnas nunca tienen valores en unique_routes
    always_empty_containers = ['40NOR', '40HCOT', '40HCFR', 
                               '40RF', '45RF']
    
    # Eliminar esas columnas del procesamiento
    container_columns = [col for col in container_columns if col not in always_empty_containers]

    print(f"Se eliminaron estas columnas por estar vacías en todas las rutas: {always_empty_containers}")

    for _, surcharge_row in df_surcharges.iterrows():
        # Limpiar datos
        charge_full = str(surcharge_row['SUBJECT SURCHARGES:'])
        charge_name = charge_full.split('(')[0].strip() if '(' in charge_full else charge_full
        
        # Filtrar solo los recargos deseados
        if charge_name not in valid_surcharges:
            continue

        type_ = str(surcharge_row['TYPE ']).strip()
        basis = str(surcharge_row['BASIS ']).strip().upper()
        rate_basis = basis_mapping.get(basis, 'PER_CONTAINER')
        remarks = str(surcharge_row['Remarks'])
        
        # Determinar Charge Type
        charge_type = 'FREIGHT'
        
        # Procesar monto
        amount = surcharge_row['AMOUNT']
        if isinstance(amount, str):
            amount = float(amount.replace('$', '').strip())

        for _, route in unique_routes.iterrows():
            origin = route['ORIGIN PORT']
            destination = route['DESTINATION PORT']

            # Obtener contenedores con valores en FREIGHT desde unique_routes (excluyendo los vacíos)
            valid_containers = [col for col in container_columns if pd.notna(route[col])]

            # Mapear contenedores afectados
            container_cols = {}
            if basis in container_columns:
                # Si el BASIS es un contenedor específico, solo asignarlo si tiene valor en FREIGHT
                if basis in valid_containers:
                    container_cols[basis] = amount
            else:
                # Determinar a qué contenedores se debe asignar el recargo
                if type_ == 'DRY':
                    targets = [c for c in valid_containers if 'DRY' in c or 'OT' in c or 'FR' in c]
                elif type_ == 'REEFER':
                    targets = [c for c in valid_containers if 'RF' in c or 'HCRF' in c]
                elif type_ == 'ALL':
                    targets = valid_containers  # Aplicar a todas las columnas con valor
                else:
                    targets = []
                
                # Asignar el monto solo a las columnas válidas
                for target in targets:
                    container_cols[target] = amount

            # Extraer límites de Remarks
            limits = ''
            weight_match = re.search(r'>= (\d+)', remarks)
            if weight_match:
                limits = weight_match.group(1)

            # Buscar si ya existe una fila con el mismo ORIGIN, DESTINATION, CHARGE y RATE BASIS
            matching_row = next(
                (row for row in surcharge_rows if 
                 row['ORIGIN PORT'] == origin and 
                 row['DESTINATION PORT'] == destination and 
                 row['CHARGE'] == charge_name and 
                 row['RATE BASIS'] == rate_basis),
                None
            )
            
            if matching_row:
                # Si ya existe una fila para este recargo y ruta, agregar los valores a las columnas correspondientes
                for container, value in container_cols.items():
                    matching_row[container] = value
            else:
                # Si no existe una fila, crear una nueva
                new_row = {
                    'ORIGIN PORT': origin,
                    'DESTINATION PORT': destination,
                    'CHARGE': charge_name,
                    'RATE BASIS': rate_basis,
                    'CURRENCY': surcharge_row['CURRENCY'],
                    'PAYMENT TERM': 'COLLECT' if charge_name == 'ESD' and origin in ['BD', 'LK'] else ('PREPAID' if charge_name == 'ESD' else ''),
                    'LIMITS': limits,
                    **{k: route[k] for k in ['START DATE', 'EXPIRATION DATE', 'VIA', 'SERVICE NAME']},
                    **container_cols
                }
                surcharge_rows.append(new_row)
    
    return pd.DataFrame(surcharge_rows)


def extract_validity_dates(file_path, sheet_name, search_range=(7, 10), col_range=("N", "Q")):

    # Cargar el archivo de Excel sin encabezado para buscar fechas
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    # Convertir nombres de columnas de letras a índices numéricos
    col_indices = [ord(col) - ord('A') for col in col_range]

    # Buscar las fechas en las filas del rango especificado
    start_date, expiration_date = None, None
    for i in range(search_range[0], search_range[1] + 1):
        row_values = df_raw.iloc[i, col_indices[0]:col_indices[1] + 1].dropna().tolist()
        if "Valid from" in row_values:
            start_date = row_values[row_values.index("Valid from") + 1] if len(row_values) > row_values.index("Valid from") + 1 else None
        if "Valid till" in row_values:
            expiration_date = row_values[row_values.index("Valid till") + 1] if len(row_values) > row_values.index("Valid till") + 1 else None

    return start_date, expiration_date


def process_arbitraries(file_path, sheet_name, search_range=(7, 15)):

    def find_header_row(df, search_range):
        """ Busca dinámicamente la fila del encabezado en un rango dado. """
        header_keywords = [
            "Service Scope", "Origin /Destination", "Point", "Origin", "Over (Rate Over)",
            "Via (Route Over)/TS", "Cargo Type", "Cur.", "20", "40", "HC", "Remarks"
        ]

        for i in range(search_range[0], search_range[1] + 1):
            row_values = df.iloc[i].astype(str).str.strip().tolist()
            matches = sum(1 for keyword in header_keywords if keyword in row_values)
            if matches >= len(header_keywords) // 2:  # Al menos la mitad de las columnas deben coincidir
                return i

        return None  # Retorna None si no encuentra el encabezado

    # Cargar el archivo de Excel sin encabezados
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    # Encontrar la fila del encabezado
    header_row = find_header_row(df_raw, search_range)
    if header_row is None:
        raise ValueError("No se encontró el encabezado en el rango especificado.")

    # Extraer la fila del encabezado y asignarla correctamente
    new_column_names = df_raw.iloc[header_row].astype(str).str.strip().tolist()
    df_arbitraries = df_raw[header_row + 1:].reset_index(drop=True)

    # Aplicar los nombres de las columnas
    df_arbitraries.columns = new_column_names

    # Eliminar columnas innecesarias como 'nan' o vacías
    df_arbitraries = df_arbitraries.loc[:, ~df_arbitraries.columns.str.lower().str.contains('nan')]

    # Renombrar las columnas según el template esperado
    rename_map = {
        "Origin": "ORIGIN",
        "Over (Rate Over)": "RATE OVER",
        "Via (Route Over)/TS": "VIA",
        "Cargo Type": "CARGO TYPE",
        "Cur.": "CURRENCY",
        "20": "20FT",
        "40": "40FT",
        "HC": "40HC",
        "Remarks": "REMARKS"
    }

    # Validar que las columnas esperadas existan antes de renombrar
    rename_map = {k: v for k, v in rename_map.items() if k in df_arbitraries.columns}
    df_arbitraries = df_arbitraries.rename(columns=rename_map)

    # Verificar si "RATE OVER" está presente ahora
    if "RATE OVER" not in df_arbitraries.columns:
        raise ValueError("RATE OVER no está presente después de la reasignación de encabezado.")

    # Extraer fechas de validez
    start_date, expiration_date = extract_validity_dates(file_path, sheet_name)
    print("Fechas de validez encontradas:", start_date, expiration_date)
    start_date = str(start_date).split(" ")[0] if start_date else None
    start_date = str(start_date).replace("-", "/") if start_date else None
    expiration_date = str(expiration_date).split(" ")[0] if expiration_date else None
    expiration_date = str(expiration_date).replace("-", "/") if expiration_date else None

    # Mapeo de Cargo Type a las columnas correspondientes
    cargo_mapping = {
        "Dry General": {"20FT": "20DRY", "40FT": "40DRY", "40HC": "40HDRY"},
        "Reefer": {"20FT": "20RF", "40FT": "40RF", "40HC": "40HCRF"},
        "Reefer Dry": {"20FT": "20DRY", "40FT": "40DRY", "40HC": "40HDRY"}
    }

    # Crear un dataframe con la estructura del template procesado
    template_columns = [
        "ORIGIN", "DESTINATION", "RATE OVER", "VIA", "CHARGE", "RATE BASIS", "MODE OF TRANSPORT",
        "CURRENCY", "20DRY", "40DRY", "40HDRY", "45HDRY", "40NOR", "20RF", "40RF", "40HCRF", "45RF",
        "20OT", "40OT", "40HCOT", "20FR", "40FR", "40HCFR", "PROVIDER", "START DATE",
        "EXPIRATION DATE", "COMMODITY", "SERVICE NAME", "REMARKS", "EXCEPTIONS ORIGIN",
        "EXCEPTIONS DESTINATION", "NOTAS", "SERVICE SCOPE", "SERVICE SCOPE 2"
    ]
    df_processed = pd.DataFrame(columns=template_columns)

    # Asignar valores a las columnas
    df_processed["ORIGIN"] = df_arbitraries["ORIGIN"]
    df_processed["RATE OVER"] = df_arbitraries["RATE OVER"]
    df_processed["VIA"] = df_arbitraries["VIA"]
    df_processed["CURRENCY"] = df_arbitraries["CURRENCY"]
    df_processed["REMARKS"] = df_arbitraries["REMARKS"]

    # Asignar valores de CHARGE y RATE BASIS
    df_processed["CHARGE"] = "TAO"
    df_processed["RATE BASIS"] = "PER_CONTAINER"

    # Aplicar la nueva lógica para "via Rail"
    df_processed.loc[df_arbitraries["REMARKS"].str.contains("via Rail", na=False, case=False), "CHARGE"] = "PRE CARRIAGE - RAIL"
    df_processed.loc[df_arbitraries["REMARKS"].str.contains("via Rail", na=False, case=False), "MODE OF TRANSPORT"] = "RAIL"

    # Asignar valores de las columnas de contenedores según Cargo Type
    for index, row in df_arbitraries.iterrows():
        cargo_type = row["CARGO TYPE"]
        if cargo_type in cargo_mapping:
            for size, column in cargo_mapping[cargo_type].items():
                df_processed.at[index, column] = row[size]

    # Filtrar rutas con Cargo Type "Dry Dangerous" (NO PROCESAR)
    df_processed = df_processed[df_arbitraries["CARGO TYPE"] != "Dry Dangerous"]

    # Asignar fechas de validez
    df_processed["START DATE"] = start_date
    df_processed["EXPIRATION DATE"] = expiration_date

    return df_processed


# Uso de la función con búsqueda dinámica del encabezado


# # Procesar el archivo con búsqueda dinámica del encabezado
# sheet_name = "OUTPORTS - ARBITRARIES"
# # Uso:
# archivo = 'C:/Users/ruso_/OneDrive/Escritorio/C5/cargofiveLCL/procesadorFCL/151896_Ibercondor_DRY_04-02-2025_FCL.xlsx'
# archivo = 'C:/Users/ruso_/OneDrive/Escritorio/C5/cargofiveLCL/procesadorFCL/128908_Bas Josa_DRY_04-02-2025_FCL.xlsx'
# xls = pd.ExcelFile(archivo)
# hojas = xls.sheet_names
# df_freights, df_surcharges = obtener_df_archivo(archivo, hojas)
# result = map_data_to_template(df_freights, df_surcharges)
# df_arbitraries = process_arbitraries(archivo, hojas[5], search_range=(7, 15))
# save_dataframe_multi(result, df_arbitraries, os.path.basename(archivo), 'C:/Users/ruso_/OneDrive/Escritorio/C5/cargofiveLCL/procesadorFCL/')



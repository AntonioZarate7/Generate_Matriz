import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

directorio_actual = Path.cwd()
file_path = directorio_actual/'Matiz Emision Pruebas 2024 v2_NUEVA.xlsx'

data = pd.read_excel(file_path, sheet_name='Sheet4')

# print(data.head())
print(data.columns) # ------> Renombrar el nombre de alguans columnas


# Obtiene la fecha actual y le agrega 2 numeros
def fecha_contacto():
    fecha_actual = datetime.now()
    fecha_nueva = fecha_actual + timedelta(days=2)
    fecha = fecha_nueva.strftime('%d-%m-%Y')
    print(fecha)
    return fecha


def fecha_siguiente_contacto():
    fecha_actual = datetime.now()
    fecha_nueva = fecha_actual + timedelta(days=2*30)
    fecha = fecha_nueva.strftime('%d-%m-%Y')
    print(fecha)
    return fecha

id_column = data.pop('ID')
data.insert(0, 'TCID', 'NA')
data.insert(1, 'Nombre_foler_evidencia', id_column)
data.insert(2, '¿Crear_Oçportunidad?', 'Si')
data.insert(3, '¿Crear_Cotizacion?', 'Si')
data.insert(4, 'Nombre(s)', 'NA')
data.insert(5, 'ApellidoPaterno', 'NA')
data.insert(6, 'ApellidoMaterno', 'NA')
data.insert(7, 'Fuente', 'Networking')
data.insert(8, 'FormaDeContacto', 'Visita')
data.insert(9, 'FechaDeContacto', fecha_contacto())
data.insert(10, 'TelefonoCelular', '5530112344')
data.insert(11, 'TipoDeTelefono', 'Celular')
data.insert(12, 'FechaNacimiento', None)
edad_column = data.pop('Edad')
data.insert(13, 'Edad', edad_column)
data.insert(14, 'Genero', 0)
data.insert(15, 'Fuma', 'No')
data.insert(16, 'IngresoMensual', '100000')
data.insert(17, 'NombreOportunidad', 'demo_1')
data.insert(18, 'FechaSiguienteContacto', fecha_siguiente_contacto())
data.insert(19, 'Oportunidad', 'demo_1')
data.insert(20, 'CP', '0')
data.insert(21, 'Estado', 'NA')
data.insert(22, 'Preferente', 0)
data.insert(23, 'Edad_titular', edad_column)
data.insert(24, 'Genero ', 0)
data.insert(25, 'Solicitante_es_igual_al_Asegurado', '')
data.insert(26, 'Esposa', 0)
data.insert(27, 'Asegurados', 0)
data.insert(28, 'TipoPlan', 0)
plan_column = data.pop('Plan')
data.insert(29, 'Plan', 0)
data.insert(30, 'PLAN_TEXTO', plan_column)
coa_column = data.pop('Coaseguro')
data.insert(31, 'Coaseguro', coa_column*100)
data.pop('Tipo de Póliza')
inc_column = data.pop('Incremento GURA')
data.insert(32, 'CHMQ', inc_column*100)
data.insert(33, 'DeducibleUnico', 0)
data.insert(34, 'Red_Copago', 0)
data.insert(45, 'DP', 0)
data.insert(51, 'CEDAP', 0)

coverage_columns = ['CPF', 'CAE', 'CEC', 'CEE', 'CEDA', 'DENTAL', 'AMCD', 'CEDA PREM', 'CRFCA']

""" Genera las columnas de Preferente para el valor que tiene el titular y hasta 11 asegurados """
preferente_columns = ['Preferente_A' + str(i) for i in range(1, 12)]
for col in preferente_columns:
    data[col] = 0  # Iniciar todas las columnas Preferente_A con 0


""" Genera las columnas de CETTE para el valor que tiene el titular y hasta 11 aseguirados
  Inicia todas las columnas con valor 0 """
cette_columns = ['cetteTit'] + [f'cette{i}' for i in range(1, 12)]
for col in cette_columns:
    data[col] = 'OFF'  # Iniciar todas las columnas CETTE con OFF


""" Crear columnas para fechas de nacimeinto """
for i in range(1, 12):
    data[f'Fecha_Nacimiento_A{i}'] = pd.NaT


""" Genera las columnas para saber si fuman """
fuma_columns = [f'Fuma_A{i}' for i in range(1, 12)]
for col in fuma_columns:
    data[col] = 0  # Le asigna el valor 0 a todas las columnas


""" Genera las columnas para los nombres de los asegurados """
for i in range(1, 12):
    data[f'Nombre_A{i}'] = ''
    data[f'Apellido_Paterno_A{i}'] = ''
    data[f'Apellido_Materno_A{i}'] = ''


""" Genera las columnas para el genero de los asegurados """
for i in range(1, 12):
    data.insert(125 + i, f'Genero_A{i}', 0)


""" Agrega las columnas de parentesco de los asegurados """
for i in range(1, 12):
    data[f'Rol_A{i}'] = ''


# Agregar columnas para cash
data['cashTit'] = 'OFF'  # Columna para el titular
for i in range(1, 12):
    data[f'cash{i}'] = 'OFF'  # Columnas para asegurados


# Agregar la columna PrimaTotalAnualizada al final del DataFrame
data['PrimaTotalAnualizada'] = data['Pma + Der']

# Insertar la columna Frecuencia en la penúltima posición
data.insert(len(data.columns) - 1, 'Frecuencia', data['Forma de pago'].str.capitalize())


data['CPF'] = data['CPF'].apply(lambda x: 1 if pd.notna(x) else 0)
data['CAE'] = data['CAE'].apply(lambda x: 1 if pd.notna(x) else 0)
data['CEC'] = data['CEC'].apply(lambda x: 1 if pd.notna(x) else 0)
data['CEE'] = data['CEE'].apply(lambda x: 1 if pd.notna(x) else 0)
data['CEDA'] = data['CEDA'].apply(lambda x: 1 if pd.notna(x) else 0)


# Diccionario de codigos postales
codigos_postales = {

    'CD1': '01000',
    'CD2': '55720',
    'CEN': '76000',
    'NL': '64000',
    'NOR': '98000',
    'JAL': '48903',
    'OCC': '36000',
    'NOE': '21000',
    'CS': '94327',
    'PEN': '97000'
}

estados = {
    'CD1': 'CIUDAD DE MÉXICO',
    'CD2': 'MÉXICO',
    'CEN': 'OAXACA',
    'NL': 'NUEVO LEON',
    'NOR': 'ZACATECAS',
    'JAL': 'JALISCO',
    'OCC': 'AGUASCALIENTES',
    'NOE': 'SONORA',
    'CS': 'VERACRUZ',
    'PEN': 'TABASCO'
}


# Funcion para distribuir coberturas y manejar los datos de CETTE, Riesgo, Edad, Sexo y CP
def propagate_coverages(group):
    # Propagar las coberturas generales
    for col in coverage_columns:
        if group[col].iloc[0] == 1:
            group[col] = 1

    # Manejar la cobertura CETTE y asignar valores en la fila del titular
    for i, row in enumerate(group.itertuples()):
        # CETTE
        if pd.notna(getattr(row, 'CETTE')):
            if i == 0:
                group.at[row.Index, 'cetteTit'] = 'ON'
            elif i < len(cette_columns):  # Solo manejar hasta el número máximo de asegurados
                group.at[group.index[0], cette_columns[i]] = 'ON'
        # Riesgo a Preferente
        if pd.notna(getattr(row, 'Riesgo')):
            if i == 0:
                group.at[row.Index, 'Preferente'] = 1
            if i > 0 and i < len(preferente_columns) + 1:
                group.at[group.index[0], preferente_columns[i - 1]] = 1

        # Edad
        if pd.notna(getattr(row, 'Edad')):
            current_year = datetime.now().year
            birth_year = current_year - getattr(row, 'Edad')
            birth_date = datetime(int(birth_year), 1, 1)
            if i > 0:
                # print(birth_date.strftime('%d/%m/%Y'))
                fecha_aseg = birth_date.strftime('%d/%m/%Y')
                group.at[group.index[0], f'Fecha_Nacimiento_A{i}'] = "'"+fecha_aseg
            elif i == 0:
                fecha_tit = "'" + datetime(int(birth_year), 1, 1).strftime('%d-%m-%Y')
                group.at[group.index[0], 'FechaNacimiento'] = fecha_tit  # Asignar fecha de nacimiento al titular

        # Sexo a Genero
        if getattr(row, 'Sexo') == 'M':
            value = 1
        elif getattr(row, 'Sexo') == 'F':
            value = 2
        else:
            value = 0
        group.at[group.index[0], f'Genero_A{i}' if i > 0 else 'Genero'] = value

        # Asignar codigo postal basado en la zona
        zona = getattr(row, 'Zona')
        group.at[group.index[0], 'CP'] = codigos_postales.get(zona, 'Desconocido')

        # Asigna Estado basado en la zona
        group.at[group.index[0], 'Estado'] = estados.get(zona, 'Desconocido')

        # Calcular el número de asegurados (excluyendo al titular)
        group['Asegurados'] = len(group) - 1  # Asignar el número de asegurados a cada fila del grupo

        # Parentesco a Rol
        if pd.notna(getattr(row, 'Parentesco')):
            if i > 0:  # Asegurar que no estamos en el titular
                group.at[group.index[0], f'Rol_A{i}'] = getattr(row, 'Parentesco')

        # Asignar el valor de DeducibleUnico basado en Tipo de Deducible
        # if getattr(row, 'Tipo de Deducible') == 'UNI':
            # group.at[group.index[0], 'DeducibleUnico'] = 1
        # elif getattr(row, 'Tipo de Deducible') == 'ANU':
            # group.at[group.index[0], 'DeducibleUnico'] = 0

        # Asignar el valor de Red_Copago basado en la columna CRFCA
        group.at[group.index[0], 'Red_Copago'] = 1 if pd.notna(getattr(row, 'CRFCA')) else 0

        # Asignar el valor de DP basado en la columna DENTAL
        group.at[group.index[0], 'DP'] = 1 if pd.notna(getattr(row, 'DENTAL')) else 0

        # Asignar el valor de CEDA PREM basado en la columna DENTAL
        group.at[group.index[0], 'CEDA PREM'] = 1 if pd.notna(getattr(row, 'CEDAP')) else 0

        # Asignar el valor de cash basado en la columna AMCD
        cash_value = 'ON' if pd.notna(getattr(row, 'AMCD')) else 'OFF'
        if i == 0:
            group.at[row.Index, 'cashTit'] = cash_value
        else:
            group.at[group.index[0], f'cash{i}'] = cash_value

    return group


# Aplicar la funcion de consolidación
grouped_data = data.groupby('Nombre_foler_evidencia').apply(propagate_coverages)


""" Se eliminan las columnas ya procesadas que ya no son necesarias """
grouped_data.pop('Sexo')
grouped_data.pop('Zona')
grouped_data.pop('Riesgo')
grouped_data.pop('Nombre ')
grouped_data.pop('Parentesco')
# grouped_data.pop('Tipo de Deducible')
grouped_data.pop('CRFCA')
grouped_data.pop('DENTAL')
grouped_data.pop('CEDA PREM')
grouped_data.pop('CETTE')
grouped_data.pop('AMCD')
grouped_data.pop('Pma + Der')
grouped_data.pop('Forma de pago')

# Eliminar duplicados basados solo en la columna ID
grouped_data = grouped_data.drop_duplicates(subset=['Nombre_foler_evidencia'], keep='first')

name_file = 'matriz.xlsx'

# Exportar dataframe a excel
grouped_data.to_excel(name_file, index=False, engine='openpyxl')
print(f'Data exportada exitosamente a {name_file}')
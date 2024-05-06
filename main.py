import pandas as pd
from pathlib import Path

directorio_actual = Path.cwd()
file_path = directorio_actual/'Matiz Emision Pruebas 2024 v2_NUEVA.xlsx'

data = pd.read_excel(file_path, sheet_name='Sheet2')

# Seleccionar solo las columnas relevantes para el procesamiento
columns_to_process = [
    'Unnamed: 1',  # ID
    'Unnamed: 2',  # Nombre
    'Unnamed: 3',  # Parentesco
    'Unnamed: 4',  # Riesgo
    'Unnamed: 5',  # Edad
    'Unnamed: 6',  # Sexo
    'DATOS DE COBERTURA BASICA',  # Plan
    'Unnamed: 11',  # Zona
    'Unnamed: 12',  # Suma Asegurada
    'Unnamed: 13',  # Deducible
    'Unnamed: 14',  # Coaseguro
    'Unnamed: 15',  # Incremento GURA
    'Unnamed: 16',  # Tipo de Póliza
    'Unnamed: 17',  # Tipo de Deducible
    'Unnamed: 18',  # Forma de pago
    'FACT_CPF',  # CPF
    'FACT_CAE',  # CAE
    'FACT_PLAN_CEC',  # CEC
    'Unnamed: 51',  # CEE
    'Unnamed: 52',  # CEDA
    'Unnamed: 53',  # Dental
    'Unnamed: 54',  # AMCD
    'Unnamed: 55',  # CEDA PREM
    'Unnamed: 56',  # CRFCA
    'Unnamed: 57'  # CETTE (Asumido como CETTE)
]
filtered_data = data[columns_to_process]

filtered_data = filtered_data.rename(columns={
    'Unnamed: 1':'ID',
    'Unnamed: 2':'Nombre',
    'Unnamed: 3':'Parentesco',
    'Unnamed: 4':'Preferente',
    'Unnamed: 5':'Edad',
    'Unnamed: 6':'Genero',
    'DATOS DE COBERTURA BASICA':'Plan',
    'Unnamed: 11':'Zona',
    'Unnamed: 12':'Suma Aseg',
    'Unnamed: 13':'Deducible',
    'Unnamed: 14':'Coaseguro',
    'Unnamed: 15':'CHMQ',
    'Unnamed: 16':'TipoPoliza',
    'Unnamed: 17':'DeducibleUnico',
    'Unnamed: 18':'Frecuencia',
    'FACT_CPF':'CPF',
    'FACT_CAE':'CAE',
    'FACT_PLAN_CEC':'CEC',
    'Unnamed: 51':'CEE',
    'Unnamed: 52':'CEDA',
    'Unnamed: 53':'DP',
    'Unnamed: 54':'AMCD',
    'Unnamed: 55':'CEDAP',
    'Unnamed: 56':'Red_Copago',
    'Unnamed: 57':'CETTE'
})

coverage_columns = ['CPF', 'CAE', 'CEC', 'CEE', 'CEDA', 'DP', 'AMCD', 'CEDAP', 'Red_Copago']  # Incluyendo CEE


# Crear columnas para CETTE para titular y hasta nueve asegurados
cette_columns = ['Tit_cette'] + [f'cette{i}' for i in range(1, 12)]

# Inicializar las nuevas columnas en el DataFrame filtrado
for col in cette_columns:
    filtered_data[col] = 0  # Iniciar todas las columnas CETTE con 0


# Definir una función para procesar cada grupo de ID
# Asumir que los nombres de columnas ya están correctamente etiquetados
coverage_columns = ['CPF', 'CAE', 'CEC', 'CEE', 'CEDA', 'Dental', 'AMCD', 'CEDA PREM', 'CRFCA']

# Crear columnas para CETTE para titular y hasta 11 asegurados
cette_columns = ['Tit_cette'] + [f'cette{i}' for i in range(1, 12)]
for col in cette_columns:
    data[col] = 0  # Iniciar todas las columnas CETTE con 0

# Función para propagar coberturas y manejar CETTE
def propagate_coverages(group):
    # Propagar las coberturas generales
    for col in coverage_columns:
        if group[col].iloc[0] == 1:
            group[col] = 1

    # Manejar la cobertura CETTE y asignar valores en la fila del titular
    for i, row in enumerate(group.itertuples()):
        if pd.notna(getattr(row, 'CETTE')):
            if i == 0:
                group.at[row.Index, 'Tit_cette'] = 1
            elif i < len(cette_columns):  # Solo manejar hasta el número máximo de asegurados
                group.at[group.index[0], cette_columns[i]] = 1

    return group

# Aplicar la función de consolidación
consolidated_data = data.groupby('ID').apply(propagate_coverages)

# Mostrar los resultados consolidados
print(consolidated_with_cette.head(20))  # Mostramos las primeras 20 filas para verificar las transformaciones

name_file = 'matriz.xlsx'

# Exportar dataframe a excel
consolidated_with_cette.to_excel(name_file, index=False, engine='openpyxl')

print(f'Data exportada exitosamente a {name_file}')



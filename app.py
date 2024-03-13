from flask import Flask, render_template, request
import pandas as pd
import os
import openpyxl

app = Flask(__name__, static_url_path='/static')

# Lógica de búsqueda por listado de rutas
def buscar_usrnoms_por_rutas(df, rutas):
    resultados = []

    for ruta in rutas:
        resultado_ruta = df[df['RutaNombre'].str.strip().str.upper() == ruta][['RutaNombre', 'UsrNom', 'UsrPersona']].drop_duplicates(subset='UsrNom')
        resultado_ruta['RutaNombre'] = ruta
        resultados.append(resultado_ruta)

    df_resultados = pd.concat(resultados)

    return df_resultados

# Define las rutas de tu aplicación
@app.route('/')
def index():
    try:
        archivo_excel = './SistemaConsultas_RutasLectores.xlsx'
        df = pd.read_excel(archivo_excel)

        # Obtener los valores únicos de la columna FacGrNombre
        opciones_facgrnombre = df['FacGrNombre'].unique().tolist()

        return render_template('index.html', opciones_facgrnombre=opciones_facgrnombre)
    except Exception as e:
        return render_template('error.html', message=f"Error inesperado: {str(e)}")

@app.route('/buscar', methods=['POST'])
def buscar():
    try:
        archivo_excel = './SistemaConsultas_RutasLectores.xlsx'
        df = pd.read_excel(archivo_excel)

        if 'FacGrNombre' in request.form and 'RutaNombre' in request.form:
            # Si se proporcionaron FacGrNombre y RutaNombre en el formulario
            fac_gr_nombre = request.form['FacGrNombre']
            ruta_nombre = request.form['RutaNombre']
            resultado = df[(df['FacGrNombre'] == fac_gr_nombre) & (df['RutaNombre'] == ruta_nombre)][['RutaNombre', 'UsrNom', 'UsrPersona']].drop_duplicates(subset='UsrNom')
        elif 'rutas' in request.form:
            # Si se proporcionó un listado de rutas en el formulario
            rutas = [ruta.strip().upper() for ruta in request.form['rutas'].split(',') if ruta.strip()]
            resultado = buscar_usrnoms_por_rutas(df, rutas)
        else:
            return render_template('error.html', message="Error: No se proporcionaron datos de búsqueda válidos")

        if resultado.empty:
            return render_template('sin_resultados.html')

        resultado = resultado.pivot(index='RutaNombre', columns='UsrNom', values='UsrPersona').fillna('')
        resultado = resultado.reset_index().rename_axis(None, axis=1)

        return render_template('resultado.html', resultados=resultado.to_dict(orient='records'))
    except Exception as e:
        return render_template('error.html', message=f"Error inesperado: {str(e)}")

if __name__ == '__main__':
    app.run(debug=True, port=5001)

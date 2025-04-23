import io
import pandas as pd
import streamlit as st

# Función para cargar el archivo Excel
def load_excel(uploaded_file):
    return pd.ExcelFile(uploaded_file)

# Función para extraer las tablas desde el archivo Excel
def extract_tables_from_excel(xls):
    tables = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        tables[sheet_name] = df
    return tables

# Función para aplicar la interpolación cúbica
def apply_cubic_spline_to_dict(dataframes, points):
    # Implementa tu lógica de interpolación aquí
    # En este ejemplo, solo hago una interpolación simple para ilustrar
    results = {}
    for name, df in dataframes.items():
        # Supón que la columna 'Plazo (días)' contiene los plazos
        if "Plazo (días)" in df.columns:
            df_interp = df.copy()
            df_interp["Plazo (días)"] = points  # Asignar nuevos puntos de interpolación
            results[name] = df_interp  # Aquí se debe aplicar la lógica de interpolación
    return results

# Función para crear el DataFrame
def create_dataframe(results):
    headers = [
        "Forward", "Plazo (días)", "Posición Compañía", "Moneda negociada", "Moneda extranjera [A]", "Moneda local [B]",
        "TC Contractual", "TC Spot", "Nominal en [A]", "Nominal en [B]", "Tasa de descuento [A]", "Tasa de descuento [B]",
        "Factor de descuento [A]", "Factor de descuento [B]", "VPN Posición spot (PEN)", "VPN Posición fwd (PEN)",
        "PD Compañía", "PD Contraparte", "Ajuste DVA (PEN)", "Ajuste CVA (PEN)", "Ajuste DVA/CVA neto (PEN)",
        "MTM ajustado (PEN)"
    ]
    
    df_kpmg = pd.DataFrame(columns=headers)
    
    # Extraer valores de los dataframes que contienen 'contrato' en su key
    for key, df in results.items():
        if 'Contrato' in key:
            for header in ["Forward", "Plazo (días)", "Posición Compañía", "Moneda negociada", "Moneda extranjera [A]", 
                           "Moneda local [B]", "TC Contractual", "Nominal en [A]", "Nominal en [B]"]:
                if header in df.columns:
                    df_kpmg[header] = df[header]
    
    # Asignar valores a la columna 'TC Spot'
    for i, row in df_kpmg.iterrows():
        moneda_extranjera = row["Moneda extranjera [A]"]
        moneda_local = row["Moneda local [B]"]
        
        for key, df in results.items():
            if 'TC_USDEUR' in key or 'TC_USDPEN' in key:
                for col in df.columns:
                    if moneda_extranjera in col and moneda_local in col:
                        tc_spot_value = df[col].min()
                        df_kpmg.at[i, 'TC Spot'] = tc_spot_value
    
    # Agregar el nuevo dataframe al diccionario results
    results["KPMG Fwd Método 2"] = df_kpmg
    
    return results

# Función para generar el diccionario de tasas
def generar_diccionario_tasas(results, BBG):
    df_kpmg = results["KPMG Fwd Método 2"]
    plazos_kpmg = df_kpmg["Plazo (días)"].unique()
    
    # Crear un nuevo dataframe para almacenar los resultados
    df_tasas_bbg = pd.DataFrame()
    
    # Iterar sobre cada plazo en el dataframe KPMG
    for plazo in plazos_kpmg:
        for key, df in BBG.items():
            if "Plazo (días)" in df.columns:
                filas_encontradas = df[df["Plazo (días)"] == plazo]
                if not filas_encontradas.empty:
                    for index, row in filas_encontradas.iterrows():
                        nueva_fila = row.to_dict()
                        nueva_fila["BBG"] = key
                        df_tasas_bbg = pd.concat([df_tasas_bbg, pd.DataFrame([nueva_fila])], ignore_index=True)
    
    # Reordenar las columnas para que 'BBG' sea la primera
    columnas = ["BBG"] + [col for col in df_tasas_bbg.columns if col != "BBG"]
    df_tasas_bbg = df_tasas_bbg[columnas]
    
    # Agregar el nuevo dataframe al diccionario results
    results["Tasas_BBG_por_dia"] = df_tasas_bbg
    
    return results

# Función de VLOOKUP para seleccionar tasas de descuento
def vlookup_tasas(results):
    df_kpmg = results["KPMG Fwd Método 2"]
    df_tasas_bbg = results["Tasas_BBG_por_dia"]
    
    def seleccionar_fila(plazo, columna_objetivo):
        filas = df_tasas_bbg[df_tasas_bbg["Plazo (días)"] == plazo]
        opciones_filas = filas["BBG"].unique()
        
        print(f"Seleccione una fila para el plazo de días {plazo} en la {columna_objetivo}:")
        for i, opcion in enumerate(opciones_filas):
            print(f"{i + 1}. {opcion}")
        
        seleccion_fila = int(input(f"Ingrese el número de la opción deseada para el plazo de días {plazo} en la {columna_objetivo}: ")) - 1
        fila_seleccionada = opciones_filas[seleccion_fila]
        
        return filas[filas["BBG"] == fila_seleccionada].iloc[0]
    
    def seleccionar_columna(fila):
        opciones_columnas = [col for col in fila.index if col != "BBG" and col != "Plazo (días)"]
        
        print("Seleccione una columna:")
        for i, opcion in enumerate(opciones_columnas):
            print(f"{i + 1}. {opcion}")
        
        seleccion_columna = int(input("Ingrese el número de la opción deseada: ")) - 1
        columna_seleccionada = opciones_columnas[seleccion_columna]
        
        return columna_seleccionada
    
    for i, row in df_kpmg.iterrows():
        plazo = row["Plazo (días)"]
        
        # Asignar valor a "Tasa de descuento [A]"
        fila_seleccionada_A = seleccionar_fila(plazo, "Tasa de descuento [A]")
        columna_seleccionada_A = seleccionar_columna(fila_seleccionada_A)
        tasa_A = fila_seleccionada_A[columna_seleccionada_A] / 100
        df_kpmg.at[i, "Tasa de descuento [A]"] = tasa_A
        
        # Asignar valor a "Tasa de descuento [B]"
        fila_seleccionada_B = seleccionar_fila(plazo, "Tasa de descuento [B]")
        columna_seleccionada_B = seleccionar_columna(fila_seleccionada_B)
        tasa_B = fila_seleccionada_B[columna_seleccionada_B] / 100
        df_kpmg.at[i, "Tasa de descuento [B]"] = tasa_B
    
    results["KPMG Fwd Método 2"] = df_kpmg
    
    return results

# Función para exportar los resultados a Excel
def export_results_to_excel(results_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in results_dict.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], startrow=8, index=False)
        writer.close()
    output.seek(0)
    return output

# Interfaz Streamlit
st.title("Análisis de Forward y Tasas de Descuento")

uploaded_file = st.file_uploader("Carga tu archivo Excel de Bloomberg", type=["xlsx"])

if uploaded_file:
    xls = load_excel(uploaded_file)
    dataframes = extract_tables_from_excel(xls)
    
    st.subheader("Tablas extraídas")
    selected_table = st.selectbox("Selecciona una tabla para visualizar", list(dataframes.keys()))
    st.dataframe(dataframes[selected_table], use_container_width=True)

    if st.button("Aplicar interpolación cúbica"):
        points = list(range(1, 12001))
        interpolated_results = apply_cubic_spline_to_dict(dataframes, points)

        st.subheader("Resultados Interpolados")
        for name, df in interpolated_results.items():
            st.write(f"📊 {name}")
            st.dataframe(df.head(10), use_container_width=True)
    
    # Procesar resultados
    results = create_dataframe(dataframes)
    results = generar_diccionario_tasas(results, dataframes)
    results = vlookup_tasas(results)
    
    # Descargar el archivo Excel de resultados
    excel_data = export_results_to_excel(results)
    st.download_button(
        label="📥 Descargar resultados interpolados (Excel)",
        data=excel_data,
        file_name="resultados_interpolados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")


import streamlit as st
import pdfplumber
import pandas as pd
from numpy import nan
from io import BytesIO
from datetime import datetime, timezone, timedelta

def main():
    st.title("PDF a Excel")

    if 'uploaded_file' not in st.session_state:
        st.session_state.uploaded_file = None
    
    uploaded_file = st.file_uploader("Elegi un PDF", type="pdf")
    
    if uploaded_file is not None:
        st.session_state.uploaded_file = uploaded_file
        st.write("PDF cargado con exito!")

    if st.session_state.uploaded_file is not None:
        if st.button("Que empiece la fiesta"):
            pdf_name = uploaded_file.name.strip(".pdf")

            all_data = extract_data(uploaded_file)
            df_t1 = process_data(all_data)
            df_final= add_categories(df_t1) # con nueva columna
        
            output_file = save_file(df_final)

            # set timezone to ARG
            AR_hour = -3
            AR_offset = timedelta(hours=AR_hour)
            AR_timezone = timezone(AR_offset)
            now = datetime.now()
            AR_now = now.astimezone(AR_timezone)

            AR_now_final = AR_now.strftime("%d-%m-%Y_%Hh%Mm%Ss")

            # format file
            file_name = f"{pdf_name}-{AR_now_final}.xlsx"
        
            st.write("PDF procesado con exito!")
            st.write(df_final)
            st.download_button(
                label="Descargar Excel",
                data=output_file,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def extract_data(pdf_path) -> list:
    """
    Extracts raw data from pdf, outputs a list
    """

    all_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if len(table) > 0:
                    for row in table:
                        cleaned_row = [cell.strip() if cell else '' for cell in row]
                        all_data.append(cleaned_row)
    return all_data

def process_data(all_data: list) -> pd.DataFrame:
    """
    Agarra los datos crudos, pone las columnas correctas, parsea bien los nombres 
    y datos para luego ser correctamente casteados a float, luego llama a la fun
    `checkear_y_asignar` para mover los datos desfazados y limpiar las filas
    """

    df = pd.DataFrame(all_data, columns=["temp1", "temp2", "temp3", "temp4"])
    df.columns = df.iloc[0]
    df = df.drop(0).reset_index(drop=True)
    df.columns = [name.replace("\n", " ") for name in df.columns]
    df2 = df.copy()
    df2["PRECIO PACK"] = df2["PRECIO PACK"].astype(str).str.replace("$", "").str.replace("\n", "").str.replace(".", "").str.replace(",", ".")
    df2.loc[df2["PRECIO PACK"] == "", "PRECIO PACK"] = nan
    df2["PRECIO PACK"] = pd.to_numeric(df2["PRECIO PACK"], errors='coerce')
    df2["PRECIO UNITARIO"] = df2["PRECIO UNITARIO"].astype(str).str.replace("$ ", "").str.replace(".", "").str.replace(",", ".")
    df2.loc[df2["PRECIO UNITARIO"] == "", "PRECIO UNITARIO"] = nan
    df2["PRECIO UNITARIO"] = pd.to_numeric(df2["PRECIO UNITARIO"], errors='coerce')
    df2 = df2[~df2.iloc[:, 0].str.contains('ARTÍCULO', na=False)].reset_index(drop=True)

    df2_1 = df2.copy()
    df2_1 = checkear_y_asignar(df2_1)

    return df2_1

def checkear_y_asignar(df: pd.DataFrame) -> pd.DataFrame:
    """
    Busca en la columna articulo si la fila esta vacia o no y agarra el valor
    de esa fila en la columna PRECIO PACK para pasarlo a la fila de arriba
    donde deberia estar
    luego llama la funcion `more_processing` para limpiar la fila vacia
    """

    for index in df.index.to_list():
        if (pd.isna(df.at[index, 'ARTÍCULO']) or df.at[index, 'ARTÍCULO'] == ''):
            if pd.isna(df.at[index-1, 'PRECIO PACK']) or df.at[index-1, 'PRECIO PACK'] == '':
                df.at[index-1, 'PRECIO PACK'] = df.at[index, 'PRECIO PACK']

    # drop empty since they values were transfered

    df_f = more_processing(df)

    return df_f

def more_processing(df_t1: pd.DataFrame) -> pd.DataFrame:
    """
    Limpia la fila con ARTICULO vacio ya que su valor fue extraido y ordenado
    en la funcion `checkear_y_asignar`
    """

    df_t2 = df_t1.copy()
    df_t2 = df_t2.drop(df_t2[df_t2["ARTÍCULO"] == ""].index)
    df_final = df_t2.reset_index(drop=True)
    return df_final

def add_categories(df_final: pd.DataFrame) -> pd.DataFrame:
    """
    Checkea que en la fila las columnas PRECIO PACK O CANT X PACK sean o nulas o vacias
    de ser asi, agarra el articulo, si no esta vacia agrega ese articulo previamente
    guardado y lo agrega como fila en la nueva columna Categoria
    Una vez encontradas todas las categorias borra las filas que las contenian
    en ultima instancia pone bien los data types a las columnas
    """

    dff = df_final.copy()
    # Inicializar la columna 
    dff['Categoria'] = None
    # guardar la categoria que encontramos sola
    current_category = None
    # Iterar las filas y asignar la categoria que encontremos sola, de lo contrario agregarla a la nueva columna
    for index, row in dff.iterrows():
        if (pd.isnull(row['PRECIO PACK']) or row['PRECIO PACK'] == "") and (pd.isnull(row['CANT. X PACK']) or row['CANT. X PACK'] == ""):
            current_category = row['ARTÍCULO']
        else:
            dff.at[index, 'Categoria'] = current_category   ## using df.at since only requires 1 specific row to assing
                                                           ## since df.iloc is more suitable for grouping rows
    # Dropear las columnas de las categorias ya encontradas
    dff = dff.dropna(subset=['PRECIO PACK', 'CANT. X PACK'])
    # Reordenamos
    dff = dff[['Categoria', 'ARTÍCULO', 'PRECIO PACK', 'CANT. X PACK', 'PRECIO UNITARIO']]
    # Reseteamos index
    dff2 = dff.reset_index(drop=True)
    # Make sure of types
    dff2['PRECIO PACK'] = pd.to_numeric(dff2['PRECIO PACK'])
    dff2['PRECIO UNITARIO'] = pd.to_numeric(dff2['PRECIO UNITARIO'])
    dff2['CANT. X PACK'] = dff2['CANT. X PACK'].astype(int)

    dff2.head()

    return dff2

def save_file(df_final: pd.DataFrame) -> BytesIO:
    """
    Writes the data in memory instead of saving in locally
    slighly styles the Excel adding spacing to columns
    """
    output = BytesIO()
    # writes to the bytes object
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Bebidas')
        
        workbook = writer.book
        worksheet = writer.sheets['Bebidas']
        
        for i, col in enumerate(df_final.columns):
            max_len = max(df_final[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)
        
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        for i, col in enumerate(df_final.columns):
            worksheet.write(0, i, col, header_format)
    # reset the pointer
    output.seek(0)
    return output

if __name__ == "__main__":
    main()

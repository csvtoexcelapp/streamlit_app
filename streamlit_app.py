import streamlit as st
import pandas as pd

import zipfile
import uuid
import io
import os

from PIL import Image

#----------------------------------------------------------#

st.markdown('''
            <style> .font {
            font-size: 45px;
            color: #008000;
            text-align: left;}
            </style>
            ''', unsafe_allow_html=True)

st.sidebar.markdown(
    '<p class="font">Conversor de Excel a CSV</p>', unsafe_allow_html=True)

logo_image = Image.open('imgs/excel_img.jpg')
st.sidebar.image(logo_image, use_container_width=True)

with st.sidebar.expander('**¿Cómo funciona?**', expanded=True):
    st.write('''
            Este convertidor toma archivos de Excel y los convierte directamente a archivos CSV sin modificar las columnas ni las filas de ninguna manera. 
            Asegúrese de que los archivos de entrada sean archivos de Excel de una sola hoja, no de varias hojas, y que en la primera fila del excel se encuentre
            el encabezado de las columnas. Asegúrese de utilizar únicamente archivos con formato .xlsx o .xls. Usar cualquier otro tipo de archivo generará un error.
            ''')

st.markdown('<p class="font">Subir archivo Excel</p>', unsafe_allow_html=True)

# The `label` argument cannot be an empty value. This is discouraged for accessibility reasons
# and may be disallowed in the future by raising an exception. A string is provided but it is hidden
# using `label_visibility` instead to prevent any future issues.
file_upload = st.file_uploader('Upload File(s)', type=['xlsx', 'xls'], accept_multiple_files=True,
                               label_visibility='hidden')

if file_upload is not None:
    file_list = []  # Create a list to store file data and name
    for file in file_upload:
        # Check if the uploaded file is an accepted excel file
        if file.type not in ['application/vnd.ms-excel',
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
            st.error(
                f'This file type is invalid: {file.name} must be an .xlsx or .xls file to proceed.')
        else:
            # Procesar el archivo cargado directamente desde el objeto file
            try:
                df_excel = pd.read_excel(file)
                st.success(f'Archivo procesado exitosamente: {file.name}')
                st.write(df_excel.head(5))  # Muestra el dataframe procesado
            except Exception as e:
                st.error(f'No se pudo procesar el archivo {file.name}: {str(e)}')
            csv = df_excel.to_csv(index=False)
            df_csv = pd.read_csv(io.StringIO(csv))
            # Split the file name from the file extension to make the download button distinctive
            file_name_no_ext = os.path.splitext(file.name)[0]

            # Check the shape of the original file and csv file to ensure it has the same
            # number of rows and columns
            if df_excel.shape != df_csv.shape:
                st.error(f'''
                         El CSV generado {file.name} no es idéntico al archivo original. 
                         Insepecciona tu archivo Excel si contiene celdas con formato o combinadas, fórmulas, caracteres especiales, etc,
                         e intenta nuevamente.
                         ''')
            else:
                file_list.append((csv, file_name_no_ext + '.csv'))
                # If there's only one file, show the regular download button
                st.download_button(
                    label=f'Download \'{file_name_no_ext}\' as CSV',
                    data=csv,
                    file_name=file_name_no_ext + '.csv',
                    mime='text/csv',
                    key=str(uuid.uuid4())) # Generate random key to avoid 'DuplicateWidgetID' error

    # When more than one file is uploaded, add an option to download them all as zip
    if len(file_list) > 1:
        zipped_csvs = io.BytesIO()
        with zipfile.ZipFile(zipped_csvs, mode='w') as zip_file:
            for file_data, file_name in file_list:
                zip_file.writestr(file_name, file_data)
        st.download_button(
            label='Descarga todos los archivos como zip',
            data=zipped_csvs.getvalue(),
            file_name='CSV_convertidos.zip',
            mime='application/zip',
            key=str(uuid.uuid4()))

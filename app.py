
import os
import re
import unicodedata
import pandas as pd
from datetime import datetime
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(layout="wide")
st.title("Conciliación Bancaria LRV")

banco_file = st.file_uploader("Archivo de BANCO (.xlsx)", type=["xlsx"])
sistema_file = st.file_uploader("Archivo del SISTEMA (.xlsx)", type=["xlsx"])
transferencias_file = st.file_uploader("Archivo de TRANSFERENCIAS (.xlsx)", type=["xlsx"])
base_output_path = st.text_input("Ruta base donde guardar resultados (ej: ./resultados)", value="./resultados")

if banco_file and sistema_file and transferencias_file and base_output_path:
    if st.button("Ejecutar conciliación"):  
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_dir = os.path.join(base_output_path, timestamp)
        os.makedirs(output_dir, exist_ok=True)


        ## Procesamiento de Movimientos Bancarios -- Fuente Banco
        import re
        import unicodedata
        import pandas as pd
        from datetime import datetime
        import os

        # Leer archivo .xls (ajusta la ruta local)
        df_raw = pd.read_excel("banco_jul.xlsx",header=None,dtype={"N.DOC.": str}, 
                                        engine="openpyxl")

        # Helper: chequear si toda la fila está vacía
        def is_row_empty(row):
            return row.isnull().all()

        # 1. Encontrar índice de "RESUMEN DE CHEQUES"
        cheques_title_idx = df_raw[df_raw.apply(lambda row: row.astype(str).str.upper().str.contains("RESUMEN DE CHEQUES").any(), axis=1)].index[0]

        # 2. Encabezado de cheques está en la fila siguiente
        cheques_header_idx = cheques_title_idx + 1
        cheques_headers = df_raw.iloc[cheques_header_idx].dropna().tolist()

        # 3. Tomar las filas siguientes hasta una fila completamente vacía
        cheques_data = []
        for idx in range(cheques_header_idx + 1, len(df_raw)):
            row = df_raw.iloc[idx]
            if is_row_empty(row):
                break
            cheques_data.append(row.tolist())

        # 4. Crear DataFrame de cheques reorganizado
        df_cheques_raw = pd.DataFrame(cheques_data)
        df_cheques = pd.DataFrame()

        # Combinar columnas en bloques de 3
        for i in range(0, len(cheques_headers), 3):
            subset = df_cheques_raw.iloc[:, i:i+3]
            subset.columns = ['CHEQUE', 'FECHA', 'MONTO']
            df_cheques = pd.concat([df_cheques, subset], ignore_index=True)


        # 5. Buscar índice de "DETALLE DE MOVIMIENTOS"
        mov_title_idx = df_raw[df_raw.apply(lambda row: row.astype(str).str.upper().str.contains("DETALLE DE MOVIMIENTOS").any(), axis=1)].index[0]

        # 6. Encabezado está en la siguiente fila
        mov_header_idx = mov_title_idx + 1
        mov_headers = df_raw.iloc[mov_header_idx].tolist()

        # 7. Tomar datos desde la siguiente línea hasta una vacía
        mov_data = []
        for idx in range(mov_header_idx + 1, len(df_raw)):
            row = df_raw.iloc[idx]
            if is_row_empty(row):
                break
            mov_data.append(row.tolist())

        # 8. Crear DataFrame de movimientos
        df_movimientos = pd.DataFrame(mov_data, columns=mov_headers)

        # 9. Limpieza final
        df_cheques.dropna(how='all', inplace=True)
        df_movimientos.dropna(how='all', inplace=True)
        df_cheques.reset_index(drop=True, inplace=True)
        df_movimientos.reset_index(drop=True, inplace=True)

        #fecha que se haga string
        df_cheques['FECHA'] = df_cheques['FECHA'].astype(str)
        df_cheques['CHEQUE'] = df_cheques['CHEQUE'].astype('Int64').astype(str)

        df_movimientos['FECHA']=pd.to_datetime(df_movimientos['FECHA'])
        df_cheques['FECHA']=pd.to_datetime(df_cheques['FECHA'])

        # --- Normalizar MONTO como decimal (conserva centavos) ---
        df_cheques['MONTO'] = pd.to_numeric(df_cheques['MONTO'], errors='coerce').astype('float64')

        # Lista de columnas numéricas
        cols_numericas = ['DEBITO', 'CREDITO', 'SALDO']

        # Convertir a número decimal y reemplazar NaN por 0.0
        for col in cols_numericas:
            df_movimientos[col] = (
                pd.to_numeric(df_movimientos[col], errors='coerce')  
                .fillna(0.0)                                         
                .astype('float64')                                   
            )

        # Eliminar col con nans
        df_movimientos.dropna(axis=1, how='all', inplace=True)
        df_movimientos.drop(columns=['OFIC.','SALDO'],inplace=True)


        df_movimientos.rename(columns={
            'DEBITO': 'EGRESO',
            'CREDITO': 'INGRESO',
        }, inplace=True)

        df_movimientos = df_movimientos[['FECHA', 'N.DOC.', 'DESCRIPCION', 'INGRESO', 'EGRESO']]

        import pandas as pd
        import re
        import unicodedata

        # --- Leer archivos ---
        df_transferencias = pd.read_excel(transferencias_file, dtype={"NUMERO DE DOCUMENTO": str}, engine="openpyxl")
        df_transferencias = df_transferencias.rename(columns={
            ' NUMERO DE DOCUMENTO': 'NUMERO DE DOCUMENTO',
            ' DESCRIPCION': 'DESCRIPCION'
        })
        df_transferencias['NUMERO DE DOCUMENTO'] = df_transferencias['NUMERO DE DOCUMENTO'].astype(str)

        def clean_float_to_str(x):
            if pd.isna(x):
                return ''
            try:
                return str(int(float(x)))
            except:
                return str(x)

        df_transferencias['NUMERO DE DOCUMENTO'] = df_transferencias['NUMERO DE DOCUMENTO'].apply(clean_float_to_str)

        df_transferencias['NUMERO DE DOCUMENTO'] = (
            df_transferencias['NUMERO DE DOCUMENTO']
            .str.normalize('NFKC')
            .str.strip()
            .str.replace(r'[\.\-\/\s]', '', regex=True)
        )


        df_movimientos['N.DOC.'] = df_movimientos['N.DOC.'].astype(str)

        df_movimientos['N.DOC.'] = (
            df_movimientos['N.DOC.']
            .astype('string')
            .str.normalize('NFKC')
            .str.strip()
            .str.replace(r'[\.\-\/\s]', '', regex=True)
            .str.replace(r'\D', '', regex=True)
        )


        # --- Palabras basura ---
        palabras_basura = [
            'TRANSFERENCIA', 'CHEQUE', 'DEPOSITO', 'DEP', 'REMESA',
            'COBRO', 'PAGO', 'RECIBIDO', 'INTERBANCARIA', 'INTERNET', 'TRF', 'EFECTIVIZADO'
        ]

        regex_basura = re.compile(
            r'\b(' + '|'.join(palabras_basura) + r')\b(\s*(DIRECTA|INTERBANCARIA)?\s*(DE\s|A\s)?)?',
            flags=re.IGNORECASE
        )

        # --- Funciones auxiliares ---
        def normalizar_texto(texto):
            if pd.isna(texto): return ""
            texto = str(texto).upper()
            texto = "".join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))
            texto = re.sub(r'[^A-Z\s]', '', texto)
            texto = re.sub(r'\s+', ' ', texto).strip()
            return texto

        def limpiar_texto(texto):
            texto = normalizar_texto(texto)
            return regex_basura.sub('', texto).strip()

        def contiene_patron_rodriguez_villam(texto):
            texto = normalizar_texto(texto)
            return ('RODRIGUEZ' in texto and 'VILLAM' in texto) or ('VILLAM' in texto and 'RODRIGUEZ' in texto)

        # --- Merge para obtener DESCRIPCION del TRF si coincide N.DOC. ---
        df_enriquecido = df_movimientos.merge(
            df_transferencias[['NUMERO DE DOCUMENTO', 'DESCRIPCION']],
            how='left',
            left_on='N.DOC.',
            right_on='NUMERO DE DOCUMENTO'
        ).rename(columns={'DESCRIPCION_y': 'DESCRIPCION_TRF', 'DESCRIPCION_x': 'DESCRIPCION'})

        print(
            "Coincidencias únicas:",
            len(set(df_movimientos['N.DOC.']) & set(df_transferencias['NUMERO DE DOCUMENTO']))
        )

        # --- Lógica principal para asignar beneficiario ---
        def asignar_beneficiario(row):
            ingreso = row.get('INGRESO', 0)
            egreso = row.get('EGRESO', 0)
            tiene_match = pd.notna(row.get('DESCRIPCION_TRF'))

            if tiene_match:
                desc_limpia = limpiar_texto(row['DESCRIPCION_TRF'])
                if contiene_patron_rodriguez_villam(desc_limpia):
                    return 'REVISAR TRF LRV'
                return desc_limpia

            descripcion_banco = row.get('DESCRIPCION', '')
            desc_norm = normalizar_texto(descripcion_banco)

            if ingreso > 0:
                contiene_basura = any(p in desc_norm for p in palabras_basura)
                if not contiene_basura:
                    return descripcion_banco.strip()
                else:
                    return 'REVISAR'

            if egreso > 0:
                return ''

            return 'REVISAR'  # Fallback lógico si no es ni ingreso ni egreso

        # --- Aplicar al DataFrame ---
        df_enriquecido['BENEFICIARIO'] = df_enriquecido.apply(asignar_beneficiario, axis=1)

        df_enriquecido.drop(columns=['NUMERO DE DOCUMENTO', 'DESCRIPCION_TRF'], errors='ignore', inplace=True)
        df_movimientos = df_enriquecido

        ## Procesamiento de Movimientos Bancarios -- Fuente Sistema

        df_sistema = pd.read_excel(sistema_file, skiprows=5)
        columns_gen = df_sistema.columns.to_list()
        for col in columns_gen:
            if col.startswith('Unnamed:'):
                df_sistema.drop(columns=[col], inplace=True)

        df_sistema.drop(columns=['F.Cheque','Tp'], inplace=True)

        def es_fila_vacía(row):
            return all(pd.isna(x) or x == 0 for x in row)

        # Detectar primera fila con más de 5 NaN
        fila_corte = df_sistema[df_sistema.isnull().sum(axis=1) > 5].index.min()

        # Cortar hasta esa fila si existe
        if pd.notna(fila_corte):
            df_sistema = df_sistema.iloc[:fila_corte].copy()

        df_sistema['Fecha'] = pd.to_datetime(df_sistema['Fecha'], dayfirst=True, errors='coerce')

        df_sistema.rename(columns={
            'Débito': 'INGRESO',
            'Crédito': 'EGRESO',
            'Fecha':'FECHA'}, inplace=True)

        # Col CREDITO mult por -1
        df_sistema['EGRESO'] = df_sistema['EGRESO'].apply(lambda x: -x if pd.notna(x) else x)

        # Crear nombre del directorio con timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_dir = timestamp

        # Crear el directorio si no existe
        os.makedirs(timestamp, exist_ok=True)

        # Guardar los archivos dentro del directorio
        df_cheques.to_excel(f'{output_dir}/banco_cheques_final.xlsx', index=False)
        df_movimientos.to_excel(f'{output_dir}/banco_movimientos_final.xlsx', index=False)
        df_sistema.to_excel(f'{output_dir}/sistema_movimientos_final.xlsx', index=False)


        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill

        # ----------------------------
        # 1. Validación de estructura
        # ----------------------------
        def validar_movimientos(df, nombre):
            errores = []
            for idx, row in df.iterrows():
                deb, cred = row['INGRESO'], row['EGRESO']
                if pd.isna(deb) and pd.isna(cred):
                    errores.append((idx, "Ambos vacíos"))
                elif deb != 0 and cred != 0:
                    errores.append((idx, "Ambos llenos"))
            if errores:
                print(f"⚠️ Errores en {nombre}:")
                for idx, tipo in errores:
                    print(f" - Fila {idx}: {tipo}")
            else:
                print(f"✅ {nombre}: validación de DÉBITO/CRÉDITO OK")

        validar_movimientos(df_movimientos, 'Bancos')
        validar_movimientos(df_sistema, 'Sistema')

        # ----------------------------
        # 2. Preparar DataFrames
        # ----------------------------
        def preparar_df(df):
            df = df.copy()
            df['TIPO'] = df.apply(lambda x: 'EGRESO' if x['EGRESO'] != 0 else 'INGRESO', axis=1)
            df['MONTO'] = df.apply(lambda x: x['EGRESO'] if x['TIPO'] == 'EGRESO' else x['INGRESO'], axis=1)
            df['CLAVE'] = df['FECHA'].astype(str) + '|' + df['TIPO'] + '|' + df['MONTO'].round(2).astype(str)
            return df

        df_movimientos_preparado = preparar_df(df_movimientos)
        df_sistema_preparado = preparar_df(df_sistema)

        # ----------------------------
        # 3. Conciliación por clave
        # ----------------------------
        claves_banco = set(df_movimientos_preparado['CLAVE'])
        claves_sistema = set(df_sistema_preparado['CLAVE'])

        faltan_en_sistema = claves_banco - claves_sistema
        faltan_en_banco = claves_sistema - claves_banco

        df_faltantes_sistema = df_movimientos_preparado[df_movimientos_preparado['CLAVE'].isin(faltan_en_sistema)].copy()
        df_faltantes_banco = df_sistema_preparado[df_sistema_preparado['CLAVE'].isin(faltan_en_banco)].copy()

        # ----------------------------
        # 4. Marcar coincidencias
        # ----------------------------
        df_movimientos_preparado['MATCH'] = df_movimientos_preparado['CLAVE'].isin(claves_sistema)
        df_sistema_preparado['MATCH'] = df_sistema_preparado['CLAVE'].isin(claves_banco)

        # ----------------------------
        # 5. Exportar a Excel con color
        # ----------------------------
        archivo_salida = f'{output_dir}/conciliacion_resultado.xlsx'

        df_movimientos_preparado.drop(columns=['TIPO','MONTO','CLAVE'],inplace=True)
        df_sistema_preparado.drop(columns=['TIPO','MONTO','CLAVE'],inplace=True)

        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            df_movimientos_preparado.to_excel(writer, sheet_name="BANCOS", index=False)
            df_sistema_preparado.to_excel(writer, sheet_name="SISTEMA", index=False)

        # Pintar filas con coincidencias en amarillo
        amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        wb = load_workbook(archivo_salida)

        for hoja in ["BANCOS", "SISTEMA"]:
            ws = wb[hoja]
            col_match = ws.max_column  # última columna es MATCH
            for row in range(2, ws.max_row + 1):
                if ws.cell(row=row, column=col_match).value == True:
                    for col in range(1, col_match):  # pintar toda la fila excepto MATCH
                        ws.cell(row=row, column=col).fill = amarillo

            # Eliminar columna MATCH del archivo final
            ws.delete_cols(col_match)
        
        wb.save(archivo_salida)

        # ----------------------------
        # 6. Reporte final en consola
        # ----------------------------
        print(f"\n🔍 Faltantes en sistema: {len(df_faltantes_sistema)}")
        print(f"🔍 Faltantes en banco: {len(df_faltantes_banco)}")
        print(f"✅ Archivo generado: {archivo_salida}")
        st.success("✅ Conciliación exitosa. Archivos guardados en: " + output_dir)

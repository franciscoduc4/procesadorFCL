# %%streamlit_app.py
import streamlit as st
import pandas as pd
import os
import re
import io
from contextlib import contextmanager, redirect_stdout
from io import StringIO
from one_spain import obtener_df_archivo, map_data_to_template


@contextmanager
def st_capture(output_func):
    with StringIO() as stdout, redirect_stdout(stdout):
        old_write = stdout.write
        
        def new_write(string):
            ret = old_write(string)
            output_func(stdout.getvalue())
            return ret
        
        stdout.write = new_write
        yield

def main():
    st.set_page_config(page_title="Procesador FCL ONE España", page_icon="🚢")
    
    st.title("🚢 Procesador FCL ONE España")
    st.markdown("---")
    
    uploaded_files = st.file_uploader("Selecciona archivos Excel de tarifas ONE", 
                                    type=["xlsx"], 
                                    accept_multiple_files=True)
    
    if st.button("Procesar Archivos"):
        if not uploaded_files:
            st.warning("Por favor sube al menos un archivo")
            return
            
        processed_files = {}
        log_output = st.empty()
        full_logs = []
        
        with st.spinner("Procesando archivos..."):
            for uploaded_file in uploaded_files:
                try:
                    file_log = []
                    
                    with st_capture(lambda text: file_log.append(text)):
                        print(f"\n=== Procesando {uploaded_file.name} ===")
                        
                        # Procesar el archivo
                        xls = pd.ExcelFile(uploaded_file)
                        hojas = xls.sheet_names
                        df_freights, df_surcharges = obtener_df_archivo(uploaded_file, hojas)
                        result = map_data_to_template(df_freights, df_surcharges)
                        
                        # Guardar en bytes
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            result.to_excel(writer, index=False)
                        processed_files[uploaded_file.name] = output.getvalue()
                        
                        print(f"✅ Procesado exitoso: {uploaded_file.name}")
                    
                    full_logs.extend(file_log)
                
                except Exception as e:
                    error_msg = f"❌ Error procesando {uploaded_file.name}: {str(e)}"
                    print(error_msg)
                    full_logs.append(error_msg)
        
        # Mostrar logs
        log_text = "\n".join(full_logs)
        st.markdown("### Registro de Procesamiento")
        st.markdown(f'<div style="background-color: #333; color: #fff; padding: 10px; border-radius: 5px; font-family: monospace; white-space: pre-wrap;">{log_text}</div>', 
                   unsafe_allow_html=True)
        
        # Descargas
        st.markdown("---")
        st.markdown("### Descargar Archivos Procesados")
        
        cols = st.columns(2)
        for i, (filename, data) in enumerate(processed_files.items()):
            btn = cols[i % 2].download_button(
                label=f"Descargar {filename.replace('.xlsx', '_procesado.xlsx')}",
                data=data,
                file_name=f"{os.path.splitext(filename)[0]}_procesado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
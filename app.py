# %%streamlit_app.py
import streamlit as st
import pandas as pd
import os
import re
import io
from contextlib import contextmanager, redirect_stdout
from io import StringIO
from one_spain import obtener_df_archivo, map_data_to_template, process_arbitraries
import subprocess


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



def get_last_commit():
    """Obtiene el hash y mensaje del último commit en el repositorio Git."""
    try:
        commit_hash = subprocess.check_output(["git", "rev-parse", "--short", "HEAD"]).decode("utf-8").strip()
        commit_message = subprocess.check_output(["git", "log", "-1", "--pretty=%B"]).decode("utf-8").strip()
        commit_author = subprocess.check_output(["git", "log", "-1", "--pretty=%an"]).decode("utf-8").strip()
        commit_date = subprocess.check_output(["git", "log", "-1", "--pretty=%cd", "--date=short"]).decode("utf-8").strip()
        
        return f"📝 Último commit: `{commit_hash}`\n📅 Fecha: {commit_date}\n👤 Autor: {commit_author}\n💬 Mensaje: {commit_message}"
    except Exception as e:
        return f"⚠️ No se pudo obtener la información del commit: {e}"



def main():
    st.set_page_config(page_title="Procesador FCL ONE España", page_icon="🚢")

    st.title("🚢 Procesador FCL ONE España")
    st.markdown("---")
    st.markdown("### Última versión")
    st.markdown(get_last_commit())
    st.markdown("---")

    uploaded_files = st.file_uploader(
        "Selecciona archivos Excel de tarifas ONE",
        type=["xlsx"],
        accept_multiple_files=True,
    )

    # Crear contenedor para logs con tamaño fijo
    st.markdown("### Registro de Procesamiento")
    log_container = st.empty()
    
    # Inicializar variable de logs
    full_logs = []

    def update_logs(new_log):
        """Actualiza la visualización de logs con un tamaño fijo."""
        log_container.text_area(
            label="",
            value=new_log,
            height=300,  # Altura fija en píxeles
            max_chars=5000,  # Máximo de caracteres antes de truncar (opcional)
            disabled=True,
        )

    if st.button("Procesar Archivos"):
        if not uploaded_files:
            st.warning("Por favor sube al menos un archivo")
            return

        processed_files = {}

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
                        df_arbitraries = process_arbitraries(uploaded_file, hojas[5], search_range=(7, 15))

                        # Guardar en bytes
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            result.to_excel(writer, sheet_name="Tarifario", index=False)
                            df_arbitraries.to_excel(writer, sheet_name="Arbitraries", index=False)

                        processed_files[uploaded_file.name] = output.getvalue()

                        print(f"✅ Procesado exitoso: {uploaded_file.name}")

                    full_logs.extend(file_log)

                    # Actualizar logs dinámicamente mostrando las últimas 30 líneas
                    update_logs("\n".join(full_logs[-30:]))

                except Exception as e:
                    import traceback
                    error_details = traceback.format_exc()  # Capturar el stack trace completo

                    error_msg = f"❌ Error procesando {uploaded_file.name}:\n{error_details}"
                    print(error_msg)  # Mostrarlo en la consola de Streamlit

                    full_logs.append(error_msg)
                    update_logs("\n".join(full_logs[-30:]))  # Mantener últimas 30 líneas


        # Sección de descargas
        st.markdown("---")

        st.markdown("### Descargar Archivos Procesados")

        cols = st.columns(2)
        for i, (filename, data) in enumerate(processed_files.items()):
            cols[i % 2].download_button(
                label=f"Descargar {filename.replace('.xlsx', '_procesado.xlsx')}",
                data=data,
                file_name=f"{os.path.splitext(filename)[0]}_procesado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()

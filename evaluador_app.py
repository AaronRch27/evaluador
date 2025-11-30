import flet as ft
import pandas as pd
import openpyxl as op


def es_fecha(col):
    try:
        pd.to_datetime(col, dayfirst=True, errors="raise")
        return True
    except:
        return False


def write_dfs_to_excel(dfs, filename="salida.xlsx"):
    wb = op.Workbook()
    wb.remove(wb.active)

    for sheet_name, df in dfs.items():
        ws = wb.create_sheet(title=str(sheet_name))

        ws.append(df.columns.tolist())
        for row in df.itertuples(index=False, name=None):
            ws.append(row)

    wb.save(filename)




def procesar_archivo(ruta, peso_actividades, peso_examen, log_callback):

    try:
        log_callback("Leyendo archivo...")
        excel = pd.ExcelFile(ruta)

        grupos = {}
        faltas_pendientes = {}

        for sheet in excel.sheet_names:
            try:
                grupo = int(sheet)
                log_callback(f"Procesando grupo {sheet}")

                df = excel.parse(sheet_name=sheet)

                cols_lower = [str(col).lower() for col in df.columns]

                actividad_cols = [df.columns[i] for i, col in enumerate(cols_lower)
                                  if col.startswith('act')]
                examen_cols = [df.columns[i] for i, col in enumerate(cols_lower)
                               if col.startswith('examen')]
                acl = [k.lower() for k in actividad_cols] + [k.lower() for k in examen_cols]
                no_usadas = [k for k in df.columns if k.lower() not in acl]

                df['porcentaje_actividades'] = df[actividad_cols].mean(axis=1) * peso_actividades if actividad_cols else 0
                df['porcentaje_examen'] = df[examen_cols].mean(axis=1) * peso_examen if examen_cols else 0
                df['calificación'] = df['porcentaje_actividades'] + df['porcentaje_examen']
                
                grupos[grupo] = df
                log_callback(f"Grupo {grupo} procesado, y no se usaron las columnas {no_usadas}")

                # Si hay faltas pendientes, agregarlas
                if grupo in faltas_pendientes:
                    df_faltas = faltas_pendientes[grupo]
                    date_cols = [c for c in df_faltas.columns if es_fecha(c)]
                    grupos[grupo]['faltas'] = df_faltas[date_cols].sum(axis=1)
                    del faltas_pendientes[grupo]

            except:
                # Procesamiento de faltas
                log_callback(f"Procesando faltas para {sheet}")

                try:
                    parts = sheet.replace('_', ' ').split()
                    grupo_faltas = int(parts[-1])

                    df_faltas = excel.parse(sheet_name=sheet)

                    if grupo_faltas in grupos:
                        date_cols = [c for c in df_faltas.columns if es_fecha(c)]
                        grupos[grupo_faltas]['faltas'] = df_faltas[date_cols].sum(axis=1)
                        f_no_usadas = [k for k in df_faltas.columns if k not in date_cols]
                        log_callback(f"Faltas agregadas al grupo {grupo_faltas}, y no se usaron las columnas {f_no_usadas}")
                    else:
                        faltas_pendientes[grupo_faltas] = df_faltas
                        log_callback(f"Faltas pendientes para grupo {grupo_faltas}")

                except:
                    log_callback(f"No se pudo procesar faltas para hoja {sheet}")

        # Guardar archivo final
        salida = "evaluaciones.xlsx"
        write_dfs_to_excel(grupos, salida)
        log_callback(f"Archivo generado: {salida}")

    except Exception as e:
        log_callback(f"Error al procesar archivo: {e}")



def main(page: ft.Page):
    page.title = "Evaluador de Excel"
    page.window_width = 500
    page.window_height = 600

    salida = ft.Text("")

    def log(mensaje):
        salida.value += mensaje + "\n"
        page.update()

    file_picker = ft.FilePicker()
    ruta_archivo = ft.Text("Ningún archivo seleccionado")

    def on_file_selected(e):
        if file_picker.result is not None and file_picker.result.files:
            ruta_archivo.value = file_picker.result.files[0].path
            ruta_archivo.update()

    file_picker.on_result = on_file_selected
    page.overlay.append(file_picker)

    peso_act = ft.TextField(label="Peso actividades", value="0.4")
    peso_exam = ft.TextField(label="Peso examen", value="0.6")

    def ejecutar(e):
        salida.value = ""
        salida.update()

        if ruta_archivo.value == "Ningún archivo seleccionado":
            log("Debe seleccionar un archivo Excel.")
            return

        try:
            p_act = float(peso_act.value)
            p_ex = float(peso_exam.value)

            if abs(p_act + p_ex - 1.0) > 1e-6:
                log("Los pesos deben sumar 1.0")
                return

        except:
            log("Pesos inválidos, deben ser números.")
            return

        log("Iniciando procesamiento...")
        procesar_archivo(
            ruta_archivo.value,
            p_act,
            p_ex,
            log_callback=log
        )

    boton_ejecutar = ft.ElevatedButton("Procesar", on_click=ejecutar)

    page.add(
        ft.Text("Evaluador de Excel", size=24, weight=ft.FontWeight.BOLD),
        ft.Row([
            ft.Text("Archivo:"),
            ruta_archivo,
            ft.IconButton(
                icon=ft.Icons.UPLOAD_FILE,
                on_click=lambda _: file_picker.pick_files(allow_multiple=False)
            )
        ]),
        peso_act,
        peso_exam,
        boton_ejecutar,
        ft.Text("Salida:", size=18),
        ft.Container(
            salida,
            bgcolor=ft.Colors.BLACK12
        )
    )

# Ejecutar app Flet
ft.app(target=main)
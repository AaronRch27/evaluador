"""
script principal para el programa de evaluacion

"""

import pandas as pd
import openpyxl as op

def es_fecha(col):
    #print('entrada de fecha',col,type(col))
    try:
        pd.to_datetime(col, format="%dd/%mm/%YYYY")
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

excel = pd.ExcelFile('Listasprovisionales1.xlsx')
faltas = {}
grupos = {}
peso_actividades = 0.4
peso_examen = 0.6

for sheet in excel.sheet_names:
    try:
        #si funiciona es grupo
        grupo = int(sheet)
        print('grupo', sheet)
        
        df = excel.parse(sheet_name=sheet)
        
        cols_lower = [str(col).lower() for col in df.columns]
        
        actividad_cols = [df.columns[i] for i, col in enumerate(cols_lower) if col.startswith('actividad')]
        
        if actividad_cols:

            df['porcentaje_actividades'] = df[actividad_cols].mean(axis=1) * peso_actividades
        else:
            df['porcentaje_actividades'] = 0.0

        examen_cols = [df.columns[i] for i, col in enumerate(cols_lower) if col.startswith('examen')]
        
        if examen_cols:

            df['porcentaje_examen'] = df[examen_cols].mean(axis=1) * peso_examen
        else:
            df['porcentaje_examen'] = 0.0
        df['calificación'] = df['porcentaje_actividades'] + df['porcentaje_examen']

        print(f"Procesado grupo {grupo}: {len(actividad_cols)} actividades, {len(examen_cols)} examenes")
        # print(df[['porcentaje_actividades', 'porcentaje_examen']].head()) 
        
        grupos[grupo] = df

    except Exception as e:
        #si no funciona es lista de faltas
        print('Procesando faltas para:', sheet)
        try:
            parts = sheet.replace('_', ' ').split()
            grupo_faltas = int(parts[-1])
            
            if grupo_faltas in grupos:
                df_faltas = excel.parse(sheet_name=sheet)

                date_cols = [c for c in df_faltas.columns if es_fecha(c)]
                #print(sheet,date_cols)
                total_faltas = df_faltas[date_cols].sum(axis=1)
    
                grupos[grupo_faltas]['faltas'] = total_faltas
                print(f"Faltas agregadas al grupo {grupo_faltas}")
            else:
                print(f"Advertencia: Grupo {grupo_faltas} no encontrado en 'grupos' todavía.")
                
        except Exception as inner_e:
            print(f"No se pudo procesar faltas para {sheet}: {inner_e}")


write_dfs_to_excel(grupos, "evaluaciones.xlsx")
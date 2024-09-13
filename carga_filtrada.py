import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def abrir_archivo_carga():
    ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if ruta_archivo:
        try:
            carga = pd.read_excel(ruta_archivo, dtype=str)
            messagebox.showinfo("Archivo Carga Lectiva", f"Archivo cargado exitosamente:\n{ruta_archivo}")
            return carga
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")

def abrir_archivo_recurso():
    ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if ruta_archivo:
        try:
            asignaturas_recursos = pd.read_excel(ruta_archivo, 'LISTA DETALLADA ')
            messagebox.showinfo("Archivo Recurso Asignatura", f"Archivo cargado exitosamente:\n{ruta_archivo}")
            return asignaturas_recursos
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")

def procesar_datos(carga, asignaturas_recursos):
    # Realizar el merge de los datos
    cursos_recursos = pd.merge(carga, asignaturas_recursos, on=['COD_ASIGNATURA','MODALIDAD','CURSO'])

    # Limpiar la columna 'IDEN_LIGA'
    cursos_recursos['IDEN_LIGA'] = cursos_recursos['IDEN_LIGA'].astype(str).str.strip()
    
    # Filtrar y etiquetar según modalidad y 'IDEN_LIGA'
    def filtrar_modalidad(modalidad, tipo_un_horario, tipo_hora_practica, tipo_hora_teorica=None):
        modalidad_df = cursos_recursos[cursos_recursos['MODALIDAD'] == modalidad]
        un_horario_df = modalidad_df[(modalidad_df['IDEN_LIGA'] == '') | (modalidad_df['IDEN_LIGA'] == ' ')]
        hora_practica_df = modalidad_df[~modalidad_df['IDEN_LIGA'].isin(['', ' '])]
        un_horario_df['TIPO'] = tipo_un_horario
        hora_practica_df['TIPO'] = tipo_hora_practica
        if tipo_hora_teorica:
            hora_teorica_df = modalidad_df[~modalidad_df['IDEN_LIGA'].isin(['', ' '])]
            hora_teorica_df['TIPO'] = tipo_hora_teorica
            return un_horario_df, hora_practica_df, hora_teorica_df
        return un_horario_df, hora_practica_df
    
    presenciales = filtrar_modalidad('UC-PRESENCIAL', 'UN HORARIO', 'HORA PRÁCTICA')
    semipresenciales = filtrar_modalidad('UC-SEMIPRESENCIAL', 'UN HORARIO', 'HORA PRÁCTICA')
    distancias = filtrar_modalidad('UC-A DISTANCIA', 'UN HORARIO', 'HORA TEÓRICA')
    
    horarios = pd.concat(presenciales + semipresenciales + distancias)
    horarios.to_excel('Carga Filtrada.xlsx', index=False)
    
    return cursos_recursos

def procesar_asignaturas(cursos_recursos):
    # Filtrar los datos según condiciones específicas
    cursos_recursos_1 = cursos_recursos[cursos_recursos['HOR'] == '()SIN :(-Sin asignar)']
    cursos_recursos_2 = cursos_recursos[cursos_recursos['HOR'] == '(AVIR)SIN :(-Sin asignar)']
    cursos_recursos_avir = pd.concat([cursos_recursos, cursos_recursos_2, cursos_recursos_2]).drop_duplicates(keep=False)
    
    # Extraer datos de 'HOR' usando regex
    pattern = r'\((?:[^-]+-([^()]+))\)'
    matches = cursos_recursos['HOR'].str.extractall(pattern)
    matches = matches.reset_index().pivot(index='level_0', columns='match', values=0)
    matches.columns = [f'aula{i+1}' for i in matches.columns]
    
    cursos_recursos = cursos_recursos.join(matches)
    
    return cursos_recursos

def filtrar_software(cursos_recursos):
    software = [
        'ADOBE SUITE', 'ARCGIS', 'CADE SIMU', 'CIROS', 'DIRED-CAD', 'DLT-CAD', 
        'ERWIN DATA MODELER', 'FLEXSIM', 'FLUIDSIM - HIDRÁULICA', 'FLUIDSIM - NEUMÁTICA', 
        'MATLAB', 'MESHMIXER', 'PHYSIOEX', 'POWER FACTORY ', 'PROTEUS', 'PYCHARM 2023.3.3', 
        'PYTHON 3.12.1', 'SISCONT', 'SOLIDWORKS', 'TIA PORTAL', 'VENTSIM', 
        'VMWARE WORKSTATION', 'DOCTOC'
    ]
    
    lista_software = cursos_recursos[cursos_recursos['RECURSO'].isin(software)]
    lista_software = lista_software.drop_duplicates(subset=[
        'CAMPUS', 'MODALIDAD', 'COD_ASIGNATURA', 'CURSO', 'aula1', 'aula2', 'aula3', 'aula4', 'aula5', 'aula6'
    ])
    lista_software = lista_software[['CAMPUS', 'NRC', 'MODALIDAD', 'CURSO', 'RECURSO', 'aula1', 'aula2', 'aula3', 'aula4', 'aula5', 'aula6']]
    lista_software.to_excel('cursos_recursos.xlsx', index=False)

def ejecutar_procesamiento():
    carga = abrir_archivo_carga()
    if carga is not None:
        asignaturas_recursos = abrir_archivo_recurso()
        if asignaturas_recursos is not None:
            cursos_recursos = procesar_datos(carga, asignaturas_recursos)
            cursos_recursos = procesar_asignaturas(cursos_recursos)
            filtrar_software(cursos_recursos)
            messagebox.showinfo("Proceso Completo", "El procesamiento de datos se completó exitosamente.")

def main():
    ventana = tk.Tk()
    ventana.title("Seleccionar y Procesar Archivos")
    
    boton_ejecutar = tk.Button(ventana, text="Procesar Datos", command=ejecutar_procesamiento)
    boton_ejecutar.pack(pady=20)
    
    ventana.mainloop()

if __name__ == "__main__":
    main()

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np

class ExcelProcessorApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Procesador de Archivos Excel")
        self.master.geometry("400x600")
        
        self.carga_data = None
        self.recursos_data = None
        
        self.create_widgets()

    def create_widgets(self):
        self.btn_carga = tk.Button(self.master, text="Abrir Archivo de Carga", command=self.abrir_archivo_carga)
        self.btn_carga.pack(pady=10)
        
        self.text_carga = tk.Label(self.master, text="")
        self.text_carga.pack(pady=10)
        
        self.btn_recurso = tk.Button(self.master, text="Abrir Archivo de Recurso", command=self.abrir_archivo_recurso)
        self.btn_recurso.pack(pady=10)
        
        self.text_recurso = tk.Label(self.master, text="")
        self.text_recurso.pack(pady=10)
        
        self.btn_procesar = tk.Button(self.master, text="Procesar Archivos", command=self.procesar_archivos)
        # No empaquetamos btn_procesar aquí, se hará en validation_widgets

    def validation_widgets(self):
        self.btn_procesar.pack(pady=10)

    def abrir_archivo_carga(self):
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if ruta_archivo:
            try:
                self.carga_data = pd.read_excel(ruta_archivo, dtype=str)
                messagebox.showinfo("Éxito", f"Archivo de Carga Lectiva cargado:\n{ruta_archivo}")
                self.text_carga.config(text='Archivo Carga Cargado')
                if self.carga_data is not None and self.recursos_data is not None:
                    self.validation_widgets()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")

    def abrir_archivo_recurso(self):
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if ruta_archivo:
            try:
                sheet_name = 0  # Asume que queremos la primera hoja
                self.recursos_data = pd.read_excel(ruta_archivo, sheet_name=sheet_name)
                messagebox.showinfo("Éxito", f"Archivo de Recurso Asignatura cargado:\n{ruta_archivo}")
                self.text_recurso.config(text='Archivo Recurso Asignatura Cargado')
                if self.carga_data is not None and self.recursos_data is not None:
                    self.validation_widgets()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")

    def procesar_archivos(self):
        if self.carga_data is None or self.recursos_data is None:
            messagebox.showwarning("Advertencia", "Por favor, cargue ambos archivos primero.")
            return
        
        cursos_recursos = pd.merge(self.carga_data, self.recursos_data, on=['COD_ASIGNATURA', 'MODALIDAD', 'CURSO'])
        cursos_recursos['IDEN_LIGA'] = cursos_recursos['IDEN_LIGA'].astype(str).str.strip()
        cursos_recursos['IDEN_LIGA'] = cursos_recursos['IDEN_LIGA'].replace({'': 'NULO', ' ': 'NULO', 'NA': 'NULO'})
        
        self.filtrar_modalidad(cursos_recursos)
        messagebox.showinfo("Procesamiento Completo", "Los archivos han sido procesados.")

    def filtrar_modalidad(self, cursos_recursos):
        def es_numero(x):
            return x[-1].isdigit() if isinstance(x, str) and len(x) > 0 else False

        def procesar_modalidad(df, modalidad, tipo_practica, tipo_teoria):
            df_modalidad = df[df['MODALIDAD'] == modalidad].copy()
            # Asignar "MULTIPLE HORARIO" inicialmente
            df_modalidad['TIPO'] = 'MULTIPLE HORARIO'
            
            # Asignar tipo correcto a los casos no 'NULO'
            mask_no_nulo = df_modalidad['IDEN_LIGA'] != 'NULO'
            df_modalidad.loc[mask_no_nulo, 'TIPO'] = np.where(
                df_modalidad.loc[mask_no_nulo, 'IDEN_LIGA'].apply(es_numero), tipo_practica, tipo_teoria
            )
            
            # Asignar "UN HORARIO" a los casos 'NULO'
            df_modalidad.loc[~mask_no_nulo, 'TIPO'] = 'UN HORARIO'
            
            # Asegurar que los casos (-Sin asignar) sean siempre "PRÁCTICA" solo para modalidad "UC-A DISTANCIA"
            if modalidad == "UC-A DISTANCIA":
                mask_sin_asignar = df_modalidad['HOR'].str.contains('(AVIR)SIN :(-Sin asignar)', regex=False)
                df_modalidad.loc[mask_sin_asignar, 'TIPO'] = 'PRÁCTICA'
            
            return df_modalidad

        modalidades = {
            "UC-PRESENCIAL": ("PRÁCTICA", "TEORÍA"),
            "UC-SEMIPRESENCIAL": ("TEORÍA", "PRÁCTICA"),
            "UC-A DISTANCIA": ("PRÁCTICA", "TEORÍA")
        }

        resultados = []
        for modalidad, (tipo_practica, tipo_teoria) in modalidades.items():
            resultados.append(procesar_modalidad(cursos_recursos, modalidad, tipo_practica, tipo_teoria))

        horarios = pd.concat(resultados).drop_duplicates()
        horarios.to_excel('Carga_Filtrada.xlsx', index=False)

def main():
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import pandas as pd
import subprocess

# Configuración
# Variables globales para almacenar las colecciones
coleccion1 = ""
coleccion2 = ""
archivo_salida = r'D:\JEFERSON STUDY\JO-System\JO-System-v.1.0\Status.xlsx'

# Funciones para buscar PDFs
def conEntregableIdea(coleccion1):    
    referencias_con_entregable = []
    contador = 0
    try:
        for root, dirs, files in os.walk(coleccion1):
            for file in files:                              
                if file.endswith('.pdf') and 'consumo' in file.lower():
                    ultimo_nombre_carpeta2 = os.path.basename(root)
                    if ultimo_nombre_carpeta2.startswith("#"):
                        contador += 1  
                        referencias_con_entregable.append(ultimo_nombre_carpeta2)
                        break        
        df = pd.DataFrame(referencias_con_entregable, columns=['coleccion -Con Entregable'])
        largoList= len(referencias_con_entregable)
        print(f"Se encontraron {largoList} referencias con entregable")
        return df        

    except FileNotFoundError:
        messagebox.showerror("Error", f"No se encontró el directorio: {coleccion1}")
        return pd.DataFrame(columns=['Carpetas con PDF de Consumo']), pd.DataFrame(columns=['Carpetas con PDF de Consumo'])
    except PermissionError:
        messagebox.showerror("Error", f"No tienes permisos suficientes para acceder al directorio: {coleccion1}")
        return pd.DataFrame(columns=['Carpetas con PDF de Consumo']), pd.DataFrame(columns=['Carpetas con PDF de Consumo'])
    except Exception as e:
        messagebox.showerror("Error", f"Ha ocurrido un error inesperado: {e}")
        return pd.DataFrame(columns=['Carpetas con PDF de Consumo']), pd.DataFrame(columns=['Carpetas con PDF de Consumo'])


def sinEntregableIdea(coleccion1):    
    referencias_sin_entregable = []
    contador = 0    
    for root, dirs, files in os.walk(coleccion1):
        encontrado = False
        for file in files:
            if file.endswith('.pdf') and 'consumo' in file.lower():
                    contador += 1
                    encontrado = True
                    break
        if not encontrado:
            ultimo_nombre_carpeta = os.path.basename(root)
            if ultimo_nombre_carpeta.startswith("#"):
                referencias_sin_entregable.append(ultimo_nombre_carpeta)
    df = pd.DataFrame(referencias_sin_entregable, columns=['coleccion -Sin Entregable'])
    largoList= len(referencias_sin_entregable)
    print(f"Se encontraron {largoList} referencias sin entregable")
    return df


def conTrazo(coleccion1):   
    referencias_con_trazo = []
    contador = 0
    try:
        for root, dirs, files in os.walk(coleccion1):
            for file in files:
                if file.endswith('.amkx'):
                    ultimo_nombre_carpeta2 = os.path.basename(root)
                    if ultimo_nombre_carpeta2.startswith("#"):
                        contador += 1  
                        referencias_con_trazo.append(ultimo_nombre_carpeta2)
                        break  
        df = pd.DataFrame(referencias_con_trazo, columns=['coleccion - Con trazo'])
        largoList= len(referencias_con_trazo)
        print(f"Se encontraron {largoList} referencias con trazo")
        return df
    except FileNotFoundError:
        print(f"No se encontró el directorio: {coleccion1}")
    except PermissionError:
        print(f"No tienes permisos suficientes para acceder al directorio: {coleccion1}")
    except Exception as e:
        print(f"Ha ocurrido un error inesperado: {e}")
        return pd.DataFrame(columns=['Referencias RESORT-25 con Entregable(IDEA)']), pd.DataFrame(columns=['Referencias SWIM-25 con Entregable(IDEA)'])


def sinTrazo(coleccion1):
    referencias_sin_trazo = []
    contador = 0    
    for root, dirs, files in os.walk(coleccion1):
        encontrado = False
        for file in files:
            if file.endswith('.amkx'):
                contador += 1
                encontrado = True
                break
        if not encontrado:
            ultimo_nombre_carpeta = os.path.basename(root)
            if ultimo_nombre_carpeta.startswith("#"):
                referencias_sin_trazo.append(ultimo_nombre_carpeta) 
    df = pd.DataFrame(referencias_sin_trazo, columns=['coleccion - Sin Trazo'])
    largoList= len(referencias_sin_trazo)
    print(f"Se encontraron {largoList} referencias sin trazo")
    return df


# Función para seleccionar directorio con manejo de excepciones
def seleccionar_directorio(entry_widget, coleccion_var, color):    
    try:
        directorio = filedialog.askdirectory()
        if directorio:
            entry_widget.config(state='normal')
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, directorio)
            entry_widget.config(state='readonly')
            coleccion_var.set(directorio)
            lbl_status.config(text=f"Importación de colección exitosa: {directorio}", fg=color)
    except FileNotFoundError:
        print("No se encontró el directorio especificado.")
    except PermissionError:
        print("No tienes permisos suficientes para acceder a este directorio.")
    except Exception as e:
        print(f"Ha ocurrido un error inesperado: {e}")


# Función para buscar archivos
def buscar_archivos(entrada1, entrada2):
    entradaA = entrada1
    entradaB = entrada2
    print(f"Ha ocurrido un error inesperado: {entrada1}, {entrada2}")
    if not entradaA and entradaB:        
        lbl_status.config(text="Es necesario elegir una colección o carpeta.", fg="red")
    else:        
        lbl_status.config(text="buscando en: "+ coleccion1.get(), fg="green")
        ejecutar_programa(entradaA, entradaB)


def ejecutar_programa():    
    con_entregable = chk_state1.get()       # Variables de checkboxes
    sin_entregable = chk_state2.get()
    con_trazo = chk_state3.get()
    sin_trazo = chk_state4.get()
    print(f"Directorio seleccionado: ")
    
    dfs = []                                # Listas de DataFrames para combinar en Excel
    if con_entregable:
        df = conEntregableIdea(coleccion1)        
        dfs.append(df)
    if sin_entregable:
        df = sinEntregableIdea(coleccion1)
        dfs.append(df)
    if con_trazo:
        df = conTrazo(coleccion1)
        dfs.append(df)
    if sin_trazo:
        df = sinTrazo(coleccion1)
        dfs.append(df)
    
    if dfs:
        df_combined = pd.concat(dfs, axis=1)
        df_combined.to_excel(archivo_salida, index=False, sheet_name='Resultados')
        lbl_status.config(text=f"Status generado con éxito¡", fg="green")
    else:
        lbl_status.config(text="Seleccione al menos 1 opción.", fg="red")



"""_______________________________cambios en JO_______________________________________"""

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Status Explosion")
ventana.configure(bg="#D7DBDD")
ventana.geometry("500x600")

# Estilo
style = ttk.Style()
style.configure('TButton', font=('Helvetica', 12), padding=10)
style.configure('TCheckbutton', font=('Helvetica', 12))
style.configure('TLabel', font=('Helvetica', 12))

# Variables para almacenar las colecciones
coleccion1 = tk.StringVar()
coleccion2 = tk.StringVar()

# Frame para la selección de directorio
frame_directorio = tk.Frame(ventana, bg='#DDE2E6', padx=10, pady=10, relief='solid', bd=1)
frame_directorio.pack(padx=10, pady=10)

# Entradas para directorio1 y directorio2
entrada_directorio1 = tk.Entry(frame_directorio, width=50, state='normal')
entrada_directorio1.grid(row=0, column=1, padx=10, pady=(10, 5))
entrada_directorio2 = tk.Entry(frame_directorio, width=50, state='normal')
entrada_directorio2.grid(row=1, column=1, padx=10, pady=(5, 10))

# Botones para seleccionar directorio1 y directorio2
btn_seleccionar_directorio1 = tk.Button(frame_directorio, text="Seleccionar Colección 1", command=lambda: seleccionar_directorio(entrada_directorio1, coleccion1, "green"), bg='#1C6EA4', fg='white')
btn_seleccionar_directorio1.grid(row=0, column=0, pady=(10, 5))

btn_seleccionar_directorio2 = tk.Button(frame_directorio, text="Seleccionar Colección 2", command=lambda: seleccionar_directorio(entrada_directorio2, coleccion2, "blue"), bg='#1C6EA4', fg='white')
btn_seleccionar_directorio2.grid(row=1, column=0, pady=(5, 10))

# Crear un frame para los checkboxes
frame_checkboxes = tk.Frame(ventana, bg="#D7DBDD", padx=20, pady=20)
frame_checkboxes.pack(side='top', pady=20, padx=(20, 0), fill='y')
chk_state1 = tk.BooleanVar()
chk_state2 = tk.BooleanVar()
chk_state3 = tk.BooleanVar()
chk_state4 = tk.BooleanVar()

# Create a frame for the numeric value
frame_numeric = tk.Frame(ventana, bg="#dcdcdc", bd=1, relief="groove")
frame_numeric.pack(side='right', pady=20, padx=(0, 20), fill='y')

# Crear checkboxes
chk1 = tk.Checkbutton(frame_checkboxes, text="Con Entregable", var=chk_state1, bg="#D7DBDD")
chk1.pack(anchor="w")
chk2 = tk.Checkbutton(frame_checkboxes, text="Sin Entregable", var=chk_state2, bg="#D7DBDD")
chk2.pack(anchor="w")
chk3 = tk.Checkbutton(frame_checkboxes, text="Con Trazo", var=chk_state3, bg="#D7DBDD")
chk3.pack(anchor="w")
chk4 = tk.Checkbutton(frame_checkboxes, text="Sin Trazo", var=chk_state4, bg="#D7DBDD")
chk4.pack(anchor="w")

# Crear un frame para boton
frame_buttons = tk.Frame(ventana, bg="#dcdcdc", relief="groove")
frame_buttons.pack(pady=20, padx=20, fill='x')

# Botón para buscar archivos
btn_buscar = tk.Button(frame_buttons, text="Buscar", command=lambda:buscar_archivos(entrada_directorio1, entrada_directorio2 ), bg='#1C6EA4', fg='white')
btn_buscar.pack(pady=(10, 10))

# Crear un frame para estado
frame_message = ttk.Frame(ventana, padding=10, style='TFrame')
frame_message.pack(side=tk.TOP, padx=(10,10), pady=(10,10))

# Crear una etiqueta para mostrar el estado de la ejecución
lbl_status = tk.Label(text="- - - - - >   B I E N V E N I D O   < - - - - -", bg="#dddee6")
lbl_status.pack(side=tk.TOP, fill="x", pady=5)

# Ejecutar el bucle principal de la interfaz gráfica
ventana.mainloop()
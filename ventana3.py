import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import pandas as pd
import logging

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
archivo_salida = os.path.join(os.path.expanduser('~'), 'Desktop', 'Status.xlsx')


def conEntregableIdea(coleccion):    
    referencias_con_entregable = []
    try:
        for root, _, files in os.walk(coleccion):
            for file in files:                              
                if file.endswith('.pdf') and 'consumo' in file.lower():
                    ultimo_nombre_carpeta2 = os.path.basename(root)
                    if ultimo_nombre_carpeta2.startswith("#"):
                        referencias_con_entregable.append(ultimo_nombre_carpeta2)
                        break
        
        # Ordenar las referencias en orden numérico basado en el número después del '#'
        referencias_con_entregable.sort(key=lambda x: int(x.split('#')[1].split()[0]))

        # Crear el DataFrame con las referencias ordenadas
        df = pd.DataFrame(referencias_con_entregable, columns=['coleccion -Con Entregable'])
        logging.info(f"Se encontraron {len(referencias_con_entregable)} referencias con entregable")
        return df        
    except Exception as e:
        logging.error(f"Error en conEntregableIdea: {e}")
        return pd.DataFrame(columns=['coleccion -Con Entregable'])


def sinEntregableIdea(coleccion):    
    referencias_sin_entregable = []
    try:
        for root, _, files in os.walk(coleccion):
            if not any(file.endswith('.pdf') and 'consumo' in file.lower() for file in files):
                ultimo_nombre_carpeta = os.path.basename(root)
                if ultimo_nombre_carpeta.startswith("#"):
                    referencias_sin_entregable.append(ultimo_nombre_carpeta)

        # Ordenar las referencias en orden numérico basado en el número después del '#'
        referencias_sin_entregable.sort(key=lambda x: int(x.split('#')[1].split()[0]))   
        
        df = pd.DataFrame(referencias_sin_entregable, columns=['coleccion -Sin Entregable'])
        logging.info(f"Se encontraron {len(referencias_sin_entregable)} referencias sin entregable")
        return df
    except Exception as e:
        logging.error(f"Error en sinEntregableIdea: {e}")
        return pd.DataFrame(columns=['coleccion -Sin Entregable'])

def conTrazo(coleccion):   
    referencias_con_trazo = []
    try:
        for root, _, files in os.walk(coleccion):
            for file in files:
                if file.endswith('.amkx'):
                    ultimo_nombre_carpeta2 = os.path.basename(root)
                    if ultimo_nombre_carpeta2.startswith("#"):
                        referencias_con_trazo.append(ultimo_nombre_carpeta2)
                        break  
        # Ordenar las referencias en orden numérico basado en el número después del '#'
        referencias_con_trazo.sort(key=lambda x: int(x.split('#')[1].split()[0]))
                        
        df = pd.DataFrame(referencias_con_trazo, columns=['coleccion -Con trazo'])
        logging.info(f"Se encontraron {len(referencias_con_trazo)} referencias con trazo")
        return df
    except Exception as e:
        logging.error(f"Error en conTrazo: {e}")
        return pd.DataFrame(columns=['coleccion -Con trazo'])

def sinTrazo(coleccion):
    referencias_sin_trazo = []
    try:
        for root, _, files in os.walk(coleccion):
            if not any(file.endswith('.amkx') for file in files):
                ultimo_nombre_carpeta = os.path.basename(root)
                if ultimo_nombre_carpeta.startswith("#"):
                    referencias_sin_trazo.append(ultimo_nombre_carpeta)

        # Ordenar las referencias en orden numérico basado en el número después del '#'
        referencias_sin_trazo.sort(key=lambda x: int(x.split('#')[1].split()[0]))

        df = pd.DataFrame(referencias_sin_trazo, columns=['coleccion -Sin Trazo'])
        logging.info(f"Se encontraron {len(referencias_sin_trazo)} referencias sin trazo")
        return df
    except Exception as e:
        logging.error(f"Error en sinTrazo: {e}")
        return pd.DataFrame(columns=['coleccion -Sin Trazo'])

# Función para seleccionar directorio con manejo de excepciones
def seleccionar_directorio(entry_widget, color):    
    try:
        directorio = filedialog.askdirectory()
        if directorio:
            entry_widget.config(state='normal')
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, directorio)
            entry_widget.config(state='readonly')
            lbl_status.config(text=f"Importación de colección exitosa: {directorio}", fg=color)
    except Exception as e:
        logging.error(f"Error al seleccionar directorio: {e}")
        messagebox.showerror("Error", f"Ha ocurrido un error al seleccionar el directorio: {e}")


def buscar_archivos():
    entradaA = entrada_directorio1.get()
    entradaB = entrada_directorio2.get()
    
    if not entradaA and not entradaB:        
        lbl_status.config(text="Es necesario elegir al menos una colección o carpeta.", fg="red")
    else:
        if entradaA == entradaB:
            respuesta = messagebox.askyesno("Confirmar", "Las colecciones son iguales. ¿Deseas continuar?")
            if respuesta:
                #lbl_status.config(text="Buscando en las colecciones seleccionadas...", fg="green")
                ejecutar_programa(entradaA, entradaB)            
        else:
            #lbl_status.config(text="Buscando en las colecciones seleccionadas...", fg="green")
            ejecutar_programa(entradaA, entradaB)


def ejecutar_programa(coleccion1, coleccion2):    
    con_entregable = chk_state1.get()
    sin_entregable = chk_state2.get()
    con_trazo = chk_state3.get()
    sin_trazo = chk_state4.get()
    
    dfs1 = []
    for coleccionA in [coleccion1]:
        if coleccionA:
            if con_entregable:
                dfs1.append(conEntregableIdea(coleccionA))
            if sin_entregable:
                dfs1.append(sinEntregableIdea(coleccionA))
            if con_trazo:
                dfs1.append(conTrazo(coleccionA))
            if sin_trazo:
                dfs1.append(sinTrazo(coleccionA))       

    dfs2 = []
    for coleccionB in [coleccion2]:
        if coleccionB:
            if con_entregable:
                dfs2.append(conEntregableIdea(coleccionB))
            if sin_entregable:
                dfs2.append(sinEntregableIdea(coleccionB))
            if con_trazo:
                dfs2.append(conTrazo(coleccionB))
            if sin_trazo:
                dfs2.append(sinTrazo(coleccionB))

    if dfs1 and dfs2:
        df_combined1 = pd.concat(dfs1, axis=1)
        df_combined2 = pd.concat(dfs2, axis=1)
        # Crear un escritor de Excel para escribir múltiples hojas en un solo archivo
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            df_combined1.to_excel(writer, index=False, sheet_name='Coleccion 1')
            df_combined2.to_excel(writer, index=False, sheet_name='Coleccion 2')
        
        lbl_status.config(text=f"Status generado con éxito en {archivo_salida}", fg="green")
    else:
        lbl_status.config(text="Seleccione al menos 1 opción y una colección.", fg="red")


# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Status Explosion")
ventana.configure(bg="#D7DBDD") 
ventana.geometry("500x400")

# Estilo
style = ttk.Style()
style.configure('TButton', font=('Helvetica', 12), padding=10)
style.configure('TCheckbutton', font=('Helvetica', 12))
style.configure('TLabel', font=('Helvetica', 12))

# Frame para la selección de directorio
frame_directorio = tk.Frame(ventana, bg='#D7DBDD', padx=10, pady=10, relief='solid', bd=1) 
frame_directorio.pack(padx=10, pady=10)

# Entradas para directorio1 y directorio2
entrada_directorio1 = tk.Entry(frame_directorio, width=50, state='normal')
entrada_directorio1.grid(row=0, column=1, padx=10, pady=(10, 5))
entrada_directorio2 = tk.Entry(frame_directorio, width=50, state='normal')
entrada_directorio2.grid(row=1, column=1, padx=10, pady=(5, 10))

# Botones para seleccionar directorio1 y directorio2
btn_seleccionar_directorio1 = tk.Button(frame_directorio, text="Seleccionar Colección 1", 
                                        command=lambda: seleccionar_directorio(entrada_directorio1, "green"), 
                                        bg='#1C6EA4', fg='white')
btn_seleccionar_directorio1.grid(row=0, column=0, pady=(10, 5))

btn_seleccionar_directorio2 = tk.Button(frame_directorio, text="Seleccionar Colección 2", 
                                        command=lambda: seleccionar_directorio(entrada_directorio2, "blue"), 
                                        bg='#1C6EA4', fg='white')
btn_seleccionar_directorio2.grid(row=1, column=0, pady=(5, 10))

# Crear un frame para los checkboxes
frame_checkboxes = tk.Frame(ventana, bg="#D7DBDD", padx=10, pady=10)     #morado
frame_checkboxes.pack(side='top', pady=10, padx=(10, 0), fill='y')
chk_state1 = tk.BooleanVar()
chk_state2 = tk.BooleanVar()
chk_state3 = tk.BooleanVar()
chk_state4 = tk.BooleanVar()

# Cambio: Eliminado el frame numérico no utilizado

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
frame_buttons = tk.Frame(ventana, bg="#D7DBDD", relief="groove")
frame_buttons.pack(pady=10, padx=10, fill='x')

# Botón para buscar archivos
btn_buscar = tk.Button(frame_buttons, text="Buscar", command=buscar_archivos, bg='#1C6EA4', fg='white', width=20)
btn_buscar.pack(pady=(10, 10))

# Crear un frame para estado
frame_message = ttk.Frame(ventana, padding=10, style='TFrame')
frame_message.pack(side=tk.TOP, padx=(10,10))

# Crear una etiqueta para mostrar el estado de la ejecución
lbl_status = tk.Label(text="- - - - - >   B I E N V E N I D O   < - - - - -", bg="#dddee6")
lbl_status.pack(side=tk.TOP, fill="x", pady=5)

# Ejecutar el bucle principal de la interfaz gráfica
ventana.mainloop()

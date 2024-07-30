import os
import pandas as pd

# Configuración
resort = r'Y:\\1. JO EXPORT OFICINA TECNICA\\1. ESCALADOS JOHANNA ORTIZ\\2025\\RESORT 25\\RE RTW25\\PATRONES'
swim = r"Y:\\1. JO EXPORT OFICINA TECNICA\\1. ESCALADOS JOHANNA ORTIZ\\2025\\RESORT 25\\SUN 25\\PATRONES"
archivo_salida = r'D:\\JEFERSON STUDY\\Javascript - Jonmircha\\Buscar.xlsx'

def conEntregableIdea(resort, swim):
    referencias_con_entregable_resort = []
    referencias_con_entregable_swim = []
    
    try:
        for root, dirs, files in os.walk(resort):
            for file in files:
                if file.endswith('.pdf') and 'consumo' in file.lower():
                    ultimo_nombre_carpeta1 = os.path.basename(root)
                    referencias_con_entregable_resort.append(ultimo_nombre_carpeta1)
                    break
        for root, dirs, files in os.walk(swim):
            for file in files:
                if file.endswith('.pdf') and 'consumo' in file.lower():
                    ultimo_nombre_carpeta2 = os.path.basename(root)
                    referencias_con_entregable_swim.append(ultimo_nombre_carpeta2)
                    break

        df1 = pd.DataFrame(referencias_con_entregable_resort, columns= ['RESORT'])
        df2 = pd.DataFrame(referencias_con_entregable_swim, columns=['SWIM'])
        df_combined = pd.concat([df1,df2], axis=1)
        return df_combined

    except FileNotFoundError:
        print(f"No se encontró el directorio: {resort} o {swim}")
    except PermissionError:
        print(f"No tienes permisos suficientes para acceder al directorio: {resort} o {swim}")
    except Exception as e:
        print(f"Ha ocurrido un error inesperado: {e}")
        return pd.DataFrame(columns=['Referencias RESORT-25 con Entregable(IDEA)']), pd.DataFrame(columns=['Referencias SWIM-25 con Entregable(IDEA)'])



def sinEntregableIdea(resort, swim):
    referencias_sin_entregable_resort = []
    referencias_sin_entregable_swim = []
        
    for root, dirs, files in os.walk(resort):
        encontrado = False
        for file in files:
            if file.endswith('.pdf') and 'consumo' in file.lower():
                encontrado = True
                break
        if not encontrado:
            ultimo_nombre_carpeta1 = os.path.basename(root)
            referencias_sin_entregable_resort.append(ultimo_nombre_carpeta1)
    for root, dirs, files in os.walk(swim):
        encontrado = False
        for file in files:
            if file.endswith('.pdf') and 'consumo' in file.lower():
                encontrado = True
                break
        if not encontrado:
            ultimo_nombre_carpeta2 = os.path.basename(root)
            referencias_sin_entregable_swim.append(ultimo_nombre_carpeta2)

    df1 = pd.DataFrame(referencias_sin_entregable_resort, columns=['RESORT'])
    df2 = pd.DataFrame(referencias_sin_entregable_swim, columns=['SWIM'])
    df_combined = pd.concat([df1,df2], axis=1)
    return df_combined



def conTrazo(resort, swim):    
    referencias_con_trazo_resort = []
    referencias_con_trazo_swim = []

    try:
        for root, dirs, files in os.walk(resort):
            for file in files:
                if file.endswith('amkx'):
                    ultimo_nombre_carpeta1 = os.path.basename(root)
                    referencias_con_trazo_resort.append(ultimo_nombre_carpeta1)
                    break
        for root, dirs, files in os.walk(swim):
            for file in files:
                if file.endswith('.amkx'):
                    ultimo_nombre_carpeta2 = os.path.basename(root)
                    referencias_con_trazo_swim.append(ultimo_nombre_carpeta2)
                    break

        df1 = pd.DataFrame(referencias_con_trazo_resort, columns=['RESORT'])
        df2 = pd.DataFrame(referencias_con_trazo_swim, columns=['SWIM'])
        df_combined = pd.concat([df1,df2], axis=1)
        return df_combined

    except FileNotFoundError:
        print(f"No se encontró el directorio: {resort} o {swim}")
    except PermissionError:
        print(f"No tienes permisos suficientes para acceder al directorio: {resort} o {swim}")
    except Exception as e:
        print(f"Ha ocurrido un error inesperado: {e}")
        return pd.DataFrame(columns=['Referencias RESORT-25 con Entregable(IDEA)']), pd.DataFrame(columns=['Referencias SWIM-25 con Entregable(IDEA)'])



def sinTrazo(resort, swim):    
    referencias_sin_trazo_resort = []
    referencias_sin_trazo_swim = []
    
    for root, dirs, files in os.walk(resort):
        
        encontrado = False
        for file in files:
            if file.endswith('.amkx'):
                encontrado = True
                break
        if not encontrado:
            ultimo_nombre_carpeta1 = os.path.basename(root)
            referencias_sin_trazo_resort.append(ultimo_nombre_carpeta1)
    for root, dirs, files in os.walk(swim):
        encontrado = False
        for file in files:
            if file.endswith('.amkx'):
                encontrado = True
                break
        if not encontrado:
            ultimo_nombre_carpeta2 = os.path.basename(root)
            referencias_sin_trazo_swim.append(ultimo_nombre_carpeta2)
        
    df1 = pd.DataFrame(referencias_sin_trazo_resort, columns=['RESORT'])
    df2 = pd.DataFrame(referencias_sin_trazo_swim, columns=['SWIM'])
    df_combined = pd.concat([df1,df2], axis=1)
    return df_combined


# Ejecución
df_con_entregable_combined1 = conEntregableIdea(resort, swim)
df_con_entregable_combined2 = sinEntregableIdea(resort, swim)
df_con_entregable_combined3 = conTrazo(resort, swim)
df_con_entregable_combined4 = sinTrazo(resort, swim)

# Guardar en un archivo Excel con cuatro hojas
with pd.ExcelWriter(archivo_salida) as writer:
    df_con_entregable_combined1.to_excel(writer, sheet_name='Con_Entregable', index=False)
    df_con_entregable_combined2.to_excel(writer, sheet_name='Sin_Entregable', index=False)
    df_con_entregable_combined3.to_excel(writer, sheet_name='Con_Trazo', index=False)
    df_con_entregable_combined4.to_excel(writer, sheet_name='Sin_Trazo', index=False)
    
print(f'Se ha guardado el listado en {archivo_salida}')

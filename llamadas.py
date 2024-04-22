import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from pandas import ExcelWriter
from datetime import datetime

# Inicializar DataFrame
data_columns = ['Fecha de llamada', 'Cliente', 'Número de teléfono', 'Estado del artículo', 'Comentarios de la llamada', 'Respuesta del cliente', 'Prioridad de seguimiento', 'Recogido']
data_frame = pd.DataFrame(columns=data_columns)

# Función para cargar datos de un archivo Excel
def cargar_datos():
    global data_frame
    try:
        data_frame = pd.read_excel('datos_llamadas.xlsx', parse_dates=['Fecha de llamada'])
        data_frame.sort_values('Fecha de llamada', inplace=True)
        actualizar_vista()
    except Exception as e:
        messagebox.showerror("Error al cargar", str(e))

# Función para guardar datos en un archivo Excel
def guardar_datos():
    with ExcelWriter('datos_llamadas.xlsx') as writer:
        data_frame.to_excel(writer, index=False)
    messagebox.showinfo("Guardar", "Datos guardados exitosamente en Excel")

# Función para agregar datos
def agregar_dato():
    fecha = entry_fecha.get()
    cliente = entry_cliente.get()
    telefono = entry_telefono.get()
    estado = combo_estado.get()
    comentarios = entry_comentarios.get()
    respuesta = combo_respuesta.get()
    recogido = recogido_var.get()  # Estado del checkbox
    prioridad = calcular_prioridad(respuesta, fecha)
    datos_nuevos = [fecha, cliente, telefono, estado, comentarios, respuesta, prioridad, recogido]
    try:
        data_frame.loc[len(data_frame)] = datos_nuevos
        data_frame['Fecha de llamada'] = pd.to_datetime(data_frame['Fecha de llamada'])
        actualizar_vista()
        guardar_datos()
    except Exception as e:
        messagebox.showerror("Error al agregar", str(e))

# Calcular la prioridad en función de la respuesta y la fecha
def calcular_prioridad(respuesta, fecha):
    hoy = datetime.now().strftime('%Y-%m-%d')
    if respuesta == "Mañana voy por ella" and fecha == hoy:
        return "Alta"
    elif respuesta == "Ya me habían llamado":
        return "Media"
    elif respuesta == "No contesta":
        return "Baja"
    else:
        return "Normal"
    
# Borrar llamada seleccionada
def borrar_llamada():
    selected_items = tree.selection()
    if selected_items:
        for item in selected_items:
            data_frame.drop(index=int(item), inplace=True)
            tree.delete(item)
        guardar_datos()
    else:
        messagebox.showerror("Error", "No hay llamada seleccionada")

def verificar_telefono(event):
    telefono = entry_telefono.get()
    fecha_actual = datetime.now().strftime('%Y-%m-%d')
    # Filtrar DataFrame por número de teléfono
    matches = data_frame[data_frame['Número de teléfono'] == telefono]
    # Comprobar si hay coincidencias y cuándo fue la última llamada
    if not matches.empty:
        last_call_date = matches['Fecha de llamada'].dt.strftime('%Y-%m-%d').iloc[-1]
        if last_call_date == fecha_actual:
            messagebox.showwarning("Alerta de Llamada", "Ya llamaste a este número hoy.")
        else:
            messagebox.showinfo("Información de Llamada", "Última llamada a este número: " + last_call_date)        

# Actualizar vista de datos
def actualizar_vista():
    for row in tree.get_children():
        tree.delete(row)
    for idx, row in data_frame.iterrows():
        tree.insert("", 'end', iid=idx, values=list(row), tags=('prioridad' + row['Prioridad de seguimiento'],))

# Función para exportar registros recogidos a otro Excel
def exportar_recogidos():
    recogidos_df = data_frame[data_frame['Recogido'] == True]
    with ExcelWriter('recogidos.xlsx') as writer:
        recogidos_df.to_excel(writer, index=False)
    messagebox.showinfo("Exportar Recogidos", "Registros recogidos exportados exitosamente.")

# Función para cargar solo registros recogidos del archivo Excel
def cargar_recogidos():
    try:
        recogidos_df = pd.read_excel('recogidos.xlsx', parse_dates=['Fecha de llamada'])
        for row in tree.get_children():
            tree.delete(row)
        for idx, row in recogidos_df.iterrows():
            tree.insert("", 'end', iid=idx, values=list(row), tags=('prioridad' + row['Prioridad de seguimiento'],))
    except Exception as e:
        messagebox.showerror("Error al cargar recogidos", str(e))
# Función para marcar como recogido y exportar a otro Excel
def marcar_como_recogido():
    selected_items = tree.selection()
    if selected_items:
        recogidos_df = pd.DataFrame(columns=data_columns)  # DataFrame para recogidos
        for item in selected_items:
            index = int(item)
            row = data_frame.loc[index]
            recogidos_df = recogidos_df.append(row, ignore_index=True)
            data_frame.drop(index, inplace=True)
            tree.delete(item)
        # Exportar recogidos a otro Excel
        with ExcelWriter('recogidos.xlsx') as writer:
            recogidos_df.to_excel(writer, index=False)
        guardar_datos()
        actualizar_vista()
    else:
        messagebox.showerror("Error", "No hay llamada seleccionada")

# Configuración de la ventana principal
root = Tk()
root.title("Sistema de Gestión de Llamadas")

# Configuración del widget Treeview
tree = ttk.Treeview(root, columns=data_columns, show="headings")
for col in data_columns:
    tree.heading(col, text=col)
tree.pack(expand=True, fill='both')

# Configurar colores de prioridad
tree.tag_configure('prioridadAlta', background='red')
tree.tag_configure('prioridadMedia', background='yellow')
tree.tag_configure('prioridadBaja', background='lightgreen')
tree.tag_configure('prioridadNormal', background='white')

# Campos de entrada para nuevos datos
entry_fecha = Entry(root)
entry_fecha.insert(0, datetime.now().strftime('%Y-%m-%d'))
entry_cliente = Entry(root)
entry_telefono = Entry(root)
entry_telefono.bind("<FocusOut>", verificar_telefono)  # Enlazar evento para verificar el teléfono
combo_estado = ttk.Combobox(root, values=["Reparado", "No reparado"])
entry_comentarios = Entry(root)
combo_respuesta = ttk.Combobox(root, values=["Ya me habían llamado", "Mañana voy por ella", "No contesta", "Gracias!"])



entries = [entry_fecha, entry_cliente, entry_telefono, combo_estado, entry_comentarios, combo_respuesta]
labels = ['Fecha de llamada', 'Cliente', 'Número de teléfono', 'Estado del artículo', 'Comentarios de la llamada', 'Respuesta del cliente']
for i, (entry, label_text) in enumerate(zip(entries, labels), start=1):
    label = Label(root, text=label_text)
    label.pack()
    entry.pack(fill=X)

# Botones para acciones
btn_cargar = Button(root, text="Cargar Datos", command=cargar_datos)
btn_cargar.pack(fill=X)
btn_agregar = Button(root, text="Agregar Llamada", command=agregar_dato)
btn_agregar.pack(fill=X)
btn_borrar = Button(root, text="Borrar Llamada", command=borrar_llamada)
btn_borrar.pack(fill=X)
btn_cargar_recogidos = Button(root, text="Cargar Recogidos", command=cargar_recogidos)
btn_cargar_recogidos.pack(fill=X)
btn_exportar_recogidos = Button(root, text="Exportar Recogidos", command=exportar_recogidos)
btn_exportar_recogidos.pack(fill=X)
btn_marcar_recogido = Button(root, text="Marcar Como Recogido", command=marcar_como_recogido)
btn_marcar_recogido.pack(fill=X)



root.mainloop()

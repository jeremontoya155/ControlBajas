import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk, StringVar
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def load_excel():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            messagebox.showinfo("Carga exitosa", "Archivo Excel cargado correctamente.")
            load_filters()
            update_summary_table(df)
            plot_data(df)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo Excel.\n{e}")

def load_filters():
    try:
        unique_sucursales = df['sucursal'].unique().tolist()
        unique_motivos = df['motivo'].unique().tolist()
        unique_observaciones = df['Observacion'].unique().tolist()  # Nuevo: Filtro de observación
        
        if not unique_sucursales or not unique_motivos or not unique_observaciones:
            messagebox.showerror("Error", "No se encontraron datos en las columnas 'sucursal', 'motivo' u 'Observacion'.")
            return
        
        sucursal_filter.configure(values=["Seleccionar Sucursal"] + unique_sucursales)
        motivo_filter.configure(values=["Seleccionar Motivo"] + unique_motivos)
        observacion_filter.configure(values=["Seleccionar Observación"] + unique_observaciones)  # Nuevo: Combo de observación
    except KeyError as e:
        messagebox.showerror("Error", f"No se encontró la columna esperada en el archivo Excel.\n{e}")
    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar los filtros.\n{e}")

def filter_data():
    try:
        sucursal = sucursal_filter.get()
        motivo = motivo_filter.get()
        observacion = observacion_filter.get()  # Nuevo: Obtener filtro de observación
        
        filtered_df = df.copy()
        
        if sucursal and sucursal != "Seleccionar Sucursal":
            filtered_df = filtered_df[filtered_df['sucursal'] == sucursal]
        if motivo and motivo != "Seleccionar Motivo":
            filtered_df = filtered_df[filtered_df['motivo'] == motivo]
        if observacion and observacion != "Seleccionar Observación":  # Nuevo: Filtrar por observación
            filtered_df = filtered_df[filtered_df['Observacion'] == observacion]
        
        if filtered_df.empty:
            messagebox.showinfo("Sin datos", "No se encontraron datos con los filtros aplicados.")
        else:
            update_summary_table(filtered_df)
            plot_data(filtered_df)
    except Exception as e:
        messagebox.showerror("Error", f"Error al filtrar los datos.\n{e}")

def update_summary_table(filtered_df):
    clear_summary_table()
    
    total_label = total_var.get()  # Nuevo: Obtener el tipo de totalizado seleccionado
    
    if total_label == "Total Costo":
        total_column = 'TotalCosto'
    elif total_label == "Total Cantidad":
        total_column = 'Cantidad'
    else:
        total_column = 'TotalCosto'  # Por defecto, totalizar por costo
    
    for index, row in filtered_df.iterrows():
        # Redondear el total a entero
        total_entero = round(row[total_column])
        summary_tree.insert("", "end", values=(row['sucursal'], row['motivo'], row['Observacion'], row['producto'], total_entero))

    total_totalizado = round(filtered_df[total_column].sum())  # Totalizar según la columna seleccionada y redondear a entero
    summary_label.configure(text=f"{total_label}: {total_totalizado}")

def clear_summary_table():
    for row in summary_tree.get_children():
        summary_tree.delete(row)

def plot_data(filtered_df):
    fig, ax = plt.subplots(figsize=(8, 5))  # Reducido el tamaño del gráfico (ancho, alto)
    
    total_label = total_var.get()  # Obtener el tipo de totalizado seleccionado
    
    if total_label == "Total Costo":
        y_column = 'TotalCosto'
        title = 'Total Costo por Motivo'
        ylabel = 'Total Costo'
    elif total_label == "Total Cantidad":
        y_column = 'Cantidad'
        title = 'Total Cantidad por Motivo'
        ylabel = 'Total Cantidad'
    else:
        y_column = 'TotalCosto'
        title = 'Total Costo por Motivo'
        ylabel = 'Total Costo'
    
    # Redondear la columna y_column a entero
    filtered_df[y_column] = filtered_df[y_column].apply(round)
    
    filtered_df.groupby('motivo')[y_column].sum().plot(kind='bar', ax=ax)
    ax.set_title(title)
    ax.set_xlabel('Motivo')
    ax.set_ylabel(ylabel)
    ax.tick_params(axis='x', labelrotation=5, labelsize=6)  # Rotación y tamaño de etiquetas en el eje X
    
    for widget in plot_frame.winfo_children():
        widget.destroy()
    
    canvas = FigureCanvasTkAgg(fig, master=plot_frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

def export_to_excel(filtered_df):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            filtered_df.to_excel(file_path, index=False)
            messagebox.showinfo("Exportación exitosa", "Datos exportados correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar el archivo Excel.\n{e}")

def clear_filters():
    sucursal_filter.set("Seleccionar Sucursal")
    motivo_filter.set("Seleccionar Motivo")
    observacion_filter.set("Seleccionar Observación")
    update_summary_table(df)
    plot_data(df)
            
def export_summary_to_excel(filtered_df):
    try:
        # Crear resumen por sucursal
        summary_df = filtered_df.groupby('sucursal').agg({
            'TotalCosto': 'sum',
            'Cantidad': 'sum'  # Ajusta según las columnas que necesites sumarizar
        }).reset_index()
        
        # Crear detalles por sucursal y motivo
        details_df = filtered_df.groupby(['sucursal', 'motivo', 'producto']).agg({
            'TotalCosto': 'sum',
            'Cantidad': 'sum'  # Ajusta según las columnas que necesites sumarizar
        }).reset_index()

        # Exportar a Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                summary_df.to_excel(writer, sheet_name='Resumen por Sucursal', index=False)
                details_df.to_excel(writer, sheet_name='Detalles por Sucursal y Motivo', index=False)
            messagebox.showinfo("Exportación exitosa", "Datos exportados correctamente como resumen y detalles.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo exportar el archivo Excel como resumen y detalles.\n{e}")

root = ctk.CTk()
root.title("Análisis de Datos de Excel")
root.geometry("1000x600")

# Frame para filtros
filters_frame = ctk.CTkFrame(root)
filters_frame.pack(side="left", fill="y", padx=20)

load_button = ctk.CTkButton(filters_frame, text="Cargar Excel", command=load_excel)
load_button.pack(pady=10)

sucursal_filter = ctk.CTkComboBox(filters_frame, width=200)
sucursal_filter.set("Seleccionar Sucursal")
sucursal_filter.pack(pady=10)

motivo_filter = ctk.CTkComboBox(filters_frame, width=200)
motivo_filter.set("Seleccionar Motivo")
motivo_filter.pack(pady=10)

observacion_filter = ctk.CTkComboBox(filters_frame, width=200)  # Nuevo: Combo de observación
observacion_filter.set("Seleccionar Observación")
observacion_filter.pack(pady=10)

filter_button = ctk.CTkButton(filters_frame, text="Filtrar Datos", command=filter_data)
filter_button.pack(pady=10)

clear_filters_button = ctk.CTkButton(filters_frame, text="Borrar Filtros", command=clear_filters)
clear_filters_button.pack(pady=10)

# Opciones para totalizado
total_var = StringVar()
total_var.set("Total Costo")  # Valor por defecto
total_radiobutton_costo = ctk.CTkRadioButton(filters_frame, text="Total Costo   ", variable=total_var, value="Total Costo", width=15, command=lambda: plot_data(df))
total_radiobutton_costo.pack(pady=5)
total_radiobutton_cantidad = ctk.CTkRadioButton(filters_frame, text="Total Cantidad", variable=total_var, value="Total Cantidad", width=15, command=lambda: plot_data(df))
total_radiobutton_cantidad.pack(pady=5)

# Boton de exportacion
export_button = ctk.CTkButton(filters_frame, text="Resumen por Sucursal", command=lambda: export_summary_to_excel(df))
export_button.pack(pady=10, side="bottom")

# Frame para resumen y gráfico
summary_plot_frame = ctk.CTkFrame(root)
summary_plot_frame.pack(fill="both", expand=True)

# Tabla resumen
summary_tree = ttk.Treeview(summary_plot_frame, columns=("Sucursal", "Motivo", "Observacion", "Producto", "Total"), show="headings")
summary_tree.heading("Sucursal", text="Sucursal")
summary_tree.heading("Motivo", text="Motivo")
summary_tree.heading("Observacion", text="Observacion")  # Nueva columna de observación
summary_tree.heading("Producto", text="Producto")  # Nueva columna de producto
summary_tree.heading("Total", text="Total")
summary_tree.pack(side="top", fill="both", expand=True, padx=20, pady=10)  # Añadido margen

# Scrollbar para la tabla resumen
summary_scrollbar = ttk.Scrollbar(summary_plot_frame, orient="vertical", command=summary_tree.yview)
summary_scrollbar.pack(side="right", fill="y")
summary_tree.configure(yscrollcommand=summary_scrollbar.set)

# Etiqueta para mostrar totalizados
summary_label = ctk.CTkLabel(summary_plot_frame, text="")
summary_label.pack(pady=10)

# Frame para gráfico
plot_frame = ctk.CTkFrame(root, width=800, height=300)  # Ajuste de tamaño del frame del gráfico
plot_frame.pack(side="right", fill="both", expand=True, padx=20, pady=20)

root.mainloop()

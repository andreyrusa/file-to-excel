import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import ezodf

# Cargar archivo y extraer nombres de hojas
def load_file():
    filepath = filedialog.askopenfilename(filetypes=[("ODS files", "*.ods"), ("Excel files", "*.xlsx *.xls")])
    if not filepath:
        return
    file_label.config(text=filepath)

    global sheet_dict
    sheet_dict = {}

    if filepath.endswith('.ods'):
        ezodf.config.set_table_expand_strategy('all')
        doc = ezodf.opendoc(filepath)
        for sheet in doc.sheets:
            df = pd.DataFrame([[cell.value for cell in row] for row in sheet.rows()])
            df.columns = df.iloc[0]
            df = df[1:]
            sheet_dict[sheet.name] = df
    else:
        xls = pd.ExcelFile(filepath)
        for sheet_name in xls.sheet_names:
            sheet_dict[sheet_name] = xls.parse(sheet_name)

    update_sheet_dropdowns()

def update_sheet_dropdowns():
    sheet_names = list(sheet_dict.keys())
    combo_sheet1['values'] = sheet_names
    combo_sheet2['values'] = sheet_names
    combo_sheet1.current(0)
    combo_sheet2.current(1 if len(sheet_names) > 1 else 0)

def compare_sheets():
    sheet1_name = combo_sheet1.get()
    sheet2_name = combo_sheet2.get()
    if not sheet1_name or not sheet2_name:
        messagebox.showwarning("Aviso", "Selecciona dos hojas.")
        return

    df1 = sheet_dict[sheet1_name]
    df2 = sheet_dict[sheet2_name]

    try:
        df1 = df1.reset_index(drop=True)
        df2 = df2.reset_index(drop=True)

        comparison_result = df1.compare(df2, keep_shape=True, keep_equal=False)
    except Exception as e:
        messagebox.showerror("Error al comparar", str(e))
        return

    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, f"Diferencias entre '{sheet1_name}' y '{sheet2_name}':\n")
    result_text.insert(tk.END, comparison_result.to_string())

    # Guardar a ODS
    save_to_ods(comparison_result)

def save_to_ods(df, output_path="resultado_comparacion.ods"):
    doc = ezodf.newdoc(doctype="spreadsheet", filename=output_path)
    sheet = ezodf.Sheet('Diferencias', size=(len(df)+1, len(df.columns)))
    doc.sheets += sheet

    # Cabeceras
    for c, col in enumerate(df.columns):
        sheet[(0, c)].set_value(str(col))

    # Datos
    for r, row in enumerate(df.itertuples(index=False), start=1):
        for c, val in enumerate(row):
            sheet[(r, c)].set_value(str(val))

    doc.save()
    messagebox.showinfo("Guardado", f"Resultado guardado en {output_path}")

# Interfaz gráfica
root = tk.Tk()
root.title("Comparador de Hojas de Cálculo")

frame = ttk.Frame(root, padding=10)
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

file_btn = ttk.Button(frame, text="Cargar archivo", command=load_file)
file_btn.grid(row=0, column=0, sticky=tk.W)

file_label = ttk.Label(frame, text="Ningún archivo cargado")
file_label.grid(row=0, column=1, sticky=tk.W)

ttk.Label(frame, text="Hoja 1:").grid(row=1, column=0, sticky=tk.W)
combo_sheet1 = ttk.Combobox(frame, state="readonly")
combo_sheet1.grid(row=1, column=1, sticky=tk.W)

ttk.Label(frame, text="Hoja 2:").grid(row=2, column=0, sticky=tk.W)
combo_sheet2 = ttk.Combobox(frame, state="readonly")
combo_sheet2.grid(row=2, column=1, sticky=tk.W)

compare_btn = ttk.Button(frame, text="Comparar", command=compare_sheets)
compare_btn.grid(row=3, column=0, columnspan=2)

result_text = tk.Text(frame, height=20, width=100)
result_text.grid(row=4, column=0, columnspan=2, pady=10)

root.mainloop()

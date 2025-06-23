import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import numpy as np
from pathlib import Path

class SpreadsheetComparator:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Pestañas de Archivos de Cálculo")
        self.root.geometry("1000x700")

        self.file_path = None
        self.excel_file = None
        self.sheet_names = []
        self.selected_cols = []

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        title_label = ttk.Label(main_frame, text="Comparador de Pestañas",
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        ttk.Label(main_frame, text="Archivo:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.file_label = ttk.Label(main_frame, text="Ningún archivo seleccionado", foreground="gray")
        self.file_label.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)

        load_button = ttk.Button(main_frame, text="Cargar Archivo", command=self.load_file)
        load_button.grid(row=1, column=2, padx=(10, 0), pady=5)

        ttk.Label(main_frame, text="Pestaña 1:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.sheet1_combo = ttk.Combobox(main_frame, state="disabled")
        self.sheet1_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)

        ttk.Label(main_frame, text="Pestaña 2:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.sheet2_combo = ttk.Combobox(main_frame, state="disabled")
        self.sheet2_combo.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)

        # Selección de columnas clave
        ttk.Label(main_frame, text="Columnas Identificadoras:").grid(row=4, column=0, sticky=tk.W, pady=5)
        col_select_frame = ttk.Frame(main_frame)
        col_select_frame.grid(row=4, column=1, sticky=tk.W)

        self.col_selector = ttk.Combobox(col_select_frame, state="readonly", width=5)
        self.col_selector.grid(row=0, column=0)

        self.col_list_label = ttk.Label(col_select_frame, text="[]")
        self.col_list_label.grid(row=0, column=2, padx=10)

        add_col_button = ttk.Button(col_select_frame, text="Añadir", command=self.add_selected_column)
        add_col_button.grid(row=0, column=1, padx=5)

        compare_button = ttk.Button(main_frame, text="Comparar Pestañas", command=self.compare_sheets)
        compare_button.grid(row=5, column=0, columnspan=3, pady=20)

        results_frame = ttk.LabelFrame(main_frame, text="Resultados de la Comparación", padding="10")
        results_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(1, weight=1)

        self.results_info = ttk.Label(results_frame, text="")
        self.results_info.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))

        table_frame = ttk.Frame(results_frame)
        table_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.results_tree = ttk.Treeview(table_frame, show="tree headings")

        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.results_tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.results_tree.xview)

        self.results_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        self.results_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        export_button = ttk.Button(results_frame, text="Exportar Diferencias a ODS", command=self.export_differences, state="disabled")
        export_button.grid(row=2, column=0, pady=(10, 0))
        self.export_button = export_button

        main_frame.rowconfigure(6, weight=1)

    def add_selected_column(self):
        col = self.col_selector.get()
        if col and int(col) not in self.selected_cols:
            self.selected_cols.append(int(col))
            self.col_list_label.config(text=str(self.selected_cols))

    def load_file(self):
        file_types = [
            ("Archivos de cálculo", "*.xlsx *.xls *.ods"),
            ("Archivos Excel", "*.xlsx *.xls"),
            ("Archivos ODS", "*.ods"),
            ("Todos los archivos", "*.*")
        ]

        file_path = filedialog.askopenfilename(title="Seleccionar archivo de cálculo", filetypes=file_types)

        if file_path:
            try:
                self.file_path = file_path
                self.file_label.config(text=Path(file_path).name, foreground="black")

                engine = 'odf' if file_path.endswith('.ods') else None
                self.excel_file = pd.ExcelFile(file_path, engine=engine)
                self.sheet_names = self.excel_file.sheet_names

                self.sheet1_combo.config(values=self.sheet_names, state="readonly")
                self.sheet2_combo.config(values=self.sheet_names, state="readonly")
                self.sheet1_combo.set("")
                self.sheet2_combo.set("")

                num_cols = max([len(pd.read_excel(self.excel_file, sheet_name=sh).columns) for sh in self.sheet_names])
                self.col_selector.config(values=[str(i) for i in range(num_cols)], state="readonly")
                self.selected_cols = []
                self.col_list_label.config(text="[]")

                messagebox.showinfo("Éxito", f"Archivo cargado correctamente.\nPestañas encontradas: {len(self.sheet_names)}")

            except Exception as e:
                messagebox.showerror("Error", f"Error al cargar el archivo:\n{str(e)}")
                self.file_label.config(text="Error al cargar archivo", foreground="red")

    def extract_identificadores(self, df1, df2, row_idx):
        identificadores = {}
        for col_idx in self.selected_cols:
            val = None
            try:
                if row_idx is not None:
                    if row_idx < len(df1) and col_idx < df1.shape[1]:
                        val = df1.iloc[row_idx, col_idx]
                    elif row_idx < len(df2) and col_idx < df2.shape[1]:
                        val = df2.iloc[row_idx, col_idx]
            except Exception:
                val = None
            identificadores[f'Columna_{col_idx}'] = val
        return identificadores

    def compare_sheets(self):
        if not self.excel_file:
            messagebox.showwarning("Advertencia", "Primero debe cargar un archivo")
            return

        sheet1 = self.sheet1_combo.get()
        sheet2 = self.sheet2_combo.get()

        if not sheet1 or not sheet2:
            messagebox.showwarning("Advertencia", "Debe seleccionar ambas pestañas")
            return

        if sheet1 == sheet2:
            messagebox.showwarning("Advertencia", "Debe seleccionar pestañas diferentes")
            return

        try:
            df1 = pd.read_excel(self.excel_file, sheet_name=sheet1)
            df2 = pd.read_excel(self.excel_file, sheet_name=sheet2)

            self.df1 = df1
            self.df2 = df2
            self.sheet1_name = sheet1
            self.sheet2_name = sheet2

            differences = self.find_differences(df1, df2, sheet1, sheet2)
            self.display_results(differences)
            self.export_button.config(state="normal")

        except Exception as e:
            messagebox.showerror("Error", f"Error al comparar pestañas:\n{str(e)}")

    def find_differences(self, df1, df2, sheet1_name, sheet2_name):
        differences = []

        if df1.shape != df2.shape:
            differences.append({
                'Tipo': 'Dimensiones',
                'Descripción': f'{sheet1_name}: {df1.shape} vs {sheet2_name}: {df2.shape}',
                'Fila': '-',
                'Columna': '-',
                f'Valor en {sheet1_name}': f'{df1.shape[0]} filas, {df1.shape[1]} columnas',
                f'Valor en {sheet2_name}': f'{df2.shape[0]} filas, {df2.shape[1]} columnas',
                **self.extract_identificadores(df1, df2, None)
            })

        cols1 = set(df1.columns)
        cols2 = set(df2.columns)
        common_cols = cols1.intersection(cols2)
        only_in_1 = cols1 - cols2
        only_in_2 = cols2 - cols1

        for col in only_in_1:
            differences.append({
                'Tipo': 'Columna faltante',
                'Descripción': f'Columna "{col}" solo existe en {sheet1_name}',
                'Fila': '-',
                'Columna': col,
                f'Valor en {sheet1_name}': 'Existe',
                f'Valor en {sheet2_name}': 'No existe',
                **self.extract_identificadores(df1, df2, None)
            })

        for col in only_in_2:
            differences.append({
                'Tipo': 'Columna faltante',
                'Descripción': f'Columna "{col}" solo existe en {sheet2_name}',
                'Fila': '-',
                'Columna': col,
                f'Valor en {sheet1_name}': 'No existe',
                f'Valor en {sheet2_name}': 'Existe',
                **self.extract_identificadores(df1, df2, None)
            })

        min_rows = min(len(df1), len(df2))

        for col in common_cols:
            for i in range(min_rows):
                val1 = df1.iloc[i][col]
                val2 = df2.iloc[i][col]

                if pd.isna(val1) and pd.isna(val2):
                    continue
                elif pd.isna(val1) or pd.isna(val2) or val1 != val2:
                    identificadores = self.extract_identificadores(df1, df2, i)
                    differences.append({
                        'Tipo': 'Valor diferente',
                        'Descripción': f'Diferencia en fila {i+1}, columna "{col}"',
                        'Fila': i+1,
                        'Columna': col,
                        f'Valor en {sheet1_name}': str(val1) if not pd.isna(val1) else 'NaN',
                        f'Valor en {sheet2_name}': str(val2) if not pd.isna(val2) else 'NaN',
                        **identificadores
                    })

        for i in range(min_rows, len(df1)):
            differences.append({
                'Tipo': 'Fila adicional',
                'Descripción': f'Fila {i+1} solo existe en {sheet1_name}',
                'Fila': i+1,
                'Columna': '-',
                f'Valor en {sheet1_name}': 'Existe',
                f'Valor en {sheet2_name}': 'No existe',
                **self.extract_identificadores(df1, df2, i)
            })

        for i in range(min_rows, len(df2)):
            differences.append({
                'Tipo': 'Fila adicional',
                'Descripción': f'Fila {i+1} solo existe en {sheet2_name}',
                'Fila': i+1,
                'Columna': '-',
                f'Valor en {sheet1_name}': 'No existe',
                f'Valor en {sheet2_name}': 'Existe',
                **self.extract_identificadores(df1, df2, i)
            })

        return differences

    def display_results(self, differences):
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)

        if not differences:
            self.results_info.config(text="¡Las pestañas son idénticas!")
            self.results_tree.config(columns=())
            return

        columns = list(differences[0].keys())
        self.results_tree.config(columns=columns)
        self.results_tree.heading("#0", text="", anchor="w")
        self.results_tree.column("#0", width=0, stretch=False)

        for col in columns:
            self.results_tree.heading(col, text=col, anchor="w")
            self.results_tree.column(col, width=120, anchor="w")

        for i, diff in enumerate(differences):
            values = [diff.get(col, '') for col in columns]
            self.results_tree.insert("", "end", iid=i, values=values)

        self.results_info.config(text=f"Se encontraron {len(differences)} diferencias")

    def export_differences(self):
        if not hasattr(self, 'df1') or not hasattr(self, 'df2'):
            messagebox.showwarning("Advertencia", "No hay datos para exportar")
            return

        try:
            differences = self.find_differences(self.df1, self.df2, self.sheet1_name, self.sheet2_name)
            if not differences:
                messagebox.showinfo("Información", "No hay diferencias para exportar")
                return

            diff_df = pd.DataFrame(differences)

            output_file = filedialog.asksaveasfilename(
                title="Guardar diferencias como",
                defaultextension=".ods",
                filetypes=[("Archivos ODS", "*.ods"), ("Archivos Excel", "*.xlsx")]
            )

            if output_file:
                engine = 'odf' if output_file.endswith('.ods') else 'openpyxl'
                with pd.ExcelWriter(output_file, engine=engine) as writer:
                    diff_df.to_excel(writer, sheet_name='Diferencias', index=False)
                    self.df1.to_excel(writer, sheet_name=self.sheet1_name, index=False)
                    self.df2.to_excel(writer, sheet_name=self.sheet2_name, index=False)

                messagebox.showinfo("Éxito", f"Diferencias exportadas correctamente a:\n{output_file}")

        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar:\n{str(e)}")


def main():
    root = tk.Tk()
    app = SpreadsheetComparator(root)
    root.mainloop()

if __name__ == "__main__":
    main()

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import json
from main import operate

class MergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Fusionar Calificaciones")
        self.root.geometry("600x500")

        # --- SCROLL AREA ---
        canvas = tk.Canvas(root)
        scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- Variables ---
        self.source_file = tk.StringVar()
        self.dest_file = tk.StringVar()
        self.json_file = tk.StringVar()
        self.zone_columns = tk.StringVar()
        self.exam_final_column = tk.StringVar()
        self.total_zone = tk.DoubleVar(value=20)
        self.total_final = tk.DoubleVar(value=10)

        # --- Widgets dentro de scroll_frame ---
        ttk.Label(scroll_frame, text="Archivo de origen (ODS/XLSX):").pack(pady=5)
        ttk.Entry(scroll_frame, textvariable=self.source_file, width=60).pack()
        ttk.Button(scroll_frame, text="Seleccionar", command=self.browse_source).pack()

        ttk.Label(scroll_frame, text="Archivo de destino (ODS/XLSX):").pack(pady=5)
        ttk.Entry(scroll_frame, textvariable=self.dest_file, width=60).pack()
        ttk.Button(scroll_frame, text="Seleccionar", command=self.browse_dest).pack()

        ttk.Label(scroll_frame, text="Columnas Zona (separadas por coma):").pack(pady=5)
        ttk.Entry(scroll_frame, textvariable=self.zone_columns, width=60).pack()

        ttk.Label(scroll_frame, text="Columna Examen Final:").pack(pady=5)
        ttk.Entry(scroll_frame, textvariable=self.exam_final_column, width=60).pack()

        ttk.Label(scroll_frame, text="Total Zona:").pack(pady=5)
        ttk.Entry(scroll_frame, textvariable=self.total_zone, width=20).pack()

        ttk.Label(scroll_frame, text="Total Final:").pack(pady=5)
        ttk.Entry(scroll_frame, textvariable=self.total_final, width=20).pack()

        ttk.Label(scroll_frame, text="Archivo JSON de parámetros (opcional):").pack(pady=5)
        ttk.Entry(scroll_frame, textvariable=self.json_file, width=60).pack()
        ttk.Button(scroll_frame, text="Seleccionar", command=self.browse_json).pack()

        ttk.Button(scroll_frame, text="Generar JSON", style="Accent.TButton",
                   command=self.generate_json).pack(pady=10)
        ttk.Button(scroll_frame, text="Ejecutar", style="Accent.TButton",
                   command=self.run_operation).pack(pady=10)

        # --- Estilos ttk ---
        style = ttk.Style()
        style.theme_use('clam')  # Opciones: 'clam', 'default', 'alt'
        style.configure("Accent.TButton", foreground="white", background="#0078D7", padding=6)
        style.map("Accent.TButton",
                  background=[("active", "#005a9e")])

    # --- FUNCIONES ORIGINALES ---
    def browse_source(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archivos ODS/XLSX", "*.ods *.xlsx")])
        if file_path:
            self.source_file.set(file_path)

    def browse_dest(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archivos ODS/XLSX", "*.ods *.xlsx")])
        if file_path:
            self.dest_file.set(file_path)

    def browse_json(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archivos JSON", "*.json")])
        if file_path:
            self.json_file.set(file_path)

    def generate_json(self):
        try:
            cols_zone = [c.strip() for c in self.zone_columns.get().split(",") if c.strip()]
            exam_final = self.exam_final_column.get().strip()

            if not cols_zone or not exam_final:
                messagebox.showerror("Error", "Debes especificar columnas de zona y examen final.")
                return

            config = {
                "file1": {
                    "colZone": {
                        "cols": [{"col": c} for c in cols_zone]
                    },
                    "colExam": {
                        "colA": exam_final
                    },
                    "totalZone": self.total_zone.get(),
                    "totalFinal": self.total_final.get()
                }
            }

            json_path = Path("generated_config.json")
            json_path.write_text(json.dumps(config, indent=4), encoding="utf-8")
            self.json_file.set(str(json_path))
            messagebox.showinfo("Éxito", f"JSON generado: {json_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run_operation(self):
        try:
            src = Path(self.source_file.get())
            dst = Path(self.dest_file.get())
            config_path = Path(self.json_file.get())

            if not src.exists() or not dst.exists() or not config_path.exists():
                messagebox.showerror("Error", "Debes seleccionar todos los archivos y generar el JSON.")
                return

            config = json.loads(config_path.read_text())
            operate(src, dst, config)
            messagebox.showinfo("Éxito", "Operación completada con éxito.")
        except Exception as e:
            messagebox.showerror("Error", str(e))


def main_gui():
    root = tk.Tk()
    MergeApp(root)
    root.mainloop()

if __name__ == "__main__":
    main_gui()

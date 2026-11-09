import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import re
from pathlib import Path

import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── Extracción ───────────────────────────────────────────────────────────────

def extract_pdf(pdf_path: str) -> dict:
    """Extrae texto, tablas y metadatos de un PDF."""
    result = {"path": pdf_path, "pages": [], "tables": [], "metadata": {}}

    with pdfplumber.open(pdf_path) as pdf:
        result["metadata"] = pdf.metadata or {}
        result["total_pages"] = len(pdf.pages)

        for i, page in enumerate(pdf.pages):
            page_text = page.extract_text() or ""
            result["pages"].append({"page": i + 1, "text": page_text})

            for table in page.extract_tables():
                if table:
                    result["tables"].append({"page": i + 1, "data": table})

    return result


def parse_invoice_fields(pages: list) -> dict:
    """Detecta campos comunes de factura en el texto libre."""
    full_text = "\n".join(p["text"] for p in pages)
    patterns = {
        "Número de Factura": r"(?:factura|invoice|n[uú]mero|no\.?|#)\s*[:\-]?\s*([A-Z0-9\-]+)",
        "Fecha":             r"(?:fecha|date)\s*[:\-]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
        "Total":             r"(?:total|amount due|monto total)\s*[:\$]?\s*([\d,\.]+)",
        "Subtotal":          r"(?:subtotal|sub-total)\s*[:\$]?\s*([\d,\.]+)",
        "IVA / Tax":         r"(?:iva|tax|impuesto)\s*[:\$]?\s*([\d,\.]+)",
        "Proveedor":         r"(?:de:|from:|proveedor|supplier)\s*(.+)",
        "Cliente":           r"(?:para:|to:|cliente|customer|bill to)\s*(.+)",
    }
    fields = {}
    for label, pattern in patterns.items():
        match = re.search(pattern, full_text, re.IGNORECASE)
        fields[label] = match.group(1).strip() if match else ""
    return fields


def build_dataframes(extracted: dict) -> dict:
    """Construye DataFrames a partir de los datos extraídos."""
    frames = {}

    # Hoja de resumen
    meta = extracted["metadata"]
    info = {
        "Archivo":        [os.path.basename(extracted["path"])],
        "Páginas":        [extracted.get("total_pages", "")],
        "Autor":          [meta.get("Author", "")],
        "Creador":        [meta.get("Creator", "")],
        "Fecha creación": [meta.get("CreationDate", "")],
    }
    # Campos de factura detectados
    invoice_fields = parse_invoice_fields(extracted["pages"])
    info.update({k: [v] for k, v in invoice_fields.items()})
    frames["Resumen"] = pd.DataFrame(info).T.reset_index()
    frames["Resumen"].columns = ["Campo", "Valor"]

    # Hoja de texto por página
    page_rows = [{"Página": p["page"], "Texto": p["text"]} for p in extracted["pages"]]
    frames["Texto por Página"] = pd.DataFrame(page_rows)

    # Una hoja por tabla encontrada
    for idx, table_info in enumerate(extracted["tables"], 1):
        raw = table_info["data"]
        if not raw:
            continue
        header = raw[0]
        rows = raw[1:]
        # Sanitiza encabezados vacíos
        header = [h if h else f"Col{i}" for i, h in enumerate(header)]
        try:
            df = pd.DataFrame(rows, columns=header)
        except Exception:
            df = pd.DataFrame(raw)
        frames[f"Tabla {idx} (p.{table_info['page']})"] = df

    return frames


# ─── Exportación a Excel ──────────────────────────────────────────────────────

HEADER_FILL   = PatternFill("solid", fgColor="2E4057")
HEADER_FONT   = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
ALT_FILL      = PatternFill("solid", fgColor="EAF0FB")
BORDER_SIDE   = Side(style="thin", color="BFCDE0")
CELL_BORDER   = Border(left=BORDER_SIDE, right=BORDER_SIDE,
                       top=BORDER_SIDE, bottom=BORDER_SIDE)


def style_sheet(ws):
    """Aplica estilo al encabezado y filas alternas."""
    for cell in ws[1]:
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = CELL_BORDER

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        fill = ALT_FILL if row_idx % 2 == 0 else PatternFill()
        for cell in row:
            cell.fill   = fill
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Ajuste automático de columnas
    for col in ws.columns:
        max_len = max(
            (len(str(c.value)) if c.value is not None else 0 for c in col), default=0
        )
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

    ws.freeze_panes = "A2"


def export_to_excel(frames: dict, output_path: str):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in frames.items():
            safe_name = sheet_name[:31]  # Excel límita a 31 caracteres
            df.to_excel(writer, sheet_name=safe_name, index=False)

    wb = load_workbook(output_path)
    for ws in wb.worksheets:
        style_sheet(ws)
    wb.save(output_path)


# ─── Interfaz gráfica ─────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Extractor de PDFs a Excel")
        self.geometry("720x540")
        self.resizable(True, True)
        self.configure(bg="#F0F4FA")
        self._pdf_files: list[str] = []
        self._output_dir = tk.StringVar(value=str(Path.home() / "Documents"))
        self._build_ui()

    # ── Construcción de la UI ──────────────────────────────────────────────────

    def _build_ui(self):
        # Título
        tk.Label(
            self, text="Extractor de PDFs a Excel",
            font=("Segoe UI", 16, "bold"), bg="#F0F4FA", fg="#2E4057"
        ).pack(pady=(18, 4))
        tk.Label(
            self, text="Lee facturas y reportes PDF · Exporta tablas y campos a .xlsx",
            font=("Segoe UI", 9), bg="#F0F4FA", fg="#6B7A99"
        ).pack(pady=(0, 12))

        # Zona de archivos
        frame_files = tk.LabelFrame(
            self, text="  Archivos PDF  ", bg="#F0F4FA", fg="#2E4057",
            font=("Segoe UI", 10, "bold"), padx=10, pady=8
        )
        frame_files.pack(fill="x", padx=20, pady=(0, 8))

        btn_row = tk.Frame(frame_files, bg="#F0F4FA")
        btn_row.pack(fill="x")
        self._btn(btn_row, "Agregar PDFs", self._add_files, "#2E7D32").pack(side="left", padx=(0, 8))
        self._btn(btn_row, "Limpiar lista", self._clear_files, "#B71C1C").pack(side="left")

        self.listbox = tk.Listbox(
            frame_files, height=6, selectmode="extended",
            font=("Consolas", 9), bg="#FFFFFF", fg="#333333",
            selectbackground="#2E4057", activestyle="none"
        )
        self.listbox.pack(fill="both", expand=True, pady=(8, 0))

        sb = ttk.Scrollbar(frame_files, command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        # Destino
        frame_out = tk.LabelFrame(
            self, text="  Carpeta de destino  ", bg="#F0F4FA", fg="#2E4057",
            font=("Segoe UI", 10, "bold"), padx=10, pady=8
        )
        frame_out.pack(fill="x", padx=20, pady=(0, 8))

        out_row = tk.Frame(frame_out, bg="#F0F4FA")
        out_row.pack(fill="x")
        tk.Entry(out_row, textvariable=self._output_dir, font=("Segoe UI", 9), width=55).pack(side="left", expand=True, fill="x")
        self._btn(out_row, "Examinar", self._browse_output, "#1565C0").pack(side="left", padx=(8, 0))

        # Opciones
        frame_opt = tk.LabelFrame(
            self, text="  Opciones  ", bg="#F0F4FA", fg="#2E4057",
            font=("Segoe UI", 10, "bold"), padx=10, pady=8
        )
        frame_opt.pack(fill="x", padx=20, pady=(0, 8))

        self._one_file = tk.BooleanVar(value=True)
        self._open_after = tk.BooleanVar(value=True)
        tk.Checkbutton(
            frame_opt, text="Un archivo Excel por PDF",
            variable=self._one_file, bg="#F0F4FA", font=("Segoe UI", 9)
        ).pack(side="left", padx=(0, 20))
        tk.Checkbutton(
            frame_opt, text="Abrir carpeta al finalizar",
            variable=self._open_after, bg="#F0F4FA", font=("Segoe UI", 9)
        ).pack(side="left")

        # Botón principal + progreso
        self._btn_run = self._btn(self, "Extraer y Exportar", self._run, "#2E4057", width=22, size=12)
        self._btn_run.pack(pady=(4, 6))

        self.progress = ttk.Progressbar(self, mode="determinate", length=680)
        self.progress.pack(padx=20, pady=(0, 4))

        self.status = tk.StringVar(value="Listo")
        tk.Label(self, textvariable=self.status, bg="#F0F4FA", fg="#555", font=("Segoe UI", 9)).pack()

    def _btn(self, parent, text, cmd, color, width=16, size=10):
        return tk.Button(
            parent, text=text, command=cmd,
            bg=color, fg="white", relief="flat", cursor="hand2",
            font=("Segoe UI", size, "bold"), width=width,
            activebackground=color, activeforeground="white",
            padx=6, pady=5
        )

    # ── Handlers ──────────────────────────────────────────────────────────────

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar PDFs",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos", "*.*")]
        )
        for f in files:
            if f not in self._pdf_files:
                self._pdf_files.append(f)
                self.listbox.insert("end", os.path.basename(f))

    def _clear_files(self):
        self._pdf_files.clear()
        self.listbox.delete(0, "end")

    def _browse_output(self):
        d = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if d:
            self._output_dir.set(d)

    def _run(self):
        if not self._pdf_files:
            messagebox.showwarning("Sin archivos", "Agrega al menos un PDF.")
            return
        self._btn_run.configure(state="disabled")
        threading.Thread(target=self._process, daemon=True).start()

    def _process(self):
        total = len(self._pdf_files)
        self.progress["maximum"] = total
        errors = []
        output_dir = self._output_dir.get()
        os.makedirs(output_dir, exist_ok=True)

        for i, pdf_path in enumerate(self._pdf_files, 1):
            name = os.path.basename(pdf_path)
            self.status.set(f"Procesando {i}/{total}: {name}")
            try:
                extracted = extract_pdf(pdf_path)
                frames    = build_dataframes(extracted)
                stem      = Path(pdf_path).stem
                out_path  = os.path.join(output_dir, f"{stem}.xlsx")
                export_to_excel(frames, out_path)
            except Exception as e:
                errors.append(f"{name}: {e}")
            self.progress["value"] = i

        self._btn_run.configure(state="normal")
        if errors:
            self.status.set("Completado con errores")
            messagebox.showerror("Errores", "\n".join(errors))
        else:
            self.status.set(f"¡Listo! {total} archivo(s) exportado(s).")
            if self._open_after.get():
                os.startfile(output_dir)


if __name__ == "__main__":
    app = App()
    app.mainloop()

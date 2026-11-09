"""Genera dos PDFs de demostración: una factura y un reporte de ventas."""
from fpdf import FPDF, XPos, YPos
from pathlib import Path

OUTPUT_DIR = Path(__file__).parent / "demo"
OUTPUT_DIR.mkdir(exist_ok=True)

NL = {"new_x": XPos.LMARGIN, "new_y": YPos.NEXT}


# ─── Factura ──────────────────────────────────────────────────────────────────

class FacturaPDF(FPDF):
    def header(self):
        self.set_fill_color(46, 64, 87)
        self.rect(0, 0, 210, 32, "F")
        self.set_font("Helvetica", "B", 20)
        self.set_text_color(255, 255, 255)
        self.cell(0, 20, "TECNOSOLUCIONES S.A. DE C.V.", align="C", **NL)
        self.set_font("Helvetica", "", 10)
        self.cell(0, 8, "RFC: TSC930215AB3  |  tecnos@ejemplo.com  |  +52 55 1234 5678", align="C", **NL)
        self.set_text_color(0, 0, 0)
        self.ln(4)

    def footer(self):
        self.set_y(-14)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, "Este documento es una factura de demostracion generada automaticamente.", align="C")


def generar_factura():
    pdf = FacturaPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(46, 64, 87)
    pdf.cell(0, 8, "FACTURA", **NL)
    pdf.set_draw_color(46, 64, 87)
    pdf.set_line_width(0.6)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_text_color(0, 0, 0)
    datos_factura = [
        ("Numero de Factura:", "F-2024-00892"),
        ("Fecha:",             "03/11/2026"),
        ("Vencimiento:",       "03/12/2026"),
        ("Moneda:",            "MXN - Peso Mexicano"),
    ]
    for label, valor in datos_factura:
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(55, 7, label)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 7, valor, **NL)
    pdf.ln(4)

    # Emisor / receptor
    y_top = pdf.get_y()
    pdf.set_fill_color(234, 240, 251)
    pdf.rect(10, y_top, 88, 32, "F")
    pdf.rect(112, y_top, 88, 32, "F")

    for x, titulo, lineas in [
        (12, "EMISOR",  ["Tecnosoluciones S.A. de C.V.", "Av. Reforma 150, CDMX", "RFC: TSC930215AB3"]),
        (114, "CLIENTE", ["Distribuidora Lopez y Hnos.", "Calle Juarez 42, Monterrey, NL", "RFC: DLH881103QW2"]),
    ]:
        pdf.set_xy(x, y_top + 2)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(46, 64, 87)
        pdf.cell(84, 6, titulo, **NL)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(0, 0, 0)
        for linea in lineas:
            pdf.set_x(x)
            pdf.cell(84, 5, linea, **NL)

    pdf.set_y(y_top + 36)
    pdf.ln(2)

    # Tabla de conceptos
    cols    = [10, 85, 25, 35, 35]
    headers = ["#", "Descripcion", "Cant.", "P. Unitario", "Importe"]
    pdf.set_fill_color(46, 64, 87)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    for w, h in zip(cols, headers):
        pdf.cell(w, 8, h, border=0, fill=True, align="C")
    pdf.ln()

    items = [
        (1, "Licencia Software ERP Empresarial (anual)",         2,  12500.00),
        (2, "Implementacion y configuracion del sistema",        1,  18000.00),
        (3, "Capacitacion de usuarios (8 horas)",                3,   2200.00),
        (4, "Soporte tecnico prioritario - paquete trimestral",  1,   4800.00),
        (5, "Migracion de base de datos legado",                 1,   9500.00),
    ]

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 9)
    for idx, (num, desc, qty, unit) in enumerate(items):
        fill = idx % 2 == 0
        pdf.set_fill_color(234, 240, 251) if fill else pdf.set_fill_color(255, 255, 255)
        total_item = qty * unit
        pdf.cell(10,  7, str(num),               border="B", fill=fill, align="C")
        pdf.cell(85,  7, desc,                   border="B", fill=fill)
        pdf.cell(25,  7, str(qty),               border="B", fill=fill, align="C")
        pdf.cell(35,  7, f"${unit:,.2f}",        border="B", fill=fill, align="R")
        pdf.cell(35,  7, f"${total_item:,.2f}",  border="B", fill=fill, align="R")
        pdf.ln()

    pdf.ln(3)
    subtotal = sum(q * u for _, _, q, u in items)
    iva      = subtotal * 0.16
    total    = subtotal + iva

    for label, monto, bold in [
        ("Subtotal:", subtotal, False),
        ("IVA (16%):", iva,     False),
        ("TOTAL:", total,       True),
    ]:
        pdf.set_x(130)
        if bold:
            pdf.set_fill_color(46, 64, 87)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(40, 8, label,             fill=True, align="R")
            pdf.cell(40, 8, f"${monto:,.2f}",  fill=True, align="R")
            pdf.set_text_color(0, 0, 0)
        else:
            pdf.set_font("Helvetica", "", 9)
            pdf.cell(40, 7, label,            align="R")
            pdf.cell(40, 7, f"${monto:,.2f}", align="R")
        pdf.ln()

    pdf.ln(6)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(46, 64, 87)
    pdf.cell(0, 6, "Condiciones de pago:", **NL)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 5,
        "Transferencia bancaria a 30 dias. CLABE: 012180001234567890. "
        "Banco: BBVA. En caso de mora se aplicara un cargo del 2% mensual.",
    )

    path = OUTPUT_DIR / "factura_demo.pdf"
    pdf.output(str(path))
    print(f"Generado: {path}")


# ─── Reporte de ventas ────────────────────────────────────────────────────────

class ReportePDF(FPDF):
    def header(self):
        self.set_fill_color(21, 101, 192)
        self.rect(0, 0, 210, 28, "F")
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(255, 255, 255)
        self.cell(0, 16, "REPORTE DE VENTAS - Q4 2026", align="C", **NL)
        self.set_font("Helvetica", "", 9)
        self.cell(0, 8, "Periodo: Octubre - Diciembre 2026  |  Generado: 03/11/2026", align="C", **NL)
        self.set_text_color(0, 0, 0)
        self.ln(4)


def generar_reporte():
    pdf = ReportePDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    ventas = [
        ("Octubre",   "Electronica",  "Laptop ProBook 15",          45,  18500,  832500),
        ("Octubre",   "Electronica",  "Monitor UltraWide 34in",     30,   9800,  294000),
        ("Octubre",   "Accesorios",   "Teclado Mecanico RGB",       80,   1200,   96000),
        ("Noviembre", "Electronica",  "Laptop ProBook 15",          62,  18500, 1147000),
        ("Noviembre", "Software",     "Licencia Office 365",       120,   1800,  216000),
        ("Noviembre", "Accesorios",   "Mouse Inalambrico Ergo",    150,    850,  127500),
        ("Diciembre", "Electronica",  "Smartphone Galaxy S25",      88,  22000, 1936000),
        ("Diciembre", "Electronica",  "Tablet Pro 12",              34,  15500,  527000),
        ("Diciembre", "Software",     "Licencia Office 365",       200,   1800,  360000),
        ("Diciembre", "Accesorios",   "Auriculares Noise Cancel",   95,   3200,  304000),
    ]

    total_ventas = sum(r[5] for r in ventas)
    total_unids  = sum(r[3] for r in ventas)

    # KPIs
    kpis = [
        ("Ingresos Totales",  f"${total_ventas:,.0f} MXN"),
        ("Unidades Vendidas", f"{total_unids:,}"),
        ("Meses del Periodo", "3"),
        ("Categorias",        "3"),
    ]
    pdf.set_font("Helvetica", "B", 10)
    for label, _ in kpis:
        pdf.set_fill_color(46, 64, 87)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(44, 8, label, fill=True, align="C")
        pdf.set_fill_color(234, 240, 251)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(4, 8, "", fill=False)
    pdf.ln()
    for _, val in kpis:
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(21, 101, 192)
        pdf.cell(44, 9, val, align="C")
        pdf.cell(4, 9, "")
    pdf.ln(14)

    # Tabla ventas
    headers = ["Mes", "Categoria", "Producto", "Unids.", "P. Unit.", "Total"]
    widths  = [24, 28, 68, 16, 24, 30]
    pdf.set_fill_color(21, 101, 192)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    for w, h in zip(widths, headers):
        pdf.cell(w, 8, h, fill=True, align="C")
    pdf.ln()

    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(0, 0, 0)
    for i, (mes, cat, prod, unids, punit, tot) in enumerate(ventas):
        fill = i % 2 == 0
        pdf.set_fill_color(234, 240, 251) if fill else pdf.set_fill_color(255, 255, 255)
        pdf.cell(24, 7, mes,           border="B", fill=fill, align="C")
        pdf.cell(28, 7, cat,           border="B", fill=fill, align="C")
        pdf.cell(68, 7, prod,          border="B", fill=fill)
        pdf.cell(16, 7, str(unids),    border="B", fill=fill, align="C")
        pdf.cell(24, 7, f"${punit:,}", border="B", fill=fill, align="R")
        pdf.cell(30, 7, f"${tot:,}",   border="B", fill=fill, align="R")
        pdf.ln()

    pdf.set_x(130)
    pdf.set_fill_color(21, 101, 192)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(60, 8, "TOTAL PERIODO:", fill=True, align="R")
    pdf.cell(0,  8, f"  ${total_ventas:,.0f}", fill=True)
    pdf.ln(14)

    # Resumen por categoría
    from collections import defaultdict
    por_cat = defaultdict(int)
    for _, cat, _, _, _, tot in ventas:
        por_cat[cat] += tot

    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(46, 64, 87)
    pdf.cell(0, 7, "Resumen por Categoria", **NL)
    pdf.set_line_width(0.5)
    pdf.set_draw_color(21, 101, 192)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_fill_color(21, 101, 192)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(60, 7, "Categoria",   fill=True, align="C")
    pdf.cell(40, 7, "Total MXN",   fill=True, align="C")
    pdf.cell(30, 7, "% del Total", fill=True, align="C")
    pdf.ln()

    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(0, 0, 0)
    for i, (cat, tot) in enumerate(sorted(por_cat.items(), key=lambda x: -x[1])):
        fill = i % 2 == 0
        pdf.set_fill_color(234, 240, 251) if fill else pdf.set_fill_color(255, 255, 255)
        pct = tot / total_ventas * 100
        pdf.cell(60, 7, cat,             fill=fill, align="C")
        pdf.cell(40, 7, f"${tot:,.0f}",  fill=fill, align="R")
        pdf.cell(30, 7, f"{pct:.1f}%",   fill=fill, align="C")
        pdf.ln()

    path = OUTPUT_DIR / "reporte_ventas_demo.pdf"
    pdf.output(str(path))
    print(f"Generado: {path}")


if __name__ == "__main__":
    generar_factura()
    generar_reporte()
    print("\nPDFs de demostracion listos en la carpeta 'demo/'")

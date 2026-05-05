from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    import customtkinter as ctk
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Falta instalar customtkinter. Ejecuta: pip install -r requirements.txt"
    ) from exc

from impresiones_app.core import (
    DEFAULT_ALIADO_COMERCIAL,
    VISIBLE_COLUMNS,
    generate_pdfs,
    group_records,
    load_excel_records,
)


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
DEFAULT_LOGO = RESOURCE_DIR / "Logo" / "logo-afinia.png"
DEFAULT_OUTPUT = APP_DIR / "Salidad Impresiones"
PAGE_SIZE = 50


class PrintFormatterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Clientes Especiales - Generador de Formatos")
        self.geometry("1280x760")
        self.minsize(1080, 650)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.records = []
        self.current_page = 1
        self.generated_files: list[Path] = []

        self.excel_path = ctk.StringVar(value="")
        self.output_path = ctk.StringVar(value=str(DEFAULT_OUTPUT))
        self.logo_path = ctk.StringVar(value=str(DEFAULT_LOGO if DEFAULT_LOGO.exists() else ""))
        self.excel_display = ctk.StringVar(value="Ningún archivo seleccionado")
        self.output_display = ctk.StringVar(value=self._display_path(DEFAULT_OUTPUT))
        self.logo_display = ctk.StringVar(value="Logo predeterminado" if DEFAULT_LOGO.exists() else "Ningún logo seleccionado")
        self.status_text = ctk.StringVar(value="Seleccione un archivo Excel para comenzar.")
        self.summary_text = ctk.StringVar(value="0 registros | 0 formatos")
        self.printer_name = ctk.StringVar(value="")
        self.print_after_generate = ctk.BooleanVar(value=False)

        self._build_ui()
        self._load_printers()

    @staticmethod
    def _display_path(path: str | Path, empty_text: str = "Sin seleccionar") -> str:
        if not path:
            return empty_text
        path_obj = Path(path)
        return path_obj.name or str(path_obj)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        top = ctk.CTkFrame(self, corner_radius=0)
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(top, text="Archivo Excel").grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")
        ctk.CTkEntry(top, textvariable=self.excel_display, state="disabled").grid(row=0, column=1, padx=8, pady=(12, 6), sticky="ew")
        ctk.CTkButton(top, text="Seleccionar", command=self.select_excel).grid(row=0, column=2, padx=8, pady=(12, 6))
        ctk.CTkButton(top, text="Cargar", command=self.load_excel).grid(row=0, column=3, padx=12, pady=(12, 6))

        ctk.CTkLabel(top, text="Carpeta salida").grid(row=1, column=0, padx=12, pady=6, sticky="w")
        ctk.CTkEntry(top, textvariable=self.output_display, state="disabled").grid(row=1, column=1, padx=8, pady=6, sticky="ew")
        ctk.CTkButton(top, text="Cambiar", command=self.select_output).grid(row=1, column=2, padx=8, pady=6)

        ctk.CTkLabel(top, text="Logo").grid(row=2, column=0, padx=12, pady=(6, 12), sticky="w")
        ctk.CTkEntry(top, textvariable=self.logo_display, state="disabled").grid(row=2, column=1, padx=8, pady=(6, 12), sticky="ew")
        ctk.CTkButton(top, text="Cambiar", command=self.select_logo).grid(row=2, column=2, padx=8, pady=(6, 12))

        actions = ctk.CTkFrame(self, corner_radius=0)
        actions.grid(row=1, column=0, sticky="ew")
        actions.grid_columnconfigure(5, weight=1)

        ctk.CTkLabel(actions, textvariable=self.summary_text).grid(row=0, column=0, padx=12, pady=10, sticky="w")
        ctk.CTkButton(actions, text="Generar PDFs", command=self.generate).grid(row=0, column=1, padx=8, pady=10)
        ctk.CTkButton(actions, text="Abrir salida", command=self.open_output).grid(row=0, column=2, padx=8, pady=10)
        ctk.CTkCheckBox(actions, text="Enviar a impresora", variable=self.print_after_generate).grid(row=0, column=3, padx=8, pady=10)
        self.printer_menu = ctk.CTkOptionMenu(actions, variable=self.printer_name, values=["No disponible"])
        self.printer_menu.grid(row=0, column=4, padx=8, pady=10)

        grid_frame = ctk.CTkFrame(self, corner_radius=0)
        grid_frame.grid(row=2, column=0, sticky="nsew")
        grid_frame.grid_rowconfigure(0, weight=1)
        grid_frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Treeview", rowheight=24)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self.tree = ttk.Treeview(grid_frame, columns=VISIBLE_COLUMNS, show="headings")
        self.tree.tag_configure("even_row", background="#FFFFFF")
        self.tree.tag_configure("odd_row", background="#EEF6FF")
        for column in VISIBLE_COLUMNS:
            self.tree.heading(column, text=column)
            width = 135
            if column in {"DIRECCION DE ENTREGA REAL", "DESCRIPCIÓN DE SUMINISTRO", "TITULAR PAGO REAL"}:
                width = 245
            self.tree.column(column, width=width, minwidth=90, anchor="w", stretch=True)

        y_scroll = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        bottom = ctk.CTkFrame(self, corner_radius=0)
        bottom.grid(row=3, column=0, sticky="ew")
        bottom.grid_columnconfigure(3, weight=1)
        ctk.CTkButton(bottom, text="Anterior", command=self.previous_page).grid(row=0, column=0, padx=12, pady=10)
        ctk.CTkButton(bottom, text="Siguiente", command=self.next_page).grid(row=0, column=1, padx=4, pady=10)
        self.page_label = ctk.CTkLabel(bottom, text="Página 0 de 0")
        self.page_label.grid(row=0, column=2, padx=12, pady=10)
        ctk.CTkLabel(bottom, textvariable=self.status_text).grid(row=0, column=3, padx=12, pady=10, sticky="e")
        ctk.CTkButton(bottom, text="Salir", command=self.destroy, fg_color="#6B7280", hover_color="#4B5563").grid(row=0, column=4, padx=12, pady=10)

    def select_excel(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Seleccionar libro de Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos los archivos", "*.*")],
        )
        if file_path:
            self.excel_path.set(file_path)
            self.excel_display.set(self._display_path(file_path))

    def select_output(self) -> None:
        folder = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if folder:
            self.output_path.set(folder)
            self.output_display.set(self._display_path(folder))

    def select_logo(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Seleccionar logo",
            filetypes=[("Imagen", "*.png *.jpg *.jpeg"), ("Todos los archivos", "*.*")],
        )
        if file_path:
            self.logo_path.set(file_path)
            self.logo_display.set(self._display_path(file_path))

    def load_excel(self) -> None:
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showwarning("Archivo requerido", "Seleccione primero un archivo de Excel.")
            return

        try:
            self.status_text.set("Cargando archivo...")
            self.update_idletasks()
            self.records = load_excel_records(path)
            self.current_page = 1
            self._refresh_grid()
            self._refresh_summary()
            self.status_text.set("Archivo cargado correctamente.")
        except Exception as exc:
            self.records = []
            self._refresh_grid()
            self._refresh_summary()
            messagebox.showerror("Error al cargar", str(exc))
            self.status_text.set("No se pudo cargar el archivo.")

    def generate(self) -> None:
        if not self.records:
            messagebox.showwarning("Sin datos", "Cargue un Excel antes de generar los PDFs.")
            return

        try:
            self.status_text.set("Generando PDFs...")
            self.update_idletasks()
            self.generated_files = generate_pdfs(
                self.records,
                self.output_path.get().strip() or DEFAULT_OUTPUT,
                self.logo_path.get().strip() or None,
                DEFAULT_ALIADO_COMERCIAL,
            )
            if self.print_after_generate.get():
                self._print_generated_files()
            self.status_text.set(f"Generados {len(self.generated_files)} PDF(s).")
            messagebox.showinfo("Proceso terminado", f"Se generaron {len(self.generated_files)} PDF(s).")
        except Exception as exc:
            messagebox.showerror("Error al generar", str(exc))
            self.status_text.set("No se pudieron generar los PDFs.")

    def previous_page(self) -> None:
        if self.current_page > 1:
            self.current_page -= 1
            self._refresh_grid()

    def next_page(self) -> None:
        if self.current_page < self._page_count():
            self.current_page += 1
            self._refresh_grid()

    def open_output(self) -> None:
        folder = Path(self.output_path.get().strip() or DEFAULT_OUTPUT)
        folder.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(folder)
        else:
            subprocess.Popen(["xdg-open", str(folder)])

    def _refresh_grid(self) -> None:
        self.tree.delete(*self.tree.get_children())
        start = (self.current_page - 1) * PAGE_SIZE
        end = start + PAGE_SIZE
        for index, record in enumerate(self.records[start:end], start=start):
            tag = "odd_row" if index % 2 else "even_row"
            self.tree.insert("", "end", values=record.visible_row(), tags=(tag,))
        self.page_label.configure(text=f"Página {self.current_page if self.records else 0} de {self._page_count()}")

    def _refresh_summary(self) -> None:
        groups = group_records(self.records) if self.records else []
        pages = sum(group.page_count for group in groups)
        self.summary_text.set(f"{len(self.records)} registros | {len(groups)} formatos | {pages} páginas PDF")

    def _page_count(self) -> int:
        return max(1, (len(self.records) + PAGE_SIZE - 1) // PAGE_SIZE) if self.records else 0

    def _load_printers(self) -> None:
        try:
            import win32print

            printers = [printer[2] for printer in win32print.EnumPrinters(2)]
            if printers:
                self.printer_menu.configure(values=printers)
                self.printer_name.set(win32print.GetDefaultPrinter() if win32print.GetDefaultPrinter() in printers else printers[0])
                return
        except Exception:
            pass
        self.printer_menu.configure(values=["No disponible"])
        self.printer_name.set("No disponible")

    def _print_generated_files(self) -> None:
        printer = self.printer_name.get()
        if not printer or printer == "No disponible":
            messagebox.showwarning(
                "Impresora no disponible",
                "Se generaron los PDFs, pero no se encontró soporte de impresoras. Instale pywin32 o imprima desde la carpeta de salida.",
            )
            return

        if not sys.platform.startswith("win"):
            messagebox.showwarning("Impresión no disponible", "La impresión automática está implementada para Windows.")
            return

        for pdf in self.generated_files:
            os.startfile(str(pdf), "print")


if __name__ == "__main__":
    app = PrintFormatterApp()
    app.mainloop()

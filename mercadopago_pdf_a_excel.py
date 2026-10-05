#!/usr/bin/env python3
"""Convierte resumenes PDF de Mercado Pago a Excel filtrado."""

from __future__ import annotations

import argparse
import re
import sys
import threading
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")
DATE_START_RE = re.compile(r"^(\d{2}-\d{2}-\d{4})(?:\s+(.*))?$")
ID_RE = re.compile(r"^\d{10,15}$")
ID_ANYWHERE_RE = re.compile(r"\b\d{10,15}\b")
MONEY_RE = re.compile(r"\$\s*-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}(?!\d)")
PAGE_RE = re.compile(r"^\d+/\d+$")
HEADER_LINES = {
    "fecha",
    "descripcion",
    "descripción",
    "id de la",
    "operacion",
    "operación",
    "valor",
    "saldo",
    "detalle de movimientos",
}


@dataclass
class Movement:
    titular: str
    periodo: str
    fecha: str
    tipo: str
    descripcion: str
    contraparte: str
    valor: float


class MercadoPagoPdfError(Exception):
    """Error controlado para mensajes claros al usuario."""


def require_dependencies() -> None:
    missing = []
    try:
        import fitz  # noqa: F401
    except ImportError:
        missing.append("pymupdf")
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        missing.append("openpyxl")

    if missing:
        deps = " ".join(missing)
        raise MercadoPagoPdfError(
            f"Faltan dependencias: {deps}. Instalalas con: pip install pymupdf openpyxl"
        )


def extract_text_from_pdf(pdf_path: Path) -> str:
    try:
        import fitz

        with fitz.open(pdf_path) as doc:
            if doc.page_count == 0:
                raise MercadoPagoPdfError(f"PDF sin paginas: {pdf_path}")
            return "\n".join(page.get_text("text") for page in doc)
    except MercadoPagoPdfError:
        raise
    except Exception as exc:
        raise MercadoPagoPdfError(f"No se pudo leer el PDF '{pdf_path}': {exc}") from exc


def normalize_text(text: str) -> str:
    text = " ".join(str(text).split())
    return text.strip()


def normalize_for_compare(text: str) -> str:
    text = normalize_text(text).casefold()
    text = "".join(
        ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn"
    )
    return text


def parse_money(value: str) -> float:
    cleaned = value.replace("$", "").replace(" ", "").replace(".", "").replace(",", ".")
    return float(cleaned)


def extract_header_data(text: str) -> tuple[str, str]:
    lines = [normalize_text(line) for line in text.splitlines() if normalize_text(line)]
    titular = ""
    periodo = ""

    for index, line in enumerate(lines):
        if "CVU:" in line.upper():
            for previous in reversed(lines[:index]):
                previous_cmp = normalize_for_compare(previous)
                if previous_cmp and previous_cmp not in {
                    "resumen de cuenta en pesos",
                    "periodo:",
                } and not PAGE_RE.match(previous):
                    titular = previous
                    break
            break

    for index, line in enumerate(lines):
        if normalize_for_compare(line).startswith("periodo:"):
            after = normalize_text(line.split(":", 1)[1]) if ":" in line else ""
            if after:
                periodo = after
            else:
                candidates = []
                if index + 1 < len(lines):
                    candidates.append(lines[index + 1])
                if index > 0:
                    candidates.append(lines[index - 1])
                for candidate in candidates:
                    cmp_candidate = normalize_for_compare(candidate)
                    if cmp_candidate.startswith("del "):
                        periodo = candidate
                        break
            break

    return titular, periodo


def _clean_extracted_lines(text: str) -> list[str]:
    lines: list[str] = []
    for raw_line in text.splitlines():
        line = normalize_text(raw_line)
        if not line:
            continue
        cmp_line = normalize_for_compare(line)
        if PAGE_RE.match(line) or cmp_line in HEADER_LINES:
            continue
        lines.append(line)
    return lines


def split_movements(text: str) -> list[str]:
    lines = _clean_extracted_lines(text)
    blocks: list[list[str]] = []
    current: list[str] = []
    in_detail = False

    for line in lines:
        if normalize_for_compare(line) == "detalle de movimientos":
            in_detail = True
            continue
        date_match = DATE_START_RE.match(line)
        if date_match:
            in_detail = True
            if current:
                blocks.append(current)
            current = [date_match.group(1)]
            if date_match.group(2):
                current.append(date_match.group(2))
        elif in_detail and current:
            current.append(line)

    if current:
        blocks.append(current)

    return ["\n".join(block) for block in blocks]


def _parse_movement_block(block: str, titular: str, periodo: str) -> Movement | None:
    lines = [normalize_text(line) for line in block.splitlines() if normalize_text(line)]
    if not lines or not DATE_RE.match(lines[0]):
        return None

    fecha = lines[0]
    rest = lines[1:]
    id_index = next((i for i, line in enumerate(rest) if ID_RE.match(line)), None)
    if id_index is None:
        joined_rest = normalize_text(" ".join(rest))
        id_match = ID_ANYWHERE_RE.search(joined_rest)
        if not id_match:
            return None
        descripcion = normalize_text(joined_rest[: id_match.start()])
        money_values = MONEY_RE.findall(joined_rest[id_match.end() :])
    else:
        descripcion = normalize_text(" ".join(rest[:id_index]))
        money_values = MONEY_RE.findall(" ".join(rest[id_index + 1 :]))
    if not descripcion or not money_values:
        return None

    valor = parse_money(money_values[0])
    cmp_desc = normalize_for_compare(descripcion)

    if cmp_desc == "liquidacion de dinero":
        return Movement(
            titular=titular,
            periodo=periodo,
            fecha=fecha,
            tipo="Liquidación de dinero",
            descripcion="Liquidación de dinero",
            contraparte="",
            valor=valor,
        )

    transfer_prefix = "transferencia recibida"
    if cmp_desc.startswith(transfer_prefix):
        prefix_len = len("Transferencia recibida")
        contraparte = normalize_text(descripcion[prefix_len:])
        return Movement(
            titular=titular,
            periodo=periodo,
            fecha=fecha,
            tipo="Transferencia recibida",
            descripcion=descripcion,
            contraparte=contraparte,
            valor=valor,
        )

    return None


def parse_movements(text: str, titular: str, periodo: str) -> list[Movement]:
    return [
        movement
        for block in split_movements(text)
        if (movement := _parse_movement_block(block, titular, periodo)) is not None
    ]


def filter_valid_movements(movements: Iterable[Movement]) -> list[Movement]:
    valid_types = {"Liquidación de dinero", "Transferencia recibida"}
    return [movement for movement in movements if movement.tipo in valid_types]


def extract_entradas_total(text: str) -> float | None:
    match = re.search(r"Entradas:\s*(" + MONEY_RE.pattern + r")", text, re.IGNORECASE)
    return parse_money(match.group(1)) if match else None


def process_pdf(pdf_path: Path) -> tuple[list[Movement], list[str]]:
    if not pdf_path.exists():
        raise MercadoPagoPdfError(f"No existe el archivo: {pdf_path}")

    text = extract_text_from_pdf(pdf_path)
    titular, periodo = extract_header_data(text)
    if not split_movements(text):
        raise MercadoPagoPdfError(f"PDF sin movimientos detectados: {pdf_path.name}")

    movements = filter_valid_movements(parse_movements(text, titular, periodo))
    if not movements:
        raise MercadoPagoPdfError(f"No se encontraron movimientos válidos en: {pdf_path.name}")

    warnings: list[str] = []
    entradas_total = extract_entradas_total(text)
    if entradas_total is not None:
        extracted_total = sum(movement.valor for movement in movements)
        diff = extracted_total - entradas_total
        if abs(diff) > 0.01:
            warnings.append(
                f"{pdf_path.name}: Entradas del resumen {entradas_total:.2f}; "
                f"movimientos exportados {extracted_total:.2f}; diferencia {diff:.2f}."
            )

    return movements, warnings


def get_safe_output_path(pdf_path: Path, output_folder: Path) -> Path:
    output_folder = Path(output_folder)
    base_name = f"{pdf_path.stem}_movimientos"
    candidate = output_folder / f"{base_name}.xlsx"
    counter = 2
    while candidate.exists():
        candidate = output_folder / f"{base_name}_{counter}.xlsx"
        counter += 1
    return candidate


def _common_value(movements: list[Movement], field_name: str) -> str:
    values = {
        normalize_text(str(getattr(movement, field_name)))
        for movement in movements
        if normalize_text(str(getattr(movement, field_name)))
    }
    if not values:
        return ""
    if len(values) == 1:
        return next(iter(values))
    return "Varios"


def _add_movement_sheet(wb, title: str, movements: list[Movement], header_fill, header_font) -> None:
    from openpyxl.styles import Alignment, Font

    ws = wb.active if wb.active.title == "Sheet" else wb.create_sheet(title)
    ws.title = title

    ws.merge_cells("A1:E1")
    ws.merge_cells("A2:E2")
    ws["A1"] = f"Titular: {_common_value(movements, 'titular')}"
    ws["A2"] = f"Periodo: {_common_value(movements, 'periodo')}"
    for cell in (ws["A1"], ws["A2"]):
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="left")

    headers = ["Fecha", "Tipo", "Descripci\u00f3n", "Contraparte", "Valor"]
    header_row = 4
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for movement in movements:
        ws.append(
            [
                movement.fecha,
                movement.tipo,
                movement.descripcion,
                movement.contraparte,
                movement.valor,
            ]
        )

    ws.auto_filter.ref = f"A{header_row}:E{ws.max_row}"
    ws.freeze_panes = "A5"
    for row in range(header_row + 1, ws.max_row + 1):
        ws.cell(row=row, column=5).number_format = '"$"#,##0.00'

    widths = {
        "A": 14,
        "B": 24,
        "C": 56,
        "D": 38,
        "E": 16,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width


def build_excel(movements: list[Movement], output_path: Path) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError as exc:
        raise MercadoPagoPdfError(
            "Falta openpyxl. Instalalo con: pip install openpyxl"
        ) from exc

    if not movements:
        raise MercadoPagoPdfError("No se detectaron movimientos validos para exportar.")

    wb = Workbook()
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    _add_movement_sheet(wb, "Movimientos", movements, header_fill, header_font)
    _add_movement_sheet(
        wb,
        "Liquidación de dinero",
        [movement for movement in movements if movement.tipo == "Liquidación de dinero"],
        header_fill,
        header_font,
    )
    _add_movement_sheet(
        wb,
        "Transferencia recibida",
        [movement for movement in movements if movement.tipo == "Transferencia recibida"],
        header_fill,
        header_font,
    )

    summary = wb.create_sheet("Resumen")
    summary.append(["Tipo", "Cantidad", "Total"])
    grouped: dict[str, dict[str, float]] = defaultdict(lambda: {"cantidad": 0, "total": 0.0})
    for movement in movements:
        grouped[movement.tipo]["cantidad"] += 1
        grouped[movement.tipo]["total"] += movement.valor

    for tipo in ["Liquidación de dinero", "Transferencia recibida"]:
        values = grouped[tipo]
        summary.append([tipo, int(values["cantidad"]), values["total"]])

    summary.append(
        [
            "Total general",
            len(movements),
            sum(movement.valor for movement in movements),
        ]
    )

    for cell in summary[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    summary.auto_filter.ref = summary.dimensions
    summary.freeze_panes = "A2"
    for row in range(2, summary.max_row + 1):
        summary.cell(row=row, column=3).number_format = '"$"#,##0.00'
    for col_idx in range(1, summary.max_column + 1):
        max_len = max(
            len(str(summary.cell(row=row, column=col_idx).value or ""))
            for row in range(1, summary.max_row + 1)
        )
        summary.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 32)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(output_path)
    except Exception as exc:
        raise MercadoPagoPdfError(f"No se pudo guardar el Excel '{output_path}': {exc}") from exc


def select_files_with_tkinter() -> list[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception as exc:
        raise MercadoPagoPdfError(f"No se pudo abrir Tkinter: {exc}") from exc

    root = tk.Tk()
    root.withdraw()
    files = filedialog.askopenfilenames(
        title="Seleccionar resumenes de Mercado Pago",
        filetypes=[("PDF", "*.pdf"), ("Todos los archivos", "*.*")],
    )
    if not files:
        messagebox.showinfo("Mercado Pago a Excel", "No se selecciono ningun archivo.")
        return []
    return [Path(file) for file in files]


def select_output_with_tkinter() -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception as exc:
        raise MercadoPagoPdfError(f"No se pudo abrir Tkinter: {exc}") from exc

    root = tk.Tk()
    root.withdraw()
    output = filedialog.asksaveasfilename(
        title="Guardar Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
    )
    if not output:
        messagebox.showinfo("Mercado Pago a Excel", "No se eligio donde guardar el Excel.")
        return None
    return Path(output)


def generate_separate_excels(pdf_paths: Iterable[Path], output_folder: Path) -> tuple[int, int, list[str]]:
    generated = 0
    errors = 0
    logs: list[str] = []
    output_folder.mkdir(parents=True, exist_ok=True)

    for pdf_path in pdf_paths:
        try:
            movements, warnings = process_pdf(pdf_path)
            output_path = get_safe_output_path(pdf_path, output_folder)
            build_excel(movements, output_path)
            generated += 1
            logs.append(f"OK: {pdf_path.name} - {len(movements)} movimientos.")
            logs.append(f"Excel generado: {output_path}")
            logs.extend(f"Advertencia: {warning}" for warning in warnings)
        except Exception as exc:
            errors += 1
            logs.append(f"Error en {pdf_path.name}: {exc}")

    return generated, errors, logs


class MercadoPagoExtractorApp:
    def __init__(self) -> None:
        try:
            import tkinter as tk
            from tkinter import ttk
        except Exception as exc:
            raise MercadoPagoPdfError(f"No se pudo abrir Tkinter: {exc}") from exc

        self.tk = tk
        self.ttk = ttk
        self.root = tk.Tk()
        self.root.title("Extractor Mercado Pago a Excel")
        self.root.geometry("820x560")
        self.pdf_paths: list[Path] = []
        self.output_folder: Path | None = None
        self.build_ui()

    def build_ui(self) -> None:
        tk = self.tk
        ttk = self.ttk

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)
        self.root.rowconfigure(6, weight=2)

        top_frame = ttk.Frame(self.root, padding=12)
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(1, weight=1)

        self.select_files_button = ttk.Button(
            top_frame, text="Seleccionar PDFs", command=self.select_pdf_files
        )
        self.select_files_button.grid(
            row=0, column=0, padx=(0, 8), sticky="w"
        )
        self.select_folder_button = ttk.Button(
            top_frame,
            text="Seleccionar carpeta destino",
            command=self.select_output_folder,
        )
        self.select_folder_button.grid(row=0, column=1, padx=(0, 8), sticky="w")
        self.clear_button = ttk.Button(
            top_frame, text="Limpiar selección", command=self.clear_selection
        )
        self.clear_button.grid(
            row=0, column=2, sticky="e"
        )

        ttk.Label(self.root, text="PDFs seleccionados:", padding=(12, 0)).grid(
            row=1, column=0, sticky="w"
        )
        self.files_list = tk.Listbox(self.root, height=7)
        self.files_list.grid(row=2, column=0, padx=12, pady=(4, 10), sticky="nsew")

        self.folder_label = ttk.Label(
            self.root,
            text="Carpeta destino: no seleccionada",
            padding=(12, 0),
        )
        self.folder_label.grid(row=3, column=0, sticky="ew")

        action_frame = ttk.Frame(self.root, padding=12)
        action_frame.grid(row=4, column=0, sticky="ew")
        action_frame.columnconfigure(0, weight=1)
        self.generate_button = ttk.Button(
            action_frame,
            text="Generar Excels",
            command=self.generate_excels,
        )
        self.generate_button.grid(row=0, column=0, sticky="w")

        self.progress = ttk.Progressbar(action_frame, mode="determinate")
        self.progress.grid(row=0, column=1, padx=(12, 0), sticky="ew")
        action_frame.columnconfigure(1, weight=1)

        ttk.Label(self.root, text="Logs:", padding=(12, 0)).grid(row=5, column=0, sticky="w")
        log_frame = ttk.Frame(self.root, padding=(12, 4, 12, 12))
        log_frame.grid(row=6, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=12, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def select_pdf_files(self) -> None:
        from tkinter import filedialog

        files = filedialog.askopenfilenames(
            title="Seleccionar PDFs de Mercado Pago",
            filetypes=[("PDF", "*.pdf"), ("Todos los archivos", "*.*")],
        )
        if not files:
            return
        self.pdf_paths = [Path(file) for file in files]
        self.files_list.delete(0, self.tk.END)
        for pdf_path in self.pdf_paths:
            self.files_list.insert(self.tk.END, str(pdf_path))
        self.log(f"PDFs seleccionados: {len(self.pdf_paths)}")

    def select_output_folder(self) -> None:
        from tkinter import filedialog

        folder = filedialog.askdirectory(title="Seleccionar carpeta destino")
        if not folder:
            return
        self.output_folder = Path(folder)
        self.folder_label.configure(text=f"Carpeta destino: {self.output_folder}")
        self.log(f"Carpeta destino seleccionada: {self.output_folder}")

    def generate_excels(self) -> None:
        from tkinter import messagebox

        if not self.pdf_paths:
            messagebox.showwarning("Faltan PDFs", "Debes seleccionar al menos un PDF.")
            return
        if self.output_folder is None:
            messagebox.showwarning(
                "Falta carpeta destino",
                "Debes seleccionar una carpeta de destino.",
            )
            return

        pdf_paths = tuple(self.pdf_paths)
        output_folder = self.output_folder
        self.set_processing_state(True)
        self.progress.configure(maximum=len(pdf_paths), value=0)
        self.log("Iniciando proceso...")
        thread = threading.Thread(
            target=self.generate_excels_from_gui,
            args=(pdf_paths, output_folder),
            daemon=True,
        )
        thread.start()

    def set_processing_state(self, processing: bool) -> None:
        state = "disabled" if processing else "normal"
        for button in (
            self.generate_button,
            self.select_files_button,
            self.select_folder_button,
            self.clear_button,
        ):
            button.configure(state=state)

    def generate_excels_from_gui(self, pdf_paths: tuple[Path, ...], output_folder: Path) -> None:
        generated = 0
        errors = 0
        total = len(pdf_paths)

        for index, pdf_path in enumerate(pdf_paths, 1):
            self.log_message(f"Procesando: {pdf_path.name}")
            try:
                movements, warnings = process_pdf(pdf_path)
                output_path = get_safe_output_path(pdf_path, output_folder)
                build_excel(movements, output_path)
                generated += 1
                self.log_message(f"Archivo procesado correctamente: {pdf_path.name}")
                self.log_message(f"Cantidad de movimientos encontrados: {len(movements)}")
                self.log_message(f"Ruta del Excel generado: {output_path}")
                for warning in warnings:
                    self.log_message(f"Advertencia: {warning}")
            except Exception as exc:
                errors += 1
                self.log_message(f"Archivo omitido por error: {pdf_path.name}")
                self.log_message(str(exc))
            self.update_progress(index, total)

        self.log_message(
            f"Proceso finalizado. Archivos generados: {generated}. Archivos con error: {errors}."
        )
        self.root.after(0, self.set_processing_state, False)

    def log(self, message: str) -> None:
        self.log_text.insert(self.tk.END, f"{message}\n")
        self.log_text.see(self.tk.END)

    def log_message(self, message: str) -> None:
        self.root.after(0, self.log, message)

    def update_progress(self, current: int, total: int) -> None:
        self.root.after(0, lambda: self.progress.configure(maximum=total, value=current))

    def clear_selection(self) -> None:
        self.pdf_paths = []
        self.output_folder = None
        self.files_list.delete(0, self.tk.END)
        self.folder_label.configure(text="Carpeta destino: no seleccionada")
        self.progress.configure(value=0)
        self.log_text.delete("1.0", self.tk.END)

    def run(self) -> None:
        self.root.mainloop()


def process_pdfs(pdf_paths: Iterable[Path]) -> tuple[list[Movement], list[str]]:
    all_movements: list[Movement] = []
    warnings: list[str] = []

    for pdf_path in pdf_paths:
        try:
            movements, pdf_warnings = process_pdf(pdf_path)
        except MercadoPagoPdfError as exc:
            warnings.append(str(exc))
            continue
        all_movements.extend(movements)
        warnings.extend(pdf_warnings)

    if not all_movements:
        raise MercadoPagoPdfError("No se detectaron movimientos validos en los PDFs seleccionados.")

    return all_movements, warnings


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Lee resumenes PDF de Mercado Pago y genera un Excel filtrado."
    )
    parser.add_argument("pdfs", nargs="*", type=Path, help="Uno o varios archivos PDF.")
    parser.add_argument("-o", "--output", type=Path, help="Ruta del archivo .xlsx de salida.")
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Carpeta donde guardar un Excel separado por cada PDF.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    try:
        require_dependencies()
        args = parse_args(argv)

        if not args.pdfs:
            MercadoPagoExtractorApp().run()
            return 0

        if args.output and len(args.pdfs) > 1:
            raise MercadoPagoPdfError("No puedes usar -o con múltiples PDFs. Usa --output-dir.")

        if args.output_dir and args.output:
            raise MercadoPagoPdfError("Usa -o o --output-dir, no ambos.")

        if args.output:
            movements, warnings = process_pdf(args.pdfs[0])
            build_excel(movements, args.output)
            print(f"Excel generado: {args.output}")
            print(f"Movimientos exportados: {len(movements)}")
            for warning in warnings:
                print(f"Advertencia: {warning}")
            return 0

        output_folder = args.output_dir or Path.cwd()
        generated, errors, logs = generate_separate_excels(args.pdfs, output_folder)
        for log in logs:
            print(log)
        print(f"Proceso finalizado. Archivos generados: {generated}. Archivos con error: {errors}.")
        if errors:
            return 1
        return 0
    except MercadoPagoPdfError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())


# Instrucciones de uso:
# pip install pymupdf openpyxl
# python mercadopago_pdf_a_excel.py archivo.pdf -o movimientos.xlsx
# python mercadopago_pdf_a_excel.py archivo1.pdf archivo2.pdf -o movimientos.xlsx
# pyinstaller --onefile --noconsole mercadopago_pdf_a_excel.py

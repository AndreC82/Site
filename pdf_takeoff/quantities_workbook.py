"""Gera a planilha de quantidades/orçamento no mesmo formato usado pelo estimador:
agrupada por tipo de forro (Fireline, Aqualine, Standard...) e altura de parede,
com 3 taxas por linha (GIB, Stopping, Pintura) em $/m², tabela de preços
(Custo x margem = Venda) e uma aba Resumo com GIB/Plaster/Painting/Total,
Margem, P&G e Contingência — espelhando a estrutura de referência do usuário.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_BOLD = Font(bold=True)

# Colunas fixas da aba "Quantities" (iguais em espírito ao arquivo de referência).
COL_DESC = 1  # A
COL_QTY = 2  # B
COL_QTY_UNIT = 3  # C
COL_HEIGHT = 4  # D
COL_TOTAL = 5  # E
COL_TOTAL_UNIT = 6  # F
COL_GIB_RATE = 8  # H
COL_GIB_TOTAL = 9  # I
COL_STOP_RATE = 11  # K
COL_STOP_TOTAL = 12  # L
COL_PAINT_RATE = 14  # N
COL_PAINT_TOTAL = 15  # O
COL_ROW_TOTAL = 17  # Q
COL_SALE_LABEL = 19  # S
COL_SALE_RATE = 20  # T
COL_COST_LABEL = 22  # V
COL_COST_RATE = 23  # W


@dataclass
class RateEntry:
    label: str
    cost_rate: float  # $/m² (ou $/m, ou $/unidade, conforme o item)


@dataclass
class LiningItem:
    description: str  # ex.: "1x 10mm Standard", "2x 19mm Fireline"
    qty: float  # metros lineares de parede com esse tipo de forro
    qty_unit: str  # 'm' (linear) ou 'ea' (contagem, p.ex. cantoneiras)
    rate_label: str  # rótulo na tabela de preços (GIB) a referenciar
    layers: int = 1  # multiplicador de camada (1 = single layer, 2 = double layer)
    stop_rate_label: str | None = "STOPPING LV4"
    paint_rate_label: str | None = "PAINTING"
    total_unit: str = "m2"


@dataclass
class HeightGroup:
    name: str  # ex.: "LININGS - 2.70m"
    height_m: float
    items: list[LiningItem] = field(default_factory=list)
    corner_trims_qty: float = 0.0  # contagem de cantoneiras (ea)
    sealant_qty: float = 0.0  # metros lineares de selante de rodapé


@dataclass
class CeilingItem:
    description: str  # ex.: "13mm Standard Gib board (C1)"
    area_m2: float
    rate_label: str
    stop_rate_label: str | None = "SQUARE STOP"
    paint_rate_label: str | None = "PAINTING"


@dataclass
class DoorPaintItem:
    description: str  # ex.: "D1 - Single Swing"
    count: int
    rate_label: str  # "Single Door" ou "Double Door"


DEFAULT_RATE_TABLE: list[RateEntry] = [
    RateEntry("10mm STANDARD", 0.0),
    RateEntry("10mm AQUALINE", 0.0),
    RateEntry("13mm STANDARD", 0.0),
    RateEntry("13mm AQUALINE", 0.0),
    RateEntry("13mm FIRELINE", 0.0),
    RateEntry("13mm NOISELINE", 0.0),
    RateEntry("16mm FIRELINE", 0.0),
    RateEntry("19mm Fireline", 0.0),
    RateEntry("GIB Sealant", 0.0),
    RateEntry("STOPPING LV4", 0.0),
    RateEntry("SQUARE STOP", 0.0),
    RateEntry("CORNER BEADS", 0.0),
    RateEntry("Painting Skirting", 0.0),
    RateEntry("PAINTING", 0.0),
    RateEntry("Single Door", 0.0),
    RateEntry("Double Door", 0.0),
]


def _style_header(ws: Worksheet, row: int, cols: list[int]) -> None:
    for col in cols:
        cell = ws.cell(row=row, column=col)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _section_row(ws: Worksheet, row: int, text: str) -> None:
    cell = ws.cell(row=row, column=COL_DESC, value=text)
    cell.font = _BOLD
    for col in range(1, 18):
        ws.cell(row=row, column=col).fill = _SECTION_FILL


class QuantitiesWorkbookBuilder:
    """Monta a aba 'Quantities' linha a linha, guardando os ranges usados nos
    totais de categoria para depois alimentar a aba 'Resumo' por fórmula.
    """

    def __init__(self, building_name: str = "BUILDING 01", subcontractor_margin: float = 0.25):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "Quantities"
        self.building_name = building_name
        self.subcontractor_margin = subcontractor_margin
        self._rate_rows: dict[str, int] = {}
        self._gib_total_cells: list[str] = []
        self._stop_total_cells: list[str] = []
        self._paint_total_cells: list[str] = []
        self._row = 1

    # -- tabela de preços (lado direito) -------------------------------------
    def _write_rate_table(self, rates: list[RateEntry]) -> None:
        ws = self.ws
        ws.cell(row=9, column=COL_SALE_LABEL, value="SALE - RATES").font = _BOLD
        ws.cell(row=9, column=COL_COST_LABEL, value="COST - RATES").font = _BOLD
        ws.cell(row=7, column=18, value="Subcontractor Margin").font = _BOLD
        margin_cell = f"${get_column_letter(19)}$7"
        ws.cell(row=7, column=19, value=self.subcontractor_margin).number_format = "0%"

        r = 11
        for entry in rates:
            ws.cell(row=r, column=COL_SALE_LABEL, value=entry.label)
            ws.cell(
                row=r,
                column=COL_SALE_RATE,
                value=f"={get_column_letter(COL_COST_RATE)}{r}*(1+{margin_cell})",
            ).number_format = "0.00"
            ws.cell(row=r, column=COL_COST_LABEL, value=entry.label)
            ws.cell(row=r, column=COL_COST_RATE, value=entry.cost_rate).number_format = "0.00"
            self._rate_rows[entry.label] = r
            r += 1

        for col, width in ((COL_SALE_LABEL, 18), (COL_SALE_RATE, 11), (COL_COST_LABEL, 18), (COL_COST_RATE, 11)):
            ws.column_dimensions[get_column_letter(col)].width = width

    def _sale_ref(self, label: str, multiplier: int = 1) -> str:
        row = self._rate_rows[label]
        cell = f"{get_column_letter(COL_SALE_RATE)}{row}"
        return f"={multiplier}*{cell}" if multiplier != 1 else f"={cell}"

    # -- cabeçalho ----------------------------------------------------------
    def start(self, rates: list[RateEntry]) -> None:
        ws = self.ws
        ws.cell(row=self._row, column=1, value=self.building_name).font = Font(bold=True, size=13)
        header_row = self._row + 1
        ws.cell(row=header_row, column=COL_GIB_RATE, value="GIB").font = _BOLD
        ws.cell(row=header_row, column=COL_STOP_RATE, value="Stopping").font = _BOLD
        ws.cell(row=header_row, column=COL_PAINT_RATE, value="Painting").font = _BOLD

        col_row = header_row + 1
        labels = {
            COL_DESC: "Description",
            COL_QTY: "Qty",
            COL_QTY_UNIT: "unit",
            COL_HEIGHT: "height",
            COL_TOTAL: "Total",
            COL_TOTAL_UNIT: "unit",
            COL_GIB_RATE: "Rate",
            COL_GIB_TOTAL: "Total",
            COL_STOP_RATE: "Rate",
            COL_STOP_TOTAL: "Total",
            COL_PAINT_RATE: "Rate",
            COL_PAINT_TOTAL: "Total",
            COL_ROW_TOTAL: "Row Total",
        }
        for col, text in labels.items():
            ws.cell(row=col_row, column=col, value=text)
        _style_header(ws, col_row, list(labels.keys()))

        self._write_rate_table(rates)
        self._row = col_row + 1

    # -- seção de forros de parede por altura ------------------------------
    def add_height_group(self, group: HeightGroup) -> None:
        ws = self.ws
        self._row += 1
        _section_row(ws, self._row, group.name)
        self._row += 1

        for item in group.items:
            r = self._row
            ws.cell(row=r, column=COL_DESC, value=item.description)
            ws.cell(row=r, column=COL_QTY, value=item.qty)
            ws.cell(row=r, column=COL_QTY_UNIT, value=item.qty_unit)
            ws.cell(row=r, column=COL_HEIGHT, value=group.height_m)
            ws.cell(
                row=r, column=COL_TOTAL,
                value=f"={get_column_letter(COL_QTY)}{r}*{get_column_letter(COL_HEIGHT)}{r}",
            ).number_format = "0"
            ws.cell(row=r, column=COL_TOTAL_UNIT, value=item.total_unit)

            gib_rate_cell = f"{get_column_letter(COL_GIB_RATE)}{r}"
            ws.cell(row=r, column=COL_GIB_RATE, value=self._sale_ref(item.rate_label, item.layers)).number_format = "0.00"
            ws.cell(
                row=r, column=COL_GIB_TOTAL,
                value=f"={get_column_letter(COL_TOTAL)}{r}*{gib_rate_cell}",
            ).number_format = "0.00"
            self._gib_total_cells.append(f"{get_column_letter(COL_GIB_TOTAL)}{r}")

            if item.stop_rate_label:
                stop_rate_cell = f"{get_column_letter(COL_STOP_RATE)}{r}"
                ws.cell(row=r, column=COL_STOP_RATE, value=self._sale_ref(item.stop_rate_label)).number_format = "0.00"
                ws.cell(
                    row=r, column=COL_STOP_TOTAL,
                    value=f"={get_column_letter(COL_TOTAL)}{r}*{stop_rate_cell}",
                ).number_format = "0.00"
                self._stop_total_cells.append(f"{get_column_letter(COL_STOP_TOTAL)}{r}")

            if item.paint_rate_label:
                paint_rate_cell = f"{get_column_letter(COL_PAINT_RATE)}{r}"
                ws.cell(row=r, column=COL_PAINT_RATE, value=self._sale_ref(item.paint_rate_label)).number_format = "0.00"
                ws.cell(
                    row=r, column=COL_PAINT_TOTAL,
                    value=f"={get_column_letter(COL_TOTAL)}{r}*{paint_rate_cell}",
                ).number_format = "0.00"
                self._paint_total_cells.append(f"{get_column_letter(COL_PAINT_TOTAL)}{r}")

            ws.cell(
                row=r, column=COL_ROW_TOTAL,
                value=(
                    f"=SUM({get_column_letter(COL_GIB_TOTAL)}{r},"
                    f"{get_column_letter(COL_STOP_TOTAL)}{r},"
                    f"{get_column_letter(COL_PAINT_TOTAL)}{r})"
                ),
            ).number_format = "0.00"
            self._row += 1

        if group.corner_trims_qty:
            r = self._row
            ws.cell(row=r, column=COL_DESC, value="Corner Trims")
            ws.cell(row=r, column=COL_QTY, value=group.corner_trims_qty)
            ws.cell(row=r, column=COL_QTY_UNIT, value="ea")
            ws.cell(row=r, column=COL_HEIGHT, value=group.height_m)
            ws.cell(
                row=r, column=COL_TOTAL,
                value=f"={get_column_letter(COL_QTY)}{r}*{get_column_letter(COL_HEIGHT)}{r}",
            ).number_format = "0"
            ws.cell(row=r, column=COL_TOTAL_UNIT, value="m")
            stop_rate_cell = f"{get_column_letter(COL_STOP_RATE)}{r}"
            ws.cell(row=r, column=COL_STOP_RATE, value=self._sale_ref("CORNER BEADS")).number_format = "0.00"
            ws.cell(
                row=r, column=COL_STOP_TOTAL,
                value=f"={get_column_letter(COL_TOTAL)}{r}*{stop_rate_cell}",
            ).number_format = "0.00"
            self._stop_total_cells.append(f"{get_column_letter(COL_STOP_TOTAL)}{r}")
            ws.cell(
                row=r, column=COL_ROW_TOTAL, value=f"={get_column_letter(COL_STOP_TOTAL)}{r}"
            ).number_format = "0.00"
            self._row += 1

        if group.sealant_qty:
            r = self._row
            ws.cell(row=r, column=COL_DESC, value="Sealant")
            ws.cell(row=r, column=COL_QTY, value=group.sealant_qty)
            ws.cell(row=r, column=COL_QTY_UNIT, value="m")
            ws.cell(row=r, column=COL_HEIGHT, value=2)
            ws.cell(
                row=r, column=COL_TOTAL,
                value=f"={get_column_letter(COL_QTY)}{r}*{get_column_letter(COL_HEIGHT)}{r}",
            ).number_format = "0"
            ws.cell(row=r, column=COL_TOTAL_UNIT, value="m")
            gib_rate_cell = f"{get_column_letter(COL_GIB_RATE)}{r}"
            ws.cell(row=r, column=COL_GIB_RATE, value=self._sale_ref("GIB Sealant")).number_format = "0.00"
            ws.cell(
                row=r, column=COL_GIB_TOTAL,
                value=f"={get_column_letter(COL_TOTAL)}{r}*{gib_rate_cell}",
            ).number_format = "0.00"
            self._gib_total_cells.append(f"{get_column_letter(COL_GIB_TOTAL)}{r}")
            ws.cell(
                row=r, column=COL_ROW_TOTAL, value=f"={get_column_letter(COL_GIB_TOTAL)}{r}"
            ).number_format = "0.00"
            self._row += 1

    # -- seção de tetos -------------------------------------------------------
    def add_ceilings(self, items: list[CeilingItem], square_stop_qty: float = 0.0) -> None:
        ws = self.ws
        self._row += 1
        _section_row(ws, self._row, "CEILINGS")
        self._row += 1

        for item in items:
            r = self._row
            ws.cell(row=r, column=COL_DESC, value=item.description)
            ws.cell(row=r, column=COL_QTY, value=item.area_m2)
            ws.cell(row=r, column=COL_QTY_UNIT, value="m²")
            ws.cell(row=r, column=COL_HEIGHT, value=1)
            ws.cell(
                row=r, column=COL_TOTAL,
                value=f"={get_column_letter(COL_QTY)}{r}*{get_column_letter(COL_HEIGHT)}{r}",
            ).number_format = "0"
            ws.cell(row=r, column=COL_TOTAL_UNIT, value="m2")

            gib_rate_cell = f"{get_column_letter(COL_GIB_RATE)}{r}"
            ws.cell(row=r, column=COL_GIB_RATE, value=self._sale_ref(item.rate_label)).number_format = "0.00"
            ws.cell(
                row=r, column=COL_GIB_TOTAL,
                value=f"={get_column_letter(COL_TOTAL)}{r}*{gib_rate_cell}",
            ).number_format = "0.00"
            self._gib_total_cells.append(f"{get_column_letter(COL_GIB_TOTAL)}{r}")

            if item.stop_rate_label:
                stop_rate_cell = f"{get_column_letter(COL_STOP_RATE)}{r}"
                ws.cell(row=r, column=COL_STOP_RATE, value=self._sale_ref(item.stop_rate_label)).number_format = "0.00"
                ws.cell(
                    row=r, column=COL_STOP_TOTAL,
                    value=f"={get_column_letter(COL_TOTAL)}{r}*{stop_rate_cell}",
                ).number_format = "0.00"
                self._stop_total_cells.append(f"{get_column_letter(COL_STOP_TOTAL)}{r}")

            if item.paint_rate_label:
                paint_rate_cell = f"{get_column_letter(COL_PAINT_RATE)}{r}"
                ws.cell(row=r, column=COL_PAINT_RATE, value=self._sale_ref(item.paint_rate_label)).number_format = "0.00"
                ws.cell(
                    row=r, column=COL_PAINT_TOTAL,
                    value=f"={get_column_letter(COL_TOTAL)}{r}*{paint_rate_cell}",
                ).number_format = "0.00"
                self._paint_total_cells.append(f"{get_column_letter(COL_PAINT_TOTAL)}{r}")

            ws.cell(
                row=r, column=COL_ROW_TOTAL,
                value=(
                    f"=SUM({get_column_letter(COL_GIB_TOTAL)}{r},"
                    f"{get_column_letter(COL_STOP_TOTAL)}{r},"
                    f"{get_column_letter(COL_PAINT_TOTAL)}{r})"
                ),
            ).number_format = "0.00"
            self._row += 1

        if square_stop_qty:
            r = self._row
            ws.cell(row=r, column=COL_DESC, value="Square Stop")
            ws.cell(row=r, column=COL_QTY, value=square_stop_qty)
            ws.cell(row=r, column=COL_QTY_UNIT, value="m")
            ws.cell(row=r, column=COL_HEIGHT, value=1)
            ws.cell(
                row=r, column=COL_TOTAL,
                value=f"={get_column_letter(COL_QTY)}{r}*{get_column_letter(COL_HEIGHT)}{r}",
            ).number_format = "0"
            ws.cell(row=r, column=COL_TOTAL_UNIT, value="m")
            stop_rate_cell = f"{get_column_letter(COL_STOP_RATE)}{r}"
            ws.cell(row=r, column=COL_STOP_RATE, value=self._sale_ref("SQUARE STOP")).number_format = "0.00"
            ws.cell(
                row=r, column=COL_STOP_TOTAL,
                value=f"={get_column_letter(COL_TOTAL)}{r}*{stop_rate_cell}",
            ).number_format = "0.00"
            self._stop_total_cells.append(f"{get_column_letter(COL_STOP_TOTAL)}{r}")
            ws.cell(
                row=r, column=COL_ROW_TOTAL, value=f"={get_column_letter(COL_STOP_TOTAL)}{r}"
            ).number_format = "0.00"
            self._row += 1

    # -- seção de pintura avulsa (rodapé, portas) ------------------------------
    def add_painting_only(self, skirting_m: float, doors: list[DoorPaintItem]) -> None:
        ws = self.ws
        self._row += 1
        _section_row(ws, self._row, "PAINTING")
        self._row += 1

        if skirting_m:
            r = self._row
            ws.cell(row=r, column=COL_DESC, value="60x10 Bevelled Edge Skirting")
            ws.cell(row=r, column=COL_QTY, value=skirting_m)
            ws.cell(row=r, column=COL_QTY_UNIT, value="m")
            ws.cell(row=r, column=COL_HEIGHT, value=1)
            ws.cell(
                row=r, column=COL_TOTAL,
                value=f"={get_column_letter(COL_QTY)}{r}*{get_column_letter(COL_HEIGHT)}{r}",
            ).number_format = "0"
            ws.cell(row=r, column=COL_TOTAL_UNIT, value="m")
            paint_rate_cell = f"{get_column_letter(COL_PAINT_RATE)}{r}"
            ws.cell(row=r, column=COL_PAINT_RATE, value=self._sale_ref("Painting Skirting")).number_format = "0.00"
            ws.cell(
                row=r, column=COL_PAINT_TOTAL,
                value=f"={get_column_letter(COL_TOTAL)}{r}*{paint_rate_cell}",
            ).number_format = "0.00"
            self._paint_total_cells.append(f"{get_column_letter(COL_PAINT_TOTAL)}{r}")
            ws.cell(
                row=r, column=COL_ROW_TOTAL, value=f"={get_column_letter(COL_PAINT_TOTAL)}{r}"
            ).number_format = "0.00"
            self._row += 1

        for door in doors:
            r = self._row
            ws.cell(row=r, column=COL_DESC, value=door.description)
            ws.cell(row=r, column=COL_QTY, value=door.count)
            ws.cell(row=r, column=COL_QTY_UNIT, value="ea")
            ws.cell(row=r, column=COL_HEIGHT, value=1)
            ws.cell(
                row=r, column=COL_TOTAL,
                value=f"={get_column_letter(COL_QTY)}{r}*{get_column_letter(COL_HEIGHT)}{r}",
            ).number_format = "0"
            ws.cell(row=r, column=COL_TOTAL_UNIT, value="each")
            paint_rate_cell = f"{get_column_letter(COL_PAINT_RATE)}{r}"
            ws.cell(row=r, column=COL_PAINT_RATE, value=self._sale_ref(door.rate_label)).number_format = "0.00"
            ws.cell(
                row=r, column=COL_PAINT_TOTAL,
                value=f"={get_column_letter(COL_TOTAL)}{r}*{paint_rate_cell}",
            ).number_format = "0.00"
            self._paint_total_cells.append(f"{get_column_letter(COL_PAINT_TOTAL)}{r}")
            ws.cell(
                row=r, column=COL_ROW_TOTAL, value=f"={get_column_letter(COL_PAINT_TOTAL)}{r}"
            ).number_format = "0.00"
            self._row += 1

    # -- aba Resumo -----------------------------------------------------------
    def add_summary_sheet(self) -> None:
        ws = self.wb.create_sheet("Summary", 0)

        def joined(cells: list[str]) -> str:
            if not cells:
                return "=0"
            refs = ",".join(f"'Quantities'!{c}" for c in cells)
            return f"=SUM({refs})"

        ws.cell(row=2, column=4, value="TOTAL PROJECT").font = Font(bold=True, size=13)
        ws.cell(row=4, column=3, value="GIB").font = _BOLD
        ws.cell(row=4, column=4, value="PLASTER").font = _BOLD
        ws.cell(row=4, column=5, value="PAINTING").font = _BOLD
        ws.cell(row=4, column=6, value="TOTAL").font = _BOLD

        gib_ref = joined(self._gib_total_cells)
        stop_ref = joined(self._stop_total_cells)
        paint_ref = joined(self._paint_total_cells)

        ws.cell(row=5, column=2, value="TOTAL").font = _BOLD
        ws.cell(row=5, column=3, value=gib_ref).number_format = "#,##0.00"
        ws.cell(row=5, column=4, value=stop_ref).number_format = "#,##0.00"
        ws.cell(row=5, column=5, value=paint_ref).number_format = "#,##0.00"
        ws.cell(row=5, column=6, value="=SUM(C5:E5)").number_format = "#,##0.00"

        ws.cell(row=8, column=2, value="GROSS MARGIN (%)")
        ws.cell(row=8, column=3, value=0.20).number_format = "0%"
        ws.cell(row=9, column=2, value="P&G ($)")
        ws.cell(row=9, column=3, value=0)
        ws.cell(row=10, column=2, value="CONTINGENCY (%)")
        ws.cell(row=10, column=3, value=0.05).number_format = "0%"
        ws.cell(row=12, column=2, value="PREÇO FINAL DE VENDA").font = _BOLD
        ws.cell(
            row=12, column=3,
            value="=F5*(1+C10)+C9",
        ).number_format = "#,##0.00"

        for col, width in ((1, 4), (2, 22), (3, 14), (4, 14), (5, 14), (6, 14)):
            ws.column_dimensions[get_column_letter(col)].width = width

    def finish(self) -> Workbook:
        for col, width in (
            (COL_DESC, 30), (COL_QTY, 9), (COL_QTY_UNIT, 7), (COL_HEIGHT, 9),
            (COL_TOTAL, 9), (COL_TOTAL_UNIT, 7), (COL_GIB_RATE, 9), (COL_GIB_TOTAL, 11),
            (COL_STOP_RATE, 9), (COL_STOP_TOTAL, 11), (COL_PAINT_RATE, 9),
            (COL_PAINT_TOTAL, 11), (COL_ROW_TOTAL, 12),
        ):
            self.ws.column_dimensions[get_column_letter(col)].width = width
        self.ws.freeze_panes = "A5"
        return self.wb

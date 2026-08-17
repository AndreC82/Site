"""Gera um exemplo da planilha de quantidades no formato do estimador (Quantities + Summary),
com números fictícios só para validar a estrutura — não são dados de nenhum projeto real.
"""

from __future__ import annotations

from .quantities_workbook import (
    CeilingItem,
    DEFAULT_RATE_TABLE,
    DoorPaintItem,
    HeightGroup,
    LiningItem,
    QuantitiesWorkbookBuilder,
    RateEntry,
)

# Taxas de custo ($/m² ou $/un) de exemplo — números redondos e fictícios,
# só para a planilha calcular algo visível; troque pelos seus valores reais.
EXAMPLE_RATES = [
    RateEntry("10mm STANDARD", 15.00),
    RateEntry("10mm AQUALINE", 22.00),
    RateEntry("13mm STANDARD", 18.00),
    RateEntry("13mm AQUALINE", 28.00),
    RateEntry("13mm FIRELINE", 24.00),
    RateEntry("13mm NOISELINE", 26.00),
    RateEntry("16mm FIRELINE", 34.00),
    RateEntry("19mm Fireline", 46.00),
    RateEntry("GIB Sealant", 7.00),
    RateEntry("STOPPING LV4", 9.00),
    RateEntry("SQUARE STOP", 7.00),
    RateEntry("CORNER BEADS", 8.00),
    RateEntry("Painting Skirting", 12.00),
    RateEntry("PAINTING", 20.00),
    RateEntry("Single Door", 200.00),
    RateEntry("Double Door", 360.00),
]


def build_example(output_path: str) -> None:
    builder = QuantitiesWorkbookBuilder(building_name="BUILDING 01 (exemplo)", subcontractor_margin=0.25)
    builder.start(EXAMPLE_RATES)

    builder.add_height_group(
        HeightGroup(
            name="LININGS - 2.70m",
            height_m=2.70,
            items=[
                LiningItem("2x 19mm Fireline", qty=12.0, qty_unit="m", rate_label="19mm Fireline", layers=2),
                LiningItem("1x 16mm Fireline", qty=8.0, qty_unit="m", rate_label="16mm FIRELINE", layers=1),
                LiningItem("1x 10mm Standard", qty=60.0, qty_unit="m", rate_label="10mm STANDARD", layers=1),
                LiningItem("1x 10mm Aqualine", qty=15.0, qty_unit="m", rate_label="10mm AQUALINE", layers=1),
            ],
            corner_trims_qty=8,
            sealant_qty=95.0,
        )
    )

    builder.add_ceilings(
        items=[
            CeilingItem("13mm Standard Gib board (C1)", area_m2=180.0, rate_label="13mm STANDARD"),
            CeilingItem("13mm Aqualine Gib board (C2)", area_m2=20.0, rate_label="13mm AQUALINE"),
            CeilingItem("1/16mm Fireline Gib board (C3)", area_m2=45.0, rate_label="13mm FIRELINE"),
        ],
        square_stop_qty=220.0,
    )

    builder.add_painting_only(
        skirting_m=95.0,
        doors=[
            DoorPaintItem("D1 - Single Swing", count=6, rate_label="Single Door"),
            DoorPaintItem("D2 - Double Swing", count=2, rate_label="Double Door"),
        ],
    )

    builder.add_summary_sheet()
    wb = builder.finish()
    wb.save(output_path)


if __name__ == "__main__":
    import sys

    build_example(sys.argv[1] if len(sys.argv) > 1 else "quantities_exemplo.xlsx")

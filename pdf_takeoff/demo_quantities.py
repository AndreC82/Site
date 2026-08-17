"""Gera um exemplo da planilha de quantidades no formato do estimador (Taxas + Quantities + Summary),
com quantidades fictícias só para validar a estrutura — não são dados de nenhum projeto real.
"""

from __future__ import annotations

from .quantities_workbook import (
    CeilingItem,
    DoorPaintItem,
    GibBoardRate,
    HeightGroup,
    LiningItem,
    QuantitiesWorkbookBuilder,
)

# Taxas de custo de exemplo (install $/m² de mão de obra + custo da chapa $) —
# números redondos e fictícios, ajuste para os seus valores reais.
EXAMPLE_GIB_RATES = [
    GibBoardRate("10mm Standard", install_cost_m2=8.00, board_cost=35.00),
    GibBoardRate("10mm Aqualine", install_cost_m2=8.00, board_cost=55.00),
    GibBoardRate("13mm Standard", install_cost_m2=9.00, board_cost=42.00),
    GibBoardRate("13mm Aqualine", install_cost_m2=9.00, board_cost=65.00),
    GibBoardRate("13mm Fireline", install_cost_m2=10.00, board_cost=58.00),
    GibBoardRate("13mm Noiseline", install_cost_m2=10.00, board_cost=62.00),
    GibBoardRate("16mm Fireline", install_cost_m2=11.00, board_cost=72.00),
    GibBoardRate("19mm Fireline", install_cost_m2=12.00, board_cost=88.00),
]


def build_example(output_path: str) -> None:
    builder = QuantitiesWorkbookBuilder(building_name="BUILDING 01 (exemplo)")

    builder.add_rates_sheet(
        margin=0.25,
        gib_rates=EXAMPLE_GIB_RATES,
        painting_cost=25.00,
        stopping_wall_cost=11.00,
        stopping_ceiling_cost=9.00,
        board_width_m=1.2,
        board_height_m=2.4,
        corner_trim_cost=8.00,
        sealant_cost=7.00,
        skirting_paint_cost=15.00,
        single_door_cost=250.00,
        double_door_cost=450.00,
    )

    builder.start()

    builder.add_height_group(
        HeightGroup(
            name="LININGS - 2.70m",
            height_m=2.70,
            items=[
                LiningItem("2x 19mm Fireline", qty=12.0, board_type="19mm Fireline", layers=2),
                LiningItem("1x 16mm Fireline", qty=8.0, board_type="16mm Fireline", layers=1),
                LiningItem("1x 10mm Standard", qty=60.0, board_type="10mm Standard", layers=1),
                LiningItem("1x 10mm Aqualine", qty=15.0, board_type="10mm Aqualine", layers=1),
            ],
            corner_trims_qty=8,
            sealant_qty=95.0,
        )
    )

    builder.add_ceilings(
        items=[
            CeilingItem("13mm Standard Gib board (C1)", area_m2=180.0, board_type="13mm Standard"),
            CeilingItem("13mm Aqualine Gib board (C2)", area_m2=20.0, board_type="13mm Aqualine"),
            CeilingItem("1/16mm Fireline Gib board (C3)", area_m2=45.0, board_type="13mm Fireline"),
        ],
        square_stop_qty=220.0,
    )

    builder.add_painting_only(
        skirting_m=95.0,
        doors=[
            DoorPaintItem("D1 - Single Swing", count=6, door_type="single"),
            DoorPaintItem("D2 - Double Swing", count=2, door_type="double"),
        ],
    )

    builder.add_summary_sheet()
    wb = builder.finish()
    wb.save(output_path)


if __name__ == "__main__":
    import sys

    build_example(sys.argv[1] if len(sys.argv) > 1 else "quantities_exemplo.xlsx")

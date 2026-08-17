"""Gera um exemplo da planilha de quantidades no formato do estimador (Taxas + Quantities + Summary),
com quantidades fictícias só para validar a estrutura — não são dados de nenhum projeto real.
As taxas usam os valores passados pelo usuário: Pintura $25/m², Stopping $10/m², Gib installer $8/m²
+ custo da chapa de 2,4x1,2m (custo da chapa em si fica como exemplo a ajustar).
"""

from __future__ import annotations

from .quantities_workbook import (
    CeilingItem,
    DoorPaintItem,
    HeightGroup,
    LiningItem,
    QuantitiesWorkbookBuilder,
)


def build_example(output_path: str) -> None:
    builder = QuantitiesWorkbookBuilder(building_name="BUILDING 01 (exemplo)")

    builder.add_rates_sheet(
        painting_rate=25.00,
        stopping_rate=10.00,
        gib_install_rate=8.00,
        board_cost=55.00,  # exemplo — ajuste para o custo real da chapa 2,4x1,2m
        board_width_m=1.2,
        board_height_m=2.4,
        corner_trim_rate=8.00,
        sealant_rate=7.00,
        skirting_paint_rate=12.00,
        single_door_rate=200.00,
        double_door_rate=360.00,
    )

    builder.start()

    builder.add_height_group(
        HeightGroup(
            name="LININGS - 2.70m",
            height_m=2.70,
            items=[
                LiningItem("2x 19mm Fireline", qty=12.0, layers=2),
                LiningItem("1x 16mm Fireline", qty=8.0, layers=1),
                LiningItem("1x 10mm Standard", qty=60.0, layers=1),
                LiningItem("1x 10mm Aqualine", qty=15.0, layers=1),
            ],
            corner_trims_qty=8,
            sealant_qty=95.0,
        )
    )

    builder.add_ceilings(
        items=[
            CeilingItem("13mm Standard Gib board (C1)", area_m2=180.0),
            CeilingItem("13mm Aqualine Gib board (C2)", area_m2=20.0),
            CeilingItem("1/16mm Fireline Gib board (C3)", area_m2=45.0),
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

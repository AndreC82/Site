"""Roda as quantidades REAIS do arquivo de referência do Q&S (Hobs Road ECE)
através do modelo de taxas por tipo/espessura de chapa e compara o total
resultante com o total real que o Q&S calculou, para validar se a estrutura
do programa reproduz o mesmo resultado.
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

# Totais reais do Q&S (aba "Summary" do arquivo enviado), em $ de VENDA:
REFERENCE_TOTALS = {
    "GIB": 116546.5625,
    "PLASTER": 48915.0,
    "PAINTING": 106418.75,
    "TOTAL": 271880.3125,
}

# Taxas de CUSTO extraídas da tabela "COST - RATES" do arquivo real do Q&S
# (install_cost_m2 = a taxa inteira dele; board_cost=0 pois ele já não separa
# mão de obra de material). Margem real = 25%.
REAL_GIB_RATES = [
    GibBoardRate("10mm Standard", install_cost_m2=18.89, board_cost=0.0),
    GibBoardRate("10mm Aqualine", install_cost_m2=28.58, board_cost=0.0),
    GibBoardRate("13mm Standard", install_cost_m2=24.38, board_cost=0.0),
    GibBoardRate("13mm Aqualine", install_cost_m2=37.82, board_cost=0.0),
    GibBoardRate("13mm Fireline", install_cost_m2=29.55, board_cost=0.0),
    GibBoardRate("13mm Noiseline", install_cost_m2=32.14, board_cost=0.0),
    GibBoardRate("16mm Fireline", install_cost_m2=43.47, board_cost=0.0),
    GibBoardRate("19mm Fireline", install_cost_m2=59.03, board_cost=0.0),
]


def build_from_real_quantities(output_path: str, rates: dict) -> None:
    builder = QuantitiesWorkbookBuilder(building_name="Hobs Road ECE (quantidades reais do Q&S)")
    builder.add_rates_sheet(**rates)
    builder.start()

    # Quantidades extraídas diretamente da aba "Quantities" do arquivo do Q&S.
    builder.add_height_group(
        HeightGroup(
            name="LININGS - 3.75m",
            height_m=3.75,
            items=[
                LiningItem("2x 19mm Fireline", qty=47.42, board_type="19mm Fireline", layers=2),
                LiningItem("1x 16mm Fireline", qty=6.02, board_type="16mm Fireline", layers=1),
                LiningItem("1x 10mm Standard", qty=188.38, board_type="10mm Standard", layers=1),
                LiningItem("1x 10mm Aqualine", qty=41.22, board_type="10mm Aqualine", layers=1),
            ],
            corner_trims_qty=19,
        )
    )
    builder.add_height_group(
        HeightGroup(
            name="LININGS - 2.75m (altura real 3.00m)",
            height_m=3.00,
            items=[
                LiningItem("1x 16mm Fireline", qty=21.03, board_type="16mm Fireline", layers=1),
                LiningItem("1x 13mm Fireline", qty=79.72, board_type="13mm Fireline", layers=1),
                LiningItem("1x 10mm Standard", qty=92.01, board_type="10mm Standard", layers=1),
                LiningItem("1x 10mm Aqualine", qty=31.37, board_type="10mm Aqualine", layers=1),
            ],
            corner_trims_qty=10,
        )
    )
    builder.add_height_group(
        HeightGroup(
            name="LININGS - 2.60m (altura real 2.70m)",
            height_m=2.70,
            items=[
                LiningItem("1x 10mm Standard", qty=85.64, board_type="10mm Standard", layers=1),
                LiningItem("1x 10mm Aqualine", qty=36.02, board_type="10mm Aqualine", layers=1),
            ],
            corner_trims_qty=6,
            sealant_qty=262.8,
        )
    )

    builder.add_ceilings(
        items=[
            CeilingItem("13mm Standard Gib board (C1)", area_m2=546.44, board_type="13mm Standard"),
            CeilingItem("13mm Aqualine Gib board (C2)", area_m2=62.76, board_type="13mm Aqualine"),
            CeilingItem("1/16mm Fireline Gib board (C3)", area_m2=145.97, board_type="16mm Fireline"),
            CeilingItem("13mm Gib board (C4)", area_m2=66.8, board_type="13mm Standard"),
        ],
        square_stop_qty=679.45,
    )

    builder.add_painting_only(
        skirting_m=338.42,
        doors=[
            DoorPaintItem("Portas simples (single swing/sliding)", count=22, door_type="single"),
            DoorPaintItem("Portas duplas (double swing/sliding)", count=5, door_type="double"),
        ],
    )

    builder.add_summary_sheet()
    wb = builder.finish()
    wb.save(output_path)


if __name__ == "__main__":
    import sys

    rates = dict(
        margin=0.25,
        gib_rates=REAL_GIB_RATES,
        painting_cost=25.00,
        stopping_wall_cost=11.00,
        stopping_ceiling_cost=9.00,
        board_width_m=1.2,
        board_height_m=2.4,
        corner_trim_cost=10.00,
        sealant_cost=9.00,
        skirting_paint_cost=15.00,
        single_door_cost=250.00,
        double_door_cost=450.00,
    )
    build_from_real_quantities(sys.argv[1] if len(sys.argv) > 1 else "comparacao.xlsx", rates)

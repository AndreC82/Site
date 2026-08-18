"""Assistente interativo de terminal: faz as mesmas perguntas de ajuste de taxas
e quantidades feitas em conversa (custo ou venda? diferenciar Gib por tipo?
separar Stopping parede/teto? etc.) e gera a planilha final sozinho.

Rodar com: python -m pdf_takeoff.wizard
"""

from __future__ import annotations

from typing import Callable

from .quantities_workbook import (
    CeilingItem,
    DoorPaintItem,
    GibBoardRate,
    GIB_BOARD_TYPES,
    HeightGroup,
    LiningItem,
    QuantitiesWorkbookBuilder,
)

InputFn = Callable[[str], str]


def _ask_float(input_fn: InputFn, prompt: str, default: float | None = None) -> float:
    suffix = f" [{default}]" if default is not None else ""
    while True:
        raw = input_fn(f"{prompt}{suffix}: ").strip().replace(",", ".")
        if not raw and default is not None:
            return default
        try:
            return float(raw)
        except ValueError:
            print("  Valor inválido, digite um número (ex.: 25.00).")


def _ask_int(input_fn: InputFn, prompt: str, default: int = 0) -> int:
    raw = input_fn(f"{prompt} [{default}]: ").strip()
    if not raw:
        return default
    try:
        return int(float(raw))
    except ValueError:
        print("  Valor inválido, usando 0.")
        return 0


def _ask_yes_no(input_fn: InputFn, prompt: str, default: bool = True) -> bool:
    suffix = " [S/n]" if default else " [s/N]"
    raw = input_fn(f"{prompt}{suffix}: ").strip().lower()
    if not raw:
        return default
    return raw.startswith("s")


def _ask_text(input_fn: InputFn, prompt: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    raw = input_fn(f"{prompt}{suffix}: ").strip()
    return raw or default


def _ask_choice(input_fn: InputFn, prompt: str, options: list[str], default: str) -> str:
    opts_text = "/".join(options)
    while True:
        raw = input_fn(f"{prompt} ({opts_text}) [{default}]: ").strip().lower()
        if not raw:
            return default
        for opt in options:
            if raw == opt.lower() or raw == opt[0].lower():
                return opt
        print(f"  Escolha uma destas opções: {opts_text}")


def collect_rates(input_fn: InputFn = input) -> dict:
    print("\n=== TAXAS ===")
    is_cost = _ask_choice(input_fn, "Os valores que você vai digitar são custo ou já são preço de venda?", ["custo", "venda"], "custo") == "custo"
    margin = 0.0
    if is_cost:
        margin_pct = _ask_float(input_fn, "Margem a aplicar sobre o custo (%)", 25.0)
        margin = margin_pct / 100.0
    else:
        print("  Ok, os valores digitados serão usados como preço de venda direto (margem = 0%).")

    painting_cost = _ask_float(input_fn, "Pintura ($/m²)", 25.00)
    stopping_wall_cost = _ask_float(input_fn, "Stopping de parede ($/m²)", 11.00)
    split_ceiling = _ask_yes_no(input_fn, "Usar uma taxa de Stopping diferente para o teto (Square Stop)?", True)
    stopping_ceiling_cost = (
        _ask_float(input_fn, "Stopping de teto - Square Stop ($/m²)", 9.00)
        if split_ceiling
        else stopping_wall_cost
    )

    print("\n--- Gib ---")
    single_gib_rate = _ask_yes_no(
        input_fn, "Usar um único valor de Gib pra qualquer tipo de chapa (mais simples)?", False
    )
    board_width_m = _ask_float(input_fn, "Largura da chapa (m)", 1.2)
    board_height_m = _ask_float(input_fn, "Altura da chapa (m)", 2.4)

    gib_rates: list[GibBoardRate] = []
    if single_gib_rate:
        install = _ask_float(input_fn, "Instalação de Gib - mão de obra ($/m²)", 8.00)
        board_cost = _ask_float(input_fn, "Custo da chapa ($)", 55.00)
        for board_type in GIB_BOARD_TYPES:
            gib_rates.append(GibBoardRate(board_type, install_cost_m2=install, board_cost=board_cost))
    else:
        print("  Pra cada tipo de chapa usado no projeto, informe instalação e custo da chapa.")
        print("  Deixe em branco (0) os tipos que você não usa.")
        for board_type in GIB_BOARD_TYPES:
            used = _ask_yes_no(input_fn, f"  Usa {board_type} neste projeto?", False)
            if used:
                install = _ask_float(input_fn, f"    {board_type} - instalação ($/m²)", 8.00)
                board_cost = _ask_float(input_fn, f"    {board_type} - custo da chapa ($)", 55.00)
            else:
                install = 0.0
                board_cost = 0.0
            gib_rates.append(GibBoardRate(board_type, install_cost_m2=install, board_cost=board_cost))

    print("\n--- Outros itens ---")
    corner_trim_cost = _ask_float(input_fn, "Cantoneira / Corner trim ($/m)", 10.00)
    sealant_cost = _ask_float(input_fn, "Selante / Sealant ($/m)", 9.00)
    skirting_paint_cost = _ask_float(input_fn, "Pintura de rodapé - Skirting ($/m)", 15.00)
    single_door_cost = _ask_float(input_fn, "Porta simples - Single door ($/porta)", 250.00)
    double_door_cost = _ask_float(input_fn, "Porta dupla - Double door ($/porta)", 450.00)

    return dict(
        margin=margin,
        gib_rates=gib_rates,
        painting_cost=painting_cost,
        stopping_wall_cost=stopping_wall_cost,
        stopping_ceiling_cost=stopping_ceiling_cost,
        board_width_m=board_width_m,
        board_height_m=board_height_m,
        corner_trim_cost=corner_trim_cost,
        sealant_cost=sealant_cost,
        skirting_paint_cost=skirting_paint_cost,
        single_door_cost=single_door_cost,
        double_door_cost=double_door_cost,
    )


def _pick_board_type(input_fn: InputFn) -> str:
    print("  Tipos de chapa disponíveis:")
    for i, board_type in enumerate(GIB_BOARD_TYPES, start=1):
        print(f"    {i}. {board_type}")
    while True:
        raw = input_fn(f"  Qual tipo de chapa (1-{len(GIB_BOARD_TYPES)})? ").strip()
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(GIB_BOARD_TYPES):
                return GIB_BOARD_TYPES[idx]
        except ValueError:
            pass
        print("  Escolha um número da lista.")


def collect_height_groups(input_fn: InputFn = input) -> list[HeightGroup]:
    print("\n=== PAREDES (LININGS) POR ALTURA ===")
    groups: list[HeightGroup] = []
    while _ask_yes_no(input_fn, "\nAdicionar um grupo de paredes com uma altura de pé-direito?", False):
        height = _ask_float(input_fn, "  Altura do pé-direito deste grupo (m)", 2.70)
        name = _ask_text(input_fn, "  Nome do grupo", f"LININGS - {height:.2f}m")
        items: list[LiningItem] = []
        while _ask_yes_no(input_fn, "  Adicionar um tipo de forro a este grupo?", False):
            board_type = _pick_board_type(input_fn)
            layers = _ask_int(input_fn, "  Camadas (1 = single layer, 2 = double layer)", 1)
            qty = _ask_float(input_fn, f"  Metros lineares de parede com {board_type} ({layers}x)", 0.0)
            desc = _ask_text(input_fn, "  Descrição da linha", f"{layers}x {board_type}")
            items.append(LiningItem(desc, qty=qty, board_type=board_type, layers=layers))
        corner_trims = _ask_float(input_fn, "  Quantidade de cantoneiras (Corner Trims) neste grupo", 0.0)
        sealant = _ask_float(input_fn, "  Metros de selante (Sealant) neste grupo", 0.0)
        groups.append(
            HeightGroup(name=name, height_m=height, items=items, corner_trims_qty=corner_trims, sealant_qty=sealant)
        )
    return groups


def collect_ceilings(input_fn: InputFn = input) -> tuple[list[CeilingItem], float]:
    print("\n=== TETOS (CEILINGS) ===")
    items: list[CeilingItem] = []
    while _ask_yes_no(input_fn, "Adicionar uma área de teto?", False):
        board_type = _pick_board_type(input_fn)
        area = _ask_float(input_fn, f"  Área de teto com {board_type} (m²)", 0.0)
        desc = _ask_text(input_fn, "  Descrição da linha", board_type)
        items.append(CeilingItem(desc, area_m2=area, board_type=board_type))
    square_stop = _ask_float(input_fn, "Metros de junta de teto (Square Stop)", 0.0)
    return items, square_stop


def collect_painting_only(input_fn: InputFn = input) -> tuple[float, list[DoorPaintItem]]:
    print("\n=== PINTURA AVULSA (RODAPÉ, PORTAS) ===")
    skirting = _ask_float(input_fn, "Metros de rodapé (Skirting) a pintar", 0.0)
    doors: list[DoorPaintItem] = []
    while _ask_yes_no(input_fn, "Adicionar um grupo de portas?", False):
        desc = _ask_text(input_fn, "  Descrição (ex.: 'Portas quarto - single swing')", "Portas")
        count = _ask_int(input_fn, "  Quantidade", 0)
        door_type = _ask_choice(input_fn, "  Tipo", ["single", "double"], "single")
        doors.append(DoorPaintItem(desc, count=count, door_type=door_type))
    return skirting, doors


def run_wizard(input_fn: InputFn = input, output_path: str | None = None) -> str:
    print("Assistente de planilha de orçamento (Gib / Pintura / Stopping)")
    building_name = _ask_text(input_fn, "\nNome do projeto/edifício", "BUILDING 01")

    builder = QuantitiesWorkbookBuilder(building_name=building_name)
    rates = collect_rates(input_fn)
    builder.add_rates_sheet(**rates)
    builder.start()

    height_groups = collect_height_groups(input_fn)
    for group in height_groups:
        builder.add_height_group(group)

    ceiling_items, square_stop_qty = collect_ceilings(input_fn)
    if ceiling_items or square_stop_qty:
        builder.add_ceilings(ceiling_items, square_stop_qty=square_stop_qty)

    skirting_m, doors = collect_painting_only(input_fn)
    if skirting_m or doors:
        builder.add_painting_only(skirting_m, doors)

    if not height_groups and not ceiling_items and not square_stop_qty and not skirting_m and not doors:
        print(
            "\nATENÇÃO: nenhuma quantidade foi informada (nenhum grupo de parede, teto, "
            "rodapé ou porta). A planilha vai sair com todos os totais em $0 — "
            "rode de novo e responda 's' nas perguntas 'Adicionar...?' para incluir dados."
        )

    builder.add_summary_sheet()
    wb = builder.finish()

    if output_path is None:
        output_path = _ask_text(input_fn, "\nNome do arquivo de saída", "orcamento.xlsx")
    if not output_path.lower().endswith(".xlsx"):
        output_path += ".xlsx"
    wb.save(output_path)
    print(f"\nPlanilha salva em: {output_path}")
    return output_path


if __name__ == "__main__":
    run_wizard()

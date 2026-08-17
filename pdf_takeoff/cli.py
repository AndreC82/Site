"""Interface de linha de comando do quantificador automático de pintura/drywall."""

from __future__ import annotations

import argparse
import sys

from .calibration import auto_detect_scale
from .export import export_workbook, reimport_workbook
from .extract import extract_pdf
from .legend import load_legend
from .takeoff import build_takeoff
from .visualize import render_review_pdf


def _cmd_extract(args: argparse.Namespace) -> int:
    legend = load_legend(args.legend)
    pages = extract_pdf(args.pdf)

    if args.scale:
        scale = args.scale
        print(f"Escala informada manualmente: {scale:.6f} m por unidade de PDF.")
    else:
        all_segments = [s for p in pages for s in p.segments]
        all_words = [w for p in pages for w in p.words]
        result = auto_detect_scale(all_segments, all_words)
        scale = result.scale_m_per_unit
        status = "confiável" if result.confident else "INCERTA — confira/ajuste com --scale"
        print(
            f"Escala detectada automaticamente: {scale:.6f} m/unidade "
            f"({len(result.candidates)} cotas usadas, {status})."
        )
        if not result.confident:
            print(
                "  Dica: se a planta tiver poucas cotas legíveis, informe a escala manualmente, "
                "ex.: --scale 0.001 (se o PDF foi exportado em milímetros)."
            )

    rooms, lines = build_takeoff(
        pages, legend, scale_m_per_unit=scale, pe_direito_m=args.pe_direito
    )

    if not rooms:
        print(
            "Nenhum ambiente fechado foi reconstruído a partir das linhas do PDF. "
            "Verifique se o PDF é vetorial (não uma imagem escaneada) e se as paredes "
            "formam contornos fechados.",
            file=sys.stderr,
        )
        return 1

    export_workbook(lines, args.output)
    print(f"Planilha de orçamento gerada em: {args.output}")
    print(f"  {len(rooms)} ambiente(s) detectado(s), {len(lines)} linha(s) de quantitativo.")

    unmatched_rooms = [r for r in rooms if not (r.wall_paint_codes or r.ceiling_paint_codes or r.drywall_codes)]
    if unmatched_rooms:
        print(
            f"  Atenção: {len(unmatched_rooms)} ambiente(s) sem nenhum código da legenda "
            "reconhecido nas proximidades — confira no PDF de revisão."
        )

    if args.review_pdf:
        render_review_pdf(args.pdf, rooms, args.review_pdf)
        print(f"PDF de conferência (desenho + quantidades) gerado em: {args.review_pdf}")

    return 0


def _cmd_reimport(args: argparse.Namespace) -> int:
    lines = reimport_workbook(args.input)
    export_workbook(lines, args.output)
    print(f"Planilha recalculada a partir dos ajustes em '{args.input}' salva em: {args.output}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdf_takeoff",
        description=(
            "Quantificação automática de pintura (parede/teto) e drywall (Fireline, "
            "Aqualine, single/double layer etc.) a partir de plantas em PDF vetorial."
        ),
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_extract = sub.add_parser(
        "extract", help="Processa um PDF de planta e gera a planilha de orçamento."
    )
    p_extract.add_argument("--pdf", required=True, help="Caminho do PDF da planta.")
    p_extract.add_argument(
        "--legend",
        required=True,
        help="Caminho do JSON de legenda (códigos -> categoria/descrição). Veja examples/legend_exemplo.json.",
    )
    p_extract.add_argument("--output", required=True, help="Caminho da planilha .xlsx de saída.")
    p_extract.add_argument(
        "--pe-direito",
        type=float,
        default=2.80,
        dest="pe_direito",
        help="Pé-direito padrão em metros, usado para calcular a área de parede (padrão: 2.80).",
    )
    p_extract.add_argument(
        "--scale",
        type=float,
        default=None,
        help=(
            "Escala manual em metros por unidade de PDF (sobrepõe a detecção automática). "
            "Use se a detecção automática marcar a escala como incerta."
        ),
    )
    p_extract.add_argument(
        "--review-pdf",
        dest="review_pdf",
        default=None,
        help="Se informado, gera também um PDF de conferência com o que foi detectado desenhado sobre a planta.",
    )
    p_extract.set_defaults(func=_cmd_extract)

    p_reimport = sub.add_parser(
        "reimport",
        help="Recalcula Resumo/abas por categoria a partir de uma planilha com a aba 'Detalhe por Ambiente' editada manualmente.",
    )
    p_reimport.add_argument("--input", required=True, help="Planilha .xlsx editada.")
    p_reimport.add_argument("--output", required=True, help="Caminho da planilha final recalculada.")
    p_reimport.set_defaults(func=_cmd_reimport)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())

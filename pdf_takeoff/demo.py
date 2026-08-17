"""Gera uma planta sintética simples e roda o pipeline completo, de ponta a ponta.

Serve tanto como demonstração (`python -m pdf_takeoff.demo`) quanto como base
para o teste de integração, já que não temos uma planta real de CAD à mão.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from reportlab.pdfgen import canvas

from .calibration import auto_detect_scale
from .export import export_workbook
from .extract import extract_pdf
from .legend import load_legend
from .takeoff import build_takeoff
from .visualize import render_review_pdf

LEGEND_PATH = str(Path(__file__).resolve().parent.parent / "examples" / "legend_exemplo.json")


def generate_sample_plant(path: str) -> None:
    """Cria um PDF vetorial com duas salas retangulares, códigos de pintura/gesso e
    três cotas (para a detecção automática de escala poder ser exercitada)."""
    c = canvas.Canvas(path, pagesize=(560, 260))

    # Sala 01: 200 x 150 unidades de PDF (canto inferior-esquerdo em 50,50)
    c.rect(50, 50, 200, 150, stroke=1, fill=0)
    c.drawCentredString(150, 125, "SALA")
    c.drawString(60, 100, "P-01")
    c.drawString(60, 85, "PT-01")

    # Sala 02: 250 x 150 unidades, parede compartilhada em x=250
    c.rect(250, 50, 250, 150, stroke=1, fill=0)
    c.drawCentredString(375, 125, "SALA")
    c.drawString(260, 100, "P-02")
    c.drawString(260, 85, "PT-01")
    c.drawString(260, 70, "GB-AQ-DL")

    # Cotas para calibração automática de escala (valores em metros)
    c.drawCentredString(150, 210, "4.00")  # largura da Sala 01
    c.drawString(20, 120, "3.00")  # altura de ambas as salas
    c.drawCentredString(375, 210, "5.00")  # largura da Sala 02

    c.showPage()
    c.save()


def run_demo(output_dir: str = "demo_output") -> dict:
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    pdf_path = str(out / "planta_exemplo.pdf")
    xlsx_path = str(out / "orcamento_exemplo.xlsx")
    review_path = str(out / "planta_conferencia.pdf")

    generate_sample_plant(pdf_path)

    legend = load_legend(LEGEND_PATH)
    pages = extract_pdf(pdf_path)

    all_segments = [s for p in pages for s in p.segments]
    all_words = [w for p in pages for w in p.words]
    calibration = auto_detect_scale(all_segments, all_words)

    rooms, lines = build_takeoff(
        pages, legend, scale_m_per_unit=calibration.scale_m_per_unit, pe_direito_m=2.80
    )

    export_workbook(lines, xlsx_path)
    render_review_pdf(pdf_path, rooms, review_path)

    return {
        "pdf": pdf_path,
        "xlsx": xlsx_path,
        "review_pdf": review_path,
        "scale": calibration.scale_m_per_unit,
        "confident": calibration.confident,
        "rooms": rooms,
        "lines": lines,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", default="demo_output")
    args = parser.parse_args()

    result = run_demo(args.output_dir)

    print(f"PDF de exemplo gerado em: {result['pdf']}")
    print(
        f"Escala detectada: {result['scale']:.6f} m/unidade "
        f"({'confiável' if result['confident'] else 'incerta'})"
    )
    print(f"{len(result['rooms'])} ambiente(s) detectado(s):")
    for room in result["rooms"]:
        print(
            f"  {room.room_id} ({room.label or 'sem rótulo'}): "
            f"teto {room.ceiling_area_m2:.2f} m², parede {room.wall_area_m2:.2f} m² | "
            f"parede={room.wall_paint_codes} teto={room.ceiling_paint_codes} "
            f"gesso={room.drywall_codes}"
        )
    print(f"\nPlanilha de orçamento: {result['xlsx']}")
    print(f"PDF de conferência (planta com o que foi detectado desenhado): {result['review_pdf']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

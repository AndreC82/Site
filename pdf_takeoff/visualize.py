"""Gera um PDF de conferência: a planta original com o que foi detectado desenhado por cima."""

from __future__ import annotations

import pymupdf as fitz

from .takeoff import RoomTakeoff

# Cor de destaque do contorno de cada ambiente detectado (RGB 0-1).
_OUTLINE_COLOR = (0.85, 0.1, 0.1)
_OUTLINE_FILL = (1.0, 0.85, 0.2)
_OUTLINE_OPACITY = 0.18

# Uma cor por categoria, usada nas etiquetas (rótulos de texto) ao lado do ambiente.
_CATEGORY_COLORS = {
    "wall_paint": (0.10, 0.35, 0.85),  # azul
    "ceiling_paint": (0.85, 0.45, 0.05),  # laranja
    "drywall": (0.10, 0.55, 0.20),  # verde
}
_CATEGORY_SHORT = {
    "wall_paint": "PAR",
    "ceiling_paint": "TETO",
    "drywall": "GESSO",
}


def render_review_pdf(source_pdf_path: str, rooms: list[RoomTakeoff], output_path: str) -> None:
    """Desenha, sobre cada página do PDF original, o contorno de cada ambiente detectado
    (polígono reconstruído), o ID do ambiente e a área de teto/parede calculada, além de
    um selo colorido por código/categoria atribuído a ele. Salva como um novo PDF, para
    conferência visual lado a lado com a planilha.
    """
    doc = fitz.open(source_pdf_path)
    try:
        rooms_by_page: dict[int, list[RoomTakeoff]] = {}
        for room in rooms:
            rooms_by_page.setdefault(room.page_number, []).append(room)

        for page_index in range(len(doc)):
            page = doc[page_index]
            page_rooms = rooms_by_page.get(page_index + 1, [])
            if not page_rooms:
                continue

            shape = page.new_shape()
            for room in page_rooms:
                coords = list(room.polygon.exterior.coords)
                shape.draw_polyline([fitz.Point(x, y) for x, y in coords])
                shape.finish(
                    color=_OUTLINE_COLOR,
                    fill=_OUTLINE_FILL,
                    fill_opacity=_OUTLINE_OPACITY,
                    width=1.2,
                    closePath=True,
                )
            shape.commit()

            for room in page_rooms:
                centroid = room.polygon.centroid
                lines = [f"{room.room_id}  {room.label}".strip()]
                lines.append(f"Teto: {room.ceiling_area_m2:.2f} m²")
                lines.append(f"Parede: {room.wall_area_m2:.2f} m² (pé-dir. {room.height_m:.2f} m)")
                for category, codes in (
                    ("wall_paint", room.wall_paint_codes),
                    ("ceiling_paint", room.ceiling_paint_codes),
                    ("drywall", room.drywall_codes),
                ):
                    if codes:
                        lines.append(f"{_CATEGORY_SHORT[category]}: {', '.join(codes)}")
                if not (room.wall_paint_codes or room.ceiling_paint_codes or room.drywall_codes):
                    lines.append("(nenhum código de legenda encontrado — revisar)")

                text = "\n".join(lines)
                origin = fitz.Point(centroid.x - 45, centroid.y - 6)
                page.insert_textbox(
                    fitz.Rect(origin.x, origin.y, origin.x + 140, origin.y + 14 * len(lines)),
                    text,
                    fontsize=6.5,
                    color=(0, 0, 0),
                    fill=(1, 1, 1),
                    fill_opacity=0.75,
                    align=0,
                )

            _draw_legend_key(page)

        doc.save(output_path)
    finally:
        doc.close()


def _draw_legend_key(page: fitz.Page) -> None:
    x0, y0 = 10, 10
    lines = [
        ("Contorno detectado (ambiente)", _OUTLINE_COLOR),
        ("PAR = pintura de parede", _CATEGORY_COLORS["wall_paint"]),
        ("TETO = pintura de teto", _CATEGORY_COLORS["ceiling_paint"]),
        ("GESSO = tipo de drywall", _CATEGORY_COLORS["drywall"]),
    ]
    box_h = 12 * len(lines) + 6
    page.draw_rect(fitz.Rect(x0, y0, x0 + 190, y0 + box_h), color=(0, 0, 0), fill=(1, 1, 1), fill_opacity=0.85)
    for i, (text, color) in enumerate(lines):
        y = y0 + 12 * i + 10
        page.draw_circle(fitz.Point(x0 + 8, y - 3), 3, color=color, fill=color)
        page.insert_text(fitz.Point(x0 + 16, y), text, fontsize=6.5, color=(0, 0, 0))

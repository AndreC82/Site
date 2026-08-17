"""Monta o quantitativo: liga polígonos de ambientes aos códigos da legenda e calcula áreas."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from shapely.geometry import Point, Polygon

from .extract import PageContent, TextWord
from .geometry import area_m2, build_room_polygons, perimeter_m
from .legend import LegendEntry, match_code

# Distância (em unidades do PDF) além do contorno do ambiente em que um
# código/rótulo ainda é considerado pertencente a ele (cobre textos
# colocados rente à parede, por fora do polígono).
_ASSIGN_BUFFER = 12.0

_LABEL_IGNORE_RE = re.compile(r"^[\d.,/\-x×]+$", re.IGNORECASE)


@dataclass
class RoomTakeoff:
    room_id: str
    page_number: int
    label: str
    polygon: Polygon
    height_m: float
    ceiling_area_m2: float
    wall_area_m2: float
    wall_paint_codes: list[str] = field(default_factory=list)
    ceiling_paint_codes: list[str] = field(default_factory=list)
    drywall_codes: list[str] = field(default_factory=list)
    unmatched_texts: list[str] = field(default_factory=list)


@dataclass
class QuantityLine:
    room_id: str
    page_number: int
    room_label: str
    category: str
    code: str
    description: str
    layers: int
    area_m2: float
    note: str = ""


def _words_near_polygon(polygon: Polygon, words: list[TextWord], buffer: float) -> list[TextWord]:
    area = polygon.buffer(buffer)
    return [w for w in words if area.contains(Point(w.center))]


def _guess_label(polygon: Polygon, words: list[TextWord], legend: dict[str, LegendEntry]) -> str:
    inside = [w for w in words if polygon.contains(Point(w.center))]
    candidates = [
        w for w in inside if match_code(w.text, legend) is None and not _LABEL_IGNORE_RE.match(w.text)
    ]
    if not candidates:
        return ""
    centroid = polygon.centroid
    candidates.sort(key=lambda w: Point(w.center).distance(centroid))
    return candidates[0].text


def build_takeoff(
    pages: list[PageContent],
    legend: dict[str, LegendEntry],
    scale_m_per_unit: float,
    pe_direito_m: float,
    min_room_area_units: float = 25.0,
) -> tuple[list[RoomTakeoff], list[QuantityLine]]:
    """Reconstrói ambientes por página e gera as linhas de quantitativo (uma por código/categoria/ambiente)."""
    rooms: list[RoomTakeoff] = []
    lines: list[QuantityLine] = []

    for page in pages:
        polygons = build_room_polygons(page.segments, min_area=min_room_area_units)
        for idx, polygon in enumerate(polygons, start=1):
            room_id = f"P{page.page_number}-{idx:02d}"
            nearby_words = _words_near_polygon(polygon, page.words, _ASSIGN_BUFFER)

            wall_codes: list[str] = []
            ceiling_codes: list[str] = []
            drywall_codes: list[str] = []
            unmatched: list[str] = []

            for word in nearby_words:
                entry = match_code(word.text, legend)
                if entry is None:
                    continue
                bucket = {
                    "wall_paint": wall_codes,
                    "ceiling_paint": ceiling_codes,
                    "drywall": drywall_codes,
                }[entry.category]
                if entry.code not in bucket:
                    bucket.append(entry.code)

            label = _guess_label(polygon, nearby_words, legend)

            ceiling_area = area_m2(polygon, scale_m_per_unit)
            wall_area = perimeter_m(polygon, scale_m_per_unit) * pe_direito_m

            room = RoomTakeoff(
                room_id=room_id,
                page_number=page.page_number,
                label=label,
                polygon=polygon,
                height_m=pe_direito_m,
                ceiling_area_m2=ceiling_area,
                wall_area_m2=wall_area,
                wall_paint_codes=wall_codes,
                ceiling_paint_codes=ceiling_codes,
                drywall_codes=drywall_codes,
                unmatched_texts=unmatched,
            )
            rooms.append(room)
            lines.extend(_room_to_lines(room, legend))

    return rooms, lines


def _room_to_lines(room: RoomTakeoff, legend: dict[str, LegendEntry]) -> list[QuantityLine]:
    lines: list[QuantityLine] = []

    def split_note(n: int) -> str:
        return (
            f"Área dividida igualmente entre {n} códigos detectados no ambiente — confira e ajuste."
            if n > 1
            else ""
        )

    if room.wall_paint_codes:
        share = room.wall_area_m2 / len(room.wall_paint_codes)
        note = split_note(len(room.wall_paint_codes))
        for code in room.wall_paint_codes:
            entry = legend[code]
            lines.append(
                QuantityLine(
                    room_id=room.room_id,
                    page_number=room.page_number,
                    room_label=room.label,
                    category="wall_paint",
                    code=code,
                    description=entry.description,
                    layers=1,
                    area_m2=share,
                    note=note,
                )
            )

    if room.ceiling_paint_codes:
        share = room.ceiling_area_m2 / len(room.ceiling_paint_codes)
        note = split_note(len(room.ceiling_paint_codes))
        for code in room.ceiling_paint_codes:
            entry = legend[code]
            lines.append(
                QuantityLine(
                    room_id=room.room_id,
                    page_number=room.page_number,
                    room_label=room.label,
                    category="ceiling_paint",
                    code=code,
                    description=entry.description,
                    layers=1,
                    area_m2=share,
                    note=note,
                )
            )

    if room.drywall_codes:
        share = room.wall_area_m2 / len(room.drywall_codes)
        note = split_note(len(room.drywall_codes))
        for code in room.drywall_codes:
            entry = legend[code]
            lines.append(
                QuantityLine(
                    room_id=room.room_id,
                    page_number=room.page_number,
                    room_label=room.label,
                    category="drywall",
                    code=code,
                    description=entry.description,
                    layers=entry.layers,
                    area_m2=share * entry.layers,
                    note=note,
                )
            )

    return lines

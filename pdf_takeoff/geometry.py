"""Reconstrução de polígonos (ambientes) a partir de segmentos de linha e cálculo de área/perímetro."""

from __future__ import annotations

from shapely.geometry import Polygon
from shapely.ops import polygonize, unary_union

from .extract import Segment

# Duas pontas são consideradas o mesmo nó se estiverem a menos deste raio
# (em unidades do PDF) uma da outra. Cobre pequenas imprecisões de traçado.
SNAP_TOLERANCE = 0.75


def _snap(point: tuple[float, float], grid: float) -> tuple[float, float]:
    return (round(point[0] / grid) * grid, round(point[1] / grid) * grid)


def build_room_polygons(segments: list[Segment], min_area: float = 25.0) -> list[Polygon]:
    """Reconstrói os polígonos fechados (ambientes) a partir dos segmentos de parede.

    Usa 'snap' de nós próximos para fechar pequenas folgas de traçado e
    `shapely.polygonize` para montar os anéis fechados. `min_area` filtra
    ruído (pequenos polígonos formados por hachuras, cotas, etc.) em
    unidades de PDF ao quadrado.
    """
    snapped = [
        (_snap(a, SNAP_TOLERANCE), _snap(b, SNAP_TOLERANCE)) for a, b in segments if a != b
    ]
    lines = [seg for seg in snapped if seg[0] != seg[1]]
    if not lines:
        return []

    merged = unary_union([_to_linestring(s) for s in lines])
    polygons = [p for p in polygonize(merged) if p.area >= min_area]
    # Remove polígonos que são apenas o "furo" interno de outro (paredes desenhadas
    # como contorno duplo geram um anel externo e um interno quase idênticos).
    polygons.sort(key=lambda p: p.area)
    kept: list[Polygon] = []
    for poly in polygons:
        if not any(_nearly_same(poly, other) for other in kept):
            kept.append(poly)
    return kept


def _to_linestring(segment: Segment):
    from shapely.geometry import LineString

    return LineString(segment)


def _nearly_same(a: Polygon, b: Polygon, tol: float = 0.03) -> bool:
    inter = a.intersection(b).area
    union = a.union(b).area
    if union == 0:
        return True
    return (inter / union) > (1 - tol)


def area_m2(polygon: Polygon, scale_m_per_unit: float) -> float:
    return polygon.area * (scale_m_per_unit**2)


def perimeter_m(polygon: Polygon, scale_m_per_unit: float) -> float:
    return polygon.length * scale_m_per_unit

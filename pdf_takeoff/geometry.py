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


def filter_title_block_noise(
    polygons: list[Polygon],
    page_width: float,
    page_height: float,
    scale_m_per_unit: float,
    words: list,
    max_span_ratio: float = 0.85,
    max_word_density_per_m2: float = 2.0,
) -> list[Polygon]:
    """Remove polígonos que quase certamente não são ambientes internos, mas sim
    a moldura do quadro de textos/legenda da prancha (comum em plantas de
    arquitetura). Dois sinais, ambos genéricos (não específicos de um projeto):

    1. Tiras que ocupam quase a largura/altura inteira da página — são a
       coluna de notas ou a faixa do carimbo do desenho, nunca um ambiente.
    2. Densidade de texto muito alta (palavras por m²) — um parágrafo de
       especificação tem dezenas de palavras espremidas numa área pequena;
       um ambiente real normalmente só tem o nome e talvez o tipo de piso.

    Isso não elimina 100% do ruído — sempre confira visualmente antes de
    confiar nos números. Veja também `exclude_exterior_areas`.
    """
    from shapely.geometry import Point

    kept = []
    for poly in polygons:
        minx, miny, maxx, maxy = poly.bounds
        span_w = (maxx - minx) / page_width if page_width else 0
        span_h = (maxy - miny) / page_height if page_height else 0
        if span_w > max_span_ratio or span_h > max_span_ratio:
            continue
        area = poly.area * (scale_m_per_unit**2)
        if area <= 0:
            continue
        word_count = sum(1 for w in words if poly.contains(Point(w.center)))
        density = word_count / area
        if density > max_word_density_per_m2:
            continue
        kept.append(poly)
    return kept


# Termos que só aparecem rotulando área externa coberta (nunca um ambiente
# interno com gesso/pintura) em plantas de arquitetura em inglês. Lista curta
# e específica de propósito — uma lista ampla de palavras já se mostrou
# perigosa (acaba excluindo ambientes reais que têm essas palavras por perto
# em notas/cotas vizinhas).
_EXTERIOR_AREA_KEYWORDS = {"deck", "portico", "verandah", "veranda", "carport", "pergola"}


def exclude_exterior_areas(
    polygons: list[Polygon], words: list, keywords: set[str] = _EXTERIOR_AREA_KEYWORDS
) -> list[Polygon]:
    """Remove polígonos cujo próprio rótulo contém uma palavra que só é usada
    pra área externa coberta (deck, portico, etc.) — essas áreas não levam
    gesso/pintura interna, mas ainda assim formam um polígono fechado no
    desenho e podem ser confundidas com um ambiente.
    """
    from shapely.geometry import Point

    kept = []
    for poly in polygons:
        words_inside = {w.text.strip(".,()").lower() for w in words if poly.contains(Point(w.center))}
        if words_inside & keywords:
            continue
        kept.append(poly)
    return kept

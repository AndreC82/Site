"""Extração de quantidades de uma "Wall Linings Plan" (planta de forro de
parede) que usa a convenção: cada trecho de parede é desenhado com uma linha
colorida (uma cor por tipo/espessura de chapa, ver "WALL LININGS KEY" na
prancha) e, opcionalmente, uma tag de keynote (ex. "5113G 4.5") ao lado,
referenciando uma tabela "Keynote Legend" que diz exatamente quantas camadas
e qual espessura/produto aquele código representa.

A cor sozinha não diferencia 1 camada de 2 camadas (nem sempre diferencia
espessura, ex. Aqualine 13mm vs 16mm) — só o texto do keynote tem essa
informação. Por isso este módulo faz o casamento espacial: para cada trecho
de parede colorido, procura o keynote mais próximo e usa a especificação dele
quando disponível (mais precisa); quando não encontra um keynote perto o
suficiente, cai para a especificação da cor (menos precisa, sinalizada como
tal no resultado) — ver `MatchedSegment.source`.

Não é específico de um projeto: a lógica (cor -> legenda de cor, código ->
tabela de keynote, casamento por proximidade) funciona em qualquer prancha
que siga esse padrão, mudando só os textos/cores lidos de cada arquivo.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import pymupdf as fitz

from .gib_spec_extract import BoardSpec, _normalize_product

# Metros por unidade do PDF (pontos), quando a prancha foi plotada em escala
# 1:N. PDF point = 1/72 polegada, sempre — independente do tamanho da folha.
_POINTS_TO_METERS_PER_INCH = 25.4 / 1000 / 72

_KEYNOTE_CODE_RE = re.compile(r"\d{3,5}[A-Z]{0,2}\s\d\.\d+")
_KEYNOTE_CODE_LINE_RE = re.compile(r"^\d{3,5}[A-Z]{0,2}\s\d\.\d+$")
_LAYER_WORD_TO_INT = {"one": 1, "two": 2, "three": 3, "four": 4}
_LAYER_SPEC_RE = re.compile(
    r"(ONE|TWO|THREE|FOUR)\s+LAYERS?\s+(\d+)\s*MM\s+GIB\S*\s+(\w+)",
    re.IGNORECASE,
)
_COLOR_KEY_SPEC_RE = re.compile(r"(\d+)\s*mm\s+Gib\s+(\w+)", re.IGNORECASE)
_SCALE_RE = re.compile(r"1\s*:\s*(\d+)(?:\s*@\s*A\d)?", re.IGNORECASE)

# larguras de traço (em pt) tipicamente usadas nestas pranchas: as linhas de
# forro coloridas ficam numa faixa estreita e previsível; abaixo/acima disso
# são geometria de arquitetura (preto) ou a amostra de cor na legenda (mais
# grossa). Ajustável via parâmetro se um projeto usar outra convenção.
DEFAULT_MIN_STROKE_WIDTH = 1.0
DEFAULT_MAX_STROKE_WIDTH = 1.8

# distância máxima (m, no mundo real) entre um trecho de parede e a tag de
# keynote mais próxima para considerar que ela se refere a esse trecho.
DEFAULT_MAX_MATCH_DIST_M = 3.5


def extract_scale_m_per_unit(page_text: str) -> float | None:
    match = _SCALE_RE.search(page_text)
    if not match:
        return None
    return int(match.group(1)) * _POINTS_TO_METERS_PER_INCH


def parse_keynote_legend(full_text: str) -> dict[str, BoardSpec]:
    """Lê a tabela 'Keynote Legend' (Key Value / Keynote Text) e devolve
    {"5113G 4.5": BoardSpec(thickness_mm=13, product="Fireline", layers=2), ...}.

    Só entram no dicionário os códigos cuja descrição bate com o padrão
    "ONE/TWO/... LAYER(S) <N>MM GIB <produto>" — outros keynotes (ex. chapa
    estrutural, sistema de porta) são ignorados aqui, pois não são forro GIB.
    """
    idx = full_text.find("Keynote Legend")
    if idx == -1:
        return {}
    lines = [line.strip() for line in full_text[idx:].split("\n") if line.strip()]

    legend: dict[str, BoardSpec] = {}
    i = 0
    while i < len(lines):
        line = lines[i]
        if not _KEYNOTE_CODE_LINE_RE.match(line):
            i += 1
            continue
        code = line
        desc_parts: list[str] = []
        j = i + 1
        while j < len(lines) and len(desc_parts) < 3:
            nxt = lines[j]
            if _KEYNOTE_CODE_LINE_RE.match(nxt) or nxt.upper().startswith("NOTE"):
                break
            desc_parts.append(nxt)
            j += 1
        desc = " ".join(desc_parts)
        spec_match = _LAYER_SPEC_RE.search(desc)
        if spec_match:
            layers = _LAYER_WORD_TO_INT.get(spec_match.group(1).lower(), 1)
            thickness = int(spec_match.group(2))
            product = _normalize_product(spec_match.group(3))
            legend[code] = BoardSpec(thickness_mm=thickness, product=product, layers=layers)
        i = j
    return legend


def _legend_table_zone(page: "fitz.Page") -> "fitz.Rect | None":
    """Área da página ocupada pela tabela 'Keynote Legend', para não confundir
    os códigos ali listados com tags realmente marcadas na planta."""
    rects = page.search_for("Keynote Legend")
    if not rects:
        return None
    r = rects[0]
    return fitz.Rect(r.x0 - 40, r.y0 - 10, page.rect.x1, page.rect.y1)


def find_placed_keynote_tags(page: "fitz.Page", legend: dict[str, BoardSpec]) -> list[tuple[str, "fitz.Rect"]]:
    """Devolve [(codigo, posicao), ...] de cada vez que um código da legenda
    aparece marcado na planta (fora da própria tabela de legenda)."""
    exclude_zone = _legend_table_zone(page)
    tags: list[tuple[str, fitz.Rect]] = []
    for code in legend:
        for rect in page.search_for(code):
            if exclude_zone is not None and exclude_zone.intersects(rect):
                continue
            tags.append((code, rect))
    return tags


def extract_wall_linings_color_key(page: "fitz.Page") -> dict[tuple[float, float, float], BoardSpec | None]:
    """Lê a legenda de cores da prancha (linha colorida + texto do tipo
    'Ntmm Gib Produto Plasterboard...') e devolve {cor_rgb_arredondada: BoardSpec}.

    A cor sozinha não sabe o número de camadas (a legenda de cor não
    diferencia) — por isso todo BoardSpec aqui vem com layers=1 (assumido);
    o casamento com o keynote (quando existir) é que corrige isso.
    """
    text_dict = page.get_text("dict")
    color_key: dict[tuple[float, float, float], BoardSpec | None] = {}
    for block in text_dict["blocks"]:
        for line in block.get("lines", []):
            txt = "".join(s["text"] for s in line["spans"])
            spec_match = _COLOR_KEY_SPEC_RE.search(txt)
            if not spec_match and "No Finish" not in txt and "Plywood" not in txt and "Lining" not in txt:
                continue
            bbox = line["bbox"]
            ymid = (bbox[1] + bbox[3]) / 2
            x0 = bbox[0]
            color = _find_swatch_color_left_of(page, x0, ymid)
            if color is None:
                continue
            if spec_match:
                thickness = int(spec_match.group(1))
                product = _normalize_product(spec_match.group(2))
                color_key[color] = BoardSpec(thickness_mm=thickness, product=product, layers=1)
            else:
                color_key[color] = None  # cor mapeada, mas não é chapa GIB (ex. "No Finish")
    return color_key


def _find_swatch_color_left_of(page: "fitz.Page", x_limit: float, y: float, max_dx: float = 60.0):
    best = None
    for d in page.get_drawings():
        color = d.get("color")
        if color is None:
            continue
        for item in d["items"]:
            if item[0] != "l":
                continue
            p1, p2 = item[1], item[2]
            lymid = (p1.y + p2.y) / 2
            lxmax = max(p1.x, p2.x)
            if abs(lymid - y) < 5 and x_limit - max_dx < lxmax < x_limit:
                best = tuple(round(c, 2) for c in color)
    return best


@dataclass
class WallSegment:
    color: tuple[float, float, float]
    length_pt: float
    midpoint: tuple[float, float]


def extract_colored_wall_segments(
    page: "fitz.Page",
    min_width: float = DEFAULT_MIN_STROKE_WIDTH,
    max_width: float = DEFAULT_MAX_STROKE_WIDTH,
) -> list[WallSegment]:
    """Cada 'drawing' colorido (traço de forro) vira um WallSegment com seu
    comprimento total e ponto médio, para depois casar com o keynote mais
    próximo. Filtra por largura de traço para não pegar a amostra da legenda
    (mais grossa) nem a linha de arquitetura em preto."""
    segments: list[WallSegment] = []
    for d in page.get_drawings():
        color = d.get("color")
        width = d.get("width")
        if color is None or width is None or not (min_width <= width <= max_width):
            continue
        pts: list[tuple[float, float]] = []
        length = 0.0
        for item in d["items"]:
            if item[0] != "l":
                continue
            p1, p2 = item[1], item[2]
            length += ((p1.x - p2.x) ** 2 + (p1.y - p2.y) ** 2) ** 0.5
            pts.append((p1.x, p1.y))
            pts.append((p2.x, p2.y))
        if not pts or length <= 0:
            continue
        mx = sum(p[0] for p in pts) / len(pts)
        my = sum(p[1] for p in pts) / len(pts)
        segments.append(WallSegment(color=tuple(round(c, 2) for c in color), length_pt=length, midpoint=(mx, my)))
    return segments


@dataclass
class FilledWallGroup:
    """Um grupo de paredes desenhadas como preenchimento sólido colorido
    (retângulo cheio), agrupadas por cor + espessura -- convenção usada em
    algumas pranchas (ex. quando cada tipo de parede vira um bloco cheio, em
    vez de uma linha fina). Diferente de `MatchedSegment`, aqui não
    identificamos o produto GIB automaticamente (o texto de legenda dessas
    pranchas varia demais entre escritórios pra confiar numa leitura
    genérica) -- só o comprimento por cor/espessura, para conferência manual
    do tipo de chapa."""

    color: tuple[float, float, float]
    thickness_mm: float
    length_m: float
    n_segments: int


_MIN_WALL_THICKNESS_MM = 60.0
_MAX_WALL_THICKNESS_MM = 260.0
_MIN_FILL_ASPECT_RATIO = 6.0
_MIN_GROUP_LENGTH_M = 0.5
_THICKNESS_CLUSTER_TOL_MM = 8.0


def detect_filled_wall_groups(
    page: "fitz.Page",
    scale_m_per_unit: float,
    min_thickness_mm: float = _MIN_WALL_THICKNESS_MM,
    max_thickness_mm: float = _MAX_WALL_THICKNESS_MM,
) -> list[FilledWallGroup]:
    """Detecta paredes desenhadas como retângulo de preenchimento sólido
    (não como linha fina) e agrupa por cor + espessura. Comprimento de cada
    grupo = soma(área de cada retângulo) / espessura do grupo -- funciona
    tanto pra um único retângulo longo quanto pra vários retângulos curtos
    que formam o mesmo trecho de parede (comum quando a parede é desenhada
    em pedaços, ex. interrompida por portas).

    Filtra por: proporção comprimento/espessura mínima (`_MIN_FILL_ASPECT_RATIO`,
    descarta preenchimentos "quadrados" -- mobília, ícones, hatching) e faixa
    de espessura plausível para parede (`min_thickness_mm`..`max_thickness_mm`
    -- descarta cotas, textos, tramas finas demais ou blocos grossos demais
    pra ser parede). Não tenta identificar o produto/tipo de chapa pela cor
    -- isso fica para conferência manual (ver aviso gerado por quem chama)."""
    by_color: dict[tuple[float, float, float], list[tuple[float, float]]] = {}
    for d in page.get_drawings():
        fill = d.get("fill")
        rect = d.get("rect")
        if fill is None or rect is None:
            continue
        if all(c > 0.92 for c in fill):
            continue  # branco/quase-branco: fundo, folha de porta, ícone -- não é parede
        w, h = rect.width, rect.height
        if w <= 0 or h <= 0:
            continue
        thin, long_ = (w, h) if w <= h else (h, w)
        if thin <= 0 or long_ / thin < _MIN_FILL_ASPECT_RATIO:
            continue
        thickness_mm = thin * scale_m_per_unit * 1000
        if not (min_thickness_mm <= thickness_mm <= max_thickness_mm):
            continue
        color = tuple(round(c, 3) for c in fill)
        by_color.setdefault(color, []).append((w * h, thin))

    groups: list[FilledWallGroup] = []
    for color, entries in by_color.items():
        # clustering 1D por proximidade de espessura (single-linkage): ordena
        # pela espessura e corta o grupo sempre que o próximo valor estiver a
        # mais de _THICKNESS_CLUSTER_TOL_MM do anterior -- evita que dois
        # retângulos com a mesma espessura real caiam em grupos diferentes só
        # por estarem perto de uma fronteira de arredondamento.
        entries.sort(key=lambda e: e[1])
        tol_pt = _THICKNESS_CLUSTER_TOL_MM / 1000 / scale_m_per_unit
        clusters: list[list[tuple[float, float]]] = []
        for area, thin in entries:
            if clusters and thin - clusters[-1][-1][1] <= tol_pt:
                clusters[-1].append((area, thin))
            else:
                clusters.append([(area, thin)])
        for cluster in clusters:
            total_area_pt2 = sum(a for a, _ in cluster)
            avg_thin_pt = sum(t for _, t in cluster) / len(cluster)
            thickness_m = avg_thin_pt * scale_m_per_unit
            if thickness_m <= 0:
                continue
            length_m = (total_area_pt2 * scale_m_per_unit**2) / thickness_m
            if length_m < _MIN_GROUP_LENGTH_M:
                continue
            groups.append(FilledWallGroup(
                color=color,
                thickness_mm=round(thickness_m * 1000, 1),
                length_m=round(length_m, 2),
                n_segments=len(cluster),
            ))
    return groups


@dataclass
class MatchedSegment:
    color: tuple[float, float, float]
    length_m: float
    board_spec: BoardSpec | None
    source: str  # "keynote" | "color_only" | "unmapped_color"
    keynote_code: str | None = None
    keynote_dist_m: float | None = None
    color_spec_mismatch: bool = False


def match_segments_to_keynotes(
    segments: list[WallSegment],
    tags: list[tuple[str, "fitz.Rect"]],
    legend: dict[str, BoardSpec],
    color_key: dict[tuple[float, float, float], BoardSpec | None],
    scale_m_per_unit: float,
    max_match_dist_m: float = DEFAULT_MAX_MATCH_DIST_M,
) -> list[MatchedSegment]:
    """Para cada segmento, acha a tag de keynote mais próxima cujo PRODUTO
    bate com o que a cor já indica (a cor é o filtro confiável de
    produto/espessura — ver `extract_wall_linings_color_key`; o keynote só
    refina dentro daquele produto, dizendo se é 1x ou 2x camada).

    Isso importa em plantas densas (várias salas pequenas perto uma da
    outra): sem esse filtro, o "keynote mais próximo" às vezes pertence à
    parede vizinha, de outro produto — contaminando o resultado com camadas
    erradas. Só quando não existe NENHUM keynote do produto certo por perto
    é que caímos para "color_only" (1 camada assumida, sinalizado como
    risco) — nunca pegamos emprestado o keynote de outro produto.
    """
    max_dist_pt = max_match_dist_m / scale_m_per_unit if scale_m_per_unit else float("inf")
    tag_points = [
        (code, ((rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2))
        for code, rect in tags
    ]

    results: list[MatchedSegment] = []
    for seg in segments:
        length_m = seg.length_pt * scale_m_per_unit
        color_spec = color_key.get(seg.color, "NOT_IN_KEY")
        wanted_product = color_spec.product if isinstance(color_spec, BoardSpec) else None

        best_code = None
        best_dist = float("inf")
        fallback_code = None
        fallback_dist = float("inf")
        for code, (tx, ty) in tag_points:
            dist = ((seg.midpoint[0] - tx) ** 2 + (seg.midpoint[1] - ty) ** 2) ** 0.5
            if dist > max_dist_pt:
                continue
            keynote_spec = legend.get(code)
            if keynote_spec is None:
                continue
            if wanted_product is not None and keynote_spec.product == wanted_product:
                if dist < best_dist:
                    best_dist = dist
                    best_code = code
            elif dist < fallback_dist:
                fallback_dist = dist
                fallback_code = code

        if best_code is not None:
            results.append(
                MatchedSegment(
                    color=seg.color,
                    length_m=length_m,
                    board_spec=legend[best_code],
                    source="keynote",
                    keynote_code=best_code,
                    keynote_dist_m=best_dist * scale_m_per_unit,
                    color_spec_mismatch=False,
                )
            )
        elif color_spec != "NOT_IN_KEY":
            results.append(
                MatchedSegment(
                    color=seg.color,
                    length_m=length_m,
                    board_spec=color_spec,
                    source="color_only",
                )
            )
        elif fallback_code is not None:
            # cor nao mapeada, mas existe um keynote perto de outro produto --
            # ainda melhor que nada, mas fica marcado como divergente.
            results.append(
                MatchedSegment(
                    color=seg.color,
                    length_m=length_m,
                    board_spec=legend[fallback_code],
                    source="keynote",
                    keynote_code=fallback_code,
                    keynote_dist_m=fallback_dist * scale_m_per_unit,
                    color_spec_mismatch=True,
                )
            )
        else:
            results.append(
                MatchedSegment(color=seg.color, length_m=length_m, board_spec=None, source="unmapped_color")
            )
    return results


def summarize_by_board_spec(matched: list[MatchedSegment]) -> dict[BoardSpec, float]:
    """Agrupa por especificação EXATA (espessura + produto + camadas) — é
    aqui que a diferença entre 1 e 2 camadas finalmente aparece separada,
    coisa que agrupar só por cor não conseguia fazer."""
    totals: dict[BoardSpec, float] = {}
    for m in matched:
        if m.board_spec is None:
            continue
        totals[m.board_spec] = totals.get(m.board_spec, 0.0) + m.length_m
    return totals


def analyze_wall_linings_plan(pdf_path: str, page_number: int = 0) -> list[MatchedSegment]:
    """Ponto de entrada: abre a prancha e devolve a lista de segmentos já
    casados com keynote/cor, prontos para `summarize_by_board_spec` ou para
    o relatório de risco (`pdf_takeoff.risk_report`)."""
    doc = fitz.open(pdf_path)
    page = doc[page_number]
    text = page.get_text()

    scale = extract_scale_m_per_unit(text)
    if scale is None:
        raise ValueError(
            f"Não achei a escala (\"1:N\") no texto da prancha {pdf_path!r} — "
            "não dá para converter pontos do PDF em metros sem ela."
        )

    legend = parse_keynote_legend(text)
    color_key = extract_wall_linings_color_key(page)
    tags = find_placed_keynote_tags(page, legend)
    segments = extract_colored_wall_segments(page)

    return match_segments_to_keynotes(segments, tags, legend, color_key, scale)


def main() -> int:
    import argparse
    import os

    from .risk_report import detect_duplicate_blocks, find_wall_segment_risks, write_bilingual_risk_report

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdfs", nargs="+", help="Uma ou mais Wall Linings Plans (.pdf).")
    parser.add_argument("--output", default="wall_linings_report.xlsx", help="Planilha de saída (.xlsx).")
    parser.add_argument("--height", type=float, default=3.0, help="Altura de parede (m) para estimar $ em risco.")
    args = parser.parse_args()

    all_findings = []
    quantities: dict[str, dict[BoardSpec, float]] = {}
    for pdf_path in args.pdfs:
        label = os.path.splitext(os.path.basename(pdf_path))[0]
        matched = analyze_wall_linings_plan(pdf_path)
        totals = summarize_by_board_spec(matched)
        quantities[label] = totals
        all_findings.extend(find_wall_segment_risks(matched, wall_height_m=args.height, sheet_label=label))
        n_unmapped = sum(1 for m in matched if m.source == "unmapped_color")
        n_color_only = sum(1 for m in matched if m.source == "color_only")
        n_total = len(matched)
        print(f"{label}: {n_total} segmentos ({n_total - n_unmapped - n_color_only} por keynote, "
              f"{n_color_only} só por cor, {n_unmapped} cor não mapeada)")

    all_specs = sorted({spec for totals in quantities.values() for spec in totals}, key=lambda s: s.board_type_label)
    named_groups = {
        label: [totals.get(spec, 0.0) for spec in all_specs] for label, totals in quantities.items()
    }
    all_findings.extend(detect_duplicate_blocks(named_groups))

    write_bilingual_risk_report(all_findings, args.output, quantities=quantities)
    print(f"Relatório salvo em: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

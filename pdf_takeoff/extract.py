"""Extração de geometria vetorial e texto de um PDF de planta (CAD/Revit)."""

from __future__ import annotations

from dataclasses import dataclass

import pymupdf as fitz

Point = tuple[float, float]
Segment = tuple[Point, Point]


@dataclass
class TextWord:
    text: str
    bbox: tuple[float, float, float, float]  # x0, y0, x1, y1

    @property
    def center(self) -> Point:
        x0, y0, x1, y1 = self.bbox
        return ((x0 + x1) / 2, (y0 + y1) / 2)


@dataclass
class PageContent:
    page_number: int
    width: float
    height: float
    segments: list[Segment]
    words: list[TextWord]


def _rect_to_segments(rect: fitz.Rect) -> list[Segment]:
    x0, y0, x1, y1 = rect.x0, rect.y0, rect.x1, rect.y1
    return [
        ((x0, y0), (x1, y0)),
        ((x1, y0), (x1, y1)),
        ((x1, y1), (x0, y1)),
        ((x0, y1), (x0, y0)),
    ]


def _extract_segments(page: fitz.Page) -> list[Segment]:
    """Converte os desenhos vetoriais da página (linhas, retângulos, curvas) em segmentos."""
    segments: list[Segment] = []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            kind = item[0]
            if kind == "l":  # linha reta
                p1, p2 = item[1], item[2]
                segments.append(((p1.x, p1.y), (p2.x, p2.y)))
            elif kind == "re":  # retângulo
                segments.extend(_rect_to_segments(item[1]))
            elif kind == "qu":  # quad (4 pontos)
                quad = item[1]
                pts = [quad.ul, quad.ur, quad.lr, quad.ll]
                for a, b in zip(pts, pts[1:] + pts[:1]):
                    segments.append(((a.x, a.y), (b.x, b.y)))
            elif kind == "c":  # curva de bézier: aproxima por segmento reto ponta-a-ponta
                p1, p4 = item[1], item[4]
                segments.append(((p1.x, p1.y), (p4.x, p4.y)))
    return segments


def _extract_words(page: fitz.Page) -> list[TextWord]:
    words = []
    for w in page.get_text("words"):
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        text = text.strip()
        if text:
            words.append(TextWord(text=text, bbox=(x0, y0, x1, y1)))
    return words


def extract_pdf(path: str) -> list[PageContent]:
    """Lê um PDF vetorial e devolve, por página, os segmentos de linha e as palavras de texto."""
    doc = fitz.open(path)
    pages = []
    try:
        for i, page in enumerate(doc):
            pages.append(
                PageContent(
                    page_number=i + 1,
                    width=page.rect.width,
                    height=page.rect.height,
                    segments=_extract_segments(page),
                    words=_extract_words(page),
                )
            )
    finally:
        doc.close()
    return pages

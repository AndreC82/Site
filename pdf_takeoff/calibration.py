"""Detecção automática de escala (metros reais por unidade do PDF) a partir de cotas na planta."""

from __future__ import annotations

import math
import re
import statistics
from dataclasses import dataclass

from .extract import Segment, TextWord

# Números de cota típicos: "3.50", "4,20", "12500", "3500"
_NUMBER_RE = re.compile(r"^\d+([.,]\d+)?$")

# Distância máxima (em unidades do PDF) entre o texto da cota e a linha
# de cota associada, para ainda considerarmos o par válido.
_MAX_LABEL_DISTANCE = 40.0


@dataclass
class ScaleCandidate:
    text: str
    real_meters: float
    segment_length_units: float
    scale_m_per_unit: float


@dataclass
class CalibrationResult:
    scale_m_per_unit: float
    candidates: list[ScaleCandidate]
    confident: bool


def _parse_dimension_value(text: str) -> float | None:
    """Interpreta o texto da cota como metros, aplicando heurística de unidade.

    Valores grandes (>= 100) são tratados como milímetros (padrão comum em
    plantas de arquitetura); valores menores são tratados como metros.
    """
    if not _NUMBER_RE.match(text):
        return None
    value = float(text.replace(",", "."))
    if value <= 0:
        return None
    return value / 1000.0 if value >= 100 else value


def _segment_length(seg: Segment) -> float:
    (x1, y1), (x2, y2) = seg
    return math.hypot(x2 - x1, y2 - y1)


def _point_to_segment_distance(point: tuple[float, float], seg: Segment) -> float:
    (x1, y1), (x2, y2) = seg
    px, py = point
    dx, dy = x2 - x1, y2 - y1
    if dx == dy == 0:
        return math.hypot(px - x1, py - y1)
    t = max(0.0, min(1.0, ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy)))
    proj = (x1 + t * dx, y1 + t * dy)
    return math.hypot(px - proj[0], py - proj[1])


def auto_detect_scale(
    segments: list[Segment], words: list[TextWord]
) -> CalibrationResult:
    """Tenta detectar a escala automaticamente casando textos de cota com segmentos próximos.

    Estratégia: para cada palavra que pareça um número de cota, procura o
    segmento de reta mais próximo (a linha de cota/chamada) e usa
    valor_real / comprimento_em_unidades_pdf como candidato de escala.
    A escala final é a mediana dos candidatos (robusta a outliers).
    """
    candidates: list[ScaleCandidate] = []
    for word in words:
        real_m = _parse_dimension_value(word.text)
        if real_m is None:
            continue
        best_seg = None
        best_dist = _MAX_LABEL_DISTANCE
        for seg in segments:
            dist = _point_to_segment_distance(word.center, seg)
            if dist < best_dist:
                best_dist = dist
                best_seg = seg
        if best_seg is None:
            continue
        length_units = _segment_length(best_seg)
        if length_units <= 0:
            continue
        candidates.append(
            ScaleCandidate(
                text=word.text,
                real_meters=real_m,
                segment_length_units=length_units,
                scale_m_per_unit=real_m / length_units,
            )
        )

    if not candidates:
        # Fallback: assume que o PDF foi exportado em pontos (1/72 pol) 1:1 com
        # o mundo real em metros via escala de impressão comum 1:50 em A1/A3
        # não é seguro adivinhar — melhor devolver 1.0 e marcar como não confiável.
        return CalibrationResult(scale_m_per_unit=1.0, candidates=[], confident=False)

    scales = [c.scale_m_per_unit for c in candidates]
    median_scale = statistics.median(scales)
    # Confiável se tivermos pelo menos 3 candidatos e eles concordarem
    # razoavelmente entre si (baixa dispersão relativa).
    spread = (max(scales) - min(scales)) / median_scale if median_scale else float("inf")
    confident = len(candidates) >= 3 and spread < 0.15
    return CalibrationResult(
        scale_m_per_unit=median_scale, candidates=candidates, confident=confident
    )

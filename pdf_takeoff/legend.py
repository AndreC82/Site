"""Carrega a legenda de códigos (pintura de parede/teto e tipos de gesso) definida pelo usuário."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

VALID_CATEGORIES = {"wall_paint", "ceiling_paint", "drywall"}


@dataclass(frozen=True)
class LegendEntry:
    code: str
    category: str  # wall_paint | ceiling_paint | drywall
    description: str
    layers: int = 1  # relevante só para drywall: 1 = single layer, 2 = double layer

    def __post_init__(self):
        if self.category not in VALID_CATEGORIES:
            raise ValueError(
                f"Categoria inválida '{self.category}' no código '{self.code}'. "
                f"Use uma de: {sorted(VALID_CATEGORIES)}"
            )
        if self.layers < 1:
            raise ValueError(f"'layers' deve ser >= 1 no código '{self.code}'")


def load_legend(path: str) -> dict[str, LegendEntry]:
    """Lê um arquivo JSON de legenda: {"CODIGO": {"category": ..., "description": ..., "layers": ...}}"""
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    legend: dict[str, LegendEntry] = {}
    for code, spec in raw.items():
        legend[code.upper()] = LegendEntry(
            code=code.upper(),
            category=spec["category"],
            description=spec.get("description", ""),
            layers=int(spec.get("layers", 1)),
        )
    return legend


def match_code(text: str, legend: dict[str, LegendEntry]) -> LegendEntry | None:
    """Casa uma palavra extraída do PDF com um código da legenda.

    Tenta correspondência exata primeiro (ignorando maiúsculas/minúsculas e
    pontuação nas bordas); depois tenta prefixo, para textos como
    "P-01/PAREDE" ou "GB-FL-SL(TIPICO)".
    """
    cleaned = text.strip().strip("().,:;").upper()
    if cleaned in legend:
        return legend[cleaned]
    for code, entry in legend.items():
        if cleaned.startswith(code):
            return entry
    return None

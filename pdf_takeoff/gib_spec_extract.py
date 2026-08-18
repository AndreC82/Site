"""Extração de especificações de gesso (Gib) a partir do texto de plantas
arquitetônicas que seguem a convenção de códigos de sistema da GIB (NZ) —
ex.: "Use GBUW120 - 2/19mm Fyreline...". Esses códigos são publicados pela
própria GIB e usados de forma padronizada por vários escritórios de
arquitetura na Nova Zelândia, não são exclusivos de um projeto — por isso
os padrões aqui são genéricos (buscam por palavra-chave/regex em qualquer
página), em vez de fixados em números de prancha de um projeto específico.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import pymupdf as fitz

_PRODUCT_NORMALIZE = {
    "fyreline": "Fireline",
    "fireline": "Fireline",
    "aqualine": "Aqualine",
    "standard": "Standard",
    "braceline": "Braceline",
    "noiseline": "Noiseline",
}

_KNOWN_BOARD_TYPES = {
    "10mm Standard", "10mm Aqualine", "13mm Standard", "13mm Aqualine",
    "13mm Fireline", "13mm Noiseline", "16mm Fireline", "19mm Fireline",
}


@dataclass(frozen=True)
class BoardSpec:
    thickness_mm: int
    product: str  # Standard | Aqualine | Fireline | Braceline | Noiseline
    layers: int = 1

    @property
    def board_type_label(self) -> str:
        return f"{self.thickness_mm}mm {self.product}"

    @property
    def is_known_board_type(self) -> bool:
        return self.board_type_label in _KNOWN_BOARD_TYPES


def _normalize_product(raw: str) -> str:
    return _PRODUCT_NORMALIZE.get(raw.strip().lower(), raw.strip().title())


# --------------------------- legenda de forro de teto -----------------------

_CEILING_LEGEND_RE = re.compile(
    r"Ceiling Linings?\s*(\d+)\s*\n([^\n]+(?:\n[^\n]+){0,2}?)(?=\nCeiling Linings|\nInsulation|\nJoinery|\n[A-Z][a-z]+\n|\Z)",
    re.IGNORECASE,
)
_THICKNESS_PRODUCT_RE = re.compile(
    r"(?:(\d+)\s*/\s*)?(\d+)\s*mm\s+(Fyreline|Fireline|Standard|Aqualine|Braceline|Noiseline)",
    re.IGNORECASE,
)


def extract_ceiling_legend(full_text: str) -> dict[str, BoardSpec]:
    """Lê o bloco de notas 'Ceiling Linings N' e devolve {"C1": BoardSpec(...), ...}.

    Funciona em qualquer planta que use essa nomenclatura (comum em
    especificações de arquitetura na NZ), não depende do número da prancha.
    """
    legend: dict[str, BoardSpec] = {}
    for match in _CEILING_LEGEND_RE.finditer(full_text):
        code = f"C{match.group(1)}"
        desc = match.group(2)
        spec_match = _THICKNESS_PRODUCT_RE.search(desc)
        if not spec_match:
            continue
        layers = int(spec_match.group(1)) if spec_match.group(1) else 1
        thickness = int(spec_match.group(2))
        product = _normalize_product(spec_match.group(3))
        legend[code] = BoardSpec(thickness_mm=thickness, product=product, layers=layers)
    return legend


# --------------------------- nota geral de parede ----------------------------

_WALL_NOTE_RE = re.compile(
    r"Wall Linings\s*\n([^\n]*Gibboard[^\n]*(?:\n[^\n]+){0,1})",
    re.IGNORECASE,
)


def extract_wall_default(full_text: str) -> tuple[BoardSpec | None, BoardSpec | None]:
    """Lê a nota geral 'Wall Linings' (ex.: '10mm Gibboard internal linings
    (Aqualine to wet areas)') e devolve (padrão, área_molhada). Qualquer um
    pode vir None se não for encontrado ou não puder ser interpretado.
    """
    match = _WALL_NOTE_RE.search(full_text)
    if not match:
        return None, None
    desc = match.group(1)
    thickness_match = re.search(r"(\d+)\s*mm", desc)
    if not thickness_match:
        return None, None
    thickness = int(thickness_match.group(1))
    default_spec = BoardSpec(thickness_mm=thickness, product="Standard", layers=1)
    wet_spec = None
    if re.search(r"aqualine", desc, re.IGNORECASE):
        wet_spec = BoardSpec(thickness_mm=thickness, product="Aqualine", layers=1)
    return default_spec, wet_spec


# --------------------------- altura de pé-direito ----------------------------

_STUD_HEIGHT_RE = re.compile(r"(\d{3,5})\s*mm\s+Stud Height", re.IGNORECASE)


def extract_stud_height_m(full_text: str) -> float | None:
    match = _STUD_HEIGHT_RE.search(full_text)
    if not match:
        return None
    return int(match.group(1)) / 1000.0


# --------------------------- callouts de parede resistente a fogo -----------

_FIRE_WALL_RE = re.compile(
    r"Use\s+(GB\w+)\s*-\s*(\d+)\s*/\s*(\d+)\s*mm\s+(Fyreline|Fireline|Standard|Aqualine|Braceline|Noiseline)"
    r"(?:\s+to\s+(interior|exterior|each side(?: of the interior walls indicated)?))?",
    re.IGNORECASE,
)


@dataclass
class FireWallCallout:
    system_code: str  # ex.: "GBUW120"
    spec: BoardSpec
    side: str  # "interior" | "exterior" | "each side" | ""
    rect: "fitz.Rect | None"  # posição na página, se encontrada


def extract_fire_wall_callouts(page: "fitz.Page") -> list[FireWallCallout]:
    """Procura callouts do tipo 'Use GBUW120 - 2/19mm Fyreline to interior'
    na página (tipicamente a planta de combate a incêndio) e devolve cada um
    com sua posição na página, para depois casar com a parede mais próxima.
    """
    text = page.get_text()
    callouts = []
    for match in _FIRE_WALL_RE.finditer(text):
        system_code, layers, thickness, product, side = match.groups()
        spec = BoardSpec(thickness_mm=int(thickness), product=_normalize_product(product), layers=int(layers))
        rects = page.search_for(f"Use {system_code}")
        rect = rects[0] if rects else None
        callouts.append(FireWallCallout(system_code=system_code, spec=spec, side=(side or "").lower(), rect=rect))
    return callouts


# --------------------------- identificação genérica de prancha --------------

def find_pages_with_title_keyword(pages_text: list[str], keyword: str) -> list[int]:
    """Devolve os índices (0-based) de páginas cujo texto contém a palavra-chave
    do título da prancha (ex.: 'Ceiling Plan', 'Fire Control Plan',
    'Floor Plan'). Genérico — não depende de numeração de prancha específica."""
    keyword_lower = keyword.lower()
    return [i for i, text in enumerate(pages_text) if keyword_lower in text.lower()]

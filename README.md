# Quantificador Automático de Pintura e Drywall (PDF → Planilha)

Programa em Python que lê uma planta em PDF vetorial (exportado de CAD/Revit/ArchiCAD),
reconhece automaticamente os ambientes (paredes fechadas) e os códigos usados na planta
para pintura de parede, pintura de teto e tipos de gesso acartonado (Fireline, Aqualine,
Standard, camada única/dupla etc.), e gera uma planilha `.xlsx` pronta para orçamento —
com resumo por código, abas por categoria e um detalhamento por ambiente que você pode
conferir e ajustar manualmente.

## Como funciona (visão geral)

1. **Extração** (`pdf_takeoff/extract.py`): lê os traçados vetoriais (linhas/retângulos das
   paredes) e todos os textos do PDF, com suas posições.
2. **Calibração de escala** (`pdf_takeoff/calibration.py`): procura cotas na planta
   (números como "4.00" ou "3500") próximas a linhas, e calcula automaticamente quantos
   metros reais correspondem a uma unidade do PDF. Se não achar cotas suficientes/consistentes,
   avisa que a escala está incerta e você pode informá-la manualmente com `--scale`.
3. **Reconstrução de ambientes** (`pdf_takeoff/geometry.py`): junta os segmentos de parede
   em polígonos fechados (um por ambiente), tolerando pequenas imprecisões de traçado.
4. **Quantitativo** (`pdf_takeoff/takeoff.py`): para cada ambiente, calcula a área de teto
   (área do polígono) e de parede (perímetro × pé-direito), e associa os códigos de pintura/gesso
   encontrados perto do ambiente, usando a legenda que você define.
5. **Planilha** (`pdf_takeoff/export.py`): gera o `.xlsx` com Resumo, uma aba por categoria
   e o **Detalhe por Ambiente** — a aba editável que serve de fonte de verdade.
6. **Conferência visual** (`pdf_takeoff/visualize.py`): opcionalmente, gera uma cópia do PDF
   original com os contornos detectados e as áreas calculadas desenhados por cima da própria
   planta, para você comparar visualmente com o desenho real.

## Instalação

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Rodando a demonstração (sem precisar de um PDF real)

```bash
python -m pdf_takeoff.demo --output-dir demo_output
```

Isso gera uma planta sintética de duas salas, roda o pipeline inteiro e mostra no
terminal as áreas calculadas, além de salvar `orcamento_exemplo.xlsx` e
`planta_conferencia.pdf` em `demo_output/`.

## Uso com uma planta real

### 1. Defina a legenda de códigos

Crie um JSON como `examples/legend_exemplo.json` mapeando cada código usado na planta
para uma categoria (`wall_paint`, `ceiling_paint` ou `drywall`) e uma descrição.
Para gesso, use `layers: 2` em códigos de camada dupla (double layer) — a planilha
já calcula a área de placa como (área de parede) × camadas.

```json
{
  "P-01": { "category": "wall_paint", "description": "Tinta acrílica branco neve" },
  "PT-01": { "category": "ceiling_paint", "description": "Tinta PVA teto branco" },
  "GB-FL-DL": { "category": "drywall", "description": "Fireline - camada dupla", "layers": 2 }
}
```

### 2. Rode a extração

```bash
python -m pdf_takeoff.cli extract \
  --pdf caminho/para/planta.pdf \
  --legend caminho/para/legenda.json \
  --output orcamento.xlsx \
  --pe-direito 2.80 \
  --review-pdf conferencia.pdf
```

- `--pe-direito`: altura do pé-direito em metros, usada para calcular m² de parede
  (padrão 2.80 m). Se as salas tiverem alturas diferentes, rode em lotes separados
  ou ajuste manualmente na aba "Detalhe por Ambiente" depois.
- `--scale`: força a escala (metros por unidade de PDF) manualmente, caso a
  detecção automática avise que está "INCERTA".
- `--review-pdf`: gera o PDF de conferência com o que foi detectado desenhado
  sobre a própria planta (contorno do ambiente, área de teto/parede e os
  códigos reconhecidos).

### 3. Confira e ajuste

Abra `conferencia.pdf` ao lado de `orcamento.xlsx` para validar visualmente se os
ambientes e códigos foram lidos corretamente. Se algo precisar de correção
(um código não reconhecido, uma área que ficou errada, um ambiente que não
fechou o polígono), edite diretamente a aba **"Detalhe por Ambiente"** da
planilha — ela é a fonte de verdade. Depois, recalcule o Resumo e as abas
por categoria a partir dela:

```bash
python -m pdf_takeoff.cli reimport --input orcamento.xlsx --output orcamento_final.xlsx
```

## Rodando os testes

```bash
python -m pytest tests/ -v
```

## Limitações conhecidas (importante ler antes de confiar no resultado)

- **Só funciona com PDF vetorial** (exportado de CAD/BIM). PDFs escaneados/imagem
  exigiriam OCR e reconhecimento de imagem para achar paredes, o que é bem menos
  confiável — não está implementado nesta primeira versão.
- **Detecção de escala automática depende de cotas legíveis na planta.** Se a
  planta tiver poucas cotas ou cotas inconsistentes, use `--scale` manualmente.
- **Um ambiente com mais de um código da mesma categoria** (ex.: uma parede de
  destaque com cor diferente das outras no mesmo cômodo) tem sua área dividida
  igualmente entre os códigos, com um aviso na coluna "Observações" — ajuste
  manualmente na aba de detalhe se a divisão real não for igual.
- **Vãos de portas/janelas não são descontados automaticamente** da área de
  parede nesta versão — a área de parede é bruta (perímetro × pé-direito).
- **Pé-direito único por planilha** (parametrizável via `--pe-direito`); alturas
  diferentes por ambiente exigem ajuste manual na aba de detalhe.

## Estrutura do projeto

```
pdf_takeoff/
  extract.py       extração de geometria e texto do PDF (PyMuPDF)
  calibration.py   detecção automática de escala a partir de cotas
  geometry.py       reconstrução de polígonos de ambientes (shapely)
  legend.py         legenda de códigos definida pelo usuário
  takeoff.py        associação código↔ambiente e cálculo de quantidades
  export.py         geração/reimportação da planilha .xlsx (openpyxl)
  visualize.py       PDF de conferência com o que foi detectado desenhado
  cli.py            linha de comando (extract / reimport)
  demo.py           gera planta sintética e roda o pipeline completo
examples/
  legend_exemplo.json
tests/
  ...
```

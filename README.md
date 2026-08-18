# Quantificador Automático de Pintura e Drywall (PDF → Planilha)

Programa em Python que lê uma planta em PDF vetorial (exportado de CAD/Revit/ArchiCAD),
reconhece automaticamente os ambientes (paredes fechadas) e os códigos usados na planta
para pintura de parede, pintura de teto e tipos de gesso acartonado (Fireline, Aqualine,
Standard, camada única/dupla etc.), e gera uma planilha `.xlsx` pronta para orçamento —
com resumo por código, abas por categoria e um detalhamento por ambiente que você pode
conferir e ajustar manualmente.

## Extração automática a partir do PDF (plantas com convenção GIB - NZ)

Para plantas de arquitetura que usam a convenção de códigos de sistema GIB
comum na Nova Zelândia (legenda "Ceiling Linings N", nota "Wall Linings",
callouts "Use GBxxx - N/Tmm Fyreline..." na planta de combate a incêndio),
o `pdf_takeoff/pdf_to_input.py` lê o PDF direto e já pré-preenche a planilha
de entrada (Paredes/Tetos) — você só confere e ajusta, em vez de digitar tudo:

```bash
python -m pdf_takeoff.pdf_to_input planta.pdf entrada_preenchida.xlsx --review-pdf conferencia.pdf
```

Isso identifica automaticamente, sem depender do número de prancha (funciona
em qualquer planta com essa convenção, não só um projeto específico):
- a escala real da planta, lendo o "1:100 @ A2" do carimbo (mais confiável
  que tentar casar cota com linha numa planta densa);
- os ambientes (reconstruindo os polígonos de parede a partir do desenho
  vetorial, com um filtro que tenta remover blocos de nota/legenda que não
  são ambientes de verdade);
- o tipo de chapa de cada parede (Standard por padrão, Aqualine se o
  ambiente tiver piso Vinyl/for banheiro/cozinha, ou o tipo lido de um
  callout de parede resistente a fogo na planta de combate a incêndio, se
  houver um perto o suficiente do ambiente);
- o tipo de chapa de cada teto, casando o código (C1, C2...) desenhado na
  planta de teto com a legenda "Ceiling Linings N".

**Limitações conhecidas — sempre confira antes de orçar:**
- A reconstrução de ambientes numa planta real e densa tem margem de erro
  (na planta de teste, ficou ~13% acima da área total declarada na própria
  planta) — pode incluir área externa (deck, portico) ou perder um ambiente
  pequeno ocasionalmente.
- A granularidade é por ambiente inteiro, não por trecho de parede — se um
  ambiente tem uma parede Fireline e outras Standard, só uma das duas fica
  registrada (a mais próxima do callout encontrado).
- Não desconta vãos de porta/janela da área de parede.
- Gera avisos numa aba "Avisos da Extração" quando algo não foi encontrado.

Use sempre o PDF de conferência (`--review-pdf`) lado a lado com a planilha
antes de gerar o orçamento final.

## Planilha de orçamento por taxa (Gib / Stopping / Pintura)

Além do fluxo por PDF acima, o módulo `pdf_takeoff/quantities_workbook.py` gera
uma planilha no formato de orçamento por m² (Gib por tipo/espessura de chapa +
Stopping parede/teto + Pintura, com margem), no mesmo padrão usado por
orçamentistas. Três formas de usar:

**Planilha de entrada (recomendado)** — preenche uma tabela normal do Excel,
com menus suspensos pra evitar erro de digitação, em vez de responder
perguntas num terminal:

```bash
# 1. Gera o modelo em branco (com uma linha de EXEMPLO em cada aba, cinza/itálico)
python -m pdf_takeoff.input_template gerar entrada.xlsx

# 2. Abra entrada.xlsx no Excel e preencha as abas Taxas, Paredes, Tetos e
#    Pintura Avulsa (apague ou ignore as linhas EXEMPLO)

# 3. Gera a planilha de orçamento final a partir do que foi preenchido
python -m pdf_takeoff.input_template construir entrada.xlsx orcamento.xlsx --nome "Nome do Projeto"
```

**Modo interativo** (perguntas em sequência no terminal — mais propenso a erro
se você apertar Enter sem querer, prefira o modo de planilha acima):

```bash
python -m pdf_takeoff.wizard
```

**Modo script**, editando os valores em `pdf_takeoff/demo_quantities.py` e rodando:

```bash
python -m pdf_takeoff.demo_quantities minha_planilha.xlsx
```

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

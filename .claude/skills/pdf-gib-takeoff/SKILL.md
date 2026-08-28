---
name: pdf-gib-takeoff
description: Gera o levantamento de quantidades de gesso (GIB) e pintura — paredes, tetos, metros lineares, m² por tipo de chapa — a partir de um PDF de planta arquitetônica (planta de piso, Wall Linings Plan, Reflected Ceiling Plan, etc.) usando o pipeline determinístico deste repositório (pdf_takeoff). Use esta skill sempre que o usuário enviar/mencionar um PDF de planta e pedir levantamento, quantitativo, metragem de parede, quanto de gib precisa, orçamento de material, ou "processa essa planta" — mesmo sem citar "GIB" ou "pdf_takeoff" explicitamente. Também use para continuar até virar orçamento em $ (planilha de taxas) quando o usuário pedir preço/custo além da quantidade.
---

# Levantamento de GIB/Pintura a partir de PDF de planta

## Por que usar isto em vez de ler o PDF na mão

Esse repositório já tem um extrator geométrico testado (50+ testes) que lê a
planta vetorialmente (linhas, preenchimentos, texto, escala) e produz
metros lineares e m² reais — muito mais preciso e rápido do que estimar a
olho a partir da imagem renderizada. Sempre prefira rodar o pipeline a
"olhar a planta e chutar as quantidades" — só caia pra leitura manual se o
pipeline genuinamente não conseguir processar o arquivo (ex.: PDF
escaneado/imagem, sem geometria vetorial nenhuma).

## Pré-requisito

As dependências (`PyMuPDF`, `shapely`, `openpyxl`, `reportlab`, `Flask`)
precisam estar instaladas. Se um `import pdf_takeoff...` falhar, rode
`pip install -r requirements.txt` na raiz do repo antes de continuar.

## Passo a passo

### 1. Salve o PDF enviado

Copie o(s) PDF(s) que o usuário anexou pra um caminho de trabalho (ex. no
seu scratchpad). Se vierem vários arquivos do mesmo projeto (planta +
especificação, ou várias pranchas separadas), processe cada PDF de planta
individualmente — o extrator já varre todas as páginas de um PDF sozinho,
não precisa dividir por prancha.

### 2. Rode o extrator

```bash
python -m pdf_takeoff.pdf_to_input <planta.pdf> <entrada.xlsx> --review-pdf <conferencia.pdf>
```

Isso tenta, nesta ordem, até achar algo: (1) ambiente reconstruído +
perímetro (convenção GB-code, com detecção de teto e parede resistente a
fogo), (2) linha colorida + keynote próximo ("Wall Linings Plan"), (3)
parede em preenchimento sólido colorido (só mede comprimento — produto
exato de cada cor fica marcado "CONFERIR" pra você resolver no passo 4).
Teto só é detectado automaticamente pelo método (1).

### 3. Se pedir altura de pé-direito

O comando sai com código de erro e imprime algo como:

```
Não achei a altura de pé-direito pra: Página 1.
Rode de novo passando --altura "NIVEL=METROS" pra cada um (ex.: --altura "Página 1=2.7").
```

**Pergunte ao usuário a altura de cada nível listado** (não invente um
valor) e rode de novo com uma flag `--altura` por nível:

```bash
python -m pdf_takeoff.pdf_to_input <planta.pdf> <entrada.xlsx> --review-pdf <conferencia.pdf> \
  --altura "Página 1=3.0" --altura "Ground=2.7"
```

### 4. Se os avisos mencionarem legenda não encontrada ou "CONFERIR"

Leia os avisos impressos pelo comando. Casos comuns e o que fazer:

- **"Não encontrei a nota 'Wall Linings'"** ou legenda de teto não achada:
  a página com a legenda pode existir mas não bater com o regex esperado
  (formato diferente do que o extrator reconhece), ou pode estar mais
  adiante no PDF do que o texto padrão de nota geral. Abra o PDF (`pymupdf`,
  `page.get_text()`) nas páginas que parecem ter legenda/tabela de tipos de
  chapa, copie o texto relevante, e rode de novo passando esse texto:

  ```bash
  python -m pdf_takeoff.pdf_to_input <planta.pdf> <entrada.xlsx> --review-pdf <conferencia.pdf> \
    --info "Wall Linings
  10mm Gibboard internal linings (Aqualine to wet areas)"
  ```

  ou, pra um bloco de texto maior, salve num arquivo e use `--info-file
  caminho.txt`. Esse texto entra na mesma leitura determinística que o
  resto da planta usa — funciona com qualquer formato que os leitores já
  reconheçam (nota "Wall Linings", "Ceiling Linings N", "Keynote Legend").
  Se mesmo colando o texto certo o padrão não bater, **não tente adivinhar
  regex nem reescrever o parser sob pressão** — ao invés disso, complete a
  coluna "Tipo de chapa" direto na planilha de entrada gerada (ver passo 5),
  que é sempre a via de escape segura.

- **Linhas marcadas "CONFERIR tipo de chapa"** (método de preenchimento
  sólido): o comprimento medido está certo, mas o extrator não arrisca
  adivinhar o produto GIB de uma cor sem uma legenda confiável pra ler
  automaticamente nesse formato. Se o usuário souber o que cada cor
  representa (ou você conseguir ler numa legenda "LEGEND"/cor + descrição
  na própria prancha), ou se ele já tiver te dito isso na conversa (ex.: "a
  parede vermelha é Noiseline"), corrija essas linhas direto na planilha
  antes de entregar — não deixe "CONFERIR" na versão final sem avisar o
  usuário.

### 5. Entregue a planilha de entrada e o PDF de conferência

Envie os dois arquivos (`entrada.xlsx` e `conferencia.pdf`) ao usuário.
Explique em português, curto:
- Quantos trechos de parede/ambientes foram detectados e por qual método.
- Os avisos relevantes (o que precisa conferir manualmente).
- Que a planilha de entrada é editável — se algo bater errado com a planta
  real, é pra corrigir ali antes de virar orçamento.

Nunca apresente o resultado como definitivo sem mencionar os avisos — a
extração automática de planta real tem margem de erro, isso é dito
explicitamente em todo o resto deste projeto (ver `README.md`, seção
"Limitações conhecidas") e vale igual aqui.

### 6. Se o usuário quiser o orçamento em $ (não só quantidade)

Você vai precisar de taxas ($/m² de instalação por tipo de chapa, taxa de
stopping, taxa de pintura, etc.). Pergunte se ele tem uma tabela de
referência (de um projeto anterior, ou um orçamento real de fornecedor) ou
se é pra usar taxas de um projeto já processado nesta conversa/repo como
base — não invente valores sem avisar que são estimativa.

Depois de a planilha de entrada estar conferida e com as taxas
preenchidas na aba "Taxas":

```bash
python -m pdf_takeoff.input_template construir <entrada.xlsx> <orcamento.xlsx> --nome "Nome do Projeto"
```

Isso gera a planilha final com GIB/Stopping/Pintura por linha, margem e
contingência (ver estrutura em `pdf_takeoff/quantities_workbook.py` se
precisar entender/ajustar alguma fórmula). Valide fórmulas rodando a
planilha de verdade (biblioteca `formulas` em Python, já usada em outros
lugares deste repo) antes de entregar como definitiva, especialmente se
você editou algo nela manualmente.

## Ferramentas avançadas (fora do fluxo padrão)

- `pdf_takeoff/wall_linings_plan.py` tem um CLI próprio
  (`python -m pdf_takeoff.wall_linings_plan`) pra gerar um **relatório de
  risco bilíngue** (PT/EN) comparando várias pranchas de uma vez e
  apontando duplicatas suspeitas / divergências prováveis por $ de impacto
  — útil quando o usuário está fazendo conferência ("conferência",
  "audit") de um levantamento já feito, não pra gerar o levantamento em si.
- O webapp (`python -m pdf_takeoff.webapp`) expõe o mesmo pipeline numa
  página local com upload arrastando o arquivo — é a via que o usuário usa
  fora desta conversa, no PC dele. Você (Claude, nesta sessão) deve rodar o
  CLI diretamente, não abrir o servidor Flask.

## Coisas que já deram errado antes (não repita)

- Não assuma altura de pé-direito nem tipo de chapa "pra não travar o
  fluxo" — sempre pergunte ou marque como "CONFERIR" explicitamente. Um
  levantamento errado silenciosamente é pior que um que pede confirmação.
- `--altura` espera metros com ponto decimal (`2.7`, não `2,7`) — se o
  usuário responder com vírgula, converta antes de montar o comando.
- Depois de qualquer mudança no código do pacote `pdf_takeoff` (não só ao
  usar a skill, mas se você também estiver desenvolvendo o repo), rode
  `python -m pytest tests/ -q` antes de considerar a mudança pronta.

# Gerador de Slides PowerPoint com Python

Este projeto gera apresentacoes do PowerPoint (`.pptx`) usando Python e a biblioteca `python-pptx`, lendo o conteudo dos slides a partir de um arquivo Markdown.

## O que este projeto faz

- Cria slides programaticamente com layout personalizado
- Aplica estilos visuais (cores, fontes, alinhamento e caixas de texto)
- Le conteudo de slides a partir de `slides.md`
- Suporta layouts `comparativo` e `conceitos`

## Requisitos

- Python 3.9+
- Biblioteca `python-pptx`

## Instalacao

```bash
python -m venv .venv
source .venv/bin/activate
pip install python-pptx
```

## Como usar

Executar com arquivos padrao:

```bash
python criar_slides_agentic_ai.py
```

Executar com arquivos customizados:

```bash
python criar_slides_agentic_ai.py --input slides.md --output Slides_Agentic_AI.pptx
```

Executar com modo premium (perfil + tema externo):

```bash
python criar_slides_agentic_ai.py \
  --input aula2b.md \
  --output aula2b.pptx \
  --profile premium \
  --theme theme.premium.json
```

Executar com modo max (layout inspirado no modelo visual):

```bash
python criar_slides_agentic_ai.py \
  --input aula2b.md \
  --output aula2b-max.pptx \
  --profile max \
  --theme theme.max.json
```

## Schema do Markdown (`slides.md`)

Cada slide deve seguir este contrato:

- Inicio do slide com `# Titulo do Slide`
- Separacao entre slides com `---`
- Subtitulo opcional com `> texto`
- Layout obrigatorio com `layout: comparativo` ou `layout: conceitos`
- Blocos de lista:
  - Cabecalho `## lista: Titulo da lista`
  - Itens `- item 1`
- Blocos de conceito:
  - Cabecalho `## conceito: Titulo do conceito`
  - Descricao em uma ou mais linhas de texto abaixo

Exemplo:

```md
# TITULO DO SLIDE
> Subtitulo opcional
layout: comparativo

## lista: Bloco 1
- Item A
- Item B

## lista: Bloco 2
- Item C

---

# OUTRO SLIDE
layout: conceitos

## conceito: ESTADO
Descricao do conceito.

## conceito: OBJETIVO
Descricao do conceito.
```

## Tema e personalizacao visual

O script permite configurar tipografia, cores, tamanho do slide, espacamento e regras de parsing via arquivo JSON.

- Arquivo de exemplo: `theme.premium.json`
- Arquivos de exemplo: `theme.premium.json` e `theme.max.json`
- Perfis embutidos: `--profile premium` e `--profile max`
- O arquivo de tema sobrescreve os defaults e permite usar fontes instaladas no macOS (ex.: `Gotham HTF`, `Roboto`)

Campos principais do tema:

- `slide`: dimensoes do slide (`width_in`, `height_in`)
- `fonts`: familias e tamanhos base
- `colors`: paleta do slide
- `layout`: margens, espacamentos e posicionamento
- `parsing`: regras como `strip_page_prefix` e `ignore_sections`

Layouts extras do perfil `max`:

- `title_left_text_right`
- `title_top_bullets`
- `title_top_grid_2x2`
- `title_top_text_block`

Knobs especificos do perfil `max` (em `theme.max.json`):

- `title.force_uppercase`: forca titulo em maiusculo
- `title.accent_words`: palavras/frases que vao para a linha de destaque (magenta)
- `max.title`: composicao do titulo (`box_height_in`, `line_spacing`, `prefer_two_lines`, `first_line_max_chars`)
- `max.body`: ritmo tipografico (`line_spacing_text`, `line_spacing_bullets`, `space_before_bullet_pt`)
- `max.density`: limites de conteudo por layout (blocos, bullets, cards e palavras maximas)
- `max.layouts.title_left_text_right`: geometria e tamanho do bloco de texto lateral
- `max.layouts.title_top_bullets`: altura/espacamento de blocos e tamanhos de heading/bullet
- `max.layouts.title_top_grid_2x2`: altura de linha, tamanhos de header/titulo/texto
- `max.layouts.title_top_text_block`: largura, offset e altura do bloco corrido
- `max.background`: fundo por imagem no postech (`image_path`, `apply_after_first_slide`, `cover_first_slide_with_image`)

Exemplo de fundo no `max` (capa sem imagem e demais slides com imagem):

```json
"max": {
  "background": {
    "image_path": "bg-default.png",
    "apply_after_first_slide": true,
    "cover_first_slide_with_image": false
  }
}
```

## Erros comuns de parsing

O script falha com mensagem clara quando encontra:

- Slide sem titulo iniciado por `# `
- Slide sem `layout:` quando nao for possivel inferir
- Layout invalido (diferente de `comparativo` ou `conceitos`)
- `## lista:` sem itens `- ...`
- `## conceito:` sem descricao
- Arquivo sem nenhum slide valido

## Estrutura

- `criar_slides_agentic_ai.py`: parser Markdown + renderizacao dos slides
- `slides.md`: arquivo de entrada com o conteudo da apresentacao
- `theme.premium.json`: exemplo de tema para modo premium
- `theme.max.json`: tema para modo max inspirado no layout de referencia

## Observacao

Os arquivos `.pptx` gerados nao sao versionados neste repositorio.

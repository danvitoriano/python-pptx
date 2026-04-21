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
- Perfil embutido: `--profile premium`
- O arquivo de tema sobrescreve os defaults e permite usar fontes instaladas no macOS (ex.: `Gotham HTF`, `Roboto`)

Campos principais do tema:

- `slide`: dimensoes do slide (`width_in`, `height_in`)
- `fonts`: familias e tamanhos base
- `colors`: paleta do slide
- `layout`: margens, espacamentos e posicionamento
- `parsing`: regras como `strip_page_prefix` e `ignore_sections`

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

## Observacao

Os arquivos `.pptx` gerados nao sao versionados neste repositorio.

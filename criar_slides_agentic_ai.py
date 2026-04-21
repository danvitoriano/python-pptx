import argparse
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Inches, Pt

# Configuracoes de estilo
COR_FUNDO = RGBColor(0, 0, 0)
COR_TITULO = RGBColor(255, 255, 255)
COR_TEXTO = RGBColor(204, 204, 204)
COR_DESTAQUE = RGBColor(102, 179, 255)
FONTE_PRINCIPAL = "Segoe UI"

LAYOUTS_SUPORTADOS = {"comparativo", "conceitos"}


def _ajustar_text_frame(tf):
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def _tamanho_titulo(texto):
    tamanho = len(texto)
    if tamanho > 85:
        return Pt(24)
    if tamanho > 65:
        return Pt(28)
    if tamanho > 45:
        return Pt(32)
    return Pt(40)


def _decorar_titulo(texto):
    mapa = [
        ("termostato", "🌡️"),
        ("climatiza", "🧠"),
        ("estado", "🧠"),
        ("objetivo", "🎯"),
        ("anatomia", "🔧"),
        ("sensores", "📡"),
        ("atuadores", "⚙️"),
        ("motor", "🔄"),
    ]
    base = texto.strip()
    low = base.lower()
    for termo, emoji in mapa:
        if termo in low and emoji not in base:
            return f"{emoji} {base}"
    return base


def configurar_slide(slide):
    """Aplica fundo preto ao slide."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COR_FUNDO


def adicionar_titulo(slide, texto, top=Inches(0.5)):
    """Adiciona titulo principal."""
    textbox = slide.shapes.add_textbox(Inches(0.8), top, Inches(11.8), Inches(1.2))
    tf = textbox.text_frame
    tf.clear()
    _ajustar_text_frame(tf)
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = texto
    run.font.name = FONTE_PRINCIPAL
    run.font.size = _tamanho_titulo(texto)
    run.font.bold = True
    run.font.color.rgb = COR_TITULO
    p.alignment = PP_ALIGN.LEFT
    return textbox


def adicionar_subtitulo(slide, texto, top=Inches(1.3)):
    """Adiciona subtitulo."""
    textbox = slide.shapes.add_textbox(Inches(0.8), top, Inches(11.8), Inches(0.9))
    tf = textbox.text_frame
    tf.clear()
    _ajustar_text_frame(tf)
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = texto
    run.font.name = FONTE_PRINCIPAL
    run.font.size = Pt(24)
    run.font.color.rgb = COR_TEXTO
    p.alignment = PP_ALIGN.LEFT
    return textbox


def adicionar_lista(slide, itens, left, top, width, height, titulo=None, fonte_titulo=20, fonte_item=16):
    """Adiciona caixa com lista de itens."""
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.clear()
    _ajustar_text_frame(tf)

    if titulo:
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = _decorar_titulo(titulo)
        run.font.name = FONTE_PRINCIPAL
        run.font.size = Pt(fonte_titulo)
        run.font.bold = True
        run.font.color.rgb = COR_TITULO

    for item in itens:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = f"• {item}"
        run.font.name = FONTE_PRINCIPAL
        run.font.size = Pt(fonte_item)
        run.font.color.rgb = COR_TEXTO

    return textbox


def adicionar_caixa_conceito(slide, titulo, descricao, left, top, width, height):
    """Adiciona caixa de conceito com titulo e descricao."""
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(26, 26, 26)
    shape.line.color.rgb = RGBColor(100, 100, 100)
    shape.line.width = Pt(1)

    tf = shape.text_frame
    tf.clear()
    _ajustar_text_frame(tf)
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = _decorar_titulo(titulo)
    run.font.name = FONTE_PRINCIPAL
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = COR_DESTAQUE

    p = tf.add_paragraph()
    p.space_before = Pt(10)
    run = p.add_run()
    run.text = descricao
    run.font.name = FONTE_PRINCIPAL
    run.font.size = Pt(14)
    run.font.color.rgb = COR_TEXTO
    p.alignment = PP_ALIGN.LEFT

    return shape


def _limpar_markdown_inline(texto):
    texto = re.sub(r"`([^`]+)`", r"\1", texto)
    texto = re.sub(r"\*\*([^*]+)\*\*", r"\1", texto)
    texto = re.sub(r"\*([^*]+)\*", r"\1", texto)
    return texto.strip()


def _remover_prefixo_pagina(texto):
    """
    Remove prefixos como 'Pagina 2:' ou 'Página 2:' para nao poluir titulos.
    """
    return re.sub(r"^\s*p[aá]gina\s*\d+\s*:\s*", "", texto, flags=re.IGNORECASE).strip()


def _descricao_secao(secao):
    itens = secao.get("itens", [])
    if not itens:
        return ""
    return " ".join(itens).strip()


def _inferir_layout_generico(slide, secoes_genericas):
    if not secoes_genericas:
        return

    titulos = [s["titulo"].lower() for s in secoes_genericas[:2]]
    parece_conceitos = len(secoes_genericas) >= 2 and "estado" in titulos[0] and "objetivo" in titulos[1]

    if parece_conceitos:
        slide["layout"] = "conceitos"
        for secao in secoes_genericas[:2]:
            descricao = _descricao_secao(secao) or secao["titulo"]
            slide["conceitos"].append({"titulo": secao["titulo"], "descricao": descricao})
        for secao in secoes_genericas[2:]:
            itens = secao.get("itens") or [secao["titulo"]]
            slide["listas"].append({"titulo": secao["titulo"], "itens": itens})
        return

    slide["layout"] = "comparativo"
    for secao in secoes_genericas:
        itens = secao.get("itens") or [secao["titulo"]]
        slide["listas"].append({"titulo": secao["titulo"], "itens": itens})


def _parse_slide_block(block, idx):
    linhas = [linha.rstrip() for linha in block.splitlines()]
    linhas = [linha for linha in linhas if linha.strip()]
    if not linhas:
        return None

    primeira_linha = linhas[0].strip()
    if primeira_linha.startswith("# "):
        titulo = primeira_linha[2:].strip()
    elif primeira_linha.startswith("## "):
        titulo = primeira_linha[3:].strip()
    else:
        raise ValueError(f"Slide {idx}: o titulo deve iniciar com '# ' ou '## '.")

    slide = {
        "title": _remover_prefixo_pagina(_limpar_markdown_inline(titulo)),
        "subtitle": "",
        "layout": "",
        "listas": [],
        "conceitos": [],
    }

    h2_indices = []
    for pos, linha in enumerate(linhas[1:], start=1):
        txt = linha.strip()
        if txt.startswith("## ") and not re.match(r"^##\s+(lista|conceito):", txt, re.IGNORECASE):
            h2_indices.append(pos)

    subtitulo_h2_idx = h2_indices[0] if len(h2_indices) == 1 else None
    secoes_genericas = []
    secao_atual = None

    i = 1
    dentro_codigo = False
    while i < len(linhas):
        linha = linhas[i].strip()

        if linha.startswith("```"):
            dentro_codigo = not dentro_codigo
            i += 1
            continue

        if linha.startswith("> "):
            slide["subtitle"] = _limpar_markdown_inline(linha[2:].strip())
            i += 1
            continue

        if linha.lower().startswith("layout:"):
            slide["layout"] = linha.split(":", 1)[1].strip().lower()
            i += 1
            continue

        if subtitulo_h2_idx == i:
            slide["subtitle"] = _limpar_markdown_inline(linha[3:].strip())
            i += 1
            continue

        m_lista = re.match(r"^##\s+lista:\s*(.+)$", linha, re.IGNORECASE)
        if m_lista:
            titulo_lista = _limpar_markdown_inline(m_lista.group(1).strip())
            itens = []
            i += 1
            while i < len(linhas):
                atual = linhas[i].strip()
                if atual.startswith("## "):
                    break
                if atual.startswith(("- ", "* ")):
                    itens.append(_limpar_markdown_inline(atual[2:].strip()))
                i += 1
            if not itens:
                raise ValueError(f"Slide {idx}, lista '{titulo_lista}': lista sem itens.")
            slide["listas"].append({"titulo": titulo_lista, "itens": itens})
            continue

        m_conceito = re.match(r"^##\s+conceito:\s*(.+)$", linha, re.IGNORECASE)
        if m_conceito:
            titulo_conceito = _limpar_markdown_inline(m_conceito.group(1).strip())
            descricao_linhas = []
            i += 1
            while i < len(linhas):
                atual = linhas[i].strip()
                if atual.startswith("## "):
                    break
                descricao_linhas.append(_limpar_markdown_inline(atual))
                i += 1
            descricao = " ".join(descricao_linhas).strip()
            if not descricao:
                raise ValueError(f"Slide {idx}, conceito '{titulo_conceito}': descricao vazia.")
            slide["conceitos"].append({"titulo": titulo_conceito, "descricao": descricao})
            continue

        m_secao = re.match(r"^(##|###|####)\s+(.+)$", linha)
        if m_secao:
            titulo_secao = _remover_prefixo_pagina(_limpar_markdown_inline(m_secao.group(2).strip()))
            if ":" in titulo_secao:
                prefixo, resto = titulo_secao.split(":", 1)
                if prefixo.strip().lower() in {
                    "coluna esquerda",
                    "coluna direita",
                    "titulo",
                    "subtitulo",
                    "subtítulo",
                    "descrição",
                    "descricao",
                }:
                    titulo_secao = resto.strip()
            secao_atual = {"titulo": titulo_secao, "itens": []}
            secoes_genericas.append(secao_atual)
            i += 1
            continue

        if linha.startswith(("- ", "* ")):
            texto_item = _limpar_markdown_inline(linha[2:].strip())
            if not secao_atual:
                secao_atual = {"titulo": "Conteudo", "itens": []}
                secoes_genericas.append(secao_atual)
            m_titulo = re.match(r"^t[íi]tulo:\s*(.+)$", texto_item, re.IGNORECASE)
            if m_titulo:
                secao_atual["titulo"] = m_titulo.group(1).strip()
                i += 1
                continue
            if re.match(r"^descri[cç][aã]o:\s*$", texto_item, re.IGNORECASE):
                i += 1
                continue
            if texto_item:
                secao_atual["itens"].append(texto_item)
            i += 1
            continue

        if linha.startswith("|") and "|" in linha[1:]:
            if not secao_atual:
                secao_atual = {"titulo": "Tabela", "itens": []}
                secoes_genericas.append(secao_atual)
            secao_atual["itens"].append(_limpar_markdown_inline(linha.replace("|", " ")))
            i += 1
            continue

        if not dentro_codigo and linha:
            if not secao_atual:
                secao_atual = {"titulo": "Conteudo", "itens": []}
                secoes_genericas.append(secao_atual)
            secao_atual["itens"].append(_limpar_markdown_inline(linha))

        i += 1

    if not slide["layout"]:
        _inferir_layout_generico(slide, secoes_genericas)

    if not slide["layout"]:
        raise ValueError(
            f"Slide {idx}: nao foi possivel inferir o layout. "
            "Use 'layout: comparativo' ou 'layout: conceitos'."
        )
    if slide["layout"] not in LAYOUTS_SUPORTADOS:
        layouts = ", ".join(sorted(LAYOUTS_SUPORTADOS))
        raise ValueError(f"Slide {idx}: layout invalido '{slide['layout']}'. Use: {layouts}.")
    if slide["layout"] == "comparativo" and not slide["listas"]:
        raise ValueError(f"Slide {idx}: layout comparativo exige ao menos uma secao '## lista:'.")
    if slide["layout"] == "conceitos" and not slide["conceitos"]:
        raise ValueError(f"Slide {idx}: layout conceitos exige ao menos uma secao '## conceito:'.")

    return slide


def carregar_slides_markdown(caminho_markdown):
    """
    Le o arquivo Markdown e retorna uma lista de slides.

    Contrato principal:
    - Cada slide comeca com '# Titulo'
    - Slides podem ser separados por '---' ou '<!-- SLIDE -->'
    - Subtitulo opcional com '> ...'
    - Layout com 'layout: comparativo|conceitos' (ou inferencia automatica)
    - Listas com '## lista: Titulo' + itens '- ...'
    - Conceitos com '## conceito: Titulo' + paragrafo de descricao
    """
    conteudo = Path(caminho_markdown).read_text(encoding="utf-8")
    blocos = [
        b.strip()
        for b in re.split(r"^\s*(?:---|<!--\s*SLIDE\s*-->)\s*$", conteudo, flags=re.MULTILINE)
    ]
    slides = []
    for idx, bloco in enumerate(blocos, start=1):
        if not bloco:
            continue
        slide = _parse_slide_block(bloco, idx)
        if slide:
            slides.append(slide)
    if not slides:
        raise ValueError("Nenhum slide valido encontrado no arquivo Markdown.")
    return slides


def render_slide_comparativo(slide, data):
    posicoes = [
        (Inches(0.8), Inches(2.0), Inches(5.6), Inches(2.4), 20, 14),
        (Inches(6.9), Inches(2.0), Inches(5.6), Inches(2.4), 20, 14),
        (Inches(0.8), Inches(4.9), Inches(11.8), Inches(2.0), 18, 13),
    ]

    for idx, secao in enumerate(data["listas"]):
        if idx < len(posicoes):
            left, top, width, height, ft_titulo, ft_item = posicoes[idx]
        else:
            left = Inches(0.8)
            top = Inches(5.0 + (idx - 2) * 1.2)
            width = Inches(11.8)
            height = Inches(1.5)
            ft_titulo = 16
            ft_item = 12

        adicionar_lista(
            slide,
            itens=secao["itens"],
            left=left,
            top=top,
            width=width,
            height=height,
            titulo=secao["titulo"],
            fonte_titulo=ft_titulo,
            fonte_item=ft_item,
        )


def render_slide_conceitos(slide, data):
    posicoes_conceitos = [
        (Inches(0.8), Inches(1.9), Inches(5.1), Inches(1.9)),
        (Inches(7.0), Inches(1.9), Inches(5.1), Inches(1.9)),
    ]

    for idx, conceito in enumerate(data["conceitos"]):
        if idx < len(posicoes_conceitos):
            left, top, width, height = posicoes_conceitos[idx]
        else:
            left = Inches(0.8)
            top = Inches(3.9 + (idx - 2) * 1.8)
            width = Inches(8.5)
            height = Inches(1.6)
        adicionar_caixa_conceito(
            slide,
            titulo=conceito["titulo"],
            descricao=conceito["descricao"],
            left=left,
            top=top,
            width=width,
            height=height,
        )

    if data["listas"]:
        adicionar_lista(
            slide,
            itens=[f"{item}" for secao in data["listas"] for item in secao["itens"]],
            left=Inches(0.8),
            top=Inches(4.8),
            width=Inches(11.8),
            height=Inches(2.2),
            titulo="Anatomia do Agente",
            fonte_titulo=18,
            fonte_item=12,
        )


def renderizar_apresentacao(slides_data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    for data in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        configurar_slide(slide)
        adicionar_titulo(slide, data["title"])
        if data["subtitle"]:
            adicionar_subtitulo(slide, data["subtitle"])

        if data["layout"] == "comparativo":
            render_slide_comparativo(slide, data)
        elif data["layout"] == "conceitos":
            render_slide_conceitos(slide, data)
        else:
            print(f"Aviso: layout desconhecido '{data['layout']}'. Slide ignorado.")

    return prs


def parse_args():
    parser = argparse.ArgumentParser(description="Gerador de slides PowerPoint a partir de Markdown.")
    parser.add_argument(
        "--input",
        default="slides.md",
        help="Arquivo markdown de entrada (padrao: slides.md).",
    )
    parser.add_argument(
        "--output",
        default="Slides_Agentic_AI.pptx",
        help="Arquivo .pptx de saida (padrao: Slides_Agentic_AI.pptx).",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    try:
        slides_data = carregar_slides_markdown(args.input)
        apresentacao = renderizar_apresentacao(slides_data)
        apresentacao.save(args.output)
        print(f"Slides criados com sucesso: '{args.output}'")
    except FileNotFoundError:
        print(f"Erro: arquivo de entrada nao encontrado: '{args.input}'")
        print("Dica: verifique o nome do arquivo ou passe o caminho completo com --input.")
        sys.exit(1)
    except ValueError as exc:
        print(f"Erro de validacao do markdown: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
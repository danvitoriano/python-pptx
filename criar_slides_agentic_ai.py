from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# Criar apresentação
prs = Presentation()

# Configurações de estilo
COR_FUNDO = RGBColor(0, 0, 0)           # Preto sólido
COR_TITULO = RGBColor(255, 255, 255)    # Branco
COR_TEXTO = RGBColor(204, 204, 204)     # Cinza claro
COR_DESTAQUE = RGBColor(102, 179, 255)  # Azul claro
FONTE_PRINCIPAL = "Segoe UI"

def configurar_slide(slide):
    """Aplica fundo preto ao slide"""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COR_FUNDO

def adicionar_titulo(slide, texto, top=Inches(0.5)):
    """Adiciona título principal"""
    textbox = slide.shapes.add_textbox(Inches(0.8), top, Inches(8.5), Inches(1))
    tf = textbox.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = texto
    run.font.name = FONTE_PRINCIPAL
    run.font.size = Pt(40)
    run.font.bold = True
    run.font.color.rgb = COR_TITULO
    p.alignment = PP_ALIGN.LEFT
    return textbox

def adicionar_subtitulo(slide, texto, top=Inches(1.3)):
    """Adiciona subtítulo"""
    textbox = slide.shapes.add_textbox(Inches(0.8), top, Inches(8.5), Inches(0.8))
    tf = textbox.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = texto
    run.font.name = FONTE_PRINCIPAL
    run.font.size = Pt(24)
    run.font.color.rgb = COR_TEXTO
    p.alignment = PP_ALIGN.LEFT
    return textbox

def adicionar_lista(slide, itens, left, top, width, height, titulo=None, fonte_titulo=20, fonte_item=16):
    """Adiciona caixa com lista de itens"""
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.clear()
    
    if titulo:
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = titulo
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
    """Adiciona caixa de conceito com título e descrição"""
    # Caixa de fundo (opcional - sutil)
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(26, 26, 26)  # Cinza muito escuro
    shape.line.color.rgb = RGBColor(100, 100, 100)
    shape.line.width = Pt(1)
    
    # Título
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = titulo
    run.font.name = FONTE_PRINCIPAL
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = COR_DESTAQUE
    
    # Descrição
    p = tf.add_paragraph()
    p.space_before = Pt(10)
    run = p.add_run()
    run.text = descricao
    run.font.name = FONTE_PRINCIPAL
    run.font.size = Pt(14)
    run.font.color.rgb = COR_TEXTO
    p.alignment = PP_ALIGN.LEFT
    
    return shape

# ============================================================
# SLIDE 1: Diferença entre IA Reativa e Agentes Inteligentes
# ============================================================
slide1 = prs.slides.add_slide(prs.slide_layouts[6])  # Layout em branco
configurar_slide(slide1)

# Título
adicionar_titulo(slide1, "DIFERENÇA ENTRE IA REATIVA E AGENTES INTELIGENTES")

# Subtítulo
adicionar_subtitulo(slide1, "Do mapeamento passivo à autonomia e execução orientada a objetivos")

# Comparativo: Duas colunas
# Coluna 1 - Termostato Comum
adicionar_lista(
    slide1,
    itens=[
        "Resposta direta e inflexível",
        "Apenas desliga ao atingir uma temperatura específica",
        "Atua sobre um único estímulo"
    ],
    left=Inches(0.8),
    top=Inches(2.3),
    width=Inches(4),
    height=Inches(2),
    titulo="🌡️ Termostato Comum"
)

# Coluna 2 - Climatização Inteligente
adicionar_lista(
    slide1,
    itens=[
        "Ajuste proativo",
        "Monitora previsão do tempo, ocupação do recinto e incidência solar",
        "Antecipa o conforto térmico com base em múltiplas variáveis"
    ],
    left=Inches(5.3),
    top=Inches(2.3),
    width=Inches(4),
    height=Inches(2),
    titulo="🧠 Climatização Inteligente"
)

# Seção: O Básico - IA Reativa
adicionar_lista(
    slide1,
    itens=[
        "Ausência de memória de longo prazo nativa",
        "Mapeamento direto de percepções imediatas para ações (Prompt in → Text out)",
        "Abordagem zero-shot: gera respostas sem intenção ou avaliação de impacto real",
        "Soluções frágeis: se o ambiente técnico muda, a IA falha e exige supervisão humana constante"
    ],
    left=Inches(0.8),
    top=Inches(4.5),
    width=Inches(8.5),
    height=Inches(2.2),
    titulo="🔹 O Básico: Inteligência Artificial Reativa",
    fonte_titulo=18,
    fonte_item=15
)

# Destacar "zero-shot" e "Prompt in → Text out" manualmente após gerar, se desejar

# ============================================================
# SLIDE 2: O Salto Evolutivo + Anatomia do Agente
# ============================================================
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
configurar_slide(slide2)

# Título da seção
adicionar_titulo(slide2, "O SALTO EVOLUTIVO", top=Inches(0.3))

# Dois blocos conceituais lado a lado
adicionar_caixa_conceito(
    slide2,
    titulo="🧠 ESTADO",
    descricao="Capacidade de manter o histórico contínuo da operação e o contexto ambiental além de uma janela de conversa isolada.",
    left=Inches(0.8),
    top=Inches(1.2),
    width=Inches(4),
    height=Inches(1.8)
)

adicionar_caixa_conceito(
    slide2,
    titulo="🎯 OBJETIVO",
    descricao="O sistema não se limita a fornecer uma resposta textual; ele dedica-se a alcançar proativamente um estado final predeterminado pelo usuário.",
    left=Inches(5.3),
    top=Inches(1.2),
    width=Inches(4),
    height=Inches(1.8)
)

# Título: Anatomia do Agente
textbox_anatomia = slide2.shapes.add_textbox(Inches(0.8), Inches(3.2), Inches(8.5), Inches(0.5))
tf = textbox_anatomia.text_frame
tf.clear()
p = tf.paragraphs[0]
run = p.add_run()
run.text = "🔧 A Anatomia do Agente"
run.font.name = FONTE_PRINCIPAL
run.font.size = Pt(18)
run.font.bold = True
run.font.color.rgb = COR_TITULO

# Diagrama: 3 caixas horizontais (Sensores → Pensar → Atuadores)
caixas = [
    ("📡 Sensores\n(Percepção)", "Integrações diretas a fontes de dados em tempo real e APIs", Inches(0.8)),
    ("🧠 Pensar", "Processar estado atual em relação ao objetivo final", Inches(3.5)),
    ("⚙️ Atuadores\n(Intervenção)", "Executar códigos, realizar cálculos, gravar arquivos, enviar e-mails", Inches(6.2))
]

for titulo, desc, left_pos in caixas:
    shape = slide2.shapes.add_shape(1, left_pos, Inches(3.7), Inches(2.3), Inches(1.8))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(30, 30, 30)
    shape.line.color.rgb = COR_DESTAQUE
    shape.line.width = Pt(1.5)
    
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    
    # Título da caixa
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = titulo
    run.font.name = FONTE_PRINCIPAL
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = COR_DESTAQUE
    p.alignment = PP_ALIGN.CENTER
    
    # Descrição
    p = tf.add_paragraph()
    p.space_before = Pt(8)
    run = p.add_run()
    run.text = desc
    run.font.name = FONTE_PRINCIPAL
    run.font.size = Pt(10)
    run.font.color.rgb = COR_TEXTO
    p.alignment = PP_ALIGN.CENTER

# Setas entre as caixas (simples, com texto)
for i in range(2):
    arrow = slide2.shapes.add_textbox(Inches(3.2 + i*2.7), Inches(4.4), Inches(0.5), Inches(0.5))
    tf = arrow.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "→"
    run.font.name = FONTE_PRINCIPAL
    run.font.size = Pt(24)
    run.font.color.rgb = COR_TEXTO
    p.alignment = PP_ALIGN.CENTER

# Motor da Autonomia: SENTIR → PENSAR → AGIR
textbox_motor = slide2.shapes.add_textbox(Inches(0.8), Inches(5.8), Inches(8.5), Inches(1.2))
tf = textbox_motor.text_frame
tf.clear()

p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
run = p.add_run()
run.text = "🔄 O MOTOR DA AUTONOMIA"
run.font.name = FONTE_PRINCIPAL
run.font.size = Pt(16)
run.font.bold = True
run.font.color.rgb = COR_TITULO

p = tf.add_paragraph()
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(8)
run = p.add_run()
run.text = "SENTIR  →  PENSAR  →  AGIR"
run.font.name = FONTE_PRINCIPAL
run.font.size = Pt(18)
run.font.bold = True
run.font.color.rgb = COR_DESTAQUE

p = tf.add_paragraph()
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(5)
run = p.add_run()
run.text = "Coletar dados  •  Processar estado  •  Modificar ambiente"
run.font.name = FONTE_PRINCIPAL
run.font.size = Pt(12)
run.font.color.rgb = COR_TEXTO

# Nota final
textbox_nota = slide2.shapes.add_textbox(Inches(1), Inches(6.9), Inches(8), Inches(0.8))
tf = textbox_nota.text_frame
tf.clear()
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
run = p.add_run()
run.text = "A grande revolução não está no tamanho do modelo gerador,\nmas sim em sua inserção neste fluxo estruturado contínuo."
run.font.name = FONTE_PRINCIPAL
run.font.size = Pt(13)
run.font.italic = True
run.font.color.rgb = COR_TEXTO

# ============================================================
# SALVAR ARQUIVO
# ============================================================
prs.save("Slides_Agentic_AI.pptx")
print("✅ Slides criados com sucesso: 'Slides_Agentic_AI.pptx'")
import os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

ROOT = Path(r"c:\Users\wilki\OneDrive\Documentos\Trabajo de Curso")
CHART_DIR = ROOT / "_work" / "charts"
IMG_ANEXO = ROOT / "_work" / "anexo_casa_assados_sofia.png"
IMG_PLANTA_PT = CHART_DIR / "planta_baixa_sofia.png"
IMG_PLANTA_ES = CHART_DIR / "planta_baixa_sofia_es.png"
IMG_BRAND = CHART_DIR / "brand_mockup_sofia.jpg"
IMG_MENU_PRINT = CHART_DIR / "cardapio_impresso_sofia.jpg"
IMG_MENU_WA = CHART_DIR / "cardapio_whatsapp_sofia.jpg"

IMG_EQUIP1 = CHART_DIR / "equip1_asadora_gas.jpg"
IMG_EQUIP2 = CHART_DIR / "equip2_churrasqueira_carvao.jpg"
IMG_EQUIP3 = CHART_DIR / "equip3_coifa_industrial.jpg"
IMG_EQUIP4 = CHART_DIR / "equip4_freezer_horizontal.jpg"
IMG_EQUIP5 = CHART_DIR / "equip5_refrigerador_inox.jpg"
IMG_EQUIP6 = CHART_DIR / "equip6_bancada_balanca.jpg"

IMG_COMBO1 = CHART_DIR / "combo1_classico_sofia.jpg"
IMG_COMBO2 = CHART_DIR / "combo2_costela_sofia.jpg"
IMG_COMBO3 = CHART_DIR / "combo3_dueto_sofia.jpg"
IMG_COMBO4 = CHART_DIR / "combo4_familia_sofia.jpg"

COLOR_NAVY = RGBColor(0x1F, 0x38, 0x64)      # #1F3864
COLOR_RED = RGBColor(0xC0, 0x39, 0x2B)       # #C0392B
COLOR_GOLD = RGBColor(0xD4, 0xAC, 0x0D)      # #D4AC0D
COLOR_DARK = RGBColor(0x2C, 0x3E, 0x50)      # #2C3E50
COLOR_BG_CARD = RGBColor(0xF4, 0xF7, 0xFA)   # #F4F7FA
COLOR_BORDER = RGBColor(0xDC, 0xE4, 0xEC)    # #DCE4EC
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)     # #FFFFFF
COLOR_GRAY = RGBColor(0x7F, 0x8C, 0x8D)      # #7F8C8D
COLOR_LIGHT_BLUE = RGBColor(0xE2, 0xEC, 0xF6)# #E2ECF6
COLOR_GREEN = RGBColor(0x27, 0xAE, 0x60)     # #27AE60

def set_slide_background(slide, color=COLOR_WHITE):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_header(slide, title_text, category_text="CASA DE ASSADOS SOFIA | TCC"):
    # Header bar shape
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = COLOR_NAVY
    top_bar.line.color.rgb = COLOR_NAVY
    
    # Accent line
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.1), Inches(13.333), Inches(0.06))
    accent.fill.solid()
    accent.fill.fore_color.rgb = COLOR_RED
    accent.line.color.rgb = COLOR_RED

    # Title text box
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.12), Inches(12.133), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p_cat = tf.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = COLOR_GOLD
    p_cat.font.name = "Arial"
    p_cat.space_after = Pt(2)
    
    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.size = Pt(20)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_WHITE
    p_title.font.name = "Arial"

def add_footer(slide, current_page=None, total_pages=18, lang="pt"):
    footer_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.15), Inches(13.333), Inches(0.35))
    footer_bar.fill.solid()
    footer_bar.fill.fore_color.rgb = COLOR_BG_CARD
    footer_bar.line.color.rgb = COLOR_BORDER
    
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(7.18), Inches(12.133), Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    
    if lang == "pt":
        p.text = "Colégio Excelência • Curso Técnico em Administração e Informática • Autor: Wilkin Barban Rosabal • 2026"
    else:
        p.text = "Colégio Excelência • Carrera Técnica en Administración e Informática • Autor: Wilkin Barban Rosabal • 2026"
        
    p.font.size = Pt(9)
    p.font.color.rgb = COLOR_GRAY
    p.font.name = "Arial"
    
    if current_page is not None:
        p_num = tf.add_paragraph()
        p_num.text = f"{current_page} / {total_pages}"
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.font.size = Pt(9)
        p_num.font.bold = True
        p_num.font.color.rgb = COLOR_NAVY
        p_num.font.name = "Arial"

def add_card(slide, left, top, width, height, bg_color=COLOR_BG_CARD, border_color=COLOR_BORDER):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(1)
    return shape

def add_bullet_list(tf, items, font_size=13, text_color=COLOR_DARK, space_after=8):
    for idx, item in enumerate(items):
        if idx == 0 and len(tf.paragraphs[0].text) == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(font_size)
        p.font.name = "Arial"
        p.font.color.rgb = text_color
        p.space_after = Pt(space_after)
        p.level = 0

def generate_presentation(lang="pt"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # ----------------------------------------------------
    # SLIDE 1: CAPA / TITLE
    # ----------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1, COLOR_NAVY)
    
    # Background accents
    acc1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.4), Inches(7.5))
    acc1.fill.solid()
    acc1.fill.fore_color.rgb = COLOR_RED
    acc1.line.color.rgb = COLOR_RED
    
    # Text card
    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.0), Inches(7.5), Inches(5.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    
    p = tf1.paragraphs[0]
    p.text = "COLÉGIO EXCELÊNCIA — CURITIBA/PR" if lang == "pt" else "COLÉGIO EXCELÊNCIA — CURITIBA/PR"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD
    p.font.name = "Arial"
    p.space_after = Pt(12)
    
    p = tf1.add_paragraph()
    p.text = "CASA DE ASSADOS SOFIA"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.font.name = "Arial"
    p.space_after = Pt(8)
    
    p = tf1.add_paragraph()
    p.text = "Plano de Negócio para Implantação de Microempresa Gastronômica de Fins de Semana Apoiada por CRM Próprio" if lang == "pt" else "Plan de Negocios para la Implantación de una Microempresa Gastronómica de Fin de Semana Apoyada por CRM Propio"
    p.font.size = Pt(16)
    p.font.color.rgb = COLOR_LIGHT_BLUE
    p.font.name = "Arial"
    p.space_after = Pt(24)
    
    p = tf1.add_paragraph()
    if lang == "pt":
        p.text = "Trabalho de Conclusão de Curso (TCC)\nCurso Técnico em Administração e Informática\n\nAutor: Wilkin Barban Rosabal\nAno Letivo: 2026"
    else:
        p.text = "Trabajo de Conclusión de Curso (TCC)\nCarrera Técnica en Administración e Informática\n\nAutor: Wilkin Barban Rosabal\nAño Lectivo: 2026"
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_WHITE
    p.font.name = "Arial"
    
    # Right Image
    if IMG_BRAND.exists():
        s1.shapes.add_picture(str(IMG_BRAND), Inches(8.7), Inches(1.2), Inches(4.0), Inches(5.1))

    # ----------------------------------------------------
    # SLIDE 2: O PROBLEMA & A OPORTUNIDADE
    # ----------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2, COLOR_WHITE)
    add_header(s2, "O Problema do Setor Tradicional e a Oportunidade de Mercado" if lang == "pt" else "El Problema del Sector Tradicional y la Oportunidad de Mercado")
    add_footer(s2, 2, 18, lang)
    
    # Card 1: Problema
    add_card(s2, 0.8, 1.5, 5.6, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s2.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(5.0), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🚨 O Dilema dos Assadores Convencionais" if lang == "pt" else "🚨 El Dilema de los Asaderos Convencionales"
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED
    p.space_after = Pt(14)
    
    prob_items = [
        "Filas Extensas e Espera Caótica: Clientes aguardam de 30 a 50 min em pé aos domingos." if lang == "pt" else "Filas Extensas y Espera Caótica: Clientes esperan de 30 a 50 min de pie los domingos.",
        "Falta de Previsibilidade de Demanda: Produção 'no escuro' gerando sobra de carne ou falta de produto às 12h30." if lang == "pt" else "Falta de Previsibilidad: Producción 'a ciegas' generando sobras o quiebre de stock a las 12:30.",
        "Desperdício Severo de Insumos: Mermas de 12% a 18% em carnes e guarnições não comercializadas." if lang == "pt" else "Desperdicio Severo: Pérdidas del 12% al 18% en carnes y guarniciones no vendidas.",
        "Zero Relacionamento com o Cliente: Vendas 100% anônimas e passivas, sem histórico nem recompra guiada." if lang == "pt" else "Cero Relación con el Cliente: Ventas 100% anónimas y pasivas, sin historial ni retención."
    ]
    add_bullet_list(tf, prob_items, font_size=12, text_color=COLOR_DARK, space_after=10)

    # Card 2: Oportunidade
    add_card(s2, 6.9, 1.5, 5.6, 5.3, COLOR_LIGHT_BLUE, COLOR_NAVY)
    tb = s2.shapes.add_textbox(Inches(7.2), Inches(1.7), Inches(5.0), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "💡 A Solução: Gastronomia Programada" if lang == "pt" else "💡 La Solución: Gastronomía Programada"
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(14)
    
    sol_items = [
        "Tradição do Churrasco de Domingo: Hábito cultural consolidado nas famílias curitibanas." if lang == "pt" else "Tradición del Asado de Domingo: Hábito cultural arraigado en las familias de Curitiba.",
        "Pré-Venda Ativa com CRM Sofia: Encomendas antecipadas na sexta-feira via WhatsApp próprio." if lang == "pt" else "Preventa Activa con CRM Sofia: Encargos anticipados los viernes vía WhatsApp propio.",
        "Retirada em 90s sem Fila: Janelas de atendimento fracionadas de 15 em 15 minutos." if lang == "pt" else "Retiro en 90s sin Filas: Franjas horarias fraccionadas de 15 en 15 minutos.",
        "Desperdício Reduzido para < 3%: Compras na CEASA calibradas na medida exata da demanda real." if lang == "pt" else "Desperdicio Reducido a < 3%: Compras en CEASA calibradas a la medida exacta de la demanda."
    ]
    add_bullet_list(tf, sol_items, font_size=12, text_color=COLOR_DARK, space_after=10)

    # ----------------------------------------------------
    # SLIDE 3: O CONCEITO DO NEGÓCIO & PROPOSTA DE VALOR
    # ----------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3, COLOR_WHITE)
    add_header(s3, "Conceito do Negócio e Proposta de Valor" if lang == "pt" else "Concepto del Negocio y Propuesta de Valor")
    add_footer(s3, 3, 18, lang)
    
    # Left text
    add_card(s3, 0.8, 1.5, 6.8, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s3.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(6.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "CASA DE ASSADOS SOFIA LTDA."
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(12)
    
    concept_items = [
        ("Slogan Oficial: ", "“O verdadeiro sabor do domingo na mesa da sua família.”" if lang == "pt" else "“El auténtico sabor del domingo en la mesa de su familia.”"),
        ("Localização Estratégica: ", "Rua Deputado Pinheiro Júnior, 1380, Umbará, Curitiba - PR." if lang == "pt" else "Rua Deputado Pinheiro Júnior, 1380, Umbará, Curitiba - PR."),
        ("Formato Operacional: ", "Dark Kitchen / Takeaway com retirada expressa e delivery próprio em raio de 5 km." if lang == "pt" else "Dark Kitchen / Takeaway con retiro exprés y delivery propio en radio de 5 km."),
        ("Dias de Funcionamento: ", "Sábados, Domingos e TODOS os Feriados Nacionais, Estaduais e Municipais." if lang == "pt" else "Sábados, Domingos y TODOS los Feriados Nacionales, Provinciales y Municipales."),
        ("Enquadramento Jurídico: ", "Sociedade Limitada Unipessoal (SLU) no Simples Nacional (Microempresa)." if lang == "pt" else "Sociedad Limitada Unipersonal (SLU) en el Simples Nacional (Microempresa)."),
        ("Diferencial-Chave: ", "4 combos familiares padronizados + Agendamento inteligente via CRM Sofia." if lang == "pt" else "4 combos familiares estandarizados + Agendamiento inteligente vía CRM Sofia.")
    ]
    for bold_lead, text in concept_items:
        p = tf.add_paragraph()
        r1 = p.add_run()
        r1.text = bold_lead
        r1.font.bold = True
        r1.font.size = Pt(12)
        r1.font.color.rgb = COLOR_NAVY
        r2 = p.add_run()
        r2.text = text
        r2.font.size = Pt(12)
        r2.font.color.rgb = COLOR_DARK
        p.space_after = Pt(8)

    # Right 3D Mockup
    if IMG_ANEXO.exists():
        s3.shapes.add_picture(str(IMG_ANEXO), Inches(8.0), Inches(1.5), Inches(4.5), Inches(5.3))

    # ----------------------------------------------------
    # SLIDE 4: O CRM SOFIA COMO COLUNA VERTEBRAL (DESTAQUE)
    # ----------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4, COLOR_WHITE)
    add_header(s4, "CRM Sofia: A Coluna Vertebral Tecnológica do Negócio" if lang == "pt" else "CRM Sofia: La Columna Vertebral Tecnológica del Negocio")
    add_footer(s4, 4, 18, lang)
    
    # 3 Column Cards
    # Col 1: Software Próprio & Stack Livre
    add_card(s4, 0.8, 1.5, 3.7, 5.3, COLOR_BG_CARD, COLOR_NAVY)
    tb = s4.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(3.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "💻 Desenvolvimento Próprio" if lang == "pt" else "💻 Desarrollo Propio In-House"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(10)
    c1_items = [
        "Criado pelo autor integrando Administração + Informática (Colégio Excelência)." if lang == "pt" else "Creado por el autor integrando Administración + Informática (Colégio Excelência).",
        "Domínio Dinâmico Grátis:\nhttps://casadeasados.duckdns.org/" if lang == "pt" else "Dominio Dinámico Gratis:\nhttps://casadeasados.duckdns.org/",
        "Certificado SSL Let's Encrypt (HTTPS seguro e criptografado)." if lang == "pt" else "Certificado SSL Let's Encrypt (HTTPS seguro y cifrado).",
        "Stack 100% Livre: Linux Ubuntu, PostgreSQL, Python/FastAPI e HTML5." if lang == "pt" else "Stack 100% Libre: Linux Ubuntu, PostgreSQL, Python/FastAPI y HTML5."
    ]
    add_bullet_list(tf, c1_items, font_size=11, space_after=8)

    # Col 2: Custo Hiper-Reduzido
    add_card(s4, 4.8, 1.5, 3.7, 5.3, COLOR_LIGHT_BLUE, COLOR_RED)
    tb = s4.shapes.add_textbox(Inches(5.0), Inches(1.7), Inches(3.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "💰 Custo Quase Zero" if lang == "pt" else "💰 Costo Casi Cero"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED
    p.space_after = Pt(10)
    c2_items = [
        "Único Custo Mensal: R$ 50,00/mês de hospedagem em servidor VPS na nuvem." if lang == "pt" else "Único Costo Mensal: R$ 50,00/mes de hospedaje en servidor VPS en la nube.",
        "Zero Licenciamento de software comercial de terceiros." if lang == "pt" else "Cero Licencias de software comercial de terceros.",
        "Economia vs SaaS Comum: Sistemas pagos cobram R$ 300 a R$ 800/mês + taxas." if lang == "pt" else "Ahorro vs SaaS Común: Sistemas pagos cobran R$ 300 a R$ 800/mes + tasas.",
        "Sem Taxas de Marketplace: Evita 20% a 27% de comissão abusiva do iFood." if lang == "pt" else "Sin Comisiones de Marketplace: Evita el 20% al 27% abusivo de iFood."
    ]
    add_bullet_list(tf, c2_items, font_size=11, space_after=8)

    # Col 3: Impacto Operacional & LGPD
    add_card(s4, 8.8, 1.5, 3.7, 5.3, COLOR_BG_CARD, COLOR_GREEN)
    tb = s4.shapes.add_textbox(Inches(9.0), Inches(1.7), Inches(3.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "⚡ Impacto & Fidelização" if lang == "pt" else "⚡ Impacto y Fidelización"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN
    p.space_after = Pt(10)
    c3_items = [
        "Janelas de 15 min: Limita a 6 pedidos/intervalo, eliminando gargalos." if lang == "pt" else "Franjas de 15 min: Limita a 6 pedidos/intervalo, eliminando cuellos de botella.",
        "Painel KDS em Cozinha: Comandas em tempo real para os assadores." if lang == "pt" else "Panel KDS en Cocina: Comandas en tiempo real para los asadores.",
        "Segmentação RFM: Disparos para VIPs e reativação automática." if lang == "pt" else "Segmentación RFM: Mensajes para VIPs y reactivación automática.",
        "Conformidade LGPD: Opt-in / Opt-out nativo e base criptografada." if lang == "pt" else "Conformidad LGPD: Opt-in / Opt-out nativo y base cifrada."
    ]
    add_bullet_list(tf, c3_items, font_size=11, space_after=8)

    # ----------------------------------------------------
    # SLIDE 5: ENGENHARIA DE CARDÁPIO (OS 4 COMBOS)
    # ----------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5, COLOR_WHITE)
    add_header(s5, "Engenharia de Cardápio: Os 4 Combos Padronizados" if lang == "pt" else "Ingeniería de Menú: Los 4 Combos Estandarizados")
    add_footer(s5, 5, 18, lang)
    
    combos_info = [
        ("Combo 1: O Clássico" if lang == "pt" else "Combo 1: El Clásico", "R$ 69,90", "CMV: R$ 26,50 (62,1%)", IMG_COMBO1, "1 Frango recheado inteiro assado + maionese caseira 300g + farofa crocante 250g. (3-4 pessoas)" if lang == "pt" else "1 Pollo relleno entero asado + mayonesa casera 300g + farofa crocante 250g. (3-4 personas)"),
        ("Combo 2: Costela Suprema" if lang == "pt" else "Combo 2: Costilla Suprema", "R$ 119,90", "CMV: R$ 48,00 (60,0%)", IMG_COMBO2, "1,0kg de Costela bovina premium ao bafo (6h) + mandioca na manteiga + vinagrete e farofa. (4 pessoas)" if lang == "pt" else "1,0kg Costilla vacuna al vapor (6h) + mandioca a la manteca + vinagreta y farofa. (4 personas)"),
        ("Combo 3: Dueto Sofia" if lang == "pt" else "Combo 3: Dueto Sofia", "R$ 94,90", "CMV: R$ 36,00 (62,1%)", IMG_COMBO3, "Meio frango dourado + 500g costelinha suína marinada + batatas rústicas e farofa. (3-4 pessoas)" if lang == "pt" else "Medio pollo dorado + 500g costilla de cerdo marinada + patatas rústicas y farofa. (3-4 personas)"),
        ("Combo 4: Kit Família" if lang == "pt" else "Combo 4: Kit Familia", "R$ 169,90", "CMV: R$ 68,00 (60,0%)", IMG_COMBO4, "1 Frango inteiro + 700g costela + 4 linguiças artesanais + 4 pães de alho + maionese e farofa grande. (5-6 pessoas)" if lang == "pt" else "1 Pollo entero + 700g costilla + 4 chorizos parrilleros + 4 panes de ajo + mayonesa y farofa grande. (5-6 personas)")
    ]
    
    for i, (title, price, cmv_str, img_p, desc) in enumerate(combos_info):
        left_pos = 0.8 + i * 2.95
        add_card(s5, left_pos, 1.5, 2.8, 5.3, COLOR_BG_CARD, COLOR_BORDER)
        if img_p.exists():
            s5.shapes.add_picture(str(img_p), Inches(left_pos + 0.15), Inches(1.65), Inches(2.5), Inches(1.8))
            
        tb = s5.shapes.add_textbox(Inches(left_pos + 0.15), Inches(3.55), Inches(2.5), Inches(3.1))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_NAVY
        p.space_after = Pt(2)
        
        p = tf.add_paragraph()
        p.text = price
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = COLOR_RED
        p.space_after = Pt(2)
        
        p = tf.add_paragraph()
        p.text = cmv_str
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GREEN
        p.space_after = Pt(6)
        
        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK

    # ----------------------------------------------------
    # SLIDE 6: IDENTIDADE VISUAL & COMUNICAÇÃO DE MARCA
    # ----------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6, COLOR_WHITE)
    add_header(s6, "Identidade Visual, Slogan e Fachada Comercial" if lang == "pt" else "Identidad Visual, Slogan y Fachada Comercial")
    add_footer(s6, 6, 18, lang)
    
    add_card(s6, 0.8, 1.5, 6.0, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s6.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(5.4), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "Posicionamento Visual de Excelência" if lang == "pt" else "Posicionamiento Visual de Excelencia"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(12)
    
    brand_bullets = [
        "Paleta de Cores Institucional: Vermelho Brasa (#C0392B), Dourado Assado (#D4AC0D), Azul Confiança (#1F3864) e Preto Carvão (#2C3E50)." if lang == "pt" else "Paleta Institucional: Rojo Brasa (#C0392B), Dorado Asado (#D4AC0D), Azul Confianza (#1F3864) y Negro Carbón (#2C3E50).",
        "Fachada Moderna em Madeira e Aço: Revestimento rústico sofisticado com painel preto fosco e letreiro luminoso em acrílico 3D." if lang == "pt" else "Fachada Moderna en Madera y Acero: Revestimiento rústico sofisticado con panel negro mate y letrero luminoso en acrílico 3D.",
        "Embalagens Térmicas Seladas: Caixas e sacolas kraft personalizadas com lacre inviolável '100% Quente'." if lang == "pt" else "Envases Térmicos Sellados: Cajas y bolsas kraft personalizadas con precinto inviolable '100% Caliente'.",
        "Sinalização de Calçada: Cavalete rústico em madeira nobre (1,0m x 0,6m) com menu do dia e QR Code direto para o WhatsApp/CRM." if lang == "pt" else "Señalización de Acera: Caballete rústico en madera (1,0m x 0,6m) con menú del día y QR Code directo al WhatsApp/CRM."
    ]
    add_bullet_list(tf, brand_bullets, font_size=11.5, space_after=8)
    
    if IMG_BRAND.exists():
        s6.shapes.add_picture(str(IMG_BRAND), Inches(7.2), Inches(1.5), Inches(5.3), Inches(5.3))

    # ----------------------------------------------------
    # SLIDE 7: CARDÁPIOS ESTRATÉGICOS (IMPRESSO E WHATSAPP)
    # ----------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7, COLOR_WHITE)
    add_header(s7, "Design dos Cardápios: Versão Impressa de Balcão e Digital WhatsApp" if lang == "pt" else "Diseño de Menús: Versión Impresa de Mostrador y Digital WhatsApp")
    add_footer(s7, 7, 18, lang)
    
    # Left Card: Impresso
    add_card(s7, 0.8, 1.5, 5.6, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s7.shapes.add_textbox(Inches(1.0), Inches(1.65), Inches(5.2), Inches(0.8))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "Cardápio Impresso para Balcão" if lang == "pt" else "Menú Impreso para Mostrador"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    if IMG_MENU_PRINT.exists():
        s7.shapes.add_picture(str(IMG_MENU_PRINT), Inches(1.1), Inches(2.25), Inches(5.0), Inches(4.3))
        
    # Right Card: WhatsApp
    add_card(s7, 6.9, 1.5, 5.6, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s7.shapes.add_textbox(Inches(7.1), Inches(1.65), Inches(5.2), Inches(0.8))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "Cardápio Interativo Mobile / WhatsApp" if lang == "pt" else "Menú Interactivo Mobile / WhatsApp"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    if IMG_MENU_WA.exists():
        s7.shapes.add_picture(str(IMG_MENU_WA), Inches(7.2), Inches(2.25), Inches(5.0), Inches(4.3))

    # ----------------------------------------------------
    # SLIDE 8: ANÁLISE DE MERCADO & PÚBLICO-ALVO
    # ----------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8, COLOR_WHITE)
    add_header(s8, "Análise de Mercado e Microrregião de Curitiba (Umbará)" if lang == "pt" else "Análisis de Mercado y Microrregión de Curitiba (Umbará)")
    add_footer(s8, 8, 18, lang)
    
    # 3 Stat Cards on Top
    stats = [
        ("1.830.795", "Habitantes em Curitiba\n(IBGE 2024/2025)" if lang == "pt" else "Habitantes en Curitiba\n(IBGE 2024/2025)", COLOR_NAVY),
        ("R$ 120,06 Bi", "PIB Municipal de Curitiba\n(1ª da Região Sul / 6ª do Brasil)" if lang == "pt" else "PIB Municipal de Curitiba\n(1ª Región Sur / 6ª de Brasil)", COLOR_RED),
        ("R$ 67.691,30", "PIB per capita oficial\n(Alto poder aquisitivo)" if lang == "pt" else "PIB per cápita oficial\n(Alto poder adquisitivo)", COLOR_GOLD)
    ]
    for i, (val, lbl, col) in enumerate(stats):
        left_pos = 0.8 + i * 4.0
        add_card(s8, left_pos, 1.45, 3.7, 1.4, COLOR_BG_CARD, COLOR_BORDER)
        tb = s8.shapes.add_textbox(Inches(left_pos + 0.1), Inches(1.55), Inches(3.5), Inches(1.2))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = val
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = col
        p.alignment = PP_ALIGN.CENTER
        p = tf.add_paragraph()
        p.text = lbl
        p.font.size = Pt(10)
        p.font.color.rgb = COLOR_DARK
        p.alignment = PP_ALIGN.CENTER

    # Bottom Content
    add_card(s8, 0.8, 3.05, 5.8, 3.8, COLOR_BG_CARD, COLOR_BORDER)
    tb = s8.shapes.add_textbox(Inches(1.0), Inches(3.2), Inches(5.4), Inches(3.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "📍 Microrregião Bairro Novo / Umbará" if lang == "pt" else "📍 Microrregión Bairro Novo / Umbará"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(8)
    mkt_bullets = [
        "População Regional: Mais de 165.000 habitantes nos bairros Umbará, Sítio Cercado e Ganchinho." if lang == "pt" else "Población Regional: Más de 165.000 habitantes en Umbará, Sítio Cercado y Ganchinho.",
        "Perfil Familiar: Classes B2, C1 e C2, famílias com 3 a 6 pessoas residentes em casas próprias e condomínios." if lang == "pt" else "Perfil Familiar: Estratos B2, C1 y C2, familias de 3 a 6 miembros en casas y condominios.",
        "Logística Privilegiada: Acesso rápido pela Rua Nicola Pellanda e Estrada do Ganchinho." if lang == "pt" else "Logística Privilegiada: Acceso rápido por Rua Nicola Pellanda y Estrada do Ganchinho."
    ]
    add_bullet_list(tf, mkt_bullets, font_size=11, space_after=6)

    # Right Chart / Mix
    add_card(s8, 6.9, 3.05, 5.6, 3.8, COLOR_WHITE, COLOR_BORDER)
    mix_chart = CHART_DIR / ("mix_pt.png" if lang == "pt" else "mix_es.png")
    if mix_chart.exists():
        s8.shapes.add_picture(str(mix_chart), Inches(7.05), Inches(3.15), Inches(5.3), Inches(3.6))

    # ----------------------------------------------------
    # SLIDE 9: PLANTA BAIXA & FLUXO SANITÁRIO (RDC 216)
    # ----------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9, COLOR_WHITE)
    add_header(s9, "Planta Baixa Técnica e Fluxo Sanitário Unidirecional (RDC 216)" if lang == "pt" else "Plano Arquitectónico y Flujo Sanitario Unidireccional (RDC 216)")
    add_footer(s9, 9, 18, lang)
    
    # Left Details
    add_card(s9, 0.8, 1.5, 4.8, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s9.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(4.4), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Estrutura Física de 60 m² (10m x 6m)" if lang == "pt" else "Estructura Física de 60 m² (10m x 6m)"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(10)
    
    layout_bullets = [
        "Conformidade Anvisa RDC 216: Fluxo linear unidirecional sem cruzamento entre cru e pronto." if lang == "pt" else "Conformidad Anvisa RDC 216: Flujo lineal unidireccional sin cruce entre crudo y cocido.",
        "Área de Cocção Isolada: Exaustão mecânica profissional com coifa em aço inox e dutos." if lang == "pt" else "Área de Cocción Aislada: Extracción mecánica profesional con campana inox y ductos.",
        "Superfícies Assépticas: Bancadas em aço inox AISI 304 e cubas gastronômicas GN." if lang == "pt" else "Superficies Asépticas: Mesadas en acero inox AISI 304 y cubas gastronómicas GN.",
        "Setor de Expedição Rápida: Balcão térmico de entrega expressa em < 90 segundos." if lang == "pt" else "Sector de Despacho Rápido: Mostrador térmico de entrega exprés en < 90 segundos."
    ]
    add_bullet_list(tf, layout_bullets, font_size=11, space_after=8)
    
    # Right Image
    planta_img = IMG_PLANTA_PT if lang == "pt" else IMG_PLANTA_ES
    if planta_img.exists():
        s9.shapes.add_picture(str(planta_img), Inches(5.9), Inches(1.5), Inches(6.6), Inches(5.3))

    # ----------------------------------------------------
    # SLIDE 10: PARQUE DE EQUIPAMENTOS ADQUIRIDOS
    # ----------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10, COLOR_WHITE)
    add_header(s10, "Catálogo Fotográfico de Maquinários e Ativos Adquiridos" if lang == "pt" else "Catálogo Fotográfico de Maquinaria y Activos Adquiridos")
    add_footer(s10, 10, 18, lang)
    
    equip_items = [
        ("Asadoras a Gás GLP" if lang == "pt" else "Asadoras a Gas GLP", "2x Queimadores infravermelho" if lang == "pt" else "2x Quemadores infrarrojos", IMG_EQUIP1),
        ("Churrasqueira Carvão" if lang == "pt" else "Parrilla a Carbón", "Grelha elevatória e bafo 6h" if lang == "pt" else "Grelha elevadora y vapor 6h", IMG_EQUIP2),
        ("Coifa Industrial Inox" if lang == "pt" else "Campana Industrial Inox", "Exaustão mecânica (VISA)" if lang == "pt" else "Extracción mecánica (VISA)", IMG_EQUIP3),
        ("Freezer Horizontal 510L" if lang == "pt" else "Congelador Horizontal 510L", "Dupla ação (-18°C)" if lang == "pt" else "Doble acción (-18°C)", IMG_EQUIP4),
        ("Refrigerador Inox 4P" if lang == "pt" else "Refrigerador Inox 4P", "Vertical comercial (+2°C)" if lang == "pt" else "Vertical comercial (+2°C)", IMG_EQUIP5),
        ("Bancada Inox + Balança" if lang == "pt" else "Mesada Inox + Balanza", "AISI 304 e Inmetro digital" if lang == "pt" else "AISI 304 e Inmetro digital", IMG_EQUIP6)
    ]
    
    for i, (title, sub, img_p) in enumerate(equip_items):
        row = i // 3
        col = i % 3
        left_pos = 0.8 + col * 4.0
        top_pos = 1.45 + row * 2.75
        add_card(s10, left_pos, top_pos, 3.7, 2.55, COLOR_BG_CARD, COLOR_BORDER)
        if img_p.exists():
            s10.shapes.add_picture(str(img_p), Inches(left_pos + 0.15), Inches(top_pos + 0.15), Inches(1.8), Inches(2.25))
        tb = s10.shapes.add_textbox(Inches(left_pos + 2.05), Inches(top_pos + 0.3), Inches(1.5), Inches(2.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_NAVY
        p.space_after = Pt(4)
        p = tf.add_paragraph()
        p.text = sub
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_DARK

    # ----------------------------------------------------
    # SLIDE 11: ESTRATÉGIA JURÍDICO-TRABALHISTA (DIARISTAS)
    # ----------------------------------------------------
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11, COLOR_WHITE)
    add_header(s11, "Estratégia Jurídico-Trabalhista e Segurança Operacional" if lang == "pt" else "Estrategia Jurídico-Laboral y Seguridad Operativa")
    add_footer(s11, 11, 18, lang)
    
    add_card(s11, 0.8, 1.5, 5.6, 5.3, COLOR_LIGHT_BLUE, COLOR_NAVY)
    tb = s11.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(5.0), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "⚖️ Contrato Intermitente (CLT 452-A)" if lang == "pt" else "⚖️ Contrato Intermitente (CLT 452-A)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(12)
    
    rh_bullets = [
        "Blindagem contra Passivo Trabalhista: Elimina riscos de vínculo clandestino ou ações na Justiça do Trabalho (TRT 9ª Região)." if lang == "pt" else "Blindaje contra Pasivos Laborales: Erradica riesgos de reclamos o juicios en la Justicia del Trabajo (TRT 9ª Región).",
        "Convocação Formal via CRM/WhatsApp: Pré-aviso de 72 horas com confirmação formal do trabalhador." if lang == "pt" else "Convocatoria Formal vía CRM/WhatsApp: Preaviso de 72 horas con aceptación expresa del trabajador.",
        "Pagamento Discriminado por Diária: Salário-base + DSR proporcional + 13º proporcional + Férias + 1/3 constitucional." if lang == "pt" else "Pago Discriminado por Jornal: Salario base + DSR + Aguinaldo proporcional + Vacaciones proporcionales + 1/3.",
        "Exames Ocupacionais (ASO): Atestados médicos admissionais com exames coprológicos para manipuladores de alimentos." if lang == "pt" else "Exámenes Ocupacionales (ASO): Aptitud médica y exámenes coprológicos obligatorios para manipuladores."
    ]
    add_bullet_list(tf, rh_bullets, font_size=11.5, space_after=8)
    
    # Right Table
    add_card(s11, 6.9, 1.5, 5.6, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s11.shapes.add_textbox(Inches(7.1), Inches(1.7), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Quadro de Funções e Diárias" if lang == "pt" else "Estructura de Funciones y Jornales"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED
    p.space_after = Pt(12)
    
    roles = [
        ("Sócio-Administrador", "Wilkin Barban", "Gestão geral, compras e CRM" if lang == "pt" else "Gestión general, compras y CRM"),
        ("Churrasqueiro Chefe", "R$ 120 / diária", "Cocção em grelha e bafo" if lang == "pt" else "Cocción en parrilla y vapor"),
        ("Auxiliar de Cozinha", "R$ 120 / diária", "Pré-preparo e guarnições" if lang == "pt" else "Preparación y guarniciones"),
        ("Auxiliar de Montagem", "R$ 120 / diária", "Embalagem e despacho KDS" if lang == "pt" else "Empaque y despacho KDS"),
        ("Entregador (Motoboy)", "R$ 120 / diária", "Delivery rápido raio 5km" if lang == "pt" else "Delivery rápido radio 5km")
    ]
    for r_title, r_pay, r_desc in roles:
        p = tf.add_paragraph()
        r1 = p.add_run()
        r1.text = f"• {r_title} ({r_pay}): "
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_DARK
        r2 = p.add_run()
        r2.text = r_desc
        r2.font.size = Pt(11)
        r2.font.color.rgb = COLOR_GRAY
        p.space_after = Pt(6)

    # ----------------------------------------------------
    # SLIDE 12: INVESTIMENTO INICIAL & FINANCIAMENTO
    # ----------------------------------------------------
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12, COLOR_WHITE)
    add_header(s12, "Investimento Inicial e Estrutura de Financiamento (R$ 38.000,00)" if lang == "pt" else "Inversión Inicial y Estructura de Financiamiento (R$ 38.000,00)")
    add_footer(s12, 12, 18, lang)
    
    # 2 Big Cards
    add_card(s12, 0.8, 1.5, 5.6, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s12.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(5.0), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "📊 Distribuição do Capital" if lang == "pt" else "📊 Distribución del Capital"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(12)
    
    inv_bullets = [
        "Investimento Fixo em Ativos: R$ 24.500,00 (64,47%)\n(Asadoras, churrasqueira, coifa industrial, freezer 510L, geladeira 4P inox e bancadas)." if lang == "pt" else "Inversión Fija en Activos: R$ 24.500,00 (64,47%)\n(Asadoras, parrilla, campana industrial, congelador 510L, refrigerador 4P inox y mesadas).",
        "Capital de Giro e Pré-Operacional: R$ 13.500,00 (35,53%)\n(Fiança locatícia 3 meses, estoque inicial, embalagens, alvarás, marketing e colchão de reserva)." if lang == "pt" else "Capital de Trabajo y Preoperativos: R$ 13.500,00 (35,53%)\n(Fianza de alquiler 3 meses, inventario inicial, envases, licencias, marketing y colchón de reserva).",
        "TOTAL DO INVESTIMENTO: R$ 38.000,00 (100%)" if lang == "pt" else "TOTAL DE LA INVERSIÓN: R$ 38.000,00 (100%)"
    ]
    add_bullet_list(tf, inv_bullets, font_size=12, space_after=10)

    # Right Card: Fontes de Financiamento
    add_card(s12, 6.9, 1.5, 5.6, 5.3, COLOR_LIGHT_BLUE, COLOR_NAVY)
    tb = s12.shapes.add_textbox(Inches(7.2), Inches(1.7), Inches(5.0), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🏛️ Origem dos Recursos" if lang == "pt" else "🏛️ Origen de los Recursos"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(12)
    
    fin_bullets = [
        "Capital Próprio do Empreendedor:\nR$ 18.000,00 (47,37% do capital total)." if lang == "pt" else "Capital Propio del Emprendedor:\nR$ 18.000,00 (47,37% del capital total).",
        "Microcrédito Fomento Paraná:\nR$ 20.000,00 (52,63% do capital total)." if lang == "pt" else "Microcrédito Fomento Paraná:\nR$ 20.000,00 (52,63% del capital total).",
        "Condições do Empréstimo:\n36 parcelas fixas de R$ 680,00 mensais com juros subsidiados para alimentação." if lang == "pt" else "Condiciones del Préstamo:\n36 cuotas fijas de R$ 680,00 mensuales con tasa subsidiada para microempresas.",
        "Integração no Custo Fixo:\nParcela de R$ 680,00 totalmente absorvida na operação sem comprometer o caixa." if lang == "pt" else "Integración en Costos Fijos:\nCuota de R$ 680,00 totalmente absorbida en la operación sin comprometer liquidez."
    ]
    add_bullet_list(tf, fin_bullets, font_size=12, space_after=10)

    # ----------------------------------------------------
    # SLIDE 13: DRE MENSAL (CENÁRIO BASE)
    # ----------------------------------------------------
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13, COLOR_WHITE)
    add_header(s13, "Demonstrativo de Resultados Mensal (DRE - Cenário Base)" if lang == "pt" else "Estado de Resultados Mensual (DRE - Escenario Base)")
    add_footer(s13, 13, 18, lang)
    
    # Left DRE Chart
    dre_chart = CHART_DIR / ("dre_pt.png" if lang == "pt" else "dre_es.png")
    if dre_chart.exists():
        s13.shapes.add_picture(str(dre_chart), Inches(0.8), Inches(1.5), Inches(5.8), Inches(5.3))
        
    # Right DRE Breakdown
    add_card(s13, 6.9, 1.5, 5.6, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s13.shapes.add_textbox(Inches(7.1), Inches(1.7), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Estrutura de Resultados (160 Combos/Mês)" if lang == "pt" else "Estructura de Resultados (160 Combos/Mes)"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(10)
    
    dre_lines = [
        ("(=) Receita Bruta Operacional", "R$ 15.809,00", "100,00%"),
        ("(-) Custos das Mercadorias (CMV)", "R$ 6.140,00", "38,84%"),
        ("(-) Simples Nacional (4,0%)", "R$ 632,36", "4,00%"),
        ("(-) Taxas de Cartão/PIX (2,0%)", "R$ 316,18", "2,00%"),
        ("(=) Margem de Contribuição Total", "R$ 8.720,46", "55,16%"),
        ("(-) Custos Fixos Operacionais", "R$ 6.870,00", "43,46%"),
        ("(=) LUCRO LÍQUIDO OPERACIONAL", "R$ 1.850,46", "11,71%")
    ]
    for lbl, val, pct in dre_lines:
        p = tf.add_paragraph()
        is_total = "LUCRO" in lbl or "Receita" in lbl or "Margem" in lbl
        r1 = p.add_run()
        r1.text = f"{lbl}: "
        r1.font.bold = is_total
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_NAVY if is_total else COLOR_DARK
        r2 = p.add_run()
        r2.text = f"{val} ({pct})"
        r2.font.bold = is_total
        r2.font.size = Pt(11)
        r2.font.color.rgb = COLOR_RED if "CMV" in lbl or "Custos Fixos" in lbl else (COLOR_GREEN if "LUCRO" in lbl else COLOR_NAVY)
        p.space_after = Pt(4)

    # ----------------------------------------------------
    # SLIDE 14: PONTO DE EQUILÍBRIO & RETORNO (PAYBACK)
    # ----------------------------------------------------
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_background(s14, COLOR_WHITE)
    add_header(s14, "Ponto de Equilíbrio Operacional e Curva de Payback" if lang == "pt" else "Punto de Equilibrio Operativo y Curva de Payback")
    add_footer(s14, 14, 18, lang)
    
    # Left Break-even Chart
    be_chart = CHART_DIR / ("breakeven_pt.png" if lang == "pt" else "breakeven_es.png")
    if be_chart.exists():
        s14.shapes.add_picture(str(be_chart), Inches(0.8), Inches(1.5), Inches(5.8), Inches(5.3))
        
    # Right KPIs Card
    add_card(s14, 6.9, 1.5, 5.6, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s14.shapes.add_textbox(Inches(7.1), Inches(1.7), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🎯 Indicadores de Solidez Financeira" if lang == "pt" else "🎯 Indicadores de Solidez Financiera"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(10)
    
    kpi_items = [
        "Ponto de Equilíbrio em Reais: R$ 12.454,37/mês (faturamento mínimo para cobrir 100% dos custos)." if lang == "pt" else "Punto de Equilibrio en Reales: R$ 12.454,37/mes (facturación mínima para cubrir costos).",
        "Ponto de Equilíbrio em Combos: 126 combos/mês (~32 por fim de semana / ~16 por dia)." if lang == "pt" else "Punto de Equilibrio en Unidades: 126 combos/mes (~32 por fin de semana / ~16 por día).",
        "Margem de Segurança: +21,25% de folga operacional acima do ponto de equilíbrio no cenário base." if lang == "pt" else "Margen de Seguridad: +21,25% de holgura operativa sobre el punto de equilibrio en base.",
        "Prazo de Retorno (Payback): Recuperação total dos R$ 38.000,00 entre o 11º e o 12º mês." if lang == "pt" else "Plazo de Recuperación (Payback): Retorno integral de los R$ 38.000,00 entre el 11º y 12º mes."
    ]
    add_bullet_list(tf, kpi_items, font_size=11.5, space_after=8)

    # ----------------------------------------------------
    # SLIDE 15: ALAVANCAGEM EXTRAORDINÁRIA EM FERIADOS (2026-2028)
    # ----------------------------------------------------
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_background(s15, COLOR_WHITE)
    add_header(s15, "Alavancagem Operacional dos Feriados Úteis (2026-2028)" if lang == "pt" else "Apalancamiento Operativo de los Feriados Hábiles (2026-2028)")
    add_footer(s15, 15, 18, lang)
    
    # 3 Stat Cards on Top
    f_stats = [
        ("28 Feriados", "Dias úteis de alavancagem\n(Ago/2026 a Dez/2028)" if lang == "pt" else "Días hábiles de operación\n(Ago/2026 a Dic/2028)", COLOR_NAVY),
        ("+ R$ 61.759,98", "Receita Bruta Extra no Biênio\n(625 combos adicionais)" if lang == "pt" else "Ingresos Brutos Extraordinarios\n(625 combos adicionales)", COLOR_GOLD),
        ("+ R$ 19.934,90", "Lucro Líquido Extraordinário\n(Colchão livre de liquidez)" if lang == "pt" else "Utilidad Neta Extraordinaria\n(Colchón libre de liquidez)", COLOR_GREEN)
    ]
    for i, (val, lbl, col) in enumerate(f_stats):
        left_pos = 0.8 + i * 4.0
        add_card(s15, left_pos, 1.45, 3.7, 1.4, COLOR_BG_CARD, COLOR_BORDER)
        tb = s15.shapes.add_textbox(Inches(left_pos + 0.1), Inches(1.55), Inches(3.5), Inches(1.2))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = val
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = col
        p.alignment = PP_ALIGN.CENTER
        p = tf.add_paragraph()
        p.text = lbl
        p.font.size = Pt(10)
        p.font.color.rgb = COLOR_DARK
        p.alignment = PP_ALIGN.CENTER

    # Bottom Breakdown
    add_card(s15, 0.8, 3.05, 11.733, 3.8, COLOR_LIGHT_BLUE, COLOR_NAVY)
    tb = s15.shapes.add_textbox(Inches(1.1), Inches(3.2), Inches(11.1), Inches(3.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "💡 A Lógica do Apalancamento em Feriados de Segunda a Sexta" if lang == "pt" else "💡 La Lógica del Apalancamiento en Feriados de Lunes a Viernes"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(8)
    
    feriado_bullets = [
        "Custos Fixos 100% Amortizados: Aluguel, contador, internet, software e parcela do empréstimo já estão pagos pelos fins de semana regulares." if lang == "pt" else "Costos Fijos 100% Amortizados: Alquiler, contador, internet, software y cuota del préstamo ya están cubiertos por los fines de semana.",
        "Economia por Feriado Operado: Cada feriado gera ~R$ 1.976,13 de faturamento com lucro líquido direto de ~R$ 625,06 a R$ 811,32." if lang == "pt" else "Economía por Feriado Operado: Cada feriado genera ~R$ 1.976,13 de facturación con lucro neto directo de ~R$ 625,06 a R$ 811,32.",
        "Conservadorismo Metodológico: Esses R$ 19,9 mil de lucro NÃO foram inflados na DRE base, funcionando como garantia de caixa e expansão acelerada." if lang == "pt" else "Conservadurismo Metodológico: Estos R$ 19,9 mil de utilidad NO se inflaron en la DRE base, sirviendo como colchón de liquidez y expansión."
    ]
    add_bullet_list(tf, feriado_bullets, font_size=12, space_after=8)

    # ----------------------------------------------------
    # SLIDE 16: ANÁLISE DE SENSIBILIDADE & MATRIZ SWOT
    # ----------------------------------------------------
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_background(s16, COLOR_WHITE)
    add_header(s16, "Análise de Sensibilidade e Matriz SWOT / FODA" if lang == "pt" else "Análisis de Sensibilidad y Matriz SWOT / FODA")
    add_footer(s16, 16, 18, lang)
    
    # Left Chart: Scenarios
    scen_chart = CHART_DIR / ("scenarios_pt.png" if lang == "pt" else "scenarios_es.png")
    if scen_chart.exists():
        s16.shapes.add_picture(str(scen_chart), Inches(0.8), Inches(1.5), Inches(5.8), Inches(5.3))
        
    # Right SWOT Summary Card
    add_card(s16, 6.9, 1.5, 5.6, 5.3, COLOR_BG_CARD, COLOR_BORDER)
    tb = s16.shapes.add_textbox(Inches(7.1), Inches(1.7), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Síntese Estratégica SWOT / FODA" if lang == "pt" else "Síntesis Estratégica SWOT / FODA"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(10)
    
    swot_items = [
        "FORÇAS (Strengths): Investimento de R$ 38k com maquinário novo, CRM próprio sem royalties, coifa industrial e ficha técnica padronizada." if lang == "pt" else "FORTALEZAS (Strengths): Inversión de R$ 38k con maquinaria nueva, CRM propio sin royalties, campana industrial y recetas estándar.",
        "OPORTUNIDADES (Opportunities): Demanda insatisfeita com filas em concorrentes e crescimento residencial do Umbará." if lang == "pt" else "OPORTUNIDADES (Opportunities): Demanda insatisfecha con filas en competidores y expansión residencial en Umbará.",
        "FRAQUEZAS (Weaknesses): Marca nova sem base inicial (superada por degustações e pré-venda ativa)." if lang == "pt" else "DEBILIDADES (Weaknesses): Marca nueva sin base inicial (superada con degustaciones y preventa activa).",
        "AMEAÇAS (Threats): Inflação de carnes (mitigada por 3 frigoríficos homologados SIF e compras CEASA)." if lang == "pt" else "AMENAZAS (Threats): Inflación de carnes (mitigada con 3 frigoríficos homologados SIF y compras CEASA)."
    ]
    add_bullet_list(tf, swot_items, font_size=11, space_after=6)

    # ----------------------------------------------------
    # SLIDE 17: SEGURANÇA TECNOLÓGICA, BACKUPS & LGPD
    # ----------------------------------------------------
    s17 = prs.slides.add_slide(blank_layout)
    set_slide_background(s17, COLOR_WHITE)
    add_header(s17, "Governança de Dados, Segurança e Conformidade com a LGPD" if lang == "pt" else "Gobernanza de Datos, Seguridad y Cumplimiento de la LGPD")
    add_footer(s17, 17, 18, lang)
    
    # 3 Cards
    add_card(s17, 0.8, 1.5, 3.7, 5.3, COLOR_BG_CARD, COLOR_NAVY)
    tb = s17.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(3.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🔒 Privacidade & LGPD" if lang == "pt" else "🔒 Privacidad y LGPD"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.space_after = Pt(10)
    lgpd_items = [
        "Consentimento Explícito: Registro de opt-in formal em cada cadastro." if lang == "pt" else "Consentimiento Explícito: Registro de opt-in formal en cada alta.",
        "Descadastramento Fácil: Comando automático 'SAIR' retira o cliente imediatamente." if lang == "pt" else "Baja Inmediata: Comando automático 'SALIR' desuscribe al cliente al instante.",
        "Minimização de Dados: Coleta apenas de nome, telefone e endereço de entrega." if lang == "pt" else "Minimización de Datos: Recolección exclusiva de nombre, teléfono y domicilio."
    ]
    add_bullet_list(tf, lgpd_items, font_size=11, space_after=8)

    add_card(s17, 4.8, 1.5, 3.7, 5.3, COLOR_BG_CARD, COLOR_RED)
    tb = s17.shapes.add_textbox(Inches(5.0), Inches(1.7), Inches(3.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🛡️ Criptografia & Backup" if lang == "pt" else "🛡️ Cifrado y Respaldos"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED
    p.space_after = Pt(10)
    sec_items = [
        "Tráfego 100% Criptografado: Protocolo TLS 1.3 e HTTPS no domínio DuckDNS." if lang == "pt" else "Tráfico 100% Cifrado: Protocolo TLS 1.3 y HTTPS en el dominio DuckDNS.",
        "Dump Diário Automatizado: Backup do PostgreSQL com cifra AES-256." if lang == "pt" else "Volcado Diario Automatizado: Respaldo de PostgreSQL con cifrado AES-256.",
        "RPO < 24h e RTO < 15 min: Recuperação quase instantânea em caso de desastre." if lang == "pt" else "RPO < 24h y RTO < 15 min: Recuperación casi instantánea ante fallas."
    ]
    add_bullet_list(tf, sec_items, font_size=11, space_after=8)

    add_card(s17, 8.8, 1.5, 3.7, 5.3, COLOR_BG_CARD, COLOR_GREEN)
    tb = s17.shapes.add_textbox(Inches(9.0), Inches(1.7), Inches(3.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "📊 Auditoria & Controle" if lang == "pt" else "📊 Auditoría y Control"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN
    p.space_after = Pt(10)
    aud_items = [
        "Trilha de Auditoria: Logs de acesso aos dados de clientes." if lang == "pt" else "Pistas de Auditoría: Logs de acceso a los datos de clientes.",
        "Autonomia Completa: Sem dependência de plataformas opacas de terceiros." if lang == "pt" else "Autonomía Completa: Sin dependencia de plataformas opacas de terceros.",
        "Conformidade Técnica: Alinhado à formação técnica em Informática." if lang == "pt" else "Conformidad Técnica: Alineado a la formación técnica en Informática."
    ]
    add_bullet_list(tf, aud_items, font_size=11, space_after=8)

    # ----------------------------------------------------
    # SLIDE 18: CONCLUSÃO & ENCERRAMENTO
    # ----------------------------------------------------
    s18 = prs.slides.add_slide(blank_layout)
    set_slide_background(s18, COLOR_NAVY)
    
    # Text
    tb = s18.shapes.add_textbox(Inches(1.5), Inches(1.2), Inches(10.333), Inches(5.2))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "CONCLUSÃO E CONSIDERAÇÕES FINAIS" if lang == "pt" else "CONCLUSIÓN Y CONSIDERACIONES FINALES"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(20)
    
    concl_items = [
        "Viabilidade Integral Comprovada: Modelo comercial, técnico, operacional e financeiro robusto com investimento de R$ 38.000,00." if lang == "pt" else "Viabilidad Integral Comprobada: Modelo comercial, técnico, operativo y financiero sólido con inversión de R$ 38.000,00.",
        "Diferencial Tecnológico Real: O CRM Sofia resolve as dores clássicas de filas e desperdício de insumos com custo marginal de apenas R$ 50/mês." if lang == "pt" else "Diferencial Tecnológico Real: El CRM Sofia resuelve los problemas de filas y desperdicio de insumos con costo marginal de R$ 50/mes.",
        "Sinergia Interdisciplinar: Aplicação prática e harmônica entre Administração e Informática aprendidas no Colégio Excelência." if lang == "pt" else "Sinergia Interdisciplinaria: Aplicación práctica y armónica entre Administración e Informática aprendidas en el Colégio Excelência.",
        "Prontidão para Implantação: Todas as fichas técnicas, planta sanitária, contratos e ferramentas digitais validados para execução em Curitiba." if lang == "pt" else "Listos para la Ejecución: Fichas técnicas, plano sanitario, contratos y herramientas digitales validados para su apertura en Curitiba."
    ]
    for it in concl_items:
        p = tf.add_paragraph()
        p.text = f"✔  {it}"
        p.font.size = Pt(14)
        p.font.color.rgb = COLOR_WHITE
        p.font.name = "Arial"
        p.space_after = Pt(12)
        
    p = tf.add_paragraph()
    p.text = "MUITO OBRIGADO! / ¡MUCHAS GRACIAS!"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(16)
    
    output_filename = "Apresentacao_Casa_de_Assados_Sofia_Portugues.pptx" if lang == "pt" else "Presentacion_Casa_de_Assados_Sofia_Espanol.pptx"
    out_path = ROOT / output_filename
    prs.save(str(out_path))
    print(f"Presentation saved successfully: {out_path}")

if __name__ == "__main__":
    generate_presentation("pt")
    generate_presentation("es")

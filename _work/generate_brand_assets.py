import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

ROOT = Path(r"c:\Users\wilki\OneDrive\Documentos\Trabajo de Curso")
CHART_DIR = ROOT / "_work" / "charts"
CHART_DIR.mkdir(parents=True, exist_ok=True)

# -------------------------------------------------------------
# 1. GENERATE BRAND IDENTITY & SIGNAGE (PORTUGUESE & SPANISH)
# -------------------------------------------------------------
def generate_brand_identity(lang="pt"):
    fig, ax = plt.subplots(figsize=(12, 8), dpi=300)
    fig.patch.set_facecolor('#F8F9FA')
    ax.set_facecolor('#F8F9FA')
    ax.axis('off')
    
    title = "MANUAL DE IDENTIDADE VISUAL E COMUNICAÇÃO DE MARCA" if lang == "pt" else "MANUAL DE IDENTIDAD VISUAL Y COMUNICACIÓN DE MARCA"
    slogan = "“O verdadeiro sabor do domingo na mesa da sua família.”" if lang == "pt" else "“El auténtico sabor del domingo en la mesa de su familia.”"
    sub_slogan = "Tradição artesanal • Reserva sem fila • Entrega pontual" if lang == "pt" else "Tradición artesanal • Reserva sin filas • Entrega puntual"
    
    # Header box
    hdr = patches.FancyBboxPatch((0.02, 0.86), 0.96, 0.12, boxstyle="round,pad=0.01", facecolor='#1F3864', edgecolor='none')
    ax.add_patch(hdr)
    ax.text(0.5, 0.94, "CASA DE ASSADOS SOFIA", ha='center', va='center', fontsize=18, fontweight='bold', color='#D4AC0D')
    ax.text(0.5, 0.89, f"{title} | Curitiba - PR", ha='center', va='center', fontsize=10, color='#FFFFFF')

    # Brandmark and Logo Section (Left Box: x=0.02..0.48, y=0.48..0.83)
    b_box = patches.FancyBboxPatch((0.02, 0.48), 0.46, 0.35, boxstyle="round,pad=0.01", facecolor='#FFFFFF', edgecolor='#BDC3C7', linewidth=1)
    ax.add_patch(b_box)
    
    logo_title = "1. LOGOMARCA E SLOGANS OFICIAIS" if lang == "pt" else "1. LOGOTIPO Y SLOGANS OFICIALES"
    ax.text(0.25, 0.80, logo_title, ha='center', va='center', fontsize=10, fontweight='bold', color='#1F3864')
    
    # Emblem badge
    circle = patches.Circle((0.25, 0.67), 0.08, facecolor='#C0392B', edgecolor='#D4AC0D', linewidth=2.5)
    ax.add_patch(circle)
    ax.text(0.25, 0.69, "SOFIA", ha='center', va='center', fontsize=12, fontweight='bold', color='#FFFFFF')
    ax.text(0.25, 0.65, "ASSADOS", ha='center', va='center', fontsize=8, fontweight='bold', color='#D4AC0D')
    ax.text(0.25, 0.62, "★ 2026 ★", ha='center', va='center', fontsize=6.5, color='#FADBD8')
    
    ax.text(0.25, 0.55, f"Slogan Principal:\n{slogan}", ha='center', va='center', fontsize=8.5, fontweight='bold', fontstyle='italic', color='#2C3E50')
    ax.text(0.25, 0.50, f"Assinatura Operacional: {sub_slogan}", ha='center', va='center', fontsize=7.5, color='#7F8C8D')

    # Color Palette Section (Right Box: x=0.52..0.98, y=0.48..0.83)
    c_box = patches.FancyBboxPatch((0.52, 0.48), 0.46, 0.35, boxstyle="round,pad=0.01", facecolor='#FFFFFF', edgecolor='#BDC3C7', linewidth=1)
    ax.add_patch(c_box)
    
    pal_title = "2. PALETA CROMÁTICA INSTITUCIONAL" if lang == "pt" else "2. PALETA CROMÁTICA INSTITUCIONAL"
    ax.text(0.75, 0.80, pal_title, ha='center', va='center', fontsize=10, fontweight='bold', color='#1F3864')
    
    colors = [
        ("#C0392B", "Vermelho Brasa\n#C0392B (Energia/Carne)" if lang == "pt" else "Rojo Brasa\n#C0392B (Energía/Carne)"),
        ("#D4AC0D", "Dourado Assado\n#D4AC0D (Sabor/Nobreza)" if lang == "pt" else "Dorado Asado\n#D4AC0D (Sabor/Calidad)"),
        ("#1F3864", "Azul Confiança\n#1F3864 (Segurança/CRM)" if lang == "pt" else "Azul Confianza\n#1F3864 (Seguridad/CRM)"),
        ("#2C3E50", "Preto Carvão\n#2C3E50 (Tradição Fogo)" if lang == "pt" else "Negro Carbón\n#2C3E50 (Tradición Fuego)")
    ]
    for idx, (hex_col, label) in enumerate(colors):
        x_pos = 0.55 + (idx % 2) * 0.22
        y_pos = 0.68 - (idx // 2) * 0.12
        swatch = patches.Rectangle((x_pos, y_pos), 0.05, 0.07, facecolor=hex_col, edgecolor='#7F8C8D', linewidth=0.5)
        ax.add_patch(swatch)
        ax.text(x_pos + 0.06, y_pos + 0.035, label, ha='left', va='center', fontsize=7, color='#2C3E50')

    # Signage Section 1: Storefront Sign (Fachada) (Bottom-Left: x=0.02..0.48, y=0.04..0.44)
    s_box = patches.FancyBboxPatch((0.02, 0.04), 0.46, 0.40, boxstyle="round,pad=0.01", facecolor='#FFFFFF', edgecolor='#BDC3C7', linewidth=1)
    ax.add_patch(s_box)
    
    sign_title = "3. CARTAZ DE FACHADA COMERCIAL (3,0m x 0,8m)" if lang == "pt" else "3. CARTEL DE FACHADA COMERCIAL (3,0m x 0,8m)"
    ax.text(0.25, 0.41, sign_title, ha='center', va='center', fontsize=9.5, fontweight='bold', color='#1F3864')
    
    # Storefront Sign graphic
    sign_rect = patches.Rectangle((0.05, 0.23), 0.40, 0.14, facecolor='#1F3864', edgecolor='#D4AC0D', linewidth=3)
    ax.add_patch(sign_rect)
    ax.text(0.25, 0.33, "CASA DE ASSADOS SOFIA", ha='center', va='center', fontsize=12, fontweight='bold', color='#D4AC0D')
    ax.text(0.25, 0.28, "CHURRASCO NO BAFO • FRANGO ASSADO • COMBOS", ha='center', va='center', fontsize=7, fontweight='bold', color='#FFFFFF')
    ax.text(0.25, 0.25, "Rua Dep. Pinheiro Júnior, 1380 • Umbará • WhatsApp (41) 9XXXX-XXXX", ha='center', va='center', fontsize=6, color='#EAECEE')
    
    desc_fac = "Lona frontlight de alta durabilidade com iluminação LED superior e contraste elegante." if lang == "pt" else "Lona frontlight de alta durabilidad con iluminación LED superior y contraste elegante."
    ax.text(0.25, 0.12, desc_fac, ha='center', va='center', fontsize=7.5, fontstyle='italic', color='#566573', wrap=True)

    # Signage Section 2: Sidewalk A-Frame Sign & Stickers (Bottom-Right: x=0.52..0.98, y=0.04..0.44)
    a_box = patches.FancyBboxPatch((0.52, 0.04), 0.46, 0.40, boxstyle="round,pad=0.01", facecolor='#FFFFFF', edgecolor='#BDC3C7', linewidth=1)
    ax.add_patch(a_box)
    
    a_title = "4. CAVALETE DE CALÇADA E LACRES DE SEGURANÇA" if lang == "pt" else "4. CABALLETE DE ACERA Y SELLOS DE SEGURIDAD"
    ax.text(0.75, 0.41, a_title, ha='center', va='center', fontsize=9.5, fontweight='bold', color='#1F3864')
    
    # A-Frame drawing
    a_frame = patches.Polygon([[0.56, 0.15], [0.65, 0.35], [0.74, 0.15]], closed=True, facecolor='#2C3E50', edgecolor='#D4AC0D', linewidth=1.5)
    ax.add_patch(a_frame)
    ax.text(0.65, 0.26, "ASSADOS\nDE HOJE!\nFrango R$ 69,90\nCostela R$ 119,90\nPeça no CRM", ha='center', va='center', fontsize=5.5, fontweight='bold', color='#FFFFFF')
    
    # Security Seal sticker
    seal = patches.Circle((0.87, 0.25), 0.07, facecolor='#C0392B', edgecolor='#D4AC0D', linewidth=1.5)
    ax.add_patch(seal)
    ax.text(0.87, 0.27, "LACRE DE\nSEGURANÇA" if lang == "pt" else "SELLO DE\nSEGURIDAD", ha='center', va='center', fontsize=6, fontweight='bold', color='#FFFFFF')
    ax.text(0.87, 0.21, "Sofia ★ 100% Quente" if lang == "pt" else "Sofia ★ 100% Caliente", ha='center', va='center', fontsize=5, color='#F9E79F')
    
    desc_extra = "Cavalete promocional de madeira tratada (1,0x0,6m) e lacre inviolável para embalagens térmicas." if lang == "pt" else "Caballete promocional de madera (1,0x0,6m) y sello inviolable para envases térmicos."
    ax.text(0.75, 0.08, desc_extra, ha='center', va='center', fontsize=7.5, fontstyle='italic', color='#566573')

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    
    out_file = CHART_DIR / f"identidade_visual_sofia_{lang}.png"
    plt.tight_layout()
    fig.savefig(out_file)
    plt.close()
    print(f"Generated: {out_file}")

# -------------------------------------------------------------
# 2. GENERATE VISUAL MENU CARD (CARDÁPIO COMERCIAL)
# -------------------------------------------------------------
def generate_menu_card(lang="pt"):
    fig, ax = plt.subplots(figsize=(10, 14), dpi=300)
    fig.patch.set_facecolor('#1F3864')
    ax.set_facecolor('#1F3864')
    ax.axis('off')
    
    # Background frame
    outer_border = patches.FancyBboxPatch((0.02, 0.02), 0.96, 0.96, boxstyle="round,pad=0.01", facecolor='#FDFEFE', edgecolor='#D4AC0D', linewidth=4)
    ax.add_patch(outer_border)
    
    # Header banner
    hdr = patches.Rectangle((0.03, 0.86), 0.94, 0.11, facecolor='#1F3864', edgecolor='none')
    ax.add_patch(hdr)
    
    ax.text(0.5, 0.93, "CASA DE ASSADOS SOFIA", ha='center', va='center', fontsize=22, fontweight='bold', color='#D4AC0D')
    sub_text = "CARDÁPIO DE FIM DE SEMANA • SABOR ARTESANAL E PONTUALIDADE" if lang == "pt" else "MENÚ DE FIN DE SEMANA • SABOR ARTESANAL Y PUNTUALIDAD"
    ax.text(0.5, 0.885, sub_text, ha='center', va='center', fontsize=9.5, fontweight='bold', color='#FFFFFF')

    # Menu Items
    items_pt = [
        ("1. O CLÁSSICO DA SOFIA", "R$ 69,90", "Serve 3 a 4 pessoas", 
         "• 1 Frango inteiro recheado com farofa especial (~1,4kg assado dourado)\n• Farofa artesanal crocante da casa (250g)\n• Maionese caseira tradicional de batatas (300g)",
         "#C0392B"),
        ("2. COSTELA SUPREMA AO BAFO", "R$ 119,90", "Serve 4 pessoas", 
         "• 1kg de Costela bovina premium assada lentamente por 6 horas\n• Mandioca macia na manteiga de garrafa (300g)\n• Vinagrete fresco artesanal e Farofa crocante da casa (250g)",
         "#922B21"),
        ("3. DUETO SOFIA (Frango + Suíno)", "R$ 94,90", "Serve 3 a 4 pessoas", 
         "• Meio frango assado suculento + 500g de Costelinha suína em ervas finas\n• Batatas rústicas douradas ao alecrim (300g)\n• Farofa artesanal especial da casa (200g)",
         "#D35400"),
        ("4. KIT CHURRASCO FAMÍLIA", "R$ 169,90", "Serve 5 a 6 pessoas", 
         "• 1 Frango recheado + 700g Costela bovina + 4 Linguiças toscanas grelhadas\n• Maionese de batata grande (500g) + Farofa grande (400g)\n• 4 Pães de alho crocantes na brasa",
         "#7B241C")
    ]
    
    items_es = [
        ("1. EL CLÁSICO DE SOFIA", "R$ 69,90", "Rinde 3 a 4 personas", 
         "• 1 Pollo entero relleno con farofa especial (~1,4kg asado dorado)\n• Farofa artesanal crocante de la casa (250g)\n• Mayonesa casera tradicional de patatas (300g)",
         "#C0392B"),
        ("2. COSTILLA SUPREMA AL VAPOR", "R$ 119,90", "Rinde 4 personas", 
         "• 1kg de Costilla vacuna premium braseada lentamente por 6 horas\n• Mandioca suave a la manteca de botella (300g)\n• Vinagreta fresca artesanal y Farofa crocante de la casa (250g)",
         "#922B21"),
        ("3. DUETO SOFIA (Pollo + Cerdo)", "R$ 94,90", "Rinde 3 a 4 personas", 
         "• Medio pollo asado jugoso + 500g de Costilla de cerdo en finas hierbas\n• Patatas rústicas doradas al romero (300g)\n• Farofa artesanal especial de la casa (200g)",
         "#D35400"),
        ("4. KIT PARRILLERO FAMILIA", "R$ 169,90", "Rinde 5 a 6 personas", 
         "• 1 Pollo relleno + 700g Costilla vacuna + 4 Chorizos criollos a la brasa\n• Mayonesa de patata grande (500g) + Farofa grande (400g)\n• 4 Panes con ajo crocantes a la parrilla",
         "#7B241C")
    ]
    
    active_items = items_pt if lang == "pt" else items_es
    
    y_starts = [0.68, 0.50, 0.32, 0.14]
    
    for idx, (name, price, serves, desc, color) in enumerate(active_items):
        y = y_starts[idx]
        
        # Item container
        box = patches.FancyBboxPatch((0.05, y), 0.90, 0.16, boxstyle="round,pad=0.01", facecolor='#FDFEFE', edgecolor='#EAEDED', linewidth=1.5)
        ax.add_patch(box)
        
        # Price badge
        p_badge = patches.FancyBboxPatch((0.72, y + 0.08), 0.21, 0.065, boxstyle="round,pad=0.005", facecolor=color, edgecolor='none')
        ax.add_patch(p_badge)
        ax.text(0.825, y + 0.112, price, ha='center', va='center', fontsize=13, fontweight='bold', color='#FFFFFF')
        
        # Item Title & Servings
        ax.text(0.08, y + 0.125, name, ha='left', va='center', fontsize=12, fontweight='bold', color=color)
        ax.text(0.08, y + 0.095, f"({serves})", ha='left', va='center', fontsize=8.5, fontstyle='italic', color='#7F8C8D')
        
        # Description
        ax.text(0.08, y + 0.045, desc, ha='left', va='center', fontsize=8.5, color='#2C3E50', linespacing=1.3)

    # Footer Call To Action
    ftr = patches.Rectangle((0.03, 0.03), 0.94, 0.09, facecolor='#1F3864', edgecolor='none')
    ax.add_patch(ftr)
    
    cta_title = "COMO ENCOMENDAR VIA CRM SOFIA / WHATSAPP:" if lang == "pt" else "CÓMO RESERVAR VÍA CRM SOFIA / WHATSAPP:"
    cta_desc = "1. Pré-venda toda Sexta-feira • 2. Escolha sua janela de 15 min • 3. Retire sem fila ou receba quentinho em casa!" if lang == "pt" else "1. Preventa todos los Viernes • 2. Elija su franja de 15 min • 3. ¡Retire sin fila o reciba caliente en su casa!"
    address_line = "Rua Dep. Pinheiro Júnior, 1380 - Umbará, Curitiba - PR | WhatsApp: (41) 9XXXX-XXXX"
    
    ax.text(0.5, 0.09, cta_title, ha='center', va='center', fontsize=9, fontweight='bold', color='#D4AC0D')
    ax.text(0.5, 0.065, cta_desc, ha='center', va='center', fontsize=7.5, color='#FFFFFF')
    ax.text(0.5, 0.045, address_line, ha='center', va='center', fontsize=6.5, color='#BDC3C7')

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    
    out_file = CHART_DIR / f"cardapio_sofia_{lang}.png"
    plt.tight_layout()
    fig.savefig(out_file)
    plt.close()
    print(f"Generated: {out_file}")

# -------------------------------------------------------------
# 3. GENERATE EQUIPMENT & TOOLS CATALOG SHEET (6 COMPONENT PANEL)
# -------------------------------------------------------------
def generate_equipment_catalog(lang="pt"):
    fig, ax = plt.subplots(figsize=(12, 8.5), dpi=300)
    fig.patch.set_facecolor('#F8F9FA')
    ax.set_facecolor('#F8F9FA')
    ax.axis('off')
    
    title = "CATÁLOGO TÉCNICO DE EQUIPAMENTOS E MÁQUINAS ADQUIRIDOS" if lang == "pt" else "CATÁLOGO TÉCNICO DE MAQUINARIA Y EQUIPAMIENTO ADQUIRIDO"
    subtitle = "Inventário de Ativos Operacionais e Patrimônio Fixo da Casa de Assados Sofia (Investimento: R$ 13.700,00)" if lang == "pt" else "Inventario de Activos Operacionales y Patrimonio Fijo de Casa de Assados Sofia (Inversión: R$ 13.700,00)"
    
    # Header
    hdr = patches.FancyBboxPatch((0.02, 0.88), 0.96, 0.10, boxstyle="round,pad=0.01", facecolor='#1F3864', edgecolor='none')
    ax.add_patch(hdr)
    ax.text(0.5, 0.94, title, ha='center', va='center', fontsize=14, fontweight='bold', color='#D4AC0D')
    ax.text(0.5, 0.90, subtitle, ha='center', va='center', fontsize=8.5, color='#FFFFFF')

    equip_pt = [
        ("1. Assadores Giratórios a Gás", "2 unidades (Capacidade: 10 frangos cada)\n• Estrutura em inox escovado AISI 430\n• Queimadores infravermelhos independentes\n• Vidros temperados e gaveta coletora de gordura\n• Custo: R$ 3.600,00 (Seminovos revisados)", "#C0392B"),
        ("2. Churrasqueira a Carvão Bafo 1,5m", "1 unidade em Aço Carbono Reforçado\n• Sistema elevatório de grelha argentina\n• Tampa articulada com termômetro analógico\n• Isolamento térmico p/ cocção lenta (6h)\n• Custo: R$ 1.200,00 (Fabricação sob medida)", "#922B21"),
        ("3. Freezer Horizontal Dupla Ação 500L", "1 unidade (Congelamento e Conservação)\n• Gabinete interno com drenagem facilitada\n• Faixa de temperatura: -18°C a +4°C\n• Baixo consumo com gás ecológico R600a\n• Custo: R$ 2.200,00 (Revisado com garantia)", "#2980B9"),
        ("4. Refrigerador Comercial Inox 4 Portas", "1 unidade Vertical em Aço Inox AISI 304\n• Capacidade volumétrica de 900 Litros\n• Refrigeração por ar forçado (+2°C a +4°C)\n• Conservação asséptica de marinadas e molhos\n• Custo: R$ 2.100,00", "#16A085"),
        ("5. Mesa Central Inox + Balança Digital", "1 Mesa AISI 304 (2,0m x 0,9m) + Balança\n• Tampo reforçado c/ paneleiro inferior\n• Balança eletrônica computadora 30kg (Inmetro)\n• Pés niveladores antivibratórios\n• Custo: R$ 1.200,00", "#D35400"),
        ("6. Terminal CRM/KDS + Caixas Térmicas", "Estação de Expedição e Delivery\n• Computador terminal c/ impressora térmica 80mm\n• 2 Caixas térmicas rígidas profissionais 45L\n• Hidrolavadora de pressão e utensílios inox\n• Custo: R$ 3.400,00", "#8E44AD")
    ]

    equip_es = [
        ("1. Asadores Giratorios a Gas", "2 unidades (Capacidad: 10 pollos c/u)\n• Estructura en acero inox cepillado\n• Quemadores infrarrojos independientes\n• Vidrios templados y bandeja recolectora de grasa\n• Costo: R$ 3.600,00 (Seminuevos revisados)", "#C0392B"),
        ("2. Parrilla a Carbón c/ Tapa al Vapor 1,5m", "1 unidad en Acero Carbono Reforzado\n• Sistema elevador con parrilla argentina\n• Tapa articulada con termómetro analógico\n• Aislamiento térmico p/ cocción lenta (6h)\n• Costo: R$ 1.200,00 (Fabricación a medida)", "#922B21"),
        ("3. Congelador Horizontal 500L Doble Función", "1 unidad (Congelamiento y Conservación)\n• Gabinete interior con desagüe de limpieza\n• Rango de temperatura: -18°C a +4°C\n• Eficiencia energética con gas R600a\n• Costo: R$ 2.200,00 (Revisado c/ garantía)", "#2980B9"),
        ("4. Refrigerador Comercial Inox 4 Puertas", "1 unidad Vertical en Acero Inoxidable\n• Capacidad volumétrica de 900 Litros\n• Frío forzado ventilado (+2°C a +4°C)\n• Conservación higiénica de marinadas y aderezos\n• Costo: R$ 2.100,00", "#16A085"),
        ("5. Mesada Central Inox + Balanza Digital", "1 Mesada AISI 304 (2,0m x 0,9m) + Balanza\n• Mesada reforzada con estante inferior\n• Balanza digital computadora 30kg (Inmetro)\n• Patas niveladoras antivibración\n• Costo: R$ 1.200,00", "#D35400"),
        ("6. Terminal CRM/KDS + Cajas Térmicas", "Estación de Despacho y Reparto\n• PC terminal con impresora térmica de comandas\n• 2 Cajas térmicas rígidas profesionales 45L\n• Hidrolavadora de presión y vajilla gastronorm\n• Costo: R$ 3.400,00", "#8E44AD")
    ]

    active_equip = equip_pt if lang == "pt" else equip_es

    for idx, (eq_title, eq_desc, eq_col) in enumerate(active_equip):
        col = idx % 2
        row = idx // 2
        
        x = 0.03 + col * 0.49
        y = 0.60 - row * 0.27
        w = 0.46
        h = 0.25
        
        # Equipment card box
        card = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.01", facecolor='#FFFFFF', edgecolor='#BDC3C7', linewidth=1.2)
        ax.add_patch(card)
        
        # Title bar
        t_bar = patches.FancyBboxPatch((x + 0.01, y + h - 0.055), w - 0.02, 0.045, boxstyle="round,pad=0.005", facecolor=eq_col, edgecolor='none')
        ax.add_patch(t_bar)
        ax.text(x + w/2, y + h - 0.032, eq_title, ha='center', va='center', fontsize=9.5, fontweight='bold', color='#FFFFFF')
        
        # Text details
        ax.text(x + 0.02, y + h/2 - 0.02, eq_desc, ha='left', va='center', fontsize=7.8, color='#2C3E50', linespacing=1.25)

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    
    out_file = CHART_DIR / f"equipamentos_sofia_{lang}.png"
    plt.tight_layout()
    fig.savefig(out_file)
    plt.close()
    print(f"Generated: {out_file}")

if __name__ == "__main__":
    generate_brand_identity("pt")
    generate_brand_identity("es")
    generate_menu_card("pt")
    generate_menu_card("es")
    generate_equipment_catalog("pt")
    generate_equipment_catalog("es")

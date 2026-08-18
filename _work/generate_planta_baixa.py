import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, ArrowStyle
from pathlib import Path

ROOT = Path(r"c:\Users\wilki\OneDrive\Documentos\Trabajo de Curso")
CHART_DIR = ROOT / "_work" / "charts"
IMG_REF_DIR = ROOT / "Casa_de_Assados_Sofia_15_Imagens_Referencia_PR_HR_v2"

def draw_planta(lang="pt"):
    # Configurar figura com alta resolução (300 DPI, 14 x 9.5 polegadas = 4200 x 2850 px)
    fig, ax = plt.subplots(figsize=(14, 9.2), dpi=300)
    ax.set_xlim(-1.2, 11.5)
    ax.set_ylim(-1.0, 7.2)
    ax.set_aspect('equal')
    ax.axis('off')

    # Paleta de Cores Arquitetônica Sofisticada
    C_WALL = '#1F3864'         # Azul Marinho Escuro para Alvenaria Externa
    C_INNER_WALL = '#5D6D7E'   # Divisórias Internas
    C_ZONE1 = '#EBF5FB'        # Recepção (Azul Claro)
    C_ZONE2 = '#E8F8F5'        # Armazenamento (Verde Água)
    C_ZONE3 = '#FEF9E7'        # Pré-preparo (Amarelo Claro)
    C_ZONE4 = '#FDEDEC'        # Cocção Isolada (Vermelho Claro / Quente)
    C_ZONE5 = '#F5EEF8'        # Montagem (Lilás Claro)
    C_ZONE6 = '#EAECEE'        # Expedição (Cinza Claro)
    C_ZONE7 = '#F2F4F4'        # Higienização (Neutro)
    
    C_FLOW = '#C0392B'         # Vermelho para Setas de Fluxo Sanitário
    C_EQUIP = '#2C3E50'        # Azul Escuro para Contorno de Equipamentos
    C_EQUIP_FILL = '#D5D8DC'   # Preenchimento Inox dos Equipamentos

    # 1. TÍTULO SUPERIOR & BANNER TÉCNICO
    title_main = "PLANTA BAIXA TÉCNICA E FLUXO SANITÁRIO UNIDIRECIONAL (60,0 m²)" if lang == "pt" else "PLANO ARQUITECTÓNICO Y FLUJO SANITARIO UNIDIRECCIONAL (60,0 m²)"
    sub_title = "Casa de Assados Sofia • Bairro Umbará, Curitiba - PR • Dimensões: 10,00 m × 6,00 m • Conforme RDC 216 Anvisa" if lang == "pt" else "Casa de Assados Sofia • Bairro Umbará, Curitiba - PR • Dimensiones: 10,00 m × 6,00 m • Conforme RDC 216 Anvisa"
    
    ax.text(5.0, 6.85, title_main, ha='center', va='center', fontsize=15, fontweight='bold', color='#1F3864', family='sans-serif')
    ax.text(5.0, 6.50, sub_title, ha='center', va='center', fontsize=10, fontweight='normal', color='#566573', family='sans-serif')

    # 2. PERÍMETRO EXTERNO DO IMÓVEL (10.0m x 6.0m)
    # Paredes externas com espessura
    outer_wall = patches.Rectangle((0, 0), 10.0, 6.0, linewidth=4.0, edgecolor=C_WALL, facecolor='#FAFAFA', zorder=1)
    ax.add_patch(outer_wall)

    # Cotas dimensionais externas
    # Cota Superior (10,00 m)
    ax.annotate('', xy=(0, 6.2), xytext=(10, 6.2), arrowprops=dict(arrowstyle='<->', color='#2C3E50', lw=1.2))
    ax.text(5.0, 6.32, "10,00 m", ha='center', va='bottom', fontsize=10, fontweight='bold', color='#2C3E50')
    # Cota Lateral Esquerda (6,00 m)
    ax.annotate('', xy=(-0.3, 0), xytext=(-0.3, 6.0), arrowprops=dict(arrowstyle='<->', color='#2C3E50', lw=1.2))
    ax.text(-0.45, 3.0, "6,00 m", ha='right', va='center', fontsize=10, fontweight='bold', color='#2C3E50', rotation=90)

    # 3. SETORES E ZONAS FUNCIONAIS (7 ZONAS)
    
    # ZONA 1: Recepção e Inspeção (0 a 3.0m, 3.2 a 6.0m)
    z1 = patches.Rectangle((0.05, 3.2), 2.9, 2.75, facecolor=C_ZONE1, edgecolor='#AED6F1', lw=1, zorder=2)
    ax.add_patch(z1)
    ax.text(1.5, 5.7, "ZONA 1: RECEPÇÃO & INSPEÇÃO" if lang == "pt" else "ZONA 1: RECEPCIÓN & INSPECCIÓN", ha='center', va='top', fontsize=9, fontweight='bold', color='#1B4F72')

    # ZONA 2: Armazenamento & Refrigeração (0 a 3.0m, 0.05 a 3.15m)
    z2 = patches.Rectangle((0.05, 0.05), 2.9, 3.1, facecolor=C_ZONE2, edgecolor='#A2D9CE', lw=1, zorder=2)
    ax.add_patch(z2)
    ax.text(1.5, 2.95, "ZONA 2: ARMAZENAMENTO" if lang == "pt" else "ZONA 2: ALMACENAMIENTO", ha='center', va='top', fontsize=9, fontweight='bold', color='#0E6251')

    # ZONA 3: Pré-Preparo e Manipulação (3.0 a 6.0m, 3.2 a 6.0m)
    z3 = patches.Rectangle((3.05, 3.2), 2.9, 2.75, facecolor=C_ZONE3, edgecolor='#F9E79F', lw=1, zorder=2)
    ax.add_patch(z3)
    ax.text(4.5, 5.7, "ZONA 3: PRÉ-PREPARO" if lang == "pt" else "ZONA 3: PREPARACIÓN PREVIA", ha='center', va='top', fontsize=9, fontweight='bold', color='#7D6608')

    # ZONA 7: Higienização e DML (3.0 a 6.0m, 0.05 a 3.15m)
    z7 = patches.Rectangle((3.05, 0.05), 2.9, 3.1, facecolor=C_ZONE7, edgecolor='#D5D8DC', lw=1, zorder=2)
    ax.add_patch(z7)
    ax.text(4.5, 2.95, "ZONA 7: HIGIENIZAÇÃO & DML" if lang == "pt" else "ZONA 7: HIGIENIZACIÓN & DML", ha='center', va='top', fontsize=9, fontweight='bold', color='#4A235A')

    # ZONA 4: Área de Cocção Isolada (6.0 a 9.95m, 2.2 a 6.0m)
    z4 = patches.Rectangle((6.05, 2.2), 3.9, 3.75, facecolor=C_ZONE4, edgecolor='#F5B7B1', lw=1, zorder=2)
    ax.add_patch(z4)
    ax.text(8.0, 5.7, "ZONA 4: ÁREA DE COCÇÃO (ISOLADA)" if lang == "pt" else "ZONA 4: ÁREA DE COCCIÓN (AISLADA)", ha='center', va='top', fontsize=9, fontweight='bold', color='#78281F')

    # ZONA 5: Montagem e Embalagem (6.0 a 7.95m, 0.05 a 2.15m)
    z5 = patches.Rectangle((6.05, 0.05), 1.9, 2.1, facecolor=C_ZONE5, edgecolor='#D2B4DE', lw=1, zorder=2)
    ax.add_patch(z5)
    ax.text(7.0, 1.95, "ZONA 5: MONTAGEM" if lang == "pt" else "ZONA 5: MONTAJE", ha='center', va='top', fontsize=8.5, fontweight='bold', color='#512E5F')

    # ZONA 6: Expedição e Atendimento Balcão (8.0 a 9.95m, 0.05 a 2.15m)
    z6 = patches.Rectangle((8.0, 0.05), 1.95, 2.1, facecolor=C_ZONE6, edgecolor='#BDC3C7', lw=1, zorder=2)
    ax.add_patch(z6)
    ax.text(9.0, 1.95, "ZONA 6: EXPEDIÇÃO" if lang == "pt" else "ZONA 6: DESPACHO", ha='center', va='top', fontsize=8.5, fontweight='bold', color='#1A5276')

    # 4. PAREDES INTERNAS E DIVISÓRIAS (Alvenaria e Vidro Sanitário)
    # Divisória Vertical X=3.0 (Zona 1/2 vs 3/7)
    ax.plot([3.0, 3.0], [0, 6.0], color=C_INNER_WALL, lw=2.5, zorder=3)
    # Divisória Horizontal Y=3.2 (Zona 1 vs 2 e Zona 3 vs 7)
    ax.plot([0, 6.0], [3.2, 3.2], color=C_INNER_WALL, lw=2.5, zorder=3)
    # Divisória Vertical X=6.0 (Zona 3/7 vs 4/5)
    ax.plot([6.0, 6.0], [0, 6.0], color=C_INNER_WALL, lw=3.0, zorder=3)
    # Divisória Horizontal Y=2.2 (Zona 4 vs 5/6)
    ax.plot([6.0, 10.0], [2.2, 2.2], color=C_INNER_WALL, lw=2.5, zorder=3)
    # Divisória Vertical X=8.0 (Zona 5 vs 6)
    ax.plot([8.0, 8.0], [0, 2.2], color=C_INNER_WALL, lw=2.0, zorder=3)

    # 5. PORTAS E ACESSOS
    # Porta Entrada Insumos (Lateral Esquerda, Y=4.5 a 5.5)
    ax.plot([0, 0], [4.4, 5.4], color='white', lw=5.0, zorder=4)
    ax.text(-0.1, 4.9, "ENTRADA DE INSUMOS ➔" if lang == "pt" else "ENTRADA INSUMOS ➔", ha='right', va='center', fontsize=8.5, fontweight='bold', color='#1F3864')
    
    # Porta Saída Balcão / Takeaway (Inferior Direita, X=8.5 a 9.5)
    ax.plot([8.4, 9.6], [0, 0], color='white', lw=5.0, zorder=4)
    ax.text(9.0, -0.25, "SAÍDA / RETIRADA / DELIVERY ⬇" if lang == "pt" else "SALIDA / RETIRO / DELIVERY ⬇", ha='center', va='top', fontsize=8.5, fontweight='bold', color='#C0392B')

    # Passagens Internas
    ax.plot([3.0, 3.0], [4.2, 5.0], color='white', lw=4.0, zorder=4) # Porta Z1 -> Z3
    ax.plot([6.0, 6.0], [4.2, 5.0], color='white', lw=4.0, zorder=4) # Porta Z3 -> Z4
    ax.plot([7.0, 7.8], [2.2, 2.2], color='white', lw=4.0, zorder=4) # Passa-Pratos Z4 -> Z5
    ax.plot([8.0, 8.0], [0.8, 1.6], color='white', lw=4.0, zorder=4) # Passagem Z5 -> Z6

    # 6. EQUIPAMENTOS DETALHADOS EM ESCALA REAL

    # ZONA 1: Mesa de Inspeção
    r_insp = FancyBboxPatch((0.4, 3.6), 1.2, 0.7, boxstyle="round,pad=0.03", fc=C_EQUIP_FILL, ec=C_EQUIP, lw=1.2, zorder=5)
    ax.add_patch(r_insp)
    ax.text(1.0, 3.95, "Mesa Inspeção\n(Balança Inox)" if lang == "pt" else "Mesa Inspección\n(Balanza Inox)", ha='center', va='center', fontsize=7, color=C_EQUIP, fontweight='bold')

    # ZONA 2: Freezer Horizontal 510L e Refrigerador Vertical Inox 4P
    r_frz = FancyBboxPatch((0.3, 0.4), 1.3, 0.9, boxstyle="round,pad=0.03", fc='#D4E6F1', ec=C_EQUIP, lw=1.2, zorder=5)
    ax.add_patch(r_frz)
    ax.text(0.95, 0.85, "Freezer 510L\n(-18°C)" if lang == "pt" else "Congelador 510L\n(-18°C)", ha='center', va='center', fontsize=7.5, color='#154360', fontweight='bold')

    r_ref = FancyBboxPatch((1.7, 0.4), 1.0, 1.0, boxstyle="round,pad=0.03", fc='#D4E6F1', ec=C_EQUIP, lw=1.2, zorder=5)
    ax.add_patch(r_ref)
    ax.text(2.2, 0.9, "Refrigerador 4P\n(+2°C a +4°C)" if lang == "pt" else "Refrigerador 4P\n(+2°C a +4°C)", ha='center', va='center', fontsize=7, color='#154360', fontweight='bold')

    # ZONA 3: Mesas de Pré-Preparo e Manipulação (AISI 304)
    r_m1 = FancyBboxPatch((3.4, 3.6), 2.1, 0.85, boxstyle="round,pad=0.03", fc=C_EQUIP_FILL, ec=C_EQUIP, lw=1.2, zorder=5)
    ax.add_patch(r_m1)
    ax.text(4.45, 4.02, "Bancada Central Inox AISI 304 (2,0x0,9m)\n[Cubas GN + Tábua Sanitária + Balança]" if lang == "pt" else "Mesada Central Inox AISI 304 (2,0x0,9m)\n[Cubas GN + Tabla Sanitaria + Balanza]", ha='center', va='center', fontsize=7, color=C_EQUIP, fontweight='bold')

    r_pia1 = FancyBboxPatch((3.4, 4.8), 0.8, 0.5, boxstyle="round,pad=0.02", fc='#AED6F1', ec=C_EQUIP, lw=1.0, zorder=5)
    ax.add_patch(r_pia1)
    ax.text(3.8, 5.05, "Pia Mãos" if lang == "pt" else "Lavamanos", ha='center', va='center', fontsize=6.5, color='#154360')

    # ZONA 7: Pias Industriais e DML
    r_pia2 = FancyBboxPatch((3.3, 0.4), 1.4, 0.8, boxstyle="round,pad=0.03", fc='#AED6F1', ec=C_EQUIP, lw=1.2, zorder=5)
    ax.add_patch(r_pia2)
    ax.text(4.0, 0.8, "Pia Dupla Industrial\n(Lavagem Utensílios)" if lang == "pt" else "Pileta Doble Inox\n(Lavado Utensilios)", ha='center', va='center', fontsize=7, color='#154360', fontweight='bold')

    r_dml = FancyBboxPatch((4.9, 0.4), 0.9, 1.0, boxstyle="round,pad=0.03", fc='#E5E7E9', ec=C_EQUIP, lw=1.0, zorder=5)
    ax.add_patch(r_dml)
    ax.text(5.35, 0.9, "Armário\nDML" if lang == "pt" else "Armario\nDML", ha='center', va='center', fontsize=7, color='#2C3E50')

    # ZONA 4: Área de Cocção Isolada (Assadoras, Churrasqueira e Coifa)
    # Coifa Industrial Superior
    r_coifa = patches.Rectangle((6.3, 2.5), 3.4, 2.8, facecolor='none', edgecolor='#C0392B', lw=2.0, ls='--', zorder=6)
    ax.add_patch(r_coifa)
    ax.text(8.0, 5.05, "SISTEMA DE COIFA INDUSTRIAL INOX c/ EXAUSTÃO MECÂNICA (VISA)" if lang == "pt" else "SISTEMA DE CAMPANA INDUSTRIAL INOX c/ EXTRACCIÓN MECÁNICA (VISA)", ha='center', va='center', fontsize=7.5, fontweight='bold', color='#922B21')

    # 2 Assadoras Giratórias a Gás GLP
    r_as1 = FancyBboxPatch((6.4, 3.7), 1.0, 0.9, boxstyle="round,pad=0.03", fc='#FADBD8', ec='#922B21', lw=1.3, zorder=7)
    ax.add_patch(r_as1)
    ax.text(6.9, 4.15, "Assadora 1\nGás GLP\n(20 Frangos)" if lang == "pt" else "Asadora 1\nGas GLP\n(20 Pollos)", ha='center', va='center', fontsize=7, color='#78281F', fontweight='bold')

    r_as2 = FancyBboxPatch((6.4, 2.6), 1.0, 0.9, boxstyle="round,pad=0.03", fc='#FADBD8', ec='#922B21', lw=1.3, zorder=7)
    ax.add_patch(r_as2)
    ax.text(6.9, 3.05, "Assadora 2\nGás GLP\n(20 Frangos)" if lang == "pt" else "Asadora 2\nGas GLP\n(20 Pollos)", ha='center', va='center', fontsize=7, color='#78281F', fontweight='bold')

    # Churrasqueira a Carvão Bafo
    r_churr = FancyBboxPatch((7.9, 2.6), 1.7, 2.0, boxstyle="round,pad=0.03", fc='#EDBB99', ec='#6E2C00', lw=1.3, zorder=7)
    ax.add_patch(r_churr)
    ax.text(8.75, 3.6, "Churrasqueira a Carvão\nCostela no Bafo (1,5m)\n[Grelha Elevatória + Tampa]" if lang == "pt" else "Parrilla a Carbón\nCostilla al Vapor (1,5m)\n[Manivela Elevador + Tapa]", ha='center', va='center', fontsize=7.5, color='#4E2200', fontweight='bold')

    # ZONA 5: Mesa de Montagem
    r_m2 = FancyBboxPatch((6.2, 0.4), 1.6, 0.8, boxstyle="round,pad=0.03", fc=C_EQUIP_FILL, ec=C_EQUIP, lw=1.2, zorder=5)
    ax.add_patch(r_m2)
    ax.text(7.0, 0.8, "Mesa Montagem Inox\n+ Estufa Térmica\n(Lacração Sacolas)" if lang == "pt" else "Mesada Montaje Inox\n+ Estufa Térmica\n(Sellado de Bolsas)", ha='center', va='center', fontsize=6.8, color=C_EQUIP, fontweight='bold')

    # ZONA 6: Balcão de Atendimento e Terminal CRM
    r_balcao = FancyBboxPatch((8.2, 0.4), 1.5, 0.9, boxstyle="round,pad=0.03", fc='#D6EAF8', ec='#1B4F72', lw=1.3, zorder=5)
    ax.add_patch(r_balcao)
    ax.text(8.95, 0.85, "Balcão Atendimento\n+ Terminal CRM / Touch\n[Retirada < 90s]" if lang == "pt" else "Mostrador Atención\n+ Terminal CRM / Touch\n[Retiro < 90s]", ha='center', va='center', fontsize=7, color='#1B4F72', fontweight='bold')

    # 7. SETAS DE FLUXO SANITÁRIO UNIDIRECIONAL (RDC 216)
    # Seta 1: Entrada -> Pré-Preparo
    ax.annotate('', xy=(3.3, 4.6), xytext=(1.8, 4.6), arrowprops=dict(arrowstyle='simple,head_width=0.6,head_length=0.7', color=C_FLOW, lw=0.5), zorder=8)
    # Seta 2: Armazenamento -> Pré-Preparo
    ax.annotate('', xy=(3.5, 3.8), xytext=(2.6, 2.2), arrowprops=dict(arrowstyle='simple,head_width=0.5,head_length=0.6', color=C_FLOW, lw=0.5), zorder=8)
    # Seta 3: Pré-Preparo -> Cocção
    ax.annotate('', xy=(6.3, 4.6), xytext=(5.6, 4.6), arrowprops=dict(arrowstyle='simple,head_width=0.6,head_length=0.7', color=C_FLOW, lw=0.5), zorder=8)
    # Seta 4: Cocção -> Montagem (Passa-pratos)
    ax.annotate('', xy=(7.0, 1.3), xytext=(7.0, 2.4), arrowprops=dict(arrowstyle='simple,head_width=0.6,head_length=0.7', color=C_FLOW, lw=0.5), zorder=8)
    # Seta 5: Montagem -> Expedição
    ax.annotate('', xy=(8.2, 1.0), xytext=(7.8, 1.0), arrowprops=dict(arrowstyle='simple,head_width=0.6,head_length=0.7', color=C_FLOW, lw=0.5), zorder=8)

    # 8. LEGENDA INFERIOR E QUADRO DE CONFORMIDADE
    leg_box = FancyBboxPatch((0.2, -0.85), 9.6, 0.65, boxstyle="round,pad=0.03", fc='#FFFFFF', ec='#B0BEC5', lw=1.0, zorder=5)
    ax.add_patch(leg_box)
    
    ax.text(0.5, -0.4, "LEGENDA DE FLUXO & NORMAS:" if lang == "pt" else "LEYENDA DE FLUJO & NORMAS:", fontsize=8.5, fontweight='bold', color='#1F3864')
    ax.annotate('', xy=(3.1, -0.4), xytext=(2.6, -0.4), arrowprops=dict(arrowstyle='simple,head_width=0.4,head_length=0.5', color=C_FLOW))
    ax.text(3.3, -0.4, "Fluxo Unidirecional Obrigatório (Sem Cruzamento Cru ➔ Pronto)" if lang == "pt" else "Flujo Unidireccional Obligatorio (Sin Cruce Crudo ➔ Listo)", fontsize=8, color='#2C3E50')
    ax.text(7.7, -0.4, "✔ VISA Curitiba / RDC 216 Anvisa", fontsize=8, fontweight='bold', color='#27AE60')

    # Rosa dos Ventos / Indicador Norte
    ax.text(10.5, 5.5, "N", ha='center', va='bottom', fontsize=12, fontweight='bold', color='#1F3864')
    ax.annotate('', xy=(10.5, 5.5), xytext=(10.5, 4.6), arrowprops=dict(arrowstyle='->', color='#1F3864', lw=2.0))

    plt.tight_layout()
    
    # Salvar em múltiplos destinos com alta resolução
    out_pt_hd = CHART_DIR / f"planta_baixa_sofia_{lang}_hd.png"
    plt.savefig(str(out_pt_hd), dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f"Generated HD floor plan: {out_pt_hd}")

if __name__ == "__main__":
    draw_planta("pt")
    draw_planta("es")

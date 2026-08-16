import matplotlib.pyplot as plt
import numpy as np
from pathlib import Path

CHART_DIR = Path(r"c:\Users\wilki\OneDrive\Documentos\Trabajo de Curso\_work\charts")
CHART_DIR.mkdir(parents=True, exist_ok=True)

plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['figure.dpi'] = 300

colors = {
    'primary': '#1F3864',
    'secondary': '#C0392B',
    'accent': '#D4AC0D',
    'dark': '#2C3E50',
    'light': '#ECF0F1',
    'success': '#27AE60',
    'gray': '#7F8C8D',
    'blue_light': '#4A90E2',
    'orange': '#E67E22'
}

# 1. Mix de Vendas
combos_pt = ['O Clássico da Sofia\n(Frango Recheado)', 'Costela Suprema\nno Bafo', 'Dueto Sofia\n(Frango+Costelinha)', 'Kit Churrasco\nFamília']
combos_es = ['El Clásico de Sofia\n(Pollo Relleno)', 'Costilla Suprema\nal Vapor', 'Dueto Sofia\n(Pollo+Costilla)', 'Kit Parrillero\nFamilia']
vols = [70, 35, 35, 20]
prices = [69.90, 119.90, 94.90, 169.90]
revs = [q * p for q, p in zip(vols, prices)]
palette = ['#C0392B', '#D4AC0D', '#E67E22', '#1F3864']

for lang, labels, fname in [('pt', combos_pt, 'mix_pt.png'), ('es', combos_es, 'mix_es.png')]:
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5.5))
    
    wedges1, texts1, autotexts1 = ax1.pie(vols, labels=labels, autopct='%1.1f%%', startangle=140, 
                                           colors=palette, textprops={'fontsize': 8.5}, wedgeprops=dict(width=0.6, edgecolor='w'))
    for at in autotexts1:
        at.set_fontsize(9)
        at.set_fontweight('bold')
    ax1.set_title('Mix por Quantidade (160 combos/mês)' if lang=='pt' else 'Mezcla por Cantidad (160 combos/mes)', 
                  fontsize=11, fontweight='bold', color=colors['primary'], pad=12)
    
    wedges2, texts2, autotexts2 = ax2.pie(revs, labels=labels, autopct='%1.1f%%', startangle=140, 
                                           colors=palette, textprops={'fontsize': 8.5}, wedgeprops=dict(width=0.6, edgecolor='w'))
    for at in autotexts2:
        at.set_fontsize(9)
        at.set_fontweight('bold')
    ax2.set_title('Mix por Receita Bruta (R$ 15.809,00)' if lang=='pt' else 'Mezcla por Ingresos Brutos (R$ 15.809,00)', 
                  fontsize=11, fontweight='bold', color=colors['primary'], pad=12)
    
    plt.tight_layout()
    plt.savefig(CHART_DIR / fname, bbox_inches='tight')
    plt.close()

# 2. DRE Waterfall
revenue = 15809.00
cmv = 6140.00
tax = revenue * 0.04
fees = revenue * 0.02
fixed = 6870.00
profit = revenue - cmv - tax - fees - fixed

for lang, fname in [('pt', 'dre_pt.png'), ('es', 'dre_es.png')]:
    fig, ax = plt.subplots(figsize=(10, 5))
    cats_pt = ['Receita\nBruta', '(-) CMV\nInsumos', '(-) Simples\n(4%)', '(-) Meios\nPgto (2%)', 'Margem de\nContribuição', '(-) Custos\nFixos', 'Lucro Líquido\nOperacional']
    cats_es = ['Ingresos\nBrutos', '(-) CMV\nInsumos', '(-) Simples\n(4%)', '(-) Medios\nPago (2%)', 'Margen de\nContribución', '(-) Costos\nFijos', 'Utilidad Neta\nOperativa']
    cats = cats_pt if lang=='pt' else cats_es
    
    vals = [revenue, -cmv, -tax, -fees, revenue-cmv-tax-fees, -fixed, profit]
    pos = range(len(cats))
    
    bar_colors = [colors['primary'], colors['secondary'], '#E74C3C', '#E67E22', colors['accent'], colors['dark'], colors['success']]
    bars = ax.bar(pos, [abs(v) for v in vals], color=bar_colors, width=0.55, edgecolor='black', linewidth=0.5)
    
    for bar, val in zip(bars, vals):
        y = bar.get_height()
        sign = "+" if val > 0 and bar != bars[0] and bar != bars[4] and bar != bars[6] else ""
        if val < 0:
            sign = "-"
        txt = f"R$ {abs(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        pct = f"({(abs(val)/revenue)*100:.1f}%)"
        ax.text(bar.get_x() + bar.get_width()/2., y + 250, f"{sign}{txt}\n{pct}", 
                ha='center', va='bottom', fontsize=8, fontweight='bold')
        
    ax.set_ylim(0, 18500)
    ax.set_ylabel('Valor em Reais (R$)' if lang=='pt' else 'Monto en Reales (R$)', fontsize=10, fontweight='bold')
    ax.set_xticks(pos)
    ax.set_xticklabels(cats, fontsize=8.5)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    ax.set_title('Composição do Resultado Mensal - Cenário-Base (160 Combos)' if lang=='pt' else 'Composición del Resultado Mensual - Escenario Base (160 Combos)', 
                 fontsize=12, fontweight='bold', color=colors['primary'], pad=15)
    plt.tight_layout()
    plt.savefig(CHART_DIR / fname, bbox_inches='tight')
    plt.close()

# 3. Ponto de Equilíbrio
cm_ratio = (revenue - cmv - tax - fees) / revenue
breakeven_rs = fixed / cm_ratio
breakeven_q = breakeven_rs / (revenue / 160.0)

for lang, fname in [('pt', 'breakeven_pt.png'), ('es', 'breakeven_es.png')]:
    q_arr = np.linspace(0, 220, 200)
    price_avg = revenue / 160.0
    var_cost_avg = (cmv + tax + fees) / 160.0
    
    rt = q_arr * price_avg
    ct = fixed + q_arr * var_cost_avg
    cf = np.full_like(q_arr, fixed)
    
    fig, ax = plt.subplots(figsize=(10, 5.5))
    ax.plot(q_arr, rt, label='Receita Total (RT)' if lang=='pt' else 'Ingreso Total (IT)', color=colors['primary'], linewidth=2.5)
    ax.plot(q_arr, ct, label='Custo Total (CT)' if lang=='pt' else 'Costo Total (CT)', color=colors['secondary'], linewidth=2.5)
    ax.plot(q_arr, cf, label='Custo Fixo (CF)' if lang=='pt' else 'Costo Fijo (CF)', color=colors['gray'], linestyle='--', linewidth=1.5)
    
    ax.fill_between(q_arr, ct, rt, where=(rt >= ct), color=colors['success'], alpha=0.15, label='Zona de Lucro' if lang=='pt' else 'Zona de Utilidad')
    ax.fill_between(q_arr, ct, rt, where=(rt < ct), color=colors['secondary'], alpha=0.15, label='Zona de Prejuízo' if lang=='pt' else 'Zona de Pérdida')
    
    ax.scatter([breakeven_q], [breakeven_rs], color='black', s=80, zorder=5)
    
    lbl_pt = f"Ponto de Equilíbrio\n{breakeven_q:.0f} combos/mês\n(R$ {breakeven_rs:,.2f})".replace(',', 'X').replace('.', ',').replace('X', '.')
    lbl_es = f"Punto de Equilibrio\n{breakeven_q:.0f} combos/mes\n(R$ {breakeven_rs:,.2f})".replace(',', 'X').replace('.', ',').replace('X', '.')
    
    ax.annotate(lbl_pt if lang=='pt' else lbl_es, 
                xy=(breakeven_q, breakeven_rs), xytext=(breakeven_q - 45, breakeven_rs + 4200),
                arrowprops=dict(facecolor='black', shrink=0.08, width=1, headwidth=6),
                fontsize=9, fontweight='bold', bbox=dict(boxstyle="round,pad=0.4", fc="#FFF9E6", ec=colors['accent'], lw=1.2))
    
    # Marcador do Cenário Base
    ax.scatter([160], [revenue], color=colors['success'], s=70, zorder=5)
    base_lbl = f"Cenário Base (160 combos)\nLucro: R$ {profit:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if lang=='pt' else f"Escenario Base (160 combos)\nUtilidad: R$ {profit:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    ax.annotate(base_lbl, xy=(160, revenue), xytext=(160 - 25, revenue - 5500),
                arrowprops=dict(facecolor=colors['success'], shrink=0.08, width=1, headwidth=6),
                fontsize=8.5, fontweight='bold', bbox=dict(boxstyle="round,pad=0.3", fc="#E8F8F5", ec=colors['success'], lw=1.2))
    
    ax.set_xlim(0, 220)
    ax.set_ylim(0, 23000)
    ax.set_xlabel('Quantidade de Combos por Mês' if lang=='pt' else 'Cantidad de Combos por Mes', fontsize=10, fontweight='bold')
    ax.set_ylabel('Valor em Reais (R$)' if lang=='pt' else 'Monto en Reales (R$)', fontsize=10, fontweight='bold')
    ax.set_title('Gráfico do Ponto de Equilíbrio Operacional' if lang=='pt' else 'Gráfico del Punto de Equilibrio Operativo', 
                 fontsize=12, fontweight='bold', color=colors['primary'], pad=15)
    ax.grid(True, linestyle='--', alpha=0.3)
    ax.legend(loc='upper left', fontsize=8.5)
    plt.tight_layout()
    plt.savefig(CHART_DIR / fname, bbox_inches='tight')
    plt.close()

# 4. Projeção de 12 Meses
months = [f"M{i}" for i in range(1, 13)]
combos_proj = [160, 168, 176, 184, 192, 200, 208, 216, 224, 232, 240, 248]
revenues_proj = [q * (revenue / 160.0) for q in combos_proj]
profits_proj = [r * cm_ratio - fixed for r in revenues_proj]

cum_cash = []
curr = -38000.0
for p_m in profits_proj:
    curr += p_m
    cum_cash.append(curr)

for lang, fname in [('pt', 'result12_pt.png'), ('es', 'result12_es.png')]:
    fig, ax1 = plt.subplots(figsize=(11, 5.5))
    
    x = np.arange(len(months))
    ax1.bar(x - 0.18, profits_proj, width=0.35, color=colors['primary'], label='Lucro Líquido Mensal' if lang=='pt' else 'Utilidad Neta Mensual', edgecolor='black', linewidth=0.4)
    ax1.set_ylabel('Lucro Mensal (R$)' if lang=='pt' else 'Utilidad Mensual (R$)', fontsize=10, fontweight='bold', color=colors['primary'])
    ax1.tick_params(axis='y', labelcolor=colors['primary'])
    ax1.set_ylim(0, 8000)
    
    for i, p_m in enumerate(profits_proj):
        ax1.text(i - 0.18, p_m + 150, f"R${p_m/1000:.1f}k", ha='center', va='bottom', fontsize=7.5, fontweight='bold')
    
    ax2 = ax1.twinx()
    ax2.plot(x + 0.18, cum_cash, color=colors['secondary'], marker='o', linewidth=2.2, label='Saldo Acumulado (Caixa)' if lang=='pt' else 'Saldo Acumulado (Caja)')
    ax2.axhline(0, color='black', linestyle=':', linewidth=1)
    ax2.set_ylabel('Saldo Acumulado (R$)' if lang=='pt' else 'Saldo Acumulado (R$)', fontsize=10, fontweight='bold', color=colors['secondary'])
    ax2.tick_params(axis='y', labelcolor=colors['secondary'])
    ax2.set_ylim(-42000, 25000)
    
    ax1.set_xticks(x)
    ax1.set_xticklabels([f"{m}\n({q}c)" for m, q in zip(months, combos_proj)], fontsize=8)
    ax1.set_xlabel('Mês de Operação (Volume em Combos)' if lang=='pt' else 'Mes de Operación (Volumen en Combos)', fontsize=10, fontweight='bold')
    
    # Anotação de Payback
    pb_idx = 10 # mês 11 / 12 cruza o zero
    ax2.annotate('Recuperação Total\n(Payback Mês 11-12)' if lang=='pt' else 'Recuperación Total\n(Payback Mes 11-12)',
                 xy=(10.18, cum_cash[10]), xytext=(7.5, cum_cash[10] + 12000),
                 arrowprops=dict(facecolor=colors['secondary'], shrink=0.08, width=1, headwidth=6),
                 fontsize=8.5, fontweight='bold', bbox=dict(boxstyle="round,pad=0.3", fc="#FDEDEC", ec=colors['secondary'], lw=1.2))
    
    ax1.set_title('Projeção Financeira e Retorno do Investimento Inicial (R$ 38.000,00)' if lang=='pt' else 'Proyección Financiera y Retorno de la Inversión Inicial (R$ 38.000,00)', 
                  fontsize=12, fontweight='bold', color=colors['primary'], pad=15)
    
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=8.5)
    
    plt.tight_layout()
    plt.savefig(CHART_DIR / fname, bbox_inches='tight')
    plt.close()

# 5. Análise de Sensibilidade
scenarios_labels_pt = ['Pessimista\n(-20% / 128 combos)', 'Cenário Base\n(160 combos)', 'Otimista\n(+30% / 208 combos)']
scenarios_labels_es = ['Pesimista\n(-20% / 128 combos)', 'Escenario Base\n(160 combos)', 'Optimista\n(+30% / 208 combos)']

scen_revs = [128 * (revenue/160.0), revenue, 208 * (revenue/160.0)]
scen_cm = [r * cm_ratio for r in scen_revs]
scen_profits = [cm - fixed for cm in scen_cm]

for lang, labels, fname in [('pt', scenarios_labels_pt, 'scenarios_pt.png'), ('es', scenarios_labels_es, 'scenarios_es.png')]:
    fig, ax = plt.subplots(figsize=(9, 5))
    x = np.arange(len(labels))
    width = 0.25
    
    b1 = ax.bar(x - width, scen_revs, width, label='Receita Bruta' if lang=='pt' else 'Ingresos Brutos', color=colors['primary'], edgecolor='black', linewidth=0.4)
    b2 = ax.bar(x, scen_cm, width, label='Margem de Contribuição' if lang=='pt' else 'Margen de Contribución', color=colors['accent'], edgecolor='black', linewidth=0.4)
    b3 = ax.bar(x + width, scen_profits, width, label='Lucro Líquido' if lang=='pt' else 'Utilidad Neta', color=colors['success'], edgecolor='black', linewidth=0.4)
    
    for bar in b1:
        y = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., y + 250, f"R${y/1000:.1f}k", ha='center', va='bottom', fontsize=7.5, fontweight='bold')
    for bar in b2:
        y = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., y + 250, f"R${y/1000:.1f}k", ha='center', va='bottom', fontsize=7.5, fontweight='bold')
    for bar in b3:
        y = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., y + 250, f"R${y:,.0f}".replace(',', '.'), ha='center', va='bottom', fontsize=7.5, fontweight='bold')
        
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=9, fontweight='bold')
    ax.set_ylabel('Valor em Reais (R$)' if lang=='pt' else 'Monto en Reales (R$)', fontsize=10, fontweight='bold')
    ax.set_ylim(0, 24000)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    ax.legend(loc='upper left', fontsize=9)
    ax.set_title('Análise de Sensibilidade e Comparação de Cenários de Demanda' if lang=='pt' else 'Análisis de Sensibilidad y Comparación de Escenarios de Demanda', 
                 fontsize=12, fontweight='bold', color=colors['primary'], pad=15)
    plt.tight_layout()
    plt.savefig(CHART_DIR / fname, bbox_inches='tight')
    plt.close()

print("Charts successfully regenerated for R$ 38.000,00 budget!")

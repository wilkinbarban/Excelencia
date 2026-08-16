# Casa de Assados Sofia — Plano de Negócio / Plan de Negocios

> **Trabalho de Conclusão de Curso (TCC)**  
> **Curso Técnico em Administração e Informática**  
> **Instituição:** [Colégio Excelência](https://colegioexcelencia.com) — Curitiba, Paraná  
> **Autor:** Wilkin Barban Rosabal  
> **Ano:** 2026  

---

## 📌 Visão Geral do Projeto / Project Overview

Este repositório contém a documentação completa, modelagem financeira, identidade visual e os geradores automatizados do Trabalho de Conclusão de Curso da **Casa de Assados Sofia Ltda.**, uma microempresa gastronômica projetada para operar aos finais de semana e feriados no bairro Umbará, em Curitiba - PR.

O diferencial do negócio consiste na integração de cortes tradicionais artesanais (frango recheado, costela ao bafo, costelinha suína e acompanhamentos) com um sistema proprietário de CRM (**CRM Sofia**) que gerencia pré-vendas semanais via WhatsApp, nivela a capacidade de produção em janelas de 15 minutos, elimina filas e mitiga perdas de insumos perecíveis.

---

## 📁 Estrutura do Repositório / Repository Structure

```text
├── Borrador_Casa_de_Assados_Sofia_Portugues.docx  # Dissertação completa em Português (ABNT)
├── Borrador_Casa_de_Assados_Sofia_Espanol.docx    # Disertación completa en Español (Normas ABNT)
├── Apresentacao_Casa_de_Assados_Sofia_Portugues.pptx # Apresentação de slides em Português (18 slides)
├── Presentacion_Casa_de_Assados_Sofia_Espanol.pptx   # Presentación de diapositivas en Español (18 slides)
├── README.md                                      # Documentação principal do repositório
├── .gitignore                                     # Arquivos e extensões ignorados pelo Git
│
├── doc/                                           # Documentação de referência e diretrizes
│   ├── Andrea Souza Mello Meirelles.pdf           # Referência acadêmica complementar
│   ├── Estrutura_Plano_de_Negocio_Passo_a_Passo.pptx # Guia estrutural de plano de negócio
│   ├── MANUAL-DE-PLANO-DE-NEGOCIO-.pdf            # Manual metodológico do SEBRAE
│   ├── Manual _plano_negocio.pdf                  # Manual de elaboração empresarial
│   ├── Normas ABNT/                               # Diretrizes de formatação ABNT
│   ├── Plano de Negocio Loja de Cosmeticos Seja Bella Cosmeticos.pdf
│   ├── Plano de Negocio Loja de Cosmeticos UP.pdf
│   └── Projeto de Oficinas Natalia e Divina (1).pptx
│
└── _work/                                         # Scripts de compilação e ativos gráficos
    ├── build_pt.py                                # Gerador do documento em Português
    ├── build_es.py                                # Gerador do documento em Espanhol
    ├── generate_presentations.py                  # Gerador das apresentações em PowerPoint (.pptx)
    ├── generate_charts.py                         # Gerador dos gráficos financeiros (Matplotlib)
    ├── run_all.py                                 # Script de execução e verificação integrada
    ├── anexo_casa_assados_sofia.png               # Conceito 3D de fachada e estação CRM
    └── charts/                                    # Figuras, plantas, identidade visual e fotos
        ├── brand_mockup_sofia.jpg                 # Identidade visual e fachada 3D
        ├── cardapio_impresso_sofia.jpg            # Cardápio impresso de balcão
        ├── cardapio_whatsapp_sofia.jpg            # Cardápio digital mobile / WhatsApp
        ├── equip1_asadora_gas.jpg                 # Máquinas giratórias de frango a gás GLP
        ├── equip2_churrasqueira_carvao.jpg        # Churrasqueira tradicional a carvão para bafo
        ├── equip3_coifa_industrial.jpg            # Sistema de coifa e exaustão industrial
        ├── equip4_freezer_horizontal.jpg          # Freezer horizontal comercial 510L
        ├── equip5_refrigerador_inox.jpg           # Refrigerador comercial vertical inox 4 portas
        ├── equip6_bancada_balanca.jpg             # Bancada inox AISI 304 com balança digital
        ├── combo1_classico_sofia.jpg              # Foto do Combo 1: O Clássico da Sofia
        ├── combo2_costela_sofia.jpg               # Foto do Combo 2: Costela Suprema no Bafo
        ├── combo3_dueto_sofia.jpg                 # Foto do Combo 3: Dueto Sofia
        ├── combo4_familia_sofia.jpg               # Foto do Combo 4: Kit Churrasco Família
        ├── planta_baixa_sofia.png                 # Planta baixa técnica (Português)
        ├── planta_baixa_sofia_es.png              # Planta baja técnica (Español)
        ├── dre_pt.png / dre_es.png                # Gráfico DRE mensal
        ├── breakeven_pt.png / breakeven_es.png    # Gráfico do ponto de equilíbrio
        ├── result12_pt.png / result12_es.png      # Projeção de fluxo de caixa em 12 meses
        ├── scenarios_pt.png / scenarios_es.png    # Análise de sensibilidade
        └── mix_pt.png / mix_es.png                # Gráfico do mix de vendas
```

---

## 📊 Síntese dos Indicadores Econômico-Financeiros

| Indicador | Valor Base (Fins de Semana) | Interpretação Gerencial |
| :--- | :--- | :--- |
| **Investimento Total** | **R$ 38.000,00** | R$ 18.000,00 Capital Próprio + R$ 20.000,00 Microcrédito Fomento Paraná (36x R$ 680,00) |
| **Investimento Fixo** | **R$ 24.500,00** | Equipamentos 100% novos com coifa industrial reglamentar VISA |
| **Capital de Giro e Pré-Operacional** | **R$ 13.500,00** | Fiança de 3 meses, estoque inicial, embalagens, alvarás e colchão de reserva (R$ 2.500,00) |
| **Receita Bruta Mensal (160 combos)** | **R$ 15.809,00** | Base de 40 combos por final de semana (4 combos padronizados) |
| **CMV Insumos e Embalagens** | **R$ 6.140,00** (38,84%) | Fichas técnicas com carnes homologadas na CEASA/PR e frigoríficos SIF |
| **Custos Fixos Mensais** | **R$ 6.870,00** | Equipe de 4 diaristas CLT intermitente, aluguel, contador, serviços e parcela de empréstimo |
| **Lucro Operacional Líquido Mensal** | **R$ 1.850,46** (11,71%) | Rentabilidade líquida saudável no cenário planejado |
| **Ponto de Equilíbrio Contábil** | **R$ 12.454,37** (~126 combos) | ~32 combos/fim de semana (~16 por dia) para cobrir todos os custos |
| **Payback Dinâmico** | **11 a 12 meses** | Amortização integral do investimento inicial no primeiro ano |
| **Alavancagem em Feriados (2026-2028)** | **+R$ 61.759,98 / +R$ 19.934,90** | 28 feriados úteis no biênio com receita bruta e lucro líquido extra |

---

## 🚀 Como Recompilar os Documentos / How to Recompile

Para gerar ou atualizar as dissertações em formato `.docx` a partir do código-fonte:

```bash
# 1. Instalar dependências necessárias
pip install python-docx matplotlib

# 2. Executar o compilador integrado
python _work/run_all.py
```

O script gerará e validará automaticamente os dois arquivos `.docx` na raiz do projeto, conferindo número de palavras, tabelas, imagens e estrutura de estilos ABNT.

---

## 📜 Licença e Direitos

Projeto desenvolvido como requisito acadêmico para o Curso Técnico em Administração e Informática do Colégio Excelência. Todos os direitos reservados.

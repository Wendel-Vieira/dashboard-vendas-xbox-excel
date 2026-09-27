"""
Análise Analítica e Auditoria de Dados de Vendas Xbox
Autor: Wendel Vieira
Projeto: Dashboard Executivo de Vendas Xbox no Excel & Python
"""

import os
import sys
import json
import argparse
from typing import Dict, Any, List
import openpyxl


def formatar_moeda(val: float) -> str:
    s = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def formatar_numero(val: int) -> str:
    return f"{val:,}".replace(",", ".")


def analisar_planilha_vendas(caminho_xlsx: str) -> Dict[str, Any]:
    if not os.path.exists(caminho_xlsx):
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {caminho_xlsx}")

    wb = openpyxl.load_workbook(caminho_xlsx, data_only=False)
    if "Base_Vendas" not in wb.sheetnames:
        raise ValueError("Aba 'Base_Vendas' não encontrada na planilha!")

    ws = wb["Base_Vendas"]

    total_receita = 0.0
    total_custo = 0.0
    total_lucro = 0.0
    total_unidades = 0

    por_categoria = {}
    por_canal = {}
    por_regiao = {}
    por_produto = {}
    por_mes = {m: {"receita": 0.0, "qtd": 0, "lucro": 0.0} for m in range(1, 13)}

    for row in ws.iter_rows(min_row=4, max_col=13, values_only=True):
        if not row[0] or str(row[0]).startswith("TOTAL"):
            continue

        id_venda = row[0]
        data_venda = row[1]
        mes = int(row[2]) if row[2] else 1
        regiao = row[3] or "Outros"
        canal = row[4] or "Outros"
        categoria = row[5] or "Outros"
        produto = row[6] or "Outros"
        qtd = int(row[7]) if row[7] is not None else 0
        preco_un = float(row[8]) if row[8] is not None else 0.0
        custo_un = float(row[10]) if row[10] is not None else 0.0

        rec = qtd * preco_un
        cst = qtd * custo_un
        luc = rec - cst

        total_receita += rec
        total_custo += cst
        total_lucro += luc
        total_unidades += qtd

        # Categoria
        if categoria not in por_categoria:
            por_categoria[categoria] = {"receita": 0.0, "qtd": 0, "lucro": 0.0}
        por_categoria[categoria]["receita"] += rec
        por_categoria[categoria]["qtd"] += qtd
        por_categoria[categoria]["lucro"] += luc

        # Canal
        if canal not in por_canal:
            por_canal[canal] = {"receita": 0.0, "qtd": 0, "lucro": 0.0}
        por_canal[canal]["receita"] += rec
        por_canal[canal]["qtd"] += qtd
        por_canal[canal]["lucro"] += luc

        # Região
        if regiao not in por_regiao:
            por_regiao[regiao] = {"receita": 0.0, "qtd": 0, "lucro": 0.0}
        por_regiao[regiao]["receita"] += rec
        por_regiao[regiao]["qtd"] += qtd
        por_regiao[regiao]["lucro"] += luc

        # Produto
        if produto not in por_produto:
            por_produto[produto] = {"categoria": categoria, "receita": 0.0, "qtd": 0, "lucro": 0.0}
        por_produto[produto]["receita"] += rec
        por_produto[produto]["qtd"] += qtd
        por_produto[produto]["lucro"] += luc

        # Mês
        if 1 <= mes <= 12:
            por_mes[mes]["receita"] += rec
            por_mes[mes]["qtd"] += qtd
            por_mes[mes]["lucro"] += luc

    margem_global = (total_lucro / total_receita * 100) if total_receita > 0 else 0.0
    ticket_medio = (total_receita / total_unidades) if total_unidades > 0 else 0.0

    # Top Produtos por Receita
    top_produtos = sorted(
        [{"produto": k, **v, "margem_pct": round(v["lucro"] / v["receita"] * 100, 2) if v["receita"] > 0 else 0.0}
         for k, v in por_produto.items()],
        key=lambda x: x["receita"],
        reverse=True
    )

    return {
        "kpis": {
            "faturamento_total": round(total_receita, 2),
            "unidades_totais": total_unidades,
            "custo_total": round(total_custo, 2),
            "lucro_bruto": round(total_lucro, 2),
            "margem_global_pct": round(margem_global, 2),
            "ticket_medio_unidade": round(ticket_medio, 2),
        },
        "categorias": {
            k: {
                "receita": round(v["receita"], 2),
                "unidades": v["qtd"],
                "lucro": round(v["lucro"], 2),
                "margem_pct": round(v["lucro"] / v["receita"] * 100, 2) if v["receita"] > 0 else 0.0,
                "share_receita_pct": round(v["receita"] / total_receita * 100, 2) if total_receita > 0 else 0.0,
            }
            for k, v in por_categoria.items()
        },
        "canais": {
            k: {
                "receita": round(v["receita"], 2),
                "unidades": v["qtd"],
                "share_pct": round(v["receita"] / total_receita * 100, 2) if total_receita > 0 else 0.0,
            }
            for k, v in por_canal.items()
        },
        "regioes": {
            k: {
                "receita": round(v["receita"], 2),
                "unidades": v["qtd"],
                "share_pct": round(v["receita"] / total_receita * 100, 2) if total_receita > 0 else 0.0,
            }
            for k, v in por_regiao.items()
        },
        "top_5_produtos": top_produtos[:5],
    }


def exibir_relatorio(analise: Dict[str, Any]):
    k = analise["kpis"]
    print("=" * 80)
    print("DASHBOARD EXECUTIVO DE VENDAS XBOX - RELATORIO ANALITICO")
    print("Autor: Wendel Vieira | Status: Consolidado")
    print("=" * 80)
    print(f"Faturamento Bruto Total : {formatar_moeda(k['faturamento_total'])}")
    print(f"Unidades Totais Vendidas: {formatar_numero(k['unidades_totais'])} unidades")
    print(f"Lucro Bruto Consolidado : {formatar_moeda(k['lucro_bruto'])}")
    print(f"Margem Bruta Media      : {k['margem_global_pct']} %")
    print(f"Ticket Medio por Unidade: {formatar_moeda(k['ticket_medio_unidade'])}")
    print("-" * 80)
    print("DESEMPENHO POR CATEGORIA:")
    print(f"{'Categoria':<15} | {'Faturamento':<16} | {'Unidades':<10} | {'Lucro':<16} | {'Margem':<8} | {'Share':<8}")
    print("-" * 80)
    for cat, v in analise["categorias"].items():
        m_str = f"{v['margem_pct']} %"
        s_str = f"{v['share_receita_pct']} %"
        print(f"{cat:<15} | {formatar_moeda(v['receita']):<16} | {v['unidades']:<10} | {formatar_moeda(v['lucro']):<16} | {m_str:<8} | {s_str:<8}")
    print("-" * 80)
    print("DESEMPENHO POR CANAL DE DISTRIBUICAO:")
    for can, v in analise["canais"].items():
        print(f" - {can:<30}: {formatar_moeda(v['receita']):<16} ({v['share_pct']} % do total)")
    print("-" * 80)
    print("TOP 5 PRODUTOS MAIS VENDIDOS:")
    for idx, p in enumerate(analise["top_5_produtos"], start=1):
        print(f" {idx}. {p['produto']:<38} | Receita: {formatar_moeda(p['receita'])} | Margem: {p['margem_pct']} %")
    print("=" * 80)


def main():
    parser = argparse.ArgumentParser(description="Auditoria e Análise de Vendas Xbox - Wendel Vieira")
    default_excel = os.path.join(os.path.dirname(__file__), "Dashboard_Vendas_Xbox.xlsx")
    parser.add_argument("--excel", type=str, default=default_excel, help="Caminho da planilha Excel")
    parser.add_argument("--json", action="store_true", help="Imprime saída em formato JSON")

    args = parser.parse_args()

    analise = analisar_planilha_vendas(args.excel)

    if args.json:
        print(json.dumps(analise, indent=2, ensure_ascii=False))
        return

    exibir_relatorio(analise)


if __name__ == "__main__":
    main()

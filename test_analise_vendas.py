"""
Testes Unitários da Análise de Vendas Xbox
Autor: Wendel Vieira
"""

import unittest
import os
from analise_vendas_xbox import analisar_planilha_vendas


class TestAnaliseVendasXbox(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.excel_path = os.path.join(os.path.dirname(__file__), "Dashboard_Vendas_Xbox.xlsx")

    def test_arquivo_existe(self):
        """Verifica se o arquivo da pasta de trabalho existe no diretório."""
        self.assertTrue(os.path.exists(self.excel_path), "Arquivo Excel deve existir.")

    def test_kpis_consistentes(self):
        """Valida que os KPIs essenciais de vendas são calculados corretamente."""
        res = analisar_planilha_vendas(self.excel_path)
        k = res["kpis"]
        self.assertGreater(k["faturamento_total"], 0, "Faturamento total deve ser positivo.")
        self.assertGreater(k["unidades_totais"], 0, "Total de unidades deve ser positivo.")
        self.assertGreater(k["lucro_bruto"], 0, "Lucro bruto deve ser positivo.")
        self.assertGreater(k["margem_global_pct"], 0, "Margem global deve ser positiva.")
        self.assertLessEqual(k["margem_global_pct"], 100, "Margem global não deve passar de 100%..")

    def test_categorias_mapeadas(self):
        """Verifica se as 4 categorias oficiais Xbox estão mapeadas."""
        res = analisar_planilha_vendas(self.excel_path)
        cats = res["categorias"].keys()
        self.assertIn("Consoles", cats)
        self.assertIn("Jogos", cats)
        self.assertIn("Acessórios", cats)
        self.assertIn("Serviços", cats)

    def test_canais_distribuicao(self):
        """Verifica presença dos canais comerciais."""
        res = analisar_planilha_vendas(self.excel_path)
        canais = res["canais"].keys()
        self.assertIn("Microsoft Store Oficial", canais)
        self.assertIn("Varejo Físico Especializado", canais)
        self.assertIn("E-commerce Parceiro", canais)

    def test_top_produtos(self):
        """Garante que a lista de top 5 produtos retorna 5 itens ordenados."""
        res = analisar_planilha_vendas(self.excel_path)
        top = res["top_5_produtos"]
        self.assertEqual(len(top), 5)
        self.assertGreaterEqual(top[0]["receita"], top[1]["receita"])


if __name__ == "__main__":
    unittest.main()

import unittest

from app import infer_sector, detect_chart_need, build_answer_prompt, extract_chart_data


class AppLogicTests(unittest.TestCase):
    def test_infer_sector_from_finance_question(self):
        sector = infer_sector("Create a budget variance analysis for FY2025")
        self.assertEqual(sector, "Finance")

    def test_infer_sector_from_hr_question(self):
        sector = infer_sector("Explain the hiring plan and attrition trend")
        self.assertEqual(sector, "HR")

    def test_detect_chart_need_for_graph_query(self):
        self.assertTrue(detect_chart_need("Show a chart of monthly revenue versus target"))

    def test_detect_chart_need_for_non_graph_query(self):
        self.assertFalse(detect_chart_need("Summarize the financial risks in this report"))

    def test_build_prompt_includes_uploaded_data_constraints(self):
        prompt = build_answer_prompt(
            sector="Finance",
            question="Compare FY2024 actual and budget",
            document_text="Actual: 100\nBudget: 120",
            active_file="budget.xlsx",
        )
        self.assertIn("uploaded file data", prompt.lower())
        self.assertIn("budget.xlsx", prompt)
        self.assertIn("compare fy2024 actual and budget", prompt.lower())

    def test_extract_chart_data_uses_real_metric_column(self):
        import tempfile
        import os

        df = __import__('pandas').DataFrame({
            'Year': [2020, 2021, 2022],
            'Revenue': [120, 180, 240],
            'Profit': [30, 60, 90],
        })

        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as temp:
            temp_path = temp.name

        try:
            df.to_csv(temp_path, index=False)
            chart = extract_chart_data(temp_path, os.path.basename(temp_path), query='show revenue trend')
            self.assertEqual(chart['categories'], ['2020', '2021', '2022'])
            self.assertEqual(chart['values'], [120, 180, 240])
            self.assertIn('Revenue', chart['title'])
        finally:
            os.remove(temp_path)

    def test_real_finance_csv_query_uses_actual_values(self):
        import os
        import pandas as pd

        csv_path = os.path.join(os.path.dirname(__file__), '..', 'uploads', 'Financial Statements.csv')
        self.assertTrue(os.path.exists(csv_path))

        df = pd.read_csv(csv_path)
        df.columns = [str(col).strip() for col in df.columns]
        aapl = df[df['Company'].astype(str).str.strip() == 'AAPL'].copy()
        row = aapl.loc[aapl['Revenue'].idxmax()]
        self.assertEqual(int(row['Year']), 2022)
        self.assertEqual(int(row['Revenue']), 394328)

    def test_answer_from_uploaded_dataframe_handles_highest_revenue_question(self):
        from app import answer_from_uploaded_dataframe
        import os

        csv_path = os.path.join(os.path.dirname(__file__), '..', 'uploads', 'Financial Statements.csv')
        answer = answer_from_uploaded_dataframe(csv_path, 'Financial Statements.csv', 'Which year had the highest revenue for AAPL?')
        self.assertIn('2022', answer)
        self.assertIn('394328', answer)


if __name__ == "__main__":
    unittest.main()

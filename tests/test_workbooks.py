from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
class WorkbookTests(unittest.TestCase):
    def test_each_script_writes_expected_summary_in_an_isolated_copy(self):
        cases = [
            ('Part A.py', 'smallAircraftData.xlsx', 'ChartForAnimals', {'Gull': 2, 'Goose': 1}),
            ('Part B.py', 'aircraftWildlifeStrikes.xlsx', 'ChartForYears', {2022: 2, 2023: 1}),
            ('part C.py', 'aircraftWildlifeStrikes.xlsx', 'ChartForMonths', {1: 2, 2: 1}),
            ('partd.py', 'aircraftWildlifeStrikes.xlsx', 'ChartForAirlines', {'Synthetic Air A': 2, 'Synthetic Air B': 1}),
        ]
        for script, filename, sheet_name, expected in cases:
            with self.subTest(script=script), tempfile.TemporaryDirectory() as directory:
                workbook = openpyxl.Workbook(); sheet = workbook.active
                headers = ['Unused'] * 32
                headers[1], headers[2], headers[5], headers[31] = 'Incident Year', 'Incident Month', 'Operator', 'Species'
                sheet.append(headers)
                for year, month, operator, species in [(2022, 1, 'Synthetic Air A', 'Herring Gull'), (2022, 1, 'Synthetic Air A', 'Herring Gull'), (2023, 2, 'Synthetic Air B', 'Canada Goose')]:
                    row = [''] * 32
                    row[1], row[2], row[5], row[31] = year, month, operator, species
                    sheet.append(row)
                target = Path(directory) / filename
                workbook.save(target); workbook.close()
                result = subprocess.run([sys.executable, str(ROOT / script)], cwd=directory, capture_output=True, text=True, timeout=20)
                self.assertEqual(result.returncode, 0, result.stderr)
                output = openpyxl.load_workbook(target)
                # Part D's original sheet title is preserved by the exercise.
                summary = output[sheet_name]
                counts = {row[0]: row[1] for row in summary.iter_rows(values_only=True) if row[0] in expected}
                self.assertEqual(counts, expected)
                self.assertEqual(len(summary._charts), 1)
                output.close()

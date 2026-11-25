# Aircraft wildlife-collision workbook exercises

Academic Python/openpyxl work from March 2023. Four scripts transform workbook columns and add count summaries/charts. Original source attribution is preserved: Part A credits Vadim Rudenko; Parts B, C, and D credit Harsh Darji. This is historical group coursework, not an aviation-risk product.

| Script | Input file in working directory | Summary |
| --- | --- | --- |
| `Part A.py` | `smallAircraftData.xlsx` | Species from column AF; remove missing/unknown values |
| `Part B.py` | `aircraftWildlifeStrikes.xlsx` | Incident year from column B |
| `part C.py` | `aircraftWildlifeStrikes.xlsx` | Incident month from column C |
| `partd.py` | `aircraftWildlifeStrikes.xlsx` | Operator from column F |

## Reproduce safely

Use Python 3.12, create a virtual environment, and run `pip install -r requirements.txt`. The original dataset is **not included**. All scripts save back to their input workbook, so work on a disposable copy in a separate directory, never an only copy. Run a script by its absolute path from that directory, for example `python /path/to/repo/"Part A.py"`.

`python -m unittest discover -s tests -v` creates isolated synthetic workbooks and checks all four scripts produce the expected counts and a chart. Tests do not download datasets or modify user workbooks. This verifies the preserved example schema; it does not establish correctness for every historical dataset or make real collision-risk claims.

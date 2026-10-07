# Lecturer files – not needed for the assessment

> **Hinweis für Studierende:** Dieser Ordner enthält nur Hilfsdateien zur Vorbereitung
> der Arbeitspakete. Er ist **nicht prüfungsrelevant**. Alles, was Sie für die Prüfung
> brauchen, liegt im Ordner [`DS_HS2026_LNW_I_Examples`](../DS_HS2026_LNW_I_Examples).

| File | Purpose |
|---|---|
| `AP01/Slides/` | Lecture slides week 01, referred to in the AP01 tasks |
| `AP02/Slides/` | Lecture slides week 04 (parts I and II) and week 05, cited in the AP02 worked solution |
| `AP03/Slides/` | Lecture slides week 06 (supervised learning), cited in the AP03 worked solution |
| `AP01/generate_cinema_db.py` | Generates `cinema.db` and `cinema_data.xlsx` (embeds `cinema_er_diagram.png`) |
| `AP01/generate_er_diagram.py` | Draws the ER diagram `cinema_er_diagram.png` (written to this folder) |
| `AP01/cinema_er_diagram.png` | ER diagram of `cinema.db` |
| `AP02/generate_smartphone_data.py` | Generates the AP02 data: `html/`, `reviews.json`, `smartphones_clean.csv`, `smartphone_data.xlsx` |
| `AP03/generate_bicycle_data.py` | Generates `bicycle_data.csv` |

The generators write their data files to the matching work package folder
`DS_HS2026_LNW_I_Examples/APxx` and can be run from any directory, e.g.

```bash
python DS_HS2026_LNW_I_Examples_Lecturer_Only/AP03/generate_bicycle_data.py
```

All generators reproduce the committed data files:

- AP02, AP03 and `cinema_er_diagram.png`: byte-identical files, except for the
  timestamps in `smartphone_data.xlsx`.
- AP01: `cinema.db` has the same schema and the same rows in the same order (identical
  SQL dump); only the internal page layout differs. `cinema_data.xlsx` has the same
  sheets, values, formats and image; only the timestamps differ.

To rebuild AP01, run `generate_er_diagram.py` first, because `generate_cinema_db.py`
embeds the PNG in `cinema_data.xlsx`.

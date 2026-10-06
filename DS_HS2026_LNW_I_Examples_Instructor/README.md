# Instructor files – not needed for the assessment

> **Hinweis für Studierende:** Dieser Ordner enthält nur Hilfsdateien zur Vorbereitung
> der Arbeitspakete. Er ist **nicht prüfungsrelevant**. Alles, was Sie für die Prüfung
> brauchen, liegt im Ordner [`DS_HS2026_LNW_I_Examples`](../DS_HS2026_LNW_I_Examples).

| File | Purpose |
|---|---|
| `AP01/generate_cinema_db.py` | Original generator for `cinema.db` and `cinema_data.xlsx` (see warning below) |
| `AP01/generate_er_diagram.py` | Draws the ER diagram `cinema_er_diagram.png` (written to this folder) |
| `AP01/cinema_er_diagram.png` | ER diagram of `cinema.db` |
| `AP02/generate_smartphone_data.py` | Generates the AP02 data: `html/`, `reviews.json`, `smartphones_clean.csv`, `smartphone_data.xlsx` |
| `AP03/generate_bicycle_data.py` | Generates `bicycle_data.csv` |

The generators write their data files to the matching work package folder
`DS_HS2026_LNW_I_Examples/APxx` and can be run from any directory, e.g.

```bash
python DS_HS2026_LNW_I_Examples_Instructor/AP03/generate_bicycle_data.py
```

The AP02 and AP03 generators reproduce the committed data files exactly.

**Warning – AP01:** do not run `generate_cinema_db.py`. The committed `cinema.db` and
`cinema_data.xlsx` were extended after generation (cinema «Kino Aurora» without
screenings, customers CU0501–CU0505 with inconsistent `membership_type`, sheets
«Description» and «ER-Diagram»). The script does not reproduce these changes and
would overwrite the exam data.

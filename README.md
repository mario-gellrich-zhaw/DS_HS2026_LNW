# Data Science HS2026 – LNW 01

Environment for performance assessment (LNW) 01: GitHub Codespaces, Python,
PostgreSQL, machine learning.

## Getting started
1. Create a new codespace from this repository (with Settings Sync turned off).
2. Wait until the setup has finished (Python packages installed, database
   container started).
3. Open the notebook of the work package (e.g.
   `DS_HS2026_LNW_I_Examples/AP01`) and run the preparation section.

## Database
Host: localhost (from the notebook) · Port: 5432 · Database: postgres ·
User: pgadmin · Password: geheim

## Example work packages

- **AP01:** Management and use of relational data (cinema database).
- **AP02:** Data acquisition, preparation and exploratory analysis (smartphone shop).
- **AP03:** Machine learning with synthetic bicycle data (used-bike dealer «VeloMarkt»).
  Open [the assessment notebook](DS_HS2026_LNW_I_Examples/AP03/AP03_Machine_Learning.ipynb).
  It follows the AP01/AP02 format: eight tasks, 40 points, about 40 minutes.
  Part A covers concept questions, Part B code diagnosis and Part C coding tasks.
  Topics are OLS regression, regression trees, random forest, classification trees,
  confusion matrix and ROC/AUC.
  Run the setup first; the variable descriptions are in the notebook.
  `bicycle_data.csv` is included.

## Instructor files

[`DS_HS2026_LNW_I_Examples_Instructor`](DS_HS2026_LNW_I_Examples_Instructor) contains
helper files for preparing the work packages (lecture slides, data generators,
ER image). They are **not needed for the assessment** and are not part of any task.

# HD-tDCS AUD cognition delivery

Brain Stimulation Original Research package: bilateral 4×1 HD-tDCS for cognitive outcomes in male inpatients with alcohol dependence (analysis set n=77).

## Contents

| Folder | Contents |
|--------|----------|
| `01_report/` | One compiled manuscript: BRS-layout DOCX |
| `02_scripts/` | Experiment and plotting scripts only |
| `03_figures/` | Figures 1–4 (PDF + PNG) |
| `04_data/processed/` | Aggregated analysis outputs (participant labels de-identified) |
| `05_submission/` | Cover letter, Highlights, submission checklist |

## Notes

- No manifest files are included.
- `01_report/` contains exactly one compiled report (BRS manuscript).
- English/Chinese working copies remain in the source project and are not duplicated here.
- `02_scripts/` excludes manuscript builders (`build_*`, `manuscript_content.py`) and reference validators.
- Raw identifiable clinical data are **not** included. Individual-level analysis tables are in the repository-root `data/deid/` folder (`study_id` only).
- Ambiguous-assignment participants in cleaning/audit logs are coded P-A…P-E.
- Author-only fields (ethics number, registration, funding, author list) remain in the manuscript Author Completion Checklist.

## Reproduction (statistics / figures)

```bash
python -m venv .venv && .venv/bin/pip install -r 02_scripts/requirements.txt
# From the source project root (individual-level cleaned table required):
# .venv/bin/python 03_analysis/clean_and_analyze.py
# .venv/bin/python 03_analysis/enhanced_analyses.py
# .venv/bin/python 03_analysis/sensitivity_analyses.py
# .venv/bin/python 03_analysis/consort_flow.py
# .venv/bin/python 03_analysis/figure_montage.py
# .venv/bin/python 03_analysis/figure_outcomes.py
# .venv/bin/python 03_analysis/figure_change.py
```

Manuscript rebuild (source project only): `.venv/bin/python 03_analysis/build_full_manuscript.py`

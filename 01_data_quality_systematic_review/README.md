# DQA Literature Review

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.XXXXXXX.svg)](https://doi.org/10.5281/zenodo.XXXXXXX)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Code, structured extraction matrices, and visualisations developed as part of the systematic review **"Data Quality Assessment of Claims-Based Healthcare Data: A Systematic Review."**

---

## Overview

This repository contains the code and structured matrices used to analyse and visualise the data quality dimensions, assessment methods, issue themes, and affected data elements reported in healthcare claims and administrative billing data.

It supports the generation of three main figures:

1. **Figure 3** — Data quality dimensions assessed across the included studies.
2. **Figure 4** — Data quality assessment methods applied in each study.
3. **Figure 5** — Relationships between data quality issue themes and affected claims-data elements.

---

## Repository Contents & Usage

| File | Produces |
|------|----------|
| `figure3_dimensions.py` | Figure 3 — frequency of DQ dimensions |
| `figure4_methods.py`    | Figure 4 — frequency of DQA methods |
| `figure5_mapping.py`    | Figure 5 — dimensions × affected elements |
| `data/extraction.csv`   | Structured extraction matrix underlying the figures |

<!-- Update the filenames above to match the actual contents of the repository. -->

**Requirements:** Python 3.11 (pandas, matplotlib)

**To reproduce the figures:**

```bash
pip install -r requirements.txt
python figure3_dimensions.py
python figure4_methods.py
python figure5_mapping.py
```

---

## Data Statement

This repository contains data extracted from published scientific studies and the code used to generate the systematic-review figures. It does **not** contain patient-level, personally identifiable, confidential, or sensitive healthcare data.

---

## Associated Publication

This code accompanies the systematic review *Data Quality Assessment of Claims-Based Healthcare Data: A Systematic Review.* The full author list, journal citation, and DOI will be added upon publication.

---

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

Copyright © 2025 Vishnu Priya

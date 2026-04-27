# Birdcage Diagram

A novel visualization technique for the evolution of hierarchical grouping
structures across multiple time slices. Each slice is laid out as a stacked
set of bands, and bands across slices show how groups split, merge, and
shift through the hierarchy over time.

This repository contains the full source code of the interactive demo
accompanying the paper **"Birdcage Diagram"** (under review at IEEE TVCG).
An online version of the demo is also available at
**https://huggingface.co/spaces/zhongjiang-licely/birdcage-diagram**.

---

## Requirements

- **Python 3.10 or newer** (tested on 3.10, 3.11, 3.12)
- A modern web browser (Chrome, Firefox, Edge, or Safari)
- ~500 MB free disk space for the dependency packages

The application launches a local web server and opens its UI in your
default browser.

---

## Installation

### Step 1 — Clone the repository

```bash
git clone https://github.com/zhongjiang-licely/Birdcage-Diagram.git
cd Birdcage-Diagram
```

### Step 2 — Create a Python virtual environment (strongly recommended)

A virtual environment isolates this project's dependencies from your
system Python and from other projects. Pick the option matching your
operating system.

**Windows (PowerShell):**

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

**Windows (Command Prompt):**

```cmd
python -m venv venv
venv\Scripts\activate.bat
```

**macOS / Linux:**

```bash
python3 -m venv venv
source venv/bin/activate
```

After activation your shell prompt should be prefixed with `(venv)`.

### Step 3 — Install the dependencies

The exact versions used during development are pinned in
`requirements.txt`:

```bash
pip install -r requirements.txt
```

Or, if you prefer to install them manually:

```bash
pip install dash==2.18.1 plotly==5.24.1 pandas==2.2.3 numpy==1.26.4 openpyxl==3.1.5 kaleido==0.2.1
```

#### Why these specific versions?

| Package    | Version  | Why pinned |
|------------|----------|-----------|
| `dash`     | 2.18.1   | The `running=` callback parameter used by the live status badge requires Dash ≥ 2.16. Newer versions (2.18.x) are tested. |
| `plotly`   | 5.24.1   | Compatible with Dash 2.18.x and stable for the figure structure used here. |
| `pandas`   | 2.2.3    | Used for the slice-table loading and fill-down logic. |
| `numpy`    | 1.26.4   | pandas 2.2.x dependency; pinned to a tested version. |
| `openpyxl` | 3.1.5    | Required by pandas to read `.xlsx` files. |
| `kaleido`  | 0.2.1    | Required by `plotly.io.to_image()` for static PNG / SVG / PDF export from the **Download** panel. **Note:** kaleido versions 1.x are incompatible — keep 0.2.1. |

---

## Running the application

With the virtual environment activated, from the repository root:

```bash
python birdcage_diagram.py --slice path/to/slice1.xlsx --slice path/to/slice2.xlsx
```

You can pass any number of `--slice` flags. Each slice file is loaded in
the order given.

Optional flags:

- `--category "Hierarchy 4"` — preselect which hierarchy column drives the
  category axis. If omitted, you can pick it from the **Category** combo
  box in the UI.

The console will print:

```
RUNNING FILE: birdcage_diagram.py  (Band panel enabled)
Dash is running on http://127.0.0.1:8050/
```

Open `http://127.0.0.1:8050/` in your browser. Press `Ctrl+C` in the
console to stop the server.

### Quick test using the bundled examples

If you have downloaded the bundled example datasets (in `examples/`), you
can launch with one of them, e.g.:

```bash
python birdcage_diagram.py \
    --slice examples/multi_file/Accommodation_and_Food_Services_2002.xlsx \
    --slice examples/multi_file/Accommodation_and_Food_Services_2007.xlsx \
    --slice examples/multi_file/Accommodation_and_Food_Services_2012.xlsx \
    --slice examples/multi_file/Accommodation_and_Food_Services_2017.xlsx \
    --slice examples/multi_file/Accommodation_and_Food_Services_2022.xlsx
```

---

## Input data format

Each slice is a tabular file (`.xlsx`, `.xls`, `.csv`, or `.ods`) with one
column per hierarchy level, plus a rightmost **element** column. Column
names are arbitrary — the rightmost column is always treated as the
element identifier. For example:

| Sector                          | Subsector       | Industry Group         | NAICS Industry           | Element                  |
|---------------------------------|-----------------|------------------------|--------------------------|--------------------------|
| Accommodation and Food Services | Accommodation   | Traveler Accommodation | Hotels and Motels        | Hotels and Motels        |
| Accommodation and Food Services | Accommodation   | Traveler Accommodation | Casino Hotels            | Casino Hotels            |
| Accommodation and Food Services | Food Services   | Restaurants            | Full-Service Restaurants | Full-Service Restaurants |
| ...                             | ...             | ...                    | ...                      | ...                      |

Rules:

- Blank cells in hierarchy columns inherit the value above them
  (fill-down behavior).
- Element values must be unique within a single slice.
- Across slices, the column set must match. Column order may differ —
  columns are auto-aligned to the first slice.

A multi-sheet Excel workbook is also accepted: pass it as a single
`--slice` argument and each sheet becomes one slice (in sheet order).

---

## User guide

A 3-page user guide PDF covering the toolbar, canvas interactions, and
common usage patterns is included at
[`assets/Birdcage_Diagram_Use_Guide.pdf`](assets/Birdcage_Diagram_Use_Guide.pdf).
The same PDF is accessible from within the app via the **Guide** button.

---

## Troubleshooting

**`ModuleNotFoundError: No module named 'dash'` (or similar)**
Your virtual environment is not activated, or the dependencies are not
installed. Re-run Step 2 then Step 3.

**`pip install` fails on `kaleido`**
On some platforms `kaleido==0.2.1` requires up-to-date `pip`. Try:
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

**The browser tab opens but the page is blank or shows a Dash error**
Ensure you are using Python 3.10+. The app relies on type annotation
syntax (`list[...]`, etc.) that older Pythons do not support.

**Static image export (PNG / PDF / SVG) fails from the Download panel**
This means `kaleido` is not installed correctly. Reinstall it:
```bash
pip install --force-reinstall kaleido==0.2.1
```

---

## Repository layout

```
Birdcage-Diagram/
├── birdcage_diagram.py                    # Main application (single-file)
├── requirements.txt                       # Pinned Python dependencies
├── README.md                              # This file
├── LICENSE                                # MIT License
├── assets/
│   └── Birdcage_Diagram_Use_Guide.pdf     # User guide (also accessible in-app)
└── examples/
    ├── multi_file/                        # NAICS hierarchy datasets
    │   ├── Accommodation_and_Food_Services_2002.xlsx
    │   ├── ...
    │   └── Information_2022.xlsx
    └── multi_sheet/                       # Synthetic multi-sheet workbooks
        ├── D0.xlsx
        ├── ...
        └── D6.xlsx
```

---

## Citation

If you use this tool in your research, please cite the accompanying paper.
A BibTeX entry will be added here once the paper is accepted.

---

## License

MIT License — see [`LICENSE`](LICENSE) for details.

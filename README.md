# Hurry Plotter

**Hurry Plotter** is a Python GUI-based data visualization tool designed for large and complex datasets. It supports spreadsheet and tabular files containing multiple sheets and many columns, allowing users to easily compare results across experiments, select custom X and Y axes, generate different types of plots, and export high-quality figures suitable for reports, presentations, and academic publications.

## Features

- GUI-based plotting tool
- Designed for large datasets with many sheets and columns
- Supports:
  - Excel files: `.xlsx`, `.xls`
  - LibreOffice Calc files: `.ods`
  - CSV files: `.csv`
  - TSV files: `.tsv`
  - Text files: `.txt`
- Automatic sheet/table detection
- User-selectable X-axis and Y-axis
- Multiple chart types:
  - Smooth Line
  - Raw Line
  - Line + Markers
  - Scatter
  - Bar
  - Step
- Multi-sheet comparison
- Automatic legend generation
- Export plots as:
  - PNG
  - PDF
  - SVG
- Publication-style plot formatting
- Simple GUI with file import support

## Screenshot

Add your screenshot here after uploading one to the repository.

```markdown
![Hurry Plotter screenshot](docs/screenshot.png)
```

## Requirements

Hurry Plotter requires **Python 3.10 or newer**.

Python packages:

- `pandas`
- `matplotlib`
- `openpyxl`
- `odfpy`

Tkinter is also required for the graphical interface.

On Windows, Tkinter is usually included with the official Python installer.

On Linux, Tkinter may need to be installed separately.

### Linux Tkinter installation

#### Ubuntu / Debian

```bash
sudo apt update
sudo apt install python3-tk
```

#### Fedora

```bash
sudo dnf install python3-tkinter
```

#### Arch Linux

```bash
sudo pacman -S tk
```

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/hurry-plotter.git
cd hurry-plotter
```

### 2. Create a virtual environment

#### Windows

```bash
python -m venv .venv
.venv\Scripts\activate
```

If PowerShell blocks activation, run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.venv\Scripts\Activate.ps1
```

#### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

## Running the application

```bash
python Hurry_Plotter.py
```

## How to use

1. Open Hurry Plotter.
2. Click **Import Data File**.
3. Select a supported file:
   - `.xlsx`
   - `.xls`
   - `.ods`
   - `.csv`
   - `.tsv`
   - `.txt`
4. Select one or more sheets/tables.
5. Select the **X-axis** column.
6. Select the **Y-axis** column.
7. Select a chart type.
8. Click **Plot**.
9. Export the plot using:
   - **Save PNG**
   - **Save PDF**
   - **Save SVG**

## Supported file formats

| Format | Extension | Notes |
|---|---|---|
| Excel | `.xlsx`, `.xls` | Multiple sheets supported |
| LibreOffice Calc | `.ods` | Multiple sheets supported |
| CSV | `.csv` | Single table |
| TSV | `.tsv` | Single table |
| Text | `.txt` | Reads tab-separated or comma-separated data |

## Expected data format

### Spreadsheet files

For `.xlsx`, `.xls`, and `.ods` files, Hurry Plotter currently expects column headers to be on the **third row**.

Internally, this is handled with:

```python
header=2
```

This means:

- Row 1 is ignored
- Row 2 is ignored
- Row 3 contains column names
- Row 4 onward contains data

### CSV / TSV / TXT files

For `.csv`, `.tsv`, and `.txt` files, the first row is expected to contain column names.

## Chart types

Hurry Plotter supports the following chart types:

| Chart type | Description |
|---|---|
| Smooth Line | Applies a rolling average for smoother curves |
| Raw Line | Plots raw line data |
| Line + Markers | Plots lines with markers |
| Scatter | Plots individual points |
| Bar | Creates bar charts |
| Step | Creates step plots |

## Export options

Plots can be exported as:

- `.png`
- `.pdf`
- `.svg`

PDF and SVG are recommended for reports, theses, presentations, and academic publications because they preserve vector quality.

## Building a Windows executable

You can create a standalone `.exe` file using PyInstaller.

### Install PyInstaller

```bash
pip install pyinstaller
```

### Build the executable

```bash
pyinstaller --onefile --windowed Hurry_Plotter.py
```

The executable will be created in:

```text
dist/Hurry_Plotter.exe
```

## Repository structure

Recommended structure:

```text
hurry-plotter/
│
├── Hurry_Plotter.py
├── README.md
├── requirements.txt
├── LICENSE
├── .gitignore
└── examples/
    └── sample_data.xlsx
```

## Notes

- For LibreOffice `.ods` support, `odfpy` is required.
- For Excel `.xlsx` support, `openpyxl` is required.
- For best results, keep column names consistent across sheets.
- If a selected column is missing in some sheets, those sheets may be skipped.
- The legend label is generated from the first part of the sheet name.

Example:

```text
RENO_set1 -> RENO
CUBIC_set1 -> CUBIC
BBR_1ST_TRY_set1 -> BBR
```

## License

This project is licensed under the MIT License.

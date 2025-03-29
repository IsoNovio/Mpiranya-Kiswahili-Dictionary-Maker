# Dictionary Redundancy Tools

A pair of Windows executables (also Python GUI tools built with Tkinter) for processing .docx files:

1. **p1gui.exe**: Underlines repeated lines in a .docx file and produces a frequency table in Excel.
2. **p2gui.exe**: Uses a .docx file and an Excel frequency table to create a “reference table” in Excel.

------

## Using the .exe Files on Windows

1. **Download** the executables:
   - `p1gui.exe`
   - `p2gui.exe` (You may also need the supporting files, like `.docx` and `.xlsx` test files, as you wish.)
2. **Double-click** the desired `.exe` to run the GUI tool: 
   - For **p1gui.exe**:
     - Click **Browse** to select an *Input .docx File*.
     - Click **Browse** to select or name the *Output .docx File*.
     - Click **Process Files** to underline repeated lines in the `.docx` and create the `frequencyTable.xlsx`.
   - For **p2gui.exe**:
     - Click **Browse** to select a *DOCX File*.
     - Click **Browse** to select the *Excel File* containing the frequency table (e.g. `frequencyTable.xlsx`).
     - Click **Process** to create the `referenceTable.xlsx`.
3. **Verify** that the new files (`.docx`, `frequencyTable.xlsx`, and `referenceTable.xlsx`) are generated in the same folder as the `.exe` or in your chosen output directory.

_Note:_ The executables are large, so they may take some time to load.
------

## Requirements (If Not Using the EXEs)

- **Python 3.7+** (recommended)
- **Libraries**:
  - [python-docx](https://pypi.org/project/python-docx/)
  - [xlsxwriter](https://pypi.org/project/XlsxWriter/)
  - [openpyxl](https://pypi.org/project/openpyxl/)
  - [tkinter](https://docs.python.org/3/library/tkinter.html) (usually included with Python on Windows)

You can install the required libraries via pip:

```bash
pip install python-docx xlsxwriter openpyxl
```


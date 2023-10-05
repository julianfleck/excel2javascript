# Excel to JavaScript Converter

This tool allows users to convert calculations from an Excel spreadsheet into JavaScript code. In addition to the conversion, the tool provides functionality to visualize cell dependencies and dependants, and to compute cell values.

## Features

- **Conversion**: Transforms Excel formulas into JavaScript syntax.
- **Visualization**: Shows dependency trees for specific cells to understand how calculations are interconnected.
- **Computation**: Uses the generated JavaScript to compute and display the value of specific cells.

## Requirements

- Python 3.x
- Required Python libraries: `openpyxl`, `js2py`, `rich`

You can install these using pip:

```bash
pip install openpyxl js2py rich
```

## Usage

```
python excel2javascript.py [path_to_excel_file] [options]
```

### Options:

- `-c, --compute [CELL]`: Compute the value of a specific cell using generated JavaScript.
- `-f, --formula [CELL]`: Display the original formula or numeric value of a specified cell from the Excel file.
- `-o, --output [PATH]`: Save the generated JavaScript to a specified file.
- `-d, --show-dependencies [CELL]`: Show the dependency tree of a specific cell or of all cells if no cell is specified.
- `-s, --show-dependants [CELL]`: Display the direct dependants of a specific cell.

## Examples

### Compute the value of a specific cell:

```bash
python excel2javascript.py sample.xlsx -c A1
```

Output:

```
The computed value of A1 is 123.45
```

### Display the original formula or numeric value of a specified cell:

```bash
python excel2javascript.py sample.xlsx -f A1
```

Output:

```
The original formula/value of A1 is =B1+C1
```

### Save the generated JavaScript to a specified file:

```bash
python excel2javascript.py sample.xlsx -o output.js
```

Output:

```
Testing JavaScript for errors...
No errors detected.
Saved JavaScript to output.js
```

### Show the dependency tree of a specific cell:

```bash
python excel2javascript.py sample.xlsx -d A1
```

Output:

```
A1 (=B1+C1 => 123.45)
├── B1 (45)
└── C1 (78.45)
```

### Display the direct dependants of a specific cell:

```bash
python excel2javascript.py sample.xlsx -s A1
```

Output:

```
A1 (=B1+C1 => 123.45)
└── D1 (=A1*2 => 246.9)
```

---

## Notes

- This tool is designed for relatively simple Excel spreadsheets. Complex functions or features from Excel might not be supported.
- Always review and test the generated JavaScript code before using it in a production environment.

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

## Options

- `-c, --compute [CELL]`: Compute the value of a specific cell using generated JavaScript.
- `-f, --formula [CELL]`: Display the original formula or numeric value of a specified cell from the Excel file.
- `-o, --output [PATH]`: Save the generated JavaScript to a specified file.
- `-d, --show-dependencies [CELL]`: Show the dependency tree of a specific cell or of all cells if no cell is specified.
- `-s, --show-dependants [CELL]`: Display the direct dependants of a specific cell.

---

# Examples

## Compute the value of a specific cell:

```bash
python excel2javascript.py sample.xlsx -c A1
```

Output:

```
The computed value of A1 is 123.45
```

## Display the original formula or numeric value of a specified cell:

```bash
python excel2javascript.py sample.xlsx -f A1
```

Output:

```
The original formula/value of A1 is =B1+C1
```

## Save the generated JavaScript to a specified file:

```bash
python excel2javascript.py sample.xlsx -o output.js
```

Output:

```
Testing JavaScript for errors...
No errors detected.
Saved JavaScript to output.js
```

## Dependency and Dependants Trees

### Dependency Tree:

The `show-dependencies` command will generate a tree-like structure that visualizes the dependencies of a given cell. A dependency is another cell that the current cell's formula references to calculate its value. 

For example, consider a cell A1 with the formula `=B1+C1`. Both B1 and C1 are dependencies for A1, because A1's value relies on the values in those cells.

Usage:

```bash
python excel2javascript.py sample.xlsx -d A1
```

Example Output:

```
A1 (=B1+C1 => 123.45)
├── B1 (45)
└── C1 (78.45)
```

**Utility**: This feature is useful for:
1. **Understanding Formula Complexity**: By visualizing dependencies, users can quickly gauge the complexity of a particular formula.
2. **Error Tracking**: If there is an unexpected value in a cell, tracing its dependencies can help in pinpointing where the error might have originated.
3. **Optimizing Spreadsheets**: By understanding dependencies, users can make informed decisions about restructuring or simplifying their spreadsheets.

### Dependants Tree:

The `show-dependants` command generates a tree-like structure that visualizes which cells directly depend on a given cell. A dependant is a cell that references the current cell in its formula.

For example, if cell D1 has the formula `=A1*2`, then D1 is a dependant of A1.

Usage:

```bash
python excel2javascript.py sample.xlsx -s A1
```

Example Output:

```
A1 (=B1+C1 => 123.45)
└── D1 (=A1*2 => 246.9)
```

**Utility**: This feature is useful for:
1. **Impact Analysis**: Before changing a cell's value or formula, users can see which other cells will be affected by this change.
2. **Spreadsheet Management**: By understanding how data flows and which cells are critical to multiple formulas, users can manage their spreadsheets more effectively.
3. **Debugging**: If an error is propagated to multiple cells, checking the dependants can help in rectifying all affected cells.

---

# Notes

- This tool is designed for relatively simple Excel spreadsheets. Complex functions or features from Excel might not be supported.
- Always review and test the generated JavaScript code before using it in a production environment.

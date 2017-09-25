# Concatenate!

This small VBA application aims to provide an easy to use interface to concatenate multiples files with the same exact format into a single one. It is currently designed to work on files with a formatted title containing a specified year.

## Usage

  1. Create a new macro-enabled Microsoft Excel Worksheet (.xlsm).
  2. Create three modules and paste the corresponding file contents:
    - procedures
    - functions
    - license
  3. Add a call to the concatenate procedure at application startup by adding the following lines in ThisWorkbook Microsoft Excel Object:
    ````vba
      Private Sub Workbook_Open()
        Procedures.consolidate
      End Sub
    ````
## License
This project is licensed under the MIT License.

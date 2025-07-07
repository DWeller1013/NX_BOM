# NX_BOM

Generate a Bill of Materials (BOM) Excel file directly from a Siemens NX assembly using custom part and assembly attributes.

## Overview

**NX_BOM** is a Visual Basic .NET NXOpen journal designed to automate the creation of BOMs for Siemens NX assemblies. It extracts user-defined attributes from all parts in the assembly, structures them into a BOM, and outputs the result into a formatted Excel file using a customizable template.

This tool is especially helpful for teams standardizing BOM reporting, automating repetitive documentation, or integrating attribute-driven data into downstream manufacturing and quoting workflows.

## Features

- Scans all components in a Siemens NX assembly and collects:
  - Detail Number
  - Quantity
  - Description
  - Material/Ordering Number
  - Dimension
  - Heat Treat
  - Comment
  - Change
  - Owning Component (Parent)
- Generates an Excel BOM file (.xlsm or .xlsx) using a template for consistent formatting.
- If a BOM already exists, adds a new sheet instead of overwriting.
- Supports attribute grouping, quantity aggregation, and parent-child hierarchy tracking.
- Compatible with the Siemens NXOpen API and Excel via COM automation.
- Output file is saved in the assembly’s directory for easy access.

## Requirements

- Siemens NX (with NXOpen support, tested with versions 1847+)
- Microsoft Excel (with macros enabled, for .xlsm output)
- Network access to the Excel template path (default: `\\server\files\lib\macros\NX\TEMPLATES\STOCKLIST-STARTER.xlsm`)
- Windows OS (required for COM automation with Excel)
- [Optional] Modify `templatePath` in `NX_BOM.vb` to point to your preferred BOM template

## Setup

1. **Download or Clone the Repository**

   ```sh
   git clone https://github.com/DWeller1013/NX_BOM.git
   ```

2. **Configure the Template Path**

   Edit `NX_BOM.vb` and update the `templatePath` variable if your template is located elsewhere.

3. **Place Script Where Needed**

   Place the `NX_BOM.vb` file in your desired scripts/journals directory accessible from NX.

4. **Set Up NX Environment**

   - Ensure you have appropriate permissions to run journals in NX.
   - Load the journal using the NX Journal Editor or via the NX menu:
     ```
     Tools > Journal > Play
     ```

## Usage

1. **Open the Target Assembly in Siemens NX**

2. **Run the Journal**

   - Play the `NX_BOM.vb` journal.
   - The script will prompt for a part file if none is open.
   - It will scan the assembly, compile attributes, and output the BOM to Excel.

3. **Find Output**

   - The BOM Excel file will be saved in the same directory as your NX assembly.
   - If a BOM already exists, a new sheet will be added to the file.
   - Sheet names follow the format `BOM-00X`.

### Example Attribute Mapping

| NX Attribute Title         | Excel BOM Column            |
|---------------------------|-----------------------------|
| Detail Number             | Detail Number (A)           |
| Description               | Description (C)             |
| Material/Ordering Number  | Material/Ordering Number (D)|
| Dimension                 | Dimension (E)               |
| Heat_Treat                | Heat_Treat (F)              |
| Comment                   | Comment (G)                 |
| Change                    | Change (H)                  |
| Parent Component Name     | Owning Comp (I)             |

## Customization

- **Excel Template**: Adjust `templatePath` to use your own BOM format.
- **Attribute Names**: Update attribute titles in the code if your organization uses different naming conventions.
- **Additional Columns**: Extend the `AssemblyComponent` class and Excel-writing routines to support more metadata.

## Troubleshooting

- Ensure Excel and network template paths are accessible from your machine.
- NXOpen API errors typically indicate version or permissions mismatches.
- For debugging, use the `Guide.InfoWriteLine` or `Echo` subroutines in the code to print diagnostic output.

---

**Author:** [DWeller1013](https://github.com/DWeller1013)  
**Project:** [NX_BOM](https://github.com/DWeller1013/NX_BOM)

# Library Data Transformation Tool

A PowerShell-based utility for transforming and validating CSV data according to mapping rules.

## Overview

This tool provides a graphical interface for:
- Loading mapping rules from a CSV file
- Loading source data from a CSV file
- Performing data validation against specified rules
- Transforming data fields based on custom functions
- Highlighting validation errors in a grid view
- Exporting processed data to CSV

## Features

- **Data Mapping**: Define source and target field names, mandatory fields, data types, and validation rules
- **Data Validation**: Validate data against regex patterns with visual error highlighting
- **Data Transformation**: Apply custom transformation functions to data fields
- **Visual Interface**: View data in a grid with alternating row colors and error highlighting
- **Filtering**: Show only rows with validation errors
- **Export**: Export processed data and error logs to CSV files

## Requirements

- Windows PowerShell 5.1 or later
- .NET Framework 4.5 or later

## Usage

1. Run the script to open the graphical interface
2. Click "Load Files" to select your mapping and data files
3. View processed data in the grid (validation errors highlighted in pink)
4. Use "Show Errors Only" to filter to rows with validation issues
5. Use "Show All Data" to revert to showing all rows
6. Export processed data with the "Export Data" button
7. Export error logs with the "Export Errors" button

## Mapping File Format

The mapping file should be a CSV with the following columns:

| Column | Description |
| ------ | ----------- |
| SourceField | Name of the field in the source data |
| NewField | Name to use in the output (leave empty to keep original name) |
| DataType | Data type for conversion (string, int, decimal, datetime) |
| Mandatory | Y/N - whether the field is required |
| Validation | Y/N - whether to validate the field |
| ValidationRule | Regex pattern for validation |
| Transformation | Y/N - whether to transform the field |
| TransformFunction | Name of the transformation function to apply |

## Example Mapping File

```csv
SourceField,NewField,DataType,Mandatory,Validation,ValidationRule,Transformation,TransformFunction
Title,BookTitle,string,Y,N,,N,
Author,AuthorName,string,Y,N,,N,
Barcode,,string,Y,Y,^\d{12}$,N,
Pages,PageCount,int,N,Y,^\d+$,N,
```

## Data Processing

The tool performs the following operations on the data:

1. **Field Mapping**: Maps source fields to new field names if specified
2. **Data Type Conversion**: Converts data to the specified type
3. **Mandatory Field Validation**: Checks if required fields are present
4. **Pattern Validation**: Validates data against regex patterns
5. **Transformation**: Applies custom transformation functions

## Extending Transformations

You can add custom transformation functions by editing the script. Each function should:

1. Accept a single input value
2. Return the transformed value
3. Be registered in the `$global:TransformFunctions` hashtable

Example transformation function:

```powershell
function GenderTransform($value) {
    if ($value -match "Mr") { return "Male" }
    elseif ($value -match "Mrs" -or $value -match "Miss") { return "Female" }
    else { return "" }
}

$global:TransformFunctions = @{
    "GenderTransform" = ${function:GenderTransform}
}
```

## Logging

The tool provides detailed logging of:
- Loaded mapping and data files
- Validation errors with specific details
- Transformation summaries including affected fields and records
- Row and column counts

## Implementation Details

The tool is built using:
- PowerShell scripting language
- Windows Forms for the GUI
- .NET Framework classes for data manipulation
- Regular expressions for validation

The main components are:
- Form layout with resizable panels
- DataGridView for data display with row numbers
- Support for row and column selection
- Alternating row colors and error cell highlighting

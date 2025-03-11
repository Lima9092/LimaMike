Library Data Transformation Tool
A PowerShell-based utility for transforming and validating CSV data according to mapping rules.
Overview
This tool provides a graphical interface for:

Loading mapping rules from a CSV file
Loading source data from a CSV file
Performing data validation against specified rules
Transforming data fields based on custom functions
Highlighting validation errors in a grid view
Exporting processed data to CSV

Features

Data Mapping: Define source and target field names, mandatory fields, data types, and validation rules
Data Validation: Validate data against regex patterns with visual error highlighting
Data Transformation: Apply custom transformation functions to data fields
Visual Interface: View data in a grid with alternating row colors and error highlighting
Filtering: Show only rows with validation errors
Export: Export processed data and error logs to CSV files

Requirements

Windows PowerShell 5.1 or later
.NET Framework 4.5 or later

Usage

Run the script to open the graphical interface
Click "Load Files" to select your mapping and data files
View processed data in the grid (validation errors highlighted in pink)
Use "Show Errors Only" to filter to rows with validation issues
Use "Show All Data" to revert to showing all rows
Export processed data with the "Export Data" button
Export error logs with the "Export Errors" button

Mapping File Format
The mapping file should be a CSV with the following columns:
ColumnDescriptionSourceFieldName of the field in the source dataNewFieldName to use in the output (leave empty to keep original name)DataTypeData type for conversion (string, int, decimal, datetime)MandatoryY/N - whether the field is requiredValidationY/N - whether to validate the fieldValidationRuleRegex pattern for validationErrorHandlingHow to handle validation errorsTransformationY/N - whether to transform the fieldTransformFunctionName of the transformation function to applyDefaultValueDefault value if error handling is set to Default
Example Mapping File
csvCopySourceField,NewField,DataType,Mandatory,Validation,ValidationRule,ErrorHandling,Transformation,TransformFunction,DefaultValue
Title,BookTitle,string,Y,N,,Error,N,,
Author,AuthorName,string,Y,N,,Error,N,,
Barcode,,string,Y,Y,^\d{12}$,Error,N,,
Pages,PageCount,int,N,Y,^\d+$,Default,N,,0
Data Processing
The tool performs the following operations on the data:

Field Mapping: Maps source fields to new field names if specified
Data Type Conversion: Converts data to the specified type
Mandatory Field Validation: Checks if required fields are present
Pattern Validation: Validates data against regex patterns
Transformation: Applies custom transformation functions

Validation Rule Examples
The ValidationRule column in the mapping file accepts regular expression patterns. Here are common validation patterns for UK formats:
Data TypeRegex PatternDescriptionEmail^[\w\.-]+@[\w\.-]+\.\w+$Validates email formatUK Phone`^(?:(?:+44\s?0)(?:1\d{8,9}Date (DD/MM/YYYY)`^(0[1-9][12][0-9]UK Postcode`^([A-Z]{1,2}\d[A-Z\d]? ?\d[A-Z]{2}GIR ?0A{2})$`UK National Insurance^[A-CEGHJ-PR-TW-Z]{1}[A-CEGHJ-NPR-TW-Z]{1}[0-9]{6}[A-D]{1}$NI number formatUK VAT Number`^GB\d{9}$^GB\d{12}$`UK Company Number`^(SCNIUK Bank Sort Code^\d{2}-\d{2}-\d{2}$Format: 12-34-56UK Bank Account^\d{8}$8-digit account numberNumeric only^\d+$Only digits allowedPrice (£)^£\d+\.\d{2}$Format: £12.99ISBN^(?:ISBN(?:-13)?:?\s)?(?=[0-9X]{10}$|(?=(?:[0-9]+[-\s]){3})[-\s0-9X]{13}$)ISBN formatURL^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \.-]*)*\/?$Web URLName^[a-zA-Z\s'-]+$Letters, spaces, hyphens, apostrophesUK Driving License^[A-Z]{5}\d{6}[A-Z]{2}\d[A-Z]{2}$UK driving license format
Additional Regex Use Cases
Validation PurposeRegex PatternDescriptionAlpha-numeric^[a-zA-Z0-9]+$Letters and numbers onlyFixed length^.{10}$Exactly 10 charactersMin-max length^.{8,16}$Between 8-16 charactersInteger range`^([1-9][1-9][0-9]Time (24-hour)`^([01]?[0-9]2[0-3]):[0-5][0-9]$`Hex color`^#([A-Fa-f0-9]{6}[A-Fa-f0-9]{3})$`IP Address^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$Format: 192.168.1.1UK Passport Number^[0-9]{9}$UK passport number (9 digits)UK NHS Number^\d{3}[ -]?\d{3}[ -]?\d{4}$Format: 123 456 7890
Error Handling Options
The ErrorHandling column in the mapping file specifies how validation errors should be treated. Valid options are:
OptionDescriptionWarningHighlights the cell, logs the error, but allows processing to continueErrorHighlights the cell, logs the error, and marks the record as invalidLogOnly logs the error without visual highlighting or affecting processingIgnoreValidation fails but no error is logged or displayedRejectThe entire record is rejected from processingNullError field is set to null/empty but record is processedDefaultError field is set to a default value (specified in DefaultValue column)
Example of mapping file with error handling:
csvCopySourceField,NewField,DataType,Mandatory,Validation,ValidationRule,ErrorHandling,Transformation,TransformFunction,DefaultValue
Title,BookTitle,string,Y,N,,Error,N,,
Author,AuthorName,string,Y,N,,Error,N,,
Barcode,,string,Y,Y,^\d{12}$,Error,N,,
PostCode,,string,Y,Y,^([A-Z]{1,2}\d[A-Z\d]? ?\d[A-Z]{2}|GIR ?0A{2})$,Warning,N,,
Pages,PageCount,int,N,Y,^\d+$,Default,N,,0
Email,,string,N,Y,^[\w\.-]+@[\w\.-]+\.\w+$,Log,N,,
PublicationDate,,datetime,Y,Y,^(0[1-9]|[12][0-9]|3[01])\/(0[1-9]|1[0-2])\/\d{4}$,Null,N,,
Extending Transformations
You can add custom transformation functions by editing the script. Each function should:

Accept a single input value
Return the transformed value
Be registered in the $global:TransformFunctions hashtable

Example transformation function:
powershellCopyfunction GenderTransform($value) {
    if ($value -match "Mr") { return "Male" }
    elseif ($value -match "Mrs" -or $value -match "Miss") { return "Female" }
    else { return "" }
}

$global:TransformFunctions = @{
    "GenderTransform" = ${function:GenderTransform}
}
Implementing Complex Validation
For more complex validation that regex alone can't handle, you can create custom validation functions similar to transformation functions:
powershellCopyfunction ValidateUKPostcode($value) {
    # First standardize by removing spaces and converting to uppercase
    $value = $value.ToUpper().Replace(" ", "")
    
    # Basic format check using regex
    if (-not ($value -match '^[A-Z]{1,2}[0-9R][0-9A-Z]?[0-9][ABD-HJLNP-UW-Z]{2}$')) {
        return $false
    }
    
    # Additional validation logic
    # For example, certain letter combinations are not used in the first position
    $invalidFirstLetters = @("QV", "X")
    $firstPart = if ($value[0..1] -join "" -match "^[A-Z]{2}$") { $value[0..1] -join "" } else { $value[0] }
    
    if ($invalidFirstLetters -contains $firstPart) {
        return $false
    }
    
    return $true
}

$global:ValidationFunctions = @{
    "ValidateUKPostcode" = ${function:ValidateUKPostcode}
}
Then in your mapping file:
csvCopySourceField,NewField,DataType,Mandatory,Validation,ValidationRule,ErrorHandling,Transformation,TransformFunction,ValidationFunction
Postcode,,string,Y,Y,,Error,N,,ValidateUKPostcode
Validation Best Practices

Start simple: Begin with basic validations and add complexity as needed
Test thoroughly: Create test data that intentionally violates validation rules
Layer validation: Use regex for format, custom functions for complex logic
Balance strictness: Overly strict validation may reject valid data
Use appropriate error handling: Choose between warning, error, and log based on data importance
Document patterns: Keep a record of regex patterns and their purposes
Consider data cleansing: Sometimes it's better to transform/cleanse than reject

Example Validation Scenarios
Library Book Catalog:
csvCopySourceField,NewField,DataType,Mandatory,Validation,ValidationRule,ErrorHandling
Title,,string,Y,N,,Error
ISBN,,string,Y,Y,^(?:ISBN(?:-13)?:?\s)?(?=[0-9X]{10}$|(?=(?:[0-9]+[-\s]){3})[-\s0-9X]{13}$),Error
PublishedYear,,int,Y,Y,^(19|20)\d{2}$,Error
PageCount,,int,N,Y,^\d{1,4}$,Warning
Publisher,,string,Y,N,,Log
UK Customer Records:
csvCopySourceField,NewField,DataType,Mandatory,Validation,ValidationRule,ErrorHandling
CustomerID,,string,Y,Y,^C\d{6}$,Error
FirstName,,string,Y,Y,^[A-Za-z\-']{2,30}$,Warning
LastName,,string,Y,Y,^[A-Za-z\-']{2,30}$,Warning
EmailAddress,,string,Y,Y,^[\w\.-]+@[\w\.-]+\.\w+$,Error
Postcode,,string,Y,Y,^([A-Z]{1,2}\d[A-Z\d]? ?\d[A-Z]{2}|GIR ?0A{2})$,Error
TelephoneNumber,,string,N,Y,^(?:(?:\+44\s?|0)(?:1\d{8,9}|[23]\d{9}|7(?:[1345789]\d{8}|624\d{6})))$,Warning
NationalInsurance,,string,N,Y,^[A-CEGHJ-PR-TW-Z]{1}[A-CEGHJ-NPR-TW-Z]{1}[0-9]{6}[A-D]{1}$,Log
UK Financial Transactions:
csvCopySourceField,NewField,DataType,Mandatory,Validation,ValidationRule,ErrorHandling
TransactionID,,string,Y,Y,^T\d{10}$,Error
Amount,,decimal,Y,Y,^\d+\.\d{2}$,Error
Currency,,string,Y,Y,^(GBP|EUR|USD)$,Error
SortCode,,string,Y,Y,^\d{2}-\d{2}-\d{2}$,Error
AccountNumber,,string,Y,Y,^\d{8}$,Error
TransactionDate,,datetime,Y,Y,^(0[1-9]|[12][0-9]|3[01])\/(0[1-9]|1[0-2])\/\d{4}$,Error
VATNumber,,string,N,Y,^GB\d{9}$|^GB\d{12}$,Log
Logging
The tool provides detailed logging of:

Loaded mapping and data files
Validation errors with specific details
Transformation summaries including affected fields and records
Row and column counts

Implementation Details
The tool is built using:

PowerShell scripting language
Windows Forms for the GUI
.NET Framework classes for data manipulation
Regular expressions for validation

The main components are:

Form layout with resizable panels
DataGridView for data display with row numbers
Support for row and column selection
Alternating row colors and error cell highlighting

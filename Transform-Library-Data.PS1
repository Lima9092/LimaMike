# Simple Library Data Transformation Tool

# Minimize the PowerShell command prompt
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GetConsoleWindow();
    }
"@
$consolePtr = [Win32]::GetConsoleWindow()
if ($consolePtr -ne [IntPtr]::Zero) {
    # 6 = Minimize the window
    [Win32]::ShowWindow($consolePtr, 6)
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$global:ErrorCells = @{}
$global:Mapping = $null
$global:ProcessedData = $null
$global:RawData = $null
$global:ShowOnlyMappedFields = $false
$global:UseMappingOrder = $false
$global:OriginalColumnOrder = @()
$global:MappingColumnOrder = @()
$global:LogData = @()
$global:SourceToNewFieldMap = @{}  # Added to track field name changes

# Transformation function
function GenderTransform($value) {
    if ($value -match "Mr") { return "Male" }
    elseif ($value -match "Mrs" -or $value -match "Miss") { return "Female" }
    else { return "" }
}

$global:TransformFunctions = @{
    "GenderTransform" = ${function:GenderTransform}
}

# Load mapping file
function Load-Mapping($path) {
    try {
        $mapping = Import-Csv -Path $path
        return $mapping
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading mapping file: $_")
        return $null
    }
}

# Load data file
function Load-Data($path) {
    try {
        # Read the raw CSV file first to get the real column order
        $rawContent = Get-Content -Path $path -Raw
        $headerLine = $rawContent.Split("`n")[0].Trim()
        
        # Process the header line to handle quoted column names
        $quotedColumnPattern = '(?:^|,)(?:"([^"]*(?:""[^"]*)*)"|([^,"]*))'
        $matches = [regex]::Matches($headerLine, $quotedColumnPattern)
        $global:OriginalColumnOrder = @()
        
        foreach ($match in $matches) {
            $columnName = if ($match.Groups[1].Success) { 
                # Handle quoted values (and double-quotes inside quotes)
                $match.Groups[1].Value.Replace('""', '"') 
            } else { 
                # Handle non-quoted values
                $match.Groups[2].Value 
            }
            $global:OriginalColumnOrder += $columnName
        }
        
        # Log detected column order for debugging
        Write-Host "Original CSV columns: $($global:OriginalColumnOrder -join ', ')"
        
        # Now load with Import-CSV
        $data = Import-Csv -Path $path
        return $data
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading data file: $_")
        return $null
    }
}

# Process data according to mapping rules
function Process-Data($data, $mapping) {
    $result = @()
    $log = @()
    $global:ErrorCells.Clear()  # Clear error cells before processing
    $global:RawData = $data  # Store raw data for reference
	
	# Make sure field mappings are captured
	$global:SourceToNewFieldMap.Clear()
	foreach ($m in $mapping) {
		if (-not [string]::IsNullOrEmpty($m.SourceField) -and -not [string]::IsNullOrEmpty($m.NewField)) {
			$global:SourceToNewFieldMap[$m.SourceField] = $m.NewField
		}
	}
    
    # Capture original column order when loading mapping
    if ($mapping.Count -gt 0) {
        $global:MappingColumnOrder = @()
        foreach ($m in $mapping) {
            if (-not [string]::IsNullOrEmpty($m.NewField)) {
                $global:MappingColumnOrder += $m.NewField
            } elseif (-not [string]::IsNullOrEmpty($m.SourceField)) {
                $global:MappingColumnOrder += $m.SourceField
            }
        }
    }

    # Create lookup dictionary from mapping
    $mapLookup = @{}
    foreach ($m in $mapping) {
        if (-not [string]::IsNullOrEmpty($m.SourceField)) {
            $mapLookup[$m.SourceField] = $m
        }
    }
    
    # Create dictionary to track which mandatory fields are in the mapping
    $mandatoryFields = @{}
    foreach ($m in $mapping) {
        if ($m.Mandatory -eq "Y") {
            $mandatoryFields[$m.SourceField] = $true
        }
    }

    # Process each row
    for ($i = 0; $i -lt $data.Count; $i++) {
        $row = $data[$i]
        $newRow = @{}
        $processedFields = @{}  # Track which fields we've processed
        
        # Process each property in the row
        foreach ($prop in $row.PSObject.Properties) {
            $sourceField = $prop.Name
            $value = $prop.Value
            
            # Check if this field is in our mapping
            if ($mapLookup.ContainsKey($sourceField)) {
                $map = $mapLookup[$sourceField]
                $dataType = $map.DataType
                $mandatory = $map.Mandatory -eq "Y"
                $doValidation = $map.Validation -eq "Y"
                $validationRule = $map.ValidationRule
                $doTransform = $map.Transformation -eq "Y"
                $transformFunction = $map.TransformFunction
                $error = $null
                
                # Track that we've processed this field
                $processedFields[$sourceField] = $true

                # Store original value before transformation for mandatory check
                $originalValue = $value
                
                # Apply transformation
                if ($doTransform -and $transformFunction -and $global:TransformFunctions.ContainsKey($transformFunction)) {
                    try {
                        $value = & $global:TransformFunctions[$transformFunction] $value
                    } catch {
                        $error = "Error in transformation: $_"
                    }
                }
                
                # Mandatory check - check BOTH original and transformed values
                # This ensures we catch empty values even after transformation
                if ($mandatory) {
                    if ([string]::IsNullOrEmpty($originalValue)) {
                        $error = "Mandatory field missing (original value)"
                    } elseif ([string]::IsNullOrEmpty($value)) {
                        $error = "Mandatory field empty after transformation"
                    }
                }
                
                # Type conversion
                if ($dataType -and $value -ne $null -and $value -ne "") {
                    try {
                        switch ($dataType.ToLower()) {
                            "int" { $value = [int]$value }
                            "decimal" { $value = [decimal]$value }
                            "datetime" { $value = [datetime]$value }
                            default { }
                        }
                    } catch {
                        $error = "Data type conversion failed for $dataType"
                    }
                }
                
                # Validation
                if ($doValidation -and $value -ne $null -and $value -ne "" -and $validationRule) {
                    if (-not ($value -match $validationRule)) {
                        $error = "Validation failed for rule $validationRule"
                    }
                }

                # Use the mapped field name
                $newField = $map.NewField
                if ([string]::IsNullOrEmpty($newField)) {
                    $newField = $sourceField
                }
                
                $newRow[$newField] = $value

                # Record errors
                if ($error) {
                    $log += "Row $($i+1) Field '$sourceField': $error"
                    $global:ErrorCells["$i-$newField"] = $error
                }
            }
            else {
                # Pass through unmapped fields (only if we're not in mapped-fields-only mode)
                if (-not $global:ShowOnlyMappedFields) {
                    $newRow[$sourceField] = $value
                }
            }
        }
        
        # Check for missing mandatory fields that weren't in the input at all
        foreach ($mandatoryField in $mandatoryFields.Keys) {
            if (-not $processedFields.ContainsKey($mandatoryField)) {
                # This mandatory field wasn't in the input data
                $map = $mapLookup[$mandatoryField]
                $newField = if (-not [string]::IsNullOrEmpty($map.NewField)) { $map.NewField } else { $mandatoryField }
                
                # Add empty value to output
                $newRow[$newField] = ""
                
                # Record the error
                $error = "Required field '$mandatoryField' is missing from input data"
                $log += "Row $($i+1): $error"
                $global:ErrorCells["$i-$newField"] = $error
            }
        }
        
        # Add the processed row
        if ($newRow.Keys.Count -gt 0) {
            $result += [PSCustomObject]$newRow
        } else {
            $log += "Row $($i+1) - No valid properties found"
        }
    }
    
    return @{ Data = $result; Log = $log }
}

# Create the main form with a clear layout
$form = New-Object System.Windows.Forms.Form
$form.Text = "Library Data Transformation Tool"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)

# Create a TableLayoutPanel for better layout control
$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.RowCount = 2
$tableLayout.ColumnCount = 1
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) # Reduced height for buttons
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # Remaining space for grid/log
$form.Controls.Add($tableLayout)

# Top button panel in first row of table
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($buttonPanel, 0, 0)
# Buttons with corrected positioning
$btnLoadFiles = New-Object System.Windows.Forms.Button
$btnLoadFiles.Text = "Load Files"
$btnLoadFiles.Width = 120
$btnLoadFiles.Location = New-Object System.Drawing.Point(10, 10)
$buttonPanel.Controls.Add($btnLoadFiles)

$btnFilter = New-Object System.Windows.Forms.Button
$btnFilter.Text = "Show Errors Only"
$btnFilter.Width = 120
$btnFilter.Location = New-Object System.Drawing.Point(140, 10)
$buttonPanel.Controls.Add($btnFilter)

$btnMappedOnly = New-Object System.Windows.Forms.Button
$btnMappedOnly.Text = "Show Mapped Only"
$btnMappedOnly.Width = 120
$btnMappedOnly.Location = New-Object System.Drawing.Point(270, 10)
$buttonPanel.Controls.Add($btnMappedOnly)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "Show All Data"
$btnReset.Width = 120
$btnReset.Location = New-Object System.Drawing.Point(400, 10)
$buttonPanel.Controls.Add($btnReset)

$btnSortByData = New-Object System.Windows.Forms.Button
$btnSortByData.Text = "Sort by Data"
$btnSortByData.Width = 120
$btnSortByData.Location = New-Object System.Drawing.Point(530, 10)
$buttonPanel.Controls.Add($btnSortByData)

$btnSortByMapping = New-Object System.Windows.Forms.Button
$btnSortByMapping.Text = "Sort by Mapping"
$btnSortByMapping.Width = 120
$btnSortByMapping.Location = New-Object System.Drawing.Point(660, 10)
$buttonPanel.Controls.Add($btnSortByMapping)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Data"
$btnExport.Width = 120
$btnExport.Location = New-Object System.Drawing.Point(790, 10)
$buttonPanel.Controls.Add($btnExport)

$btnExportErrors = New-Object System.Windows.Forms.Button
$btnExportErrors.Text = "Export Errors"
$btnExportErrors.Width = 120
$btnExportErrors.Location = New-Object System.Drawing.Point(920, 10)
$buttonPanel.Controls.Add($btnExportErrors)

# Add a split container in the second row for grid and log
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.Controls.Add($contentPanel, 0, 1)

$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitContainer.SplitterDistance = 500
$splitContainer.Panel1MinSize = 200
$splitContainer.Panel2MinSize = 100
$splitContainer.SplitterWidth = 5
$contentPanel.Controls.Add($splitContainer)

# Create a panel to contain the grid without colored bar
$gridPanel = New-Object System.Windows.Forms.Panel
$gridPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$gridPanel.BackColor = [System.Drawing.Color]::White
$splitContainer.Panel1.Controls.Add($gridPanel)

# Create the DataGridView
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
$grid.BackColor = [System.Drawing.Color]::White
# Set DataGridView properties
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
$grid.RowHeadersWidth = 60  # Wider row headers to show row numbers
$grid.RowHeadersVisible = $true
$grid.ColumnHeadersVisible = $true
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersHeight = 30
$grid.DefaultCellStyle.NullValue = ""  # Display empty string for null values

# Handle data error events to prevent exceptions from being shown to the user
$grid.Add_DataError({
    param($sender, $e)
    
    # Log the error quietly instead of showing the default error dialog
    $txtLog.AppendText("`r`nGrid display error in cell [Row $($e.RowIndex), Column $($e.ColumnIndex)]: $($e.Exception.Message)")
    
    # Tell the grid that we've handled the error
    $e.ThrowException = $false
})

# Change header text color to black
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$grid.RowHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$grid.RowHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$grid.RowHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue  # Alternating row colors
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true

# Add the grid to the panel (after the label)
$gridPanel.Controls.Add($grid)

# Create the log panel without colored bar
$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$logPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$logPanel.BackColor = [System.Drawing.Color]::White
$splitContainer.Panel2.Controls.Add($logPanel)

# Log text box
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.Text = "Ready. Click 'Load Files' to load your data files."
$logPanel.Controls.Add($txtLog)

# Event handler to display row numbers in row headers
$grid.Add_RowPostPaint({
    param($sender, $e)
    
    # Draw row numbers in the row header cells
    $rowNumber = ($e.RowIndex + 1).ToString()
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
    $size = [System.Windows.Forms.TextRenderer]::MeasureText($rowNumber, $grid.Font)
    
    # Calculate the position for the text
    $x = $e.RowBounds.Location.X + ($grid.RowHeadersWidth - $size.Width) / 2
    $y = $e.RowBounds.Location.Y + ($e.RowBounds.Height - $size.Height) / 2
    
    # Create a PointF instead of Point
    $location = New-Object System.Drawing.PointF($x, $y)
    
    # Draw the string directly with point coordinates
    $e.Graphics.DrawString($rowNumber, $grid.Font, $brush, $location)
})

# Event handler for column header click (for better column selection)
$grid.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    
    $hitTestInfo = $grid.HitTest($e.X, $e.Y)
    if ($hitTestInfo.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::ColumnHeader) {
        # Select all cells in the column
        $grid.ClearSelection()
        for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
            $grid.Rows[$i].Cells[$hitTestInfo.ColumnIndex].Selected = $true
        }
    }
})

# Event handler for row header click - Fixed to select entire row
$grid.Add_RowHeaderMouseClick({
    param($sender, $e)
    
    # Only handle clicks in row headers
    if ($e.RowIndex -ge 0) {
        # Clear any existing selection
        $grid.ClearSelection()
        
        # Select the entire row
        $grid.Rows[$e.RowIndex].Selected = $true
        
        # Log for debugging
        $txtLog.AppendText("`r`nSelected row $($e.RowIndex + 1) via row header")
    }
})

# Cell formatting event to ensure error cells stay highlighted
$grid.Add_CellFormatting({
    param($sender, $e)
    
    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
    
    $rowIndex = $e.RowIndex
    $colName = if ($e.ColumnIndex -lt $grid.Columns.Count) { 
        $grid.Columns[$e.ColumnIndex].Name 
    } else { 
        return 
    }
    
    $key = "$rowIndex-$colName"
    if ($global:ErrorCells.ContainsKey($key)) {
        $e.CellStyle.BackColor = [System.Drawing.Color]::LightPink
        $e.FormattingApplied = $true
    }
    # Don't override the alternating row style if it's not an error cell
})

function Load-Files {
    $global:ErrorCells.Clear()
    $txtLog.Text = "Starting file loading process..."

	# Build the source-to-new field map for column naming
	$global:SourceToNewFieldMap.Clear()
	foreach ($m in $mapping) {
		if (-not [string]::IsNullOrEmpty($m.SourceField) -and -not [string]::IsNullOrEmpty($m.NewField)) {
			$global:SourceToNewFieldMap[$m.SourceField] = $m.NewField
			$txtLog.Text += "`r`nField mapping: '$($m.SourceField)' => '$($m.NewField)'"
		}
	}
	$txtLog.Text += "`r`nTotal field mappings: $($global:SourceToNewFieldMap.Count)"
    
    # Clear the grid
    $grid.DataSource = $null
    $grid.Rows.Clear()
    $grid.Columns.Clear()
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    
    # Select Mapping CSV
    $openFileDialog.Title = "Select Mapping CSV File"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $txtLog.Text += "`r`nUser cancelled mapping file selection."
        return
    }
    
    $mappingPath = $openFileDialog.FileName
    $txtLog.Text += "`r`nLoading mapping file: $mappingPath"
    
    # Load mapping
    $mapping = Load-Mapping $mappingPath
    if (-not $mapping -or $mapping.Count -eq 0) {
        $txtLog.Text += "`r`nMapping file is null or empty. Loading failed."
        return
    }
    
    $global:Mapping = $mapping
    $txtLog.Text += "`r`nLoaded mapping with $($mapping.Count) row(s)"
    
    # Additional debug information 
    $txtLog.Text += "`r`nOriginal column order from CSV: $($global:OriginalColumnOrder -join ', ')"
    
    # Extract column order from mapping
    $global:MappingColumnOrder = @()
    foreach ($m in $mapping) {
        $fieldName = if (-not [string]::IsNullOrEmpty($m.NewField)) { $m.NewField } else { $m.SourceField }
        if (-not [string]::IsNullOrEmpty($fieldName) -and $global:MappingColumnOrder -notcontains $fieldName) {
            $global:MappingColumnOrder += $fieldName
        }
    }
    
    $txtLog.Text += "`r`nMapping column order: $($global:MappingColumnOrder -join ', ')"
    
    # Select Data CSV
    $openFileDialog.Title = "Select Data CSV File"
    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $txtLog.Text += "`r`nUser cancelled data file selection."
        return
    }
    
    $dataPath = $openFileDialog.FileName
    $txtLog.Text += "`r`nLoading data file: $dataPath"
    
    # Load data
    $data = Load-Data $dataPath
    if (-not $data -or $data.Count -eq 0) {
        $txtLog.Text += "`r`nData file is null or empty."
        return
    }
    
    $txtLog.Text += "`r`nLoaded data with $($data.Count) row(s)"
    
    # Process data
    $txtLog.Text += "`r`nProcessing data..."
    try {
        $result = Process-Data $data $global:Mapping
        $global:ProcessedData = $result.Data
        $global:LogData = $result.Log
        
        if ($global:LogData.Count -gt 0) {
            $txtLog.Text += "`r`n--- Validation Errors ---"
            foreach ($error in $global:LogData) {
                $txtLog.Text += "`r`n$error"
            }
        } else {
            $txtLog.Text += "`r`nNo validation errors found."
        }
    } catch {
        $txtLog.Text += "`r`nError during processing: $_"
        $global:ProcessedData = $data
    }
    
    # Display in grid with row numbers and alternating colors
    if ($global:ProcessedData.Count -gt 0) {
        $firstRow = $global:ProcessedData[0]
        $columns = $firstRow.PSObject.Properties.Name
        
        $txtLog.Text += "`r`nDetected columns: $($columns -join ', ')"
        
        # Use the Update-GridView function to populate the grid
        Update-GridView
        
        $txtLog.Text += "`r`nData loaded successfully. Found $($global:ProcessedData.Count) records."
        
        # Enhanced transformation reporting
        $transformFields = @()
        $transformRecordCount = 0
        $transformedRecords = @{}
        
        # Collect information about transforms
        foreach ($m in $global:Mapping) {
            if ($m.Transformation -eq "Y") {
                $transformFields += $m.SourceField
                
                # Count records that had this transformation applied
                $sourceField = $m.SourceField
                $targetField = if (-not [string]::IsNullOrEmpty($m.NewField)) { $m.NewField } else { $sourceField }
                
                # Count records with transformations
                for ($i = 0; $i -lt $global:ProcessedData.Count; $i++) {
                    $row = $global:ProcessedData[$i]
                    if ($row.$targetField -ne $data[$i].$sourceField) {
                        $transformedRecords[$i] = $true
                    }
                }
            }
        }
        
        # Report transformation summary
        if ($transformFields.Count -gt 0) {
            $transformRecordCount = $transformedRecords.Count
            $txtLog.Text += "`r`nTransformed $($transformFields.Count) field types during processing:"
            foreach ($field in $transformFields) {
                $txtLog.Text += "`r`n - $field"
            }
            $txtLog.Text += "`r`nTransformations affected $transformRecordCount out of $($global:ProcessedData.Count) records."
        }
    } else {
        $txtLog.Text += "`r`nNo data to display"
    }
}

# Filter to show only error rows
function Show-ErrorsOnly {
    $txtLog.Text += "`r`nFiltering to show only rows with errors..."
    
    if ($global:ErrorCells.Count -eq 0) {
        $txtLog.Text += "`r`nNo error cells detected. Nothing to filter."
        return
    }
    
    # Get unique row indices with errors
    $errorRows = @{}
    foreach ($key in $global:ErrorCells.Keys) {
        $parts = $key -split "-"
        if ($parts.Count -eq 2) {
            $rowIndex = [int]$parts[0]
            $errorRows[$rowIndex] = $true
        }
    }
    
    # Hide rows without errors
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Visible = $errorRows.ContainsKey($i)
    }
    
    $txtLog.Text += "`r`nShowing only rows with errors."
}

# Filter to show only mapped fields
function Show-MappedOnly {
    $txtLog.Text += "`r`nFiltering to show only fields specified in the mapping..."
    
    if ($null -eq $global:Mapping -or $global:Mapping.Count -eq 0) {
        $txtLog.Text += "`r`nNo mapping loaded. Nothing to filter."
        return
    }
    
    # Set flag for showing only mapped fields
    $global:ShowOnlyMappedFields = $true
    
    # Re-process data to only include mapped fields
    if ($global:RawData -and $global:Mapping) {
        try {
            $result = Process-Data $global:RawData $global:Mapping
            $global:ProcessedData = $result.Data
            
            # Update the grid
            Update-GridView
            
            $txtLog.Text += "`r`nShowing only mapped fields."
        } catch {
            $txtLog.Text += "`r`nError reprocessing data: $_"
        }
    } else {
        $txtLog.Text += "`r`nNo data loaded to reprocess."
    }
}

# Show all rows
function Show-AllData {
    $txtLog.Text += "`r`nShowing all rows..."
    
    # Reset mapped fields only flag
    $global:ShowOnlyMappedFields = $false
    
    # Make all rows visible
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Visible = $true
    }
    
    # If we have raw data, reprocess to show all fields
    if ($global:RawData -and $global:Mapping) {
        try {
            $result = Process-Data $global:RawData $global:Mapping
            $global:ProcessedData = $result.Data
            
            # Update the grid
            Update-GridView
            
            $txtLog.Text += "`r`nShowing all fields and rows."
        } catch {
            $txtLog.Text += "`r`nError reprocessing data: $_"
        }
    } else {
        $txtLog.Text += "`r`nAll rows are now visible."
    }
}

# Export processed data
function Export-Data {
    if ($null -eq $global:ProcessedData -or $global:ProcessedData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.")
        return
    }
    
    # Check for errors and provide a warning if any exist
    if ($global:ErrorCells.Count -gt 0) {
        $errorCount = $global:ErrorCells.Count
        $warningResult = [System.Windows.Forms.MessageBox]::Show(
            "Warning: Your data contains $errorCount error(s). Do you still want to proceed with the export?",
            "Export Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($warningResult -eq [System.Windows.Forms.DialogResult]::No) {
            $txtLog.Text += "`r`nExport cancelled due to validation errors."
            return
        }
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Processed Data"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:ProcessedData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Data exported successfully to $($saveFileDialog.FileName)")
            $txtLog.Text += "`r`nData exported successfully to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_")
            $txtLog.Text += "`r`nError exporting data: $_"
        }
    }
}

# Export errors
function Export-Errors {
    if ($null -eq $global:LogData -or $global:LogData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No errors to export.")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Error Log"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $errorObjects = @()
            foreach ($error in $global:LogData) {
                $errorObjects += [PSCustomObject]@{ ErrorMessage = $error }
            }
            
            $errorObjects | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Errors exported successfully to $($saveFileDialog.FileName)")
            $txtLog.Text += "`r`nErrors exported successfully to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_")
            $txtLog.Text += "`r`nError exporting errors: $_"
        }
    }
}

# Update grid view - accepts optional parameter to sort by mapping file order
function Update-GridView {
    param(
        [bool]$useMappingOrder = $false  # Default to false for data CSV column order
    )
    
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        # Clear the grid
        $grid.DataSource = $null
        $grid.Rows.Clear()
        $grid.Columns.Clear()
        
        $firstRow = $global:ProcessedData[0]
        $availableColumns = $firstRow.PSObject.Properties.Name
        
        # Determine column order based on provided parameter
        $orderedColumns = @()
        
        if ($useMappingOrder -and $global:MappingColumnOrder.Count -gt 0) {
            # Use mapping order
            $txtLog.AppendText("`r`nUsing mapping file column order")
            
            # First add all columns from mapping that exist in our data
            foreach ($col in $global:MappingColumnOrder) {
                if ($availableColumns -contains $col) {
                    $orderedColumns += $col
                }
            }
            
            # Then add any remaining columns not in mapping
            foreach ($col in $availableColumns) {
                if ($orderedColumns -notcontains $col) {
                    $orderedColumns += $col
                }
            }
        } 
        else {
            # Use original CSV order with field name translation
            $txtLog.AppendText("`r`nUsing data file column order with renamed fields")
            
            # First handle columns from original data that exist in our processed data
            # accounting for renamed fields (via mapping)
            foreach ($origCol in $global:OriginalColumnOrder) {
                $targetCol = $origCol
                
                # Check if this column was renamed via mapping
                if ($global:SourceToNewFieldMap.ContainsKey($origCol)) {
                    $targetCol = $global:SourceToNewFieldMap[$origCol]
                    $txtLog.AppendText("`r`nTranslating column '$origCol' to '$targetCol'")
                }
                
                # Add the column if it exists in processed data
                if ($availableColumns -contains $targetCol -and $orderedColumns -notcontains $targetCol) {
                    $orderedColumns += $targetCol
                }
            }
            
            # Then add any columns not already added
            foreach ($col in $availableColumns) {
                if ($orderedColumns -notcontains $col) {
                    $orderedColumns += $col
                }
            }
        }
        
        # If we somehow don't have ordered columns, fall back to whatever is available
        if ($orderedColumns.Count -eq 0) {
            $orderedColumns = $availableColumns
        }
        
        # Log column order for debugging
        $txtLog.AppendText("`r`nFinal column order: $($orderedColumns -join ', ')")
        
        # Additional debug logging
        $txtLog.AppendText("`r`nOriginal CSV order: $($global:OriginalColumnOrder -join ', ')")
        $txtLog.AppendText("`r`nMapping order: $($global:MappingColumnOrder -join ', ')")
        
        # Debug mapping information with backwards compatible string joining
        $mappings = @()
        foreach ($key in $global:SourceToNewFieldMap.Keys) {
            $mappings += "$key => $($global:SourceToNewFieldMap[$key])"
        }
        $txtLog.AppendText("`r`nField name mappings: " + ($mappings -join ', '))
        
        # Create columns
        foreach ($colName in $orderedColumns) {
            $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $column.HeaderText = $colName
            $column.Name = $colName
            $column.Width = 120
            $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
            $grid.Columns.Add($column)
        }
        
        # Add rows manually with alternating colors
        for ($rowIdx = 0; $rowIdx -lt $global:ProcessedData.Count; $rowIdx++) {
            $rowObj = $global:ProcessedData[$rowIdx]
            
            # Add new row
            $grid.Rows.Add() | Out-Null
            $currentRow = $grid.Rows[$grid.Rows.Count - 1]
            
            # Set alternating row color
            if ($rowIdx % 2 -eq 1) {
                $currentRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
            }
            
            # Add cell values with proper type handling
            for ($colIdx = 0; $colIdx -lt $orderedColumns.Count; $colIdx++) {
                $colName = $orderedColumns[$colIdx]
                
                # Handle null values and ensure proper string conversion
                if ($null -eq $rowObj.$colName) {
                    $currentRow.Cells[$colIdx].Value = [DBNull]::Value
                } else {
                    # Convert all values to strings to avoid formatting issues
                    $currentRow.Cells[$colIdx].Value = "$($rowObj.$colName)"
                }
            }
        }
        
        # Mark error cells
        foreach ($key in $global:ErrorCells.Keys) {
            $parts = $key -split "-"
            if ($parts.Count -eq 2) {
                $rowIndex = [int]$parts[0]
                $colName = $parts[1]
                
                # Find column index
                $colIndex = -1
                for ($i = 0; $i -lt $grid.Columns.Count; $i++) {
                    if ($grid.Columns[$i].Name -eq $colName) {
                        $colIndex = $i
                        break
                    }
                }
                
                # Set cell background to pink if found
                if ($rowIndex -lt $grid.Rows.Count -and $colIndex -ge 0) {
                    $grid.Rows[$rowIndex].Cells[$colIndex].Style.BackColor = [System.Drawing.Color]::LightPink
                }
            }
        }
    }
}

# Sort by Data column order
function Sort-ByDataOrder {
    $txtLog.AppendText("`r`nSorting by original data file column order")
    
    # Update the grid with data file column order
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        Update-GridView $false
    }
}

# Sort by Mapping column order
function Sort-ByMappingOrder {
    $txtLog.AppendText("`r`nSorting by mapping file column order")
    
    # Update the grid with mapping file column order
    if ($global:ProcessedData -and $global:ProcessedData.Count -gt 0) {
        Update-GridView $true
    }
}

# Button click handlers
$btnLoadFiles.Add_Click({ Load-Files })
$btnFilter.Add_Click({ Show-ErrorsOnly })
$btnMappedOnly.Add_Click({ Show-MappedOnly })
$btnReset.Add_Click({ Show-AllData })
$btnSortByData.Add_Click({ Sort-ByDataOrder })
$btnSortByMapping.Add_Click({ Sort-ByMappingOrder })
$btnExport.Add_Click({ Export-Data })
$btnExportErrors.Add_Click({ Export-Errors })

# When the form loads, simple message
$form.Add_Shown({
    $txtLog.AppendText("`r`nApplication ready. Please click 'Load Files' to begin.")
})

# Run the form
[void]$form.ShowDialog()

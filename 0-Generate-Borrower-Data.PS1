# Configurable variables
$NumberOfEntries = 50    # Default number of entries
$ExportFileName = "2-Data-Borrowers.csv"    # Default export filename
$ErrorFields = @("BorrowerId")    # Fields where errors will be introduced
$ErrorRate = 15    # Default error rate percentage (15%)
$ErrorPlusMinus = 5    # Plus/minus randomness for error rate (5%)

# Define arrays for generating random data
$firstNames = @("John", "Jane", "Alice", "Bob", "Eve", "Tom", "Lucy", "Mark", "Daisy", "Sam", "Emma", "James", "Sarah", "Robert", "Catherine")
$lastNames = @("Doe", "Smith", "Johnson", "Brown", "Davis", "Clark", "Green", "White", "Blue", "Black", "Jones", "Taylor", "Wilson", "Harris")
$titles = @("Mr", "Mrs", "Miss", "Dr", "Ms")
$streets = @("Main Rd", "Oak Ave", "Pine Rd", "Elm St", "Maple Dr", "Cedar Ln", "Birch Blvd", "Walnut Way", "Spruce Rd", "Redwood St")
$membershipTypes = @("Regular", "Premium")
$letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

# Function to generate a random date in UK format (dd/MM/yyyy)
function Get-RandomDate {
    param (
        [int]$MinYear,
        [int]$MaxYear
    )
    $year = Get-Random -Minimum $MinYear -Maximum $MaxYear
    $month = Get-Random -Minimum 1 -Maximum 13
    $day = Get-Random -Minimum 1 -Maximum 28  # Keep it simple, avoiding month-end issues
    return "{0:dd/MM/yyyy}" -f (Get-Date -Year $year -Month $month -Day $day)
}

# Function to generate a random BorrowerId in ANNNNAA format (matching ^[A-Za-z][0-9]{4}[A-Za-z]{2}$)
function Get-RandomBorrowerId {
    $letter1 = $letters[(Get-Random -Minimum 0 -Maximum $letters.Length)]
    $number = Get-Random -Minimum 0 -Maximum 10000  # 0000 to 9999
    $letter2 = $letters[(Get-Random -Minimum 0 -Maximum $letters.Length)]
    $letter3 = $letters[(Get-Random -Minimum 0 -Maximum $letters.Length)]
    return "$letter1{0:D4}$letter2$letter3" -f $number  # Ensures 4 digits with leading zeros
}

# Function to generate a random UK phone number (simplified 7-digit format)
function Get-RandomPhone {
    return Get-Random -Minimum 1000000 -Maximum 9999999
}

# Function to introduce an error by adding a random character
function Add-Error($value) {
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*"
    $randomChar = $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)]
    $position = Get-Random -Minimum 0 -Maximum ($value.Length + 1)
    return $value.Substring(0, $position) + $randomChar + $value.Substring($position)
}

# Calculate the actual error rate with plus/minus randomness
$minErrorRate = [math]::Max(0, $ErrorRate - $ErrorPlusMinus)  # Ensure it doesn't go below 0%
$maxErrorRate = [math]::Min(100, $ErrorRate + $ErrorPlusMinus)  # Ensure it doesn't exceed 100%
$script:ActualErrorRate = Get-Random -Minimum $minErrorRate -Maximum ($maxErrorRate + 1)  # +1 because Maximum is exclusive

# Generate the specified number of rows of data
$data = @()
for ($i = 0; $i -lt $NumberOfEntries; $i++) {
    $borrowerId = Get-RandomBorrowerId
    $name = "$($firstNames | Get-Random) $($lastNames | Get-Random)"
    $title = $titles | Get-Random
    $dateOfBirth = Get-RandomDate -MinYear 1950 -MaxYear 2005  # Reasonable age range
    $email = "$($name.Split(' ')[0].ToLower()).$($name.Split(' ')[1].ToLower())@example.com"
    $phone = Get-RandomPhone
    $address = "$((Get-Random -Minimum 1 -Maximum 999)) $($streets | Get-Random)"
    $membershipType = $membershipTypes | Get-Random
    $membershipStartDate = Get-RandomDate -MinYear 2015 -MaxYear 2023  # Recent memberships
    $overdueCount = Get-Random -Minimum 0 -Maximum 5

    # Apply errors to selected fields based on the actual error rate
    $introduceError = (Get-Random -Minimum 1 -Maximum 101) -le $script:ActualErrorRate
    if ($introduceError) {
        foreach ($field in $ErrorFields) {
            switch ($field) {
                "BorrowerId" { $borrowerId = Add-Error $borrowerId }
                "Name" { $name = Add-Error $name }
                "Title" { $title = Add-Error $title }
                "DateOfBirth" { $dateOfBirth = Add-Error $dateOfBirth }
                "Email" { $email = Add-Error $email }
                "Phone" { $phone = Add-Error $phone.ToString() }
                "Address" { $address = Add-Error $address }
                "MembershipType" { $membershipType = Add-Error $membershipType }
                "MembershipStartDate" { $membershipStartDate = Add-Error $membershipStartDate }
                "OverdueCount" { $overdueCount = Add-Error $overdueCount.ToString() }
            }
        }
    }

    $row = [PSCustomObject]@{
        BorrowerId = $borrowerId
        Name = $name
        Title = $title
        DateOfBirth = $dateOfBirth
        Email = $email
        Phone = $phone
        Address = $address
        MembershipType = $membershipType
        MembershipStartDate = $membershipStartDate
        OverdueCount = $overdueCount
    }
    $data += $row
}

# Export to CSV using the specified filename
$data | Export-Csv -Path $ExportFileName -NoTypeInformation

# Display first 10 rows as confirmation and report statistics
Write-Host "Generated $NumberOfEntries rows of borrower data."
Write-Host "Errors introduced to fields $($ErrorFields -join ', ') with an actual error rate of $script:ActualErrorRate% (target $ErrorRate% ± $ErrorPlusMinus%)."
Write-Host "First 10 rows:"
$data | Select-Object -First 10 | Format-Table -AutoSize
Write-Host "Data has been saved to $ExportFileName"

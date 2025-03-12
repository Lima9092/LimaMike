# Configurable variables
$NumberOfEntries = 50    # Default number of entries
$ExportFileName = "2-Data-Books.csv"    # Default export filename
$ErrorFields = @("Barcode")    # Fields where errors will be introduced
$ErrorRate = 15    # Default error rate percentage (15%)
$ErrorPlusMinus = 5    # Plus/minus randomness for error rate (5%)

# Define arrays for generating random data
$titles = @("The", "A", "An", "My", "Our", "Their", "Last", "First", "Hidden", "Secret", "Lost", "Found", "Dark", "Bright", "Silent", "Loud")
$titleNouns = @("Journey", "Tale", "Story", "Adventure", "Quest", "Mystery", "Dream", "Night", "Day", "Life", "Love", "War", "Peace", "Shadow", "Light")
$firstNames = @("John", "Jane", "William", "Elizabeth", "Thomas", "Mary", "James", "Sarah", "Robert", "Emma", "Charles", "Anne", "George", "Catherine")
$lastNames = @("Smith", "Jones", "Brown", "Taylor", "Wilson", "Davis", "Clark", "Harris", "Lewis", "Walker", "Hall", "Allen", "Young", "King")
$genres = @("Fiction", "Dystopian", "Adventure", "Romance", "Historical", "Modernist", "Epic", "Psychological", "Mystery", "Fantasy", "Sci-Fi", "Thriller")
$publishers = @("Penguin Classics", "HarperCollins", "Macmillan", "Oxford University Press", "Bloomsbury", "Faber & Faber", "Hodder & Stoughton", "Pan Macmillan", "Vintage", "Random House")
$languages = @("English", "French", "Spanish", "German", "Russian")
$formats = @("Hardcover", "Paperback", "Leatherbound")
$locations = @("Shelf A", "Shelf B", "Shelf C", "Shelf D", "Shelf E", "Shelf F", "Shelf G", "Shelf H", "Shelf I", "Shelf J")

# Function to generate a random date in UK format (dd/MM/yyyy)
function Get-RandomDate {
    $year = Get-Random -Minimum 1800 -Maximum 2023
    $month = Get-Random -Minimum 1 -Maximum 13
    $day = Get-Random -Minimum 1 -Maximum 28  # Keep it simple, avoiding month-end issues
    return "{0:dd/MM/yyyy}" -f (Get-Date -Year $year -Month $month -Day $day)
}

# Function to generate a random 12-digit barcode as a string
function Get-RandomBarcode {
    $min = 100000000000  # 11 digits + 1 to ensure 12 digits
    $max = 999999999999  # 12 digits
    $barcode = Get-Random -Minimum $min -Maximum $max
    return "{0:D12}" -f $barcode  # Forces a 12-digit string with leading zeros if needed
}

# Function to introduce an error by adding a random character (general version)
function Add-Error($value) {
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*"
    $randomChar = $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)]
    $position = Get-Random -Minimum 0 -Maximum ($value.Length + 1)
    return $value.Substring(0, $position) + $randomChar + $value.Substring($position)
}

# Function to introduce an error for Title field (only alphanumeric or space)
function Add-TitleError($value) {
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 "  # Only alphanumeric and space
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
    $title = "$($titles | Get-Random) $($titleNouns | Get-Random)"
    $author = "$($firstNames | Get-Random) $($lastNames | Get-Random)"
    $genre = $genres | Get-Random
    $barcode = Get-RandomBarcode
    $publisher = $publishers | Get-Random
    $pubDate = Get-RandomDate
    $pages = Get-Random -Minimum 100 -Maximum 1500
    $language = $languages | Get-Random
    $format = $formats | Get-Random
    $location = "$($locations | Get-Random)$((Get-Random -Minimum 1 -Maximum 10))"

    # Apply errors to selected fields based on the actual error rate
    $introduceError = (Get-Random -Minimum 1 -Maximum 101) -le $script:ActualErrorRate
    if ($introduceError) {
        foreach ($field in $ErrorFields) {
            switch ($field) {
                "Title" { $title = Add-TitleError $title }  # Use Title-specific error function
                "Author" { $author = Add-Error $author }
                "Genre" { $genre = Add-Error $genre }
                "Barcode" { $barcode = Add-Error $barcode }
                "Publisher" { $publisher = Add-Error $publisher }
                "PublicationDate" { $pubDate = Add-Error $pubDate }
                "Pages" { $pages = Add-Error $pages.ToString() }
                "Language" { $language = Add-Error $language }
                "Format" { $format = Add-Error $format }
                "Location" { $location = Add-Error $location }
            }
        }
    }

    $row = [PSCustomObject]@{
        Title = $title
        Author = $author
        Genre = $genre
        Barcode = [string]$barcode
        Publisher = $publisher
        PublicationDate = $pubDate
        Pages = $pages
        Language = $language
        Format = $format
        Location = $location
    }
    $data += $row
}

# Export to CSV using the specified filename with proper string handling
$data | ForEach-Object {
    $_ | Select-Object Title, Author, Genre, @{Name='Barcode';Expression={"$($_.Barcode)"}}, Publisher, PublicationDate, Pages, Language, Format, Location
} | Export-Csv -Path $ExportFileName -NoTypeInformation

# Display first 10 rows as confirmation and report statistics
Write-Host "Generated $NumberOfEntries rows of library data."
Write-Host "Barcodes are all 12 digits unless errors are introduced in the Barcode field."
Write-Host "Errors introduced to fields $($ErrorFields -join ', ') with an actual error rate of $script:ActualErrorRate% (target $ErrorRate% ± $ErrorPlusMinus%)."
Write-Host "First 10 rows:"
$data | Select-Object -First 10 | Format-Table -AutoSize
Write-Host "Data has been saved to $ExportFileName"

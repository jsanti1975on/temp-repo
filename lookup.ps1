# Set the root folder to start searching
$rootFolder = "D:\SalesBackups"

# Recursively check each subfolder
Get-ChildItem -Path $rootFolder -Directory -Recurse | ForEach-Object {
    $pdfCount = (Get-ChildItem -Path $_.FullName -Filter *.pdf -File -ErrorAction SilentlyContinue).Count
    if ($pdfCount -gt 3) {
        [PSCustomObject]@{
            FolderPath = $_.FullName
            PDF_Count  = $pdfCount
        }
    }
} | Format-List
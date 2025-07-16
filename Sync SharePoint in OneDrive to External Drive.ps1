Set-ExecutionPolicy -executionpolicy -Unrestricted -Scope Process
# -----------------------------
# CONFIGURATION
# -----------------------------
$excelPath = "C:\Users\<YourUserName>\OndeDrive\Desktop\Jobs list.xlsx"  # <- update this as per your case
$sheetName = "Sheet1" #Name of the sheet

# Update this to the local path of your synced SharePoint folder in OneDrive
$sourceRoot = "C:\Users\<YourUserName>\<Comapny Name>\SharePoint Name"

# Destination path on your external SSD
$destinationRoot = "E:\SharePointJobs"  # <- update this if your SSD has a different drive letter

# -----------------------------
# READ JOB FOLDERS FROM EXCEL
# -----------------------------
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item($sheetName)
$row = 1
$folderNames = @()

while ($sheet.Cells.Item($row, 1).Text -ne "") {
    $folderName = $sheet.Cells.Item($row, 1).Text.Trim()
    $folderNames += $folderName
    $row++
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel

# -----------------------------
# COPY FOLDERS TO EXTERNAL SSD
# -----------------------------
foreach ($folderName in $folderNames) {
    $sourcePath = Join-Path -Path $sourceRoot -ChildPath $folderName
    $destinationPath = Join-Path -Path $destinationRoot -ChildPath $folderName

    if (Test-Path -Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $destinationPath -Recurse -Force
        Write-Output "✅ Copied: $folderName"
    } else {
        Write-Warning "Folder not found: $folderName"
    }
}

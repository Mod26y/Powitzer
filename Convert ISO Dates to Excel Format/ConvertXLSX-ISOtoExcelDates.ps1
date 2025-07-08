Add-Type -AssemblyName System.Windows.Forms

function Select-ExcelFile {
    $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $FileDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
    $FileDialog.Title = "Select Excel File with ISO Dates"
    if ($FileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $FileDialog.FileName
    } else {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit
    }
}

function Convert-IsoDate {
    param (
        [string]$value
    )

    $formats = @(
        "yyyy-MM-ddTHH:mm:ss",
        "yyyy-MM-ddTHH:mm:ssZ",
        "yyyy-MM-ddTHH:mm:ss.fffZ",
        "yyyy-MM-dd"
    )

    foreach ($fmt in $formats) {
        try {
            $dt = [datetime]::ParseExact($value, $fmt, $null)
            return $dt.ToString("MM/dd/yyyy HH:mm:ss")
        } catch {
            continue
        }
    }

    try {
        $dt = [datetime]::Parse($value)
        return $dt.ToString("MM/dd/yyyy HH:mm:ss")
    } catch {
        return $value
    }
}

function Convert-ExcelDatesCOM {
    param (
        [string]$filePath
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($filePath, $null, $false)
        $worksheet = $workbook.Worksheets.Item(1)
        $usedRange = $worksheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        for ($row = 2; $row -le $rowCount; $row++) {
            for ($col = 1; $col -le $colCount; $col++) {
                $cell = $worksheet.Cells.Item($row, $col)
                $value = $cell.Text
                if ($value -match "^\d{4}-\d{2}-\d{2}") {
                    $newValue = Convert-IsoDate -value $value
                    $cell.Value2 = $newValue
                }
            }
        }

        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $dir = Split-Path $filePath
        $name = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
        $newFile = Join-Path $dir "Converted_${timestamp}_$name.xlsx"

        $workbook.SaveAs($newFile)
        Write-Host ""
        Write-Host "Conversion complete. File saved as:" -ForegroundColor Green
        Write-Host $newFile -ForegroundColor Green
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }

        # Ensure all COM objects are fully released
        foreach ($obj in @($worksheet, $workbook, $excel)) {
            if ($obj) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
            }
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# Main execution with error handling
try {
    $excelFile = Select-ExcelFile
    Convert-ExcelDatesCOM -filePath $excelFile
} catch {
    Write-Host ""
    Write-Host "An error occurred during execution:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

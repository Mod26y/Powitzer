Add-Type -AssemblyName System.Windows.Forms

function Select-CsvFile {
    $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $FileDialog.Filter = "CSV files (*.csv)|*.csv"
    $FileDialog.Title = "Select CSV File with ISO Dates"
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

function Convert-CsvDates {
    param (
        [string]$filePath
    )

    $fileName = Split-Path $filePath -Leaf
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputFile = "Converted_${timestamp}_$fileName"
    $outputPath = Join-Path -Path (Split-Path $filePath) -ChildPath $outputFile

    $rows = Import-Csv -Path $filePath

    $convertedRows = @()
    foreach ($row in $rows) {
        $newRow = @{}
        foreach ($key in $row.PSObject.Properties.Name) {
            $value = $row.$key
            if ($value -match "^\d{4}-\d{2}-\d{2}") {
                $newRow[$key] = Convert-IsoDate -value $value
            } else {
                $newRow[$key] = $value
            }
        }
        $convertedRows += New-Object PSObject -Property $newRow
    }

    $convertedRows | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "Conversion complete. File saved as:" -ForegroundColor Green
    Write-Host $outputPath -ForegroundColor Green
}

# Main execution with error handling
try {
    $csvFile = Select-CsvFile
    Convert-CsvDates -filePath $csvFile
} catch {
    Write-Host ""
    Write-Host "An error occurred during execution:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}


<#
DISCLAIMER AND HOLD HARMLESS

This script is provided "AS IS", without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and noninfringement.

Use of this script is at your own risk. The author assumes no responsibility or liability for any loss, damage, or disruption—including but not limited to data loss, financial loss, system failure, or any other consequence—arising out of or in connection with the use or misuse of this script.

By using this script, you agree to hold the author harmless from any and all claims, liabilities, or damages, whether in contract, tort, or otherwise, resulting from its use.

This is best-effort code intended for educational or utility purposes only.
#>

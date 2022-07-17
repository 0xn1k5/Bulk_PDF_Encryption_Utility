
# Loading the Library
Add-Type -Path "$PSScriptRoot\dll_library\BouncyCastle.Crypto.dll"
[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\dll_library\itextsharp.dll")

function PSUsing {
    param (
        [System.IDisposable] $excelSheetNumbernputObject = $(throw "The parameter -inputObject is required."),
        [ScriptBlock] $scriptBlock = $(throw "The parameter -scriptBlock is required.")
    )

    Try {
        &$scriptBlock
    }

    Finally {
        
        if ($excelSheetNumbernputObject.psbase -eq $null) {
        $excelSheetNumbernputObject.Dispose()
        } 
        
        else {
        $excelSheetNumbernputObject.psbase.Dispose()
        }

    }

}

$sourcePath = "$PSScriptRoot\OriginalPDF\"
$destinationPath = "$PSScriptRoot\EncryptedPDF\"
$excelReferenceSheet = "$PSScriptRoot\PasswdList\PDF_Password_List.xlsx"

$startRow=2
$fileNameCol=1
$PasswordCol=2
$excelSheetNumber=1

# Open the excel file having fileName (COLUMN A) & Password (COLUMN B)
$excel=new-object -com excel.application
$excel.DisplayAlerts = $false
$excel.Visible = $false
$workbook=$excel.workbooks.open($excelReferenceSheet)
$worksheet=$workbook.Sheets.Item($excelSheetNumber)
$rowIncrement=0;

# Find the Number of Rows & Columns
$WorksheetRange = $workSheet.UsedRange
$RowCount = $WorksheetRange.Rows.Count
$ColumnCount = $WorksheetRange.Columns.Count
Write-Host "RowCount:" $RowCount
Write-Host "ColumnCount" $ColumnCount

# Start loop to iterate on all the records of file name
for($i=$startRow; $i -le $RowCount; $i++) {
    $filename = $worksheet.Cells.Item($i,$fileNameCol).value2
    $SourceFilePathAndFileName= $sourcePath + $filename
    #Write-Host "File Name: " $filepath

    # Check if the given file exists 
    if (Test-Path -Path $SourceFilePathAndFileName -PathType Leaf){
        Write-Host "Row: " $i " | Input File Found: " $SourceFilePathAndFileName

        # Fetch the password
        $password= $worksheet.Cells.Item($i,$PasswordCol).value2
        # Write-Host "Encryption Password: " $password

        # Destination Path and File Name
        $destinationPathAndFileName = $destinationPath + $filename
        Write-Host "Destination Path: " $destinationPathAndFileName

        # Password Protect the Source Input File
        New-Object PSObject -Property @{Source=$SourceFilePathAndFileName;Destination=$destinationPathAndFileName;Password=$password}
        $file = New-Object System.IO.FileInfo $SourceFilePathAndFileName
        $fileWithPassword = New-Object System.IO.FileInfo $destinationPathAndFileName


        PSUsing ( $fileStreamIn = $file.OpenRead() ) {
            PSUsing ( $fileStreamOut = New-Object System.IO.FileStream($fileWithPassword.FullName,[System.IO.FileMode]::Create,[System.IO.FileAccess]::Write,[System.IO.FileShare]::None) ) {
                PSUsing ( $reader = New-Object iTextSharp.text.pdf.PdfReader $fileStreamIn ) {
                    [iTextSharp.text.pdf.PdfEncryptor]::Encrypt($reader, $fileStreamOut, $true, $password, $password, [iTextSharp.text.pdf.PdfWriter]::ALLOW_PRINTING)
                }
            }
        }

    }

    else{
            Write-Host "[!] Error Occured, File not Found: " $filePath   
    }
}


$excel.Close
$excel.Quit()

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

Remove-Variable -Name excel

function RenameFile($fileName, $newFileName)
{
    if (Test-Path -Path $fileName)
    {
        Rename-Item -Path $fileName -NewName $newFileName;
        if ((Test-Path -Path $newFileName) -eq $false)
        {
            Write-Host "File was not renamed: $fileName";
        }
    }
    else {
        Write-Host "Did not move: $fileName";   
    }
}



$FilePath = "F:\temp\Book1.xlsx";
$PathRoot = "F:\temp\Images";
$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open($FilePath)

$WorkSheet = $Workbook.Worksheets["Sheet1"];

$rowMax = $WorkSheet.UsedRange.Rows.Count;
$Column = for ($i = 2; $i -le $rowMax; $i++) #initialize at 2 to skip header
{
    $imagePath = $WorkSheet.Cells.Item($i, 5).text; #hard coded to column 5.

    $imagePath = [System.IO.Path]::Combine($PathRoot, $imagePath);

    $fileN = [System.IO.Path]::GetFileNameWithoutExtension($imagePath);   
    $fileN = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($imagePath), $fileN);
    RenameFile -fileName $fileN -newFileName $imagePath

    Write-Host("Moved $fileN to $imagePath");
}

$Excel.quit();












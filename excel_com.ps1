param(
  [Parameter(Mandatory=$true)][string]$InputFile,
  [Parameter(Mandatory=$true)][string]$OutputCsv,
  [string]$SheetName="",
  [ValidateSet("open","attach","attach-launch")][string]$Mode="open",
  [int]$WaitSeconds=120,
  [string]$Recalculate="true",
  [string]$CloseAutoOpened="true"
)

$ErrorActionPreference="Stop"
$excel=$null;$book=$null;$sheet=$null;$exportBook=$null
$ownedExcel=$false;$autoOpened=$false

function Normalize-WorkbookPath([string]$PathValue) {
  return [IO.Path]::GetFullPath($PathValue).TrimEnd('\').ToLowerInvariant()
}

function Get-RunningExcel {
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") }
  catch { return $null }
}

function Find-ExactWorkbook($Application,[string]$WantedPath) {
  if(-not $Application){return $null}
  $wanted=Normalize-WorkbookPath $WantedPath
  for($index=1;$index -le [int]$Application.Workbooks.Count;$index++){
    $candidate=$Application.Workbooks.Item($index)
    try {
      if((Normalize-WorkbookPath ([string]$candidate.FullName)) -eq $wanted){return $candidate}
    } catch {}
  }
  return $null
}

function Find-Sheet($Workbook,[string]$WantedName) {
  if(-not $WantedName){return $Workbook.Worksheets.Item(1)}
  for($index=1;$index -le [int]$Workbook.Worksheets.Count;$index++){
    $candidate=$Workbook.Worksheets.Item($index)
    if(([string]$candidate.Name) -eq $WantedName){return $candidate}
  }
  throw "Sheet not found: $WantedName"
}

function Convert-CsvCell($Value) {
  if($null -eq $Value){return '""'}
  if($Value -is [double] -or $Value -is [decimal] -or $Value -is [int]){$text=[Convert]::ToString($Value,[Globalization.CultureInfo]::InvariantCulture)}
  else {$text=[string]$Value}
  return '"'+($text -replace '"','""')+'"'
}

function Export-SheetDirect($Worksheet,[string]$Destination) {
  $used=$Worksheet.UsedRange
  $lastRow=[int]($used.Row+$used.Rows.Count-1)
  $lastColumn=[int]($used.Column+$used.Columns.Count-1)
  $encoding=New-Object Text.UTF8Encoding($true)
  $writer=New-Object IO.StreamWriter($Destination,$false,$encoding)
  try {
    $chunkSize=2000
    for($startRow=1;$startRow -le $lastRow;$startRow+=$chunkSize){
      $endRow=[Math]::Min($startRow+$chunkSize-1,$lastRow)
      $range=$Worksheet.Range($Worksheet.Cells.Item($startRow,1),$Worksheet.Cells.Item($endRow,$lastColumn))
      $values=$range.Value2
      $rowCount=$endRow-$startRow+1
      for($row=1;$row -le $rowCount;$row++){
        $line=New-Object 'System.Collections.Generic.List[string]'
        for($column=1;$column -le $lastColumn;$column++){
          if($rowCount -eq 1 -and $lastColumn -eq 1){$value=$values}else{$value=$values.GetValue($row,$column)}
          $line.Add((Convert-CsvCell $value))
        }
        $writer.WriteLine([string]::Join(',',$line))
      }
      if($range){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($range)}
    }
  } finally {
    $writer.Dispose()
    if($used){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($used)}
  }
}

try {
  $inputPath=[IO.Path]::GetFullPath($InputFile)
  if($Mode -eq "open"){
    $excel=New-Object -ComObject Excel.Application
    $ownedExcel=$true
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
    try{$excel.AutomationSecurity=3}catch{}
    $book=$excel.Workbooks.Open($inputPath,0,$true)
  } else {
    $excel=Get-RunningExcel
    $book=Find-ExactWorkbook $excel $inputPath
    if(-not $book -and $Mode -eq "attach-launch"){
      Start-Process -FilePath $inputPath | Out-Null
      $deadline=(Get-Date).AddSeconds([Math]::Max(15,$WaitSeconds))
      while((Get-Date) -lt $deadline -and -not $book){
        Start-Sleep -Milliseconds 500
        $runningExcel=Get-RunningExcel
        if($runningExcel){$excel=$runningExcel}
        $book=Find-ExactWorkbook $excel $inputPath
      }
      if($book){$autoOpened=$true}
    }
    if(-not $book){throw "The protected workbook was not authorized in Excel. Allow the Excel/NASCA prompt, then upload it again."}
  }

  $sheet=Find-Sheet $book $SheetName
  if($Recalculate -eq "true"){$excel.CalculateFullRebuild()}

  $nativeExported=$false
  try {
    $sheet.Copy()
    $exportBook=$excel.ActiveWorkbook
    $exportBook.SaveAs($OutputCsv,62)
    $nativeExported=(Test-Path $OutputCsv)
  } catch {
    if($exportBook){try{$exportBook.Close($false)}catch{};$exportBook=$null}
  }
  if(-not $nativeExported){Export-SheetDirect $sheet $OutputCsv}
  if(-not (Test-Path $OutputCsv)){throw "Excel did not produce readable data"}
} catch {
  Write-Error $_.Exception.Message
  exit 1
} finally {
  if($exportBook){try{$exportBook.Close($false)}catch{};try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($exportBook)}catch{}}
  if($book -and ($ownedExcel -or ($autoOpened -and $CloseAutoOpened -eq "true"))){try{$book.Close($false)}catch{}}
  if($ownedExcel -and $excel){try{$excel.Quit()}catch{}}
  if($sheet){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)}catch{}}
  if($book){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}catch{}}
  if($excel){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}catch{}}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

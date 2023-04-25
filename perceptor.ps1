Param(
  $csvFile = "Export.csv",
  $out = "Insulintagebuch.xlsx"
)

$values = Import-Csv -Delimiter "," -Path $csvFile
$Excel = New-Object -ComObject excel.application
$Excel.visible = $false

$workbook = $Excel.workbooks.add()
$excel.cells.item(1,1) = "Datum und Uhrzeit"
$excel.cells.item(1,2) = "Blutzucker (mg/dL)"
$excel.cells.item(1,3) = "Basal-Einheiten"
$excel.cells.item(1,4) = "Insulin (Korrektur)"
$excel.cells.item(1,5) = "Mahlzeit"
$excel.cells.item(1,6) = "Aktivität"
$excel.cells.item(1,7) = "Notizen"
$i = 2

foreach($item in $values)
{
    #Write-Output $item;
    $dateStamp = $item.Datum + " " + $item.Zeit;

    $excel.cells.item($i,1) = $dateStamp;
    $excel.cells.item($i,2) = $item.'Blutzuckermessung (mg/dL)';
    $excel.cells.item($i,3) = $item.Basalinjektionseinheiten;
    $excel.cells.item($i,4) = $item.'Insulin (Korrektur)';
    $excel.cells.item($i,5) = $item.Mahlzeitbeschreibung;
    $excel.cells.item($i,6) = $item.'Aktivitätsbeschreibung';
    $excel.cells.item($i,7) = $item.Notiz;
    $i++
} #end foreach process

# Auto-Resize columns
$workbook.Activesheet.Cells.EntireColumn.Autofit();

$workbook.saveas($out)
$Excel.Quit()
Remove-Variable -Name excel
[gc]::collect()
[gc]::WaitForPendingFinalizers() 
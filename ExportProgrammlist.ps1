<#
.DESCRIPTION
  Script to Output a Excel List with all Breaches from haveibeenpwned.com
.INPUTS
  Input Path to folder (without filename)
.OUTPUTS
 Excel List to given path with filename Breaches.xlsx
.NOTES
  Version:        1.0
  Author:         Manuel Hiller
  Creation Date:  25.04.2021
  E-Mail:         manuel.hiller@thak.de
#>


#Get All Programms from Registry
$Programms = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*  | Select-Object DisplayName, DisplayVersion, InstallDate
#Get Computername
$Computername = $env:computername



$ProgrammName = $Programms.DisplayName
$ProgrammVersion = $Programms.DisplayVersion
$Datum = $Programms.InstallDate

#Create Excel Object
$excel = New-Object -ComObject excel.application
$excel.visible = $False
$workbook = $excel.Workbooks.Add()
$diskSpacewksht= $workbook.Worksheets.Item(1)

$diskSpacewksht.Name = 'Program list for '+$Computername
$diskSpacewksht.Cells.Item(2,1) = 'Program list for '+$Computername
$diskSpacewksht.Cells.Item(3,1) = 'Porgram name'
$diskSpacewksht.Cells.Item(3,2) = 'Version'
$diskSpacewksht.Cells.Item(3,3) = 'Modification date'

$diskSpacewksht.Cells.Item(2,8).Font.Size = 18
$diskSpacewksht.Cells.Item(2,8).Font.Bold=$True
$diskSpacewksht.Cells.Item(2,8).Font.Name = "Cambria"
$diskSpacewksht.Cells.Item(2,8).Font.ThemeFont = 1
$diskSpacewksht.Cells.Item(2,8).Font.ThemeColor = 4
$diskSpacewksht.Cells.Item(2,8).Font.ColorIndex = 55
$diskSpacewksht.Cells.Item(2,8).Font.Color = 8210719

$col = 4
$col1 = 4
$col2 = 4
 foreach ($timeVal in $ProgrammName){
             $diskSpacewksht.Cells.Item($col,1) = $timeVal 
             $col++
 }
 foreach ($currentVal in $ProgrammVersion){
           $diskSpacewksht.Cells.Item($col1,2) = $currentVal 
           $col1++
 }
 foreach ($voltVal in $Datum){
          $diskSpacewksht.Cells.Item($col2,3) = $voltVal 
          $col2++
 }


#Output path -> Filename $Computername.xlsx
$filename = "$Computername.xlsx"
[string]$path = Read-Host "Input path to folder where the file should be placed (e.g C:\temp\), please"
if(Test-path -Path $path)
{
Write-Output "Path is correct, writing file to "+$path+""+$filename
$workbook.SaveAs($path+""+$filename)  
$workbook.Close
$excel.DisplayAlerts = 'False'
$excel.Quit()
}
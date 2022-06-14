param([string]$fileName)  
  
$filePath = Join-Path $pwd $fileName  
  
$excel = New-Object -ComObject Excel.Application  
  
$excel.Workbooks.Open($filePath) | % {  
  $_.VBProject.VBComponents | % {  
    $exportFileName = Join-Path $pwd ($_.Name + ".bas")  
    $_.Export($exportFileName)  
  }  
}  
  
$excel.Quit()  


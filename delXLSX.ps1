Install-Module psexcel
Import-Module psexcel

$path = Read-Host "Please input full path to .xlsx file"


$xlsxObj = new-object System.Collections.ArrayList
$dirName = [System.IO.Path]::GetDirectoryName($path)
#Import .xlsx file using module psexcel
foreach ($line in (Import-XLSX -Path $path -RowStart 1)){
    $xlsxObj.add($line) | out-null 
}
#Clear errors before the sript runs, to avoid writing previous errors to log file 
$Error.Clear()
    #Iterate through .xlsx file and remove path if it exists
    for ($x=0;$x -lt $xlsxObj.Count;$x++) { 
        if ([System.IO.File]::Exists("$($xlsxObj.Path[$x])"+"$($xlsxObj.Filename[$x])")){
            $counter = 1
            rm ("$($xlsxObj.Path[$x])"+"$($xlsxObj.Filename[$x])")
            $counter +=1
            
        } else {
            #Append to log file which line in the excel sheet could not be deleted
            "Excel line $($x+1) [ERROR]: $($xlsxObj.Path[$x])"+"$($xlsxObj.Filename[$x]) not found. Could not be deleted"| Out-File $dirName\del-info.log -Append
        }
    
    } 
    #Write to del-info.log
    @("Successfully deleted $counter files")+ (Get-Content $dirName\del-info.log) | Set-Content $dirName\del-info.log

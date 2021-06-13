Install-Module psexcel
Import-Module psexcel

$path = Read-Host "Please input full path to .xlsx file"


$xlsxObj = new-object System.Collections.ArrayList
$dirName = [System.IO.Path]::GetDirectoryName($path)
foreach ($line in (Import-XLSX -Path $path -RowStart 1)){
    $xlsxObj.add($line) | out-null 
}
$Error.Clear()

    for ($x=0;$x -lt $xlsxObj.Count;$x++) { 
        if ([System.IO.File]::Exists("$($xlsxObj.Path[$x])"+"$($xlsxObj.Filename[$x])")){
            $counter = 1
            rm ("$($xlsxObj.Path[$x])"+"$($xlsxObj.Filename[$x])")
            #"$($x+1) File ($($xlsxObj.Path[$x])"+"$($xlsxObj.Filename[$x])) succesfully deleted" | Out-File $dirName\del-info.log -Append
             #[string[]]$successCount = $($xlsxObj.Path[$x])"+"$($xlsxObj.Filename[$x])
            $counter +=1
            
        } else {
            "Excel line $($x+1) [ERROR]: $($xlsxObj.Path[$x])"+"$($xlsxObj.Filename[$x]) not found. Could not be deleted"| Out-File $dirName\del-info.log -Append
        }
    
    } 
    #"Successfully deleted $counter files" | Out-File $dirName\del-info.log -Append
    @("Successfully deleted $counter files")+ (Get-Content $dirName\del-info.log) | Set-Content $dirName\del-info.log
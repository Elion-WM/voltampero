
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
try {
    $workbook = $excel.Workbooks.Open("C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    
    Write-Host "=== VBA Components ==="
    foreach ($comp in $workbook.VBProject.VBComponents) {
        Write-Host "Component: $($comp.Name) (Type: $($comp.Type))"
    }
    
    Write-Host "`n=== VBA References ==="
    foreach ($ref in $workbook.VBProject.References) {
        $status = if ($ref.IsBroken) { "BROKEN" } else { "OK" }
        Write-Host "Reference: $($ref.Name) - $status (Path: $($ref.FullPath))"
    }
} catch {
    Write-Error "Error: $_"
}
# Keep open so I can fix it if needed in next step, or close? 
# Better close to allow safe modification.
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

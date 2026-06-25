$path = "C:\Users\user\Documents\PFE\Simulation\Ahl SOUSS 6-14-2026.xlsm"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
$wb = $excel.Workbooks.Open($path, 0, $true)  # ReadOnly = $true

foreach ($ws in $wb.Worksheets) {
    $used = $ws.UsedRange
    $rows = $used.Rows.Count
    $cols = $used.Columns.Count
    $addr = $used.Address($false, $false)
    Write-Output ("SHEET: '{0}' | Rows={1} Cols={2} | Range={3}" -f $ws.Name, $rows, $cols, $addr)
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

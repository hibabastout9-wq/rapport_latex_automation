$path = "C:\Users\user\Documents\PFE\Simulation\Ahl SOUSS 6-14-2026.xlsm"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3
$wb = $excel.Workbooks.Open($path, 0, $true)

$outDir = "C:\Users\user\Documents\PFE\rapport\Rapp\excel_dump"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

foreach ($ws in $wb.Worksheets) {
    $used = $ws.UsedRange
    $vals = $used.Value2
    $rows = $used.Rows.Count
    $cols = $used.Columns.Count
    $startAddr = $used.Cells.Item(1,1).Address($false, $false)

    $lastRow = 0
    $lastCol = 0
    for ($r = 1; $r -le $rows; $r++) {
        for ($c = 1; $c -le $cols; $c++) {
            $v = if ($rows -eq 1 -or $cols -eq 1) { $vals } else { $vals[$r, $c] }
            if ($null -ne $v -and "$v".Trim() -ne "") {
                if ($r -gt $lastRow) { $lastRow = $r }
                if ($c -gt $lastCol) { $lastCol = $c }
            }
        }
    }

    $safeName = ($ws.Name -replace '[^\w]', '_')
    $outFile = Join-Path $outDir "$safeName.tsv"

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("# Sheet '$($ws.Name)' UsedRange starts at $startAddr ; data extent rows=$lastRow cols=$lastCol")
    for ($r = 1; $r -le $lastRow; $r++) {
        $line = @()
        for ($c = 1; $c -le $lastCol; $c++) {
            $v = if ($rows -eq 1 -or $cols -eq 1) { $vals } else { $vals[$r, $c] }
            if ($null -eq $v) { $line += "" } else { $line += "$v" }
        }
        [void]$sb.AppendLine([string]::Join("`t", $line))
    }
    [System.IO.File]::WriteAllText($outFile, $sb.ToString())
    Write-Output ("Sheet '{0}' -> {1} (dataRows={2}, dataCols={3})" -f $ws.Name, $outFile, $lastRow, $lastCol)
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$ErrorActionPreference = "Stop"
$base = "C:\Users\user\Documents\PFE\rapport\Rapp"
$dump = Join-Path $base "excel_dump"

$lines = Get-Content (Join-Path $dump "Remous.tsv")
$headerIdx = -1
for ($i=0; $i -lt $lines.Count; $i++) {
  if ($lines[$i] -match "Hauteur Mur") { $headerIdx = $i; break }
}
$headerFields = $lines[$headerIdx].Split("`t") | ForEach-Object { $_.Trim() }
$idxX  = [array]::IndexOf($headerFields,"X")
$idxZ  = [array]::IndexOf($headerFields,"Z")
$idxh  = [array]::IndexOf($headerFields,"h")
$idxV  = [array]::IndexOf($headerFields,"V")
$idxFr = [array]::IndexOf($headerFields,"Fr")

$rows = @()
for ($i=$headerIdx+1; $i -lt $lines.Count; $i++) {
  $f = $lines[$i].Split("`t")
  if ($f.Length -le $idxX) { break }
  $xs = $f[$idxX].Trim()
  if ($xs -eq "") { break }
  $rows += [PSCustomObject]@{
    X=[double]$f[$idxX]; Z=[double]$f[$idxZ]; h=[double]$f[$idxh]; V=[double]$f[$idxV]; Fr=[double]$f[$idxFr]
  }
}

Write-Host "Total rows: $($rows.Count)"
$g = 9.81
$lastIdx = $rows.Count - 1
$picks = @(0,13,26,39,52,65,78,91,$lastIdx)
foreach ($idx in $picks) {
  $r = $rows[$idx]
  $E = $r.h + ($r.V*$r.V)/(2*$g)
  Write-Host ("idx={0} X={1:0.000} Z={2:0.000} h={3:0.000} V={4:0.000} Fr={5:0.000} E={6:0.000}" -f $idx,$r.X,$r.Z,$r.h,$r.V,$r.Fr,$E)
}

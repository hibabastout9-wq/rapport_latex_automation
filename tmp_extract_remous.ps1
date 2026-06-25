$ErrorActionPreference = "Stop"
$base = "C:\Users\user\Documents\PFE\rapport\Rapp"
$dump = Join-Path $base "excel_dump"
$dataDir = Join-Path $base "data"

$sheets = @(
  @{file="Remous.tsv";      out="remous_q1000.dat"; name="Q1000"; Q=794.925467002904; PHE=603.067313368364; Zdep=585.32493237924},
  @{file="Remous_10_.tsv";  out="remous_q10.dat";   name="Q10";   Q=160.735240718924; PHE=601.131834982254; Zdep=584.041342670729},
  @{file="Remous_100_.tsv"; out="remous_q100.dat";  name="Q100";  Q=507.535888342577; PHE=602.317849748864; Zdep=587.207819990252}
)

$Xr = 15.8085788190318

$summary = @()

foreach ($s in $sheets) {
  $lines = Get-Content (Join-Path $dump $s.file)
  $headerIdx = -1
  for ($i=0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match "Hauteur Mur") { $headerIdx = $i; break }
  }
  if ($headerIdx -eq -1) { throw "Header not found in $($s.file)" }
  $headerFields = $lines[$headerIdx].Split("`t") | ForEach-Object { $_.Trim() }
  $idxX  = [array]::IndexOf($headerFields,"X")
  $idxZ  = [array]::IndexOf($headerFields,"Z")
  $idxh  = [array]::IndexOf($headerFields,"h")
  $idxV  = [array]::IndexOf($headerFields,"V")
  $idxFr = [array]::IndexOf($headerFields,"Fr")
  $idxZh = [array]::IndexOf($headerFields,"Z+h")
  $idxMur= [array]::IndexOf($headerFields,"Mur")
  $idxHM = [array]::IndexOf($headerFields,"Hauteur Mur")

  $rows = @()
  for ($i=$headerIdx+1; $i -lt $lines.Count; $i++) {
    $f = $lines[$i].Split("`t")
    if ($f.Length -le $idxX) { break }
    $xs = $f[$idxX].Trim()
    if ($xs -eq "") { break }
    $x = [double]$xs
    $z = [double]$f[$idxZ]
    $h = [double]$f[$idxh]
    $v = [double]$f[$idxV]
    $fr = [double]$f[$idxFr]
    $zh = [double]$f[$idxZh]
    $mur = [double]$f[$idxMur]
    $hm = [double]$f[$idxHM]
    $rows += [PSCustomObject]@{X=$x;Z=$z;h=$h;V=$v;Fr=$fr;Zh=$zh;Mur=$mur;HM=$hm}
  }

  # write .dat file (downsample to every Nth row if too many, but keep all here)
  $datLines = New-Object System.Collections.Generic.List[string]
  $datLines.Add("X Z Zh Mur HM")
  foreach ($r in $rows) {
    $datLines.Add( ("{0:0.######} {1:0.######} {2:0.######} {3:0.######} {4:0.######}" -f $r.X,$r.Z,$r.Zh,$r.Mur,$r.HM) )
  }
  Set-Content -Path (Join-Path $dataDir $s.out) -Value $datLines -Encoding ascii

  # row0 (crest)
  $r0 = $rows[0]
  # last row
  $rl = $rows[$rows.Count-1]
  # max HauteurMur
  $maxHMrow = $rows | Sort-Object -Property HM -Descending | Select-Object -First 1

  # interpolate at X = Xr
  $interp = $null
  for ($i=0; $i -lt $rows.Count-1; $i++) {
    $a = $rows[$i]; $b = $rows[$i+1]
    if ( ($a.X -le $Xr -and $b.X -ge $Xr) -or ($a.X -ge $Xr -and $b.X -le $Xr) ) {
      $t = ($Xr - $a.X) / ($b.X - $a.X)
      $interp = [PSCustomObject]@{
        X=$Xr
        Z = $a.Z + $t*($b.Z-$a.Z)
        h = $a.h + $t*($b.h-$a.h)
        V = $a.V + $t*($b.V-$a.V)
        Fr = $a.Fr + $t*($b.Fr-$a.Fr)
        Zh = $a.Zh + $t*($b.Zh-$a.Zh)
        Mur = $a.Mur + $t*($b.Mur-$a.Mur)
        HM = $a.HM + $t*($b.HM-$a.HM)
      }
      break
    }
  }

  $summary += "=== $($s.name) ($($s.file)) ==="
  $summary += "Nb lignes: $($rows.Count)"
  $summary += ("Row0 (crete X=0): Z={0:0.######} h={1:0.######} V={2:0.######} Fr={3:0.######} Zh={4:0.######} Mur={5:0.######} HM={6:0.######}" -f $r0.Z,$r0.h,$r0.V,$r0.Fr,$r0.Zh,$r0.Mur,$r0.HM)
  $summary += ("RowLast (X={0:0.######}): Z={1:0.######} h={2:0.######} V={3:0.######} Fr={4:0.######} Zh={5:0.######} Mur={6:0.######} HM={7:0.######}" -f $rl.X,$rl.Z,$rl.h,$rl.V,$rl.Fr,$rl.Zh,$rl.Mur,$rl.HM)
  $summary += ("MaxHauteurMur: HM={0:0.######} at X={1:0.######} (Z={2:0.######} h={3:0.######} V={4:0.######} Fr={5:0.######})" -f $maxHMrow.HM,$maxHMrow.X,$maxHMrow.Z,$maxHMrow.h,$maxHMrow.V,$maxHMrow.Fr)
  if ($interp) {
    $summary += ("Interp at X=Xr={0:0.######}: Z={1:0.######} h={2:0.######} V={3:0.######} Fr={4:0.######} Zh={5:0.######} Mur={6:0.######} HM={7:0.######}" -f $Xr,$interp.Z,$interp.h,$interp.V,$interp.Fr,$interp.Zh,$interp.Mur,$interp.HM)
  } else {
    $summary += "Interp at X=Xr: OUT OF RANGE (Xmax=$($rl.X))"
  }
  $summary += "PHE=$($s.PHE)  Z_depart_jet=$($s.Zdep)  h_chute(PHE-Zdep)=$($s.PHE - $s.Zdep)"
  $summary += ""
}

$summary | Set-Content -Path (Join-Path $dataDir "remous_summary.txt") -Encoding ascii
Write-Host "Done"

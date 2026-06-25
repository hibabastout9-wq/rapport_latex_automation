$g = 9.81
$thetaDeg = 40.0
$theta = $thetaDeg * [Math]::PI / 180.0
$Zaval = 579.50  # radier du bassin de dissipation (ch4)
$R = 3.0
$L = 70.0

$scenarios = @(
  @{name="T10";  Q=160.735240718924; H=1.13183498225396; y0=0.128376; V0=17.886657; Fr0=15.938694; Z0=583.912967; qu=2.29621772455606; Hveronese=17.0904923115248; D_veronese=5.63729075460636},
  @{name="T100"; Q=507.535888342577; H=2.31784974886398; y0=0.431213; V0=16.814239; Fr0=8.175173;  Z0=586.794237; qu=7.25051269060824; Hveronese=15.1100297586119; D_veronese=10.2020489994409},
  @{name="T1000";Q=794.925467002904; H=3.06731336836435; y0=0.634125; V0=17.908252; Fr0=7.18011;   Z0=585.324932; qu=11.3560781000415; Hveronese=17.7423809891239; D_veronese=13.4773528509363}
)

foreach ($s in $scenarios) {
  $dZ = $s.Z0 - $Zaval
  $V0sin = $s.V0 * [Math]::Sin($theta)
  $V0cos = $s.V0 * [Math]::Cos($theta)
  $Lx = $V0cos * ($V0sin + [Math]::Sqrt($V0sin*$V0sin + 2*$g*$dZ)) / $g
  $Ry0 = $R / $s.y0
  $Leff = $L - 0.4*$s.H
  $C = $s.Q / ($Leff * [Math]::Pow($s.H,1.5))

  Write-Host "=== $($s.name) ==="
  Write-Host ("Q=$($s.Q)  H=$($s.H)  Leff=$Leff  C=$C")
  Write-Host ("y0=$($s.y0)  V0=$($s.V0)  Fr0=$($s.Fr0)  Z0=$($s.Z0)")
  Write-Host ("dZ = Z0 - Zaval = $dZ")
  Write-Host ("R/y0 = $Ry0")
  Write-Host ("Lx = $Lx")
  Write-Host ("q=$($s.qu)  Hveronese=$($s.Hveronese)  D=$($s.D_veronese)")
  Write-Host ""
}

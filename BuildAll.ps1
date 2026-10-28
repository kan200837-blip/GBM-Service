param(
  [Parameter(Mandatory=$true)][string]$Src,   # โฟลเดอร์ที่มี Code.gs และ Index.html เดิมของคุณ
  [string]$Dist = ''                          # โฟลเดอร์ผลลัพธ์ (ค่าเริ่มต้น: $Src\dist)
)
# ประกอบชุดไฟล์รุ่น 7 ทั้งชุด พร้อมวางแทนที่ใน Apps Script
# ใช้: powershell -ExecutionPolicy Bypass -File BuildAll.ps1 -Src "C:\kpi"
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Src = (Resolve-Path $Src).Path
if (-not $Dist) { $Dist = Join-Path $Src 'dist' }
$utf8 = New-Object Text.UTF8Encoding $false

foreach ($f in 'Code.gs','Index.html') { if (-not (Test-Path (Join-Path $Src $f))) { throw "ไม่พบ $f ใน $Src" } }
foreach ($f in 'ApplyPatches.ps1','Report.gs','ExecUpgrade.html','Upgrade7.html','appsscript.json') { if (-not (Test-Path (Join-Path $here $f))) { throw "ไม่พบ $f ข้างสคริปต์นี้ ($here)" } }
New-Item -ItemType Directory -Force $Dist | Out-Null

# 1) Code.gs — แก้ 6 จุด (ใช้ ApplyPatches.ps1)
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here 'ApplyPatches.ps1') -In (Join-Path $Src 'Code.gs') -Out (Join-Path $Dist 'Code.gs')
if ($LASTEXITCODE -ne 0) { throw 'แก้ Code.gs ไม่สำเร็จ (ดูข้อความด้านบน) ไม่สร้างชุดไฟล์' }

# 2) Index.html — ใส่ส่วนเสริมรุ่น 6 และรุ่น 7 ก่อน </body>
$h = [IO.File]::ReadAllText((Join-Path $Src 'Index.html'), [Text.Encoding]::UTF8)
if ($h.Contains('ส่วนเสริมรุ่น 7')) {
  Write-Host '[ข้าม]  Index.html มีส่วนเสริมรุ่น 7 อยู่แล้ว'
} else {
  if ($h.Contains('ส่วนเสริมรุ่น 6')) { throw 'Index.html มีส่วนเสริมรุ่น 6 อยู่แล้วแต่ไม่มีรุ่น 7 — ใช้ Index.html ต้นฉบับ (ก่อนใส่ส่วนเสริม) เป็นอินพุต' }
  $n = ([regex]::Matches($h, '</body>')).Count
  if ($n -ne 1) { throw "พบ </body> $n ที่ (ต้องมี 1 ที่) — ไฟล์ Index.html ไม่ถูกต้อง" }
  $a6 = [IO.File]::ReadAllText((Join-Path $here 'ExecUpgrade.html'), [Text.Encoding]::UTF8)
  $a7 = [IO.File]::ReadAllText((Join-Path $here 'Upgrade7.html'),  [Text.Encoding]::UTF8)
  $a6 = [regex]::Replace($a6, '^\s*<!--.*?-->\s*', '', 'Singleline')
  $a7 = [regex]::Replace($a7, '^\s*<!--.*?-->\s*', '', 'Singleline')
  $a6 = '<!-- ส่วนเสริมรุ่น 6: สรุปผู้บริหาร ค้นหาด่วน แก้ไขตรงจุด -->' + "`n" + $a6
  $a7 = '<!-- ส่วนเสริมรุ่น 7: เข้าสู่ระบบ ผู้บริหาร QR เหตุการณ์ รปภ. รายงานอีเมล -->' + "`n" + $a7
  $i = $h.IndexOf('</body>')
  $h = $h.Substring(0, $i) + $a6 + "`n" + $a7 + "`n" + $h.Substring($i)
  Write-Host '[แก้แล้ว] Index.html ใส่ส่วนเสริมรุ่น 6 + 7'
}
[IO.File]::WriteAllText((Join-Path $Dist 'Index.html'), $h, $utf8)

# 3) ไฟล์ใหม่
Copy-Item (Join-Path $here 'Report.gs') (Join-Path $Dist 'Report.gs') -Force
Copy-Item (Join-Path $here 'appsscript.json') (Join-Path $Dist 'appsscript.json') -Force

# 4) ตรวจผลลัพธ์
$ok = $true
function Chk($name, $cond) { if ($cond) { Write-Host "[ผ่าน]  $name" } else { Write-Host "[ผิด]   $name"; $script:ok = $false } }
$c = [IO.File]::ReadAllText((Join-Path $Dist 'Code.gs'), [Text.Encoding]::UTF8)
$x = [IO.File]::ReadAllText((Join-Path $Dist 'Index.html'), [Text.Encoding]::UTF8)
Chk 'Code.gs: มี ownerOnly_ 4 ที่' (([regex]::Matches($c, 'ownerOnly_\(\);')).Count -eq 4)
Chk 'Code.gs: มีบทบาท exec' ($c.Contains("'viewer', 'exec'"))
Chk 'Index.html: <script> เปิด/ปิดเท่ากัน' (([regex]::Matches($x, '<script')).Count -eq ([regex]::Matches($x, '</script>')).Count)
Chk 'Index.html: มี </body> เดียว' (([regex]::Matches($x, '</body>')).Count -eq 1)
Chk 'Index.html: มีหน้าสรุปผู้บริหาร + เข้าสู่ระบบ' ($x.Contains('rExec') -and $x.Contains('loginDlg'))
$j = Get-Content (Join-Path $Dist 'appsscript.json') -Raw | ConvertFrom-Json
Chk 'appsscript.json: อ่านได้และมีสิทธิ์ส่งอีเมล' ($j.oauthScopes -contains 'https://www.googleapis.com/auth/script.send_mail')
if (-not $ok) { Write-Host ''; Write-Host 'มีรายการที่ไม่ผ่าน อย่าใช้ชุดไฟล์นี้'; exit 1 }
Write-Host ''
Write-Host "สำเร็จ → $Dist"
Get-ChildItem $Dist | ForEach-Object { '  {0,-18} {1,10:N0} bytes' -f $_.Name, $_.Length }
Write-Host 'Seed.gs และ Seed2.gs ไม่ต้องเปลี่ยน (คงของเดิมไว้)'

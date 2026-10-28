param(
  [Parameter(Mandatory=$true)][string]$In,
  [string]$Out = ''
)
# แก้ Code.gs อัตโนมัติ 6 จุด (รุ่น 7 + ปิดช่องโหว่) — ไม่ทับไฟล์เดิม สร้างไฟล์ใหม่ให้
# ใช้: powershell -ExecutionPolicy Bypass -File ApplyPatches.ps1 -In "C:\path\Code.gs"
$ErrorActionPreference = 'Stop'
if (-not $Out) { $Out = Join-Path (Split-Path -Parent (Resolve-Path $In)) 'Code.patched.gs' }
$t = [IO.File]::ReadAllText((Resolve-Path $In), [Text.Encoding]::UTF8)
$nl = if ($t.Contains("`r`n")) { "`r`n" } else { "`n" }
$fail = @()

function Rep([string]$name, [string]$old, [string]$new) {
  if ($script:t.Contains($new)) { Write-Host "[ข้าม]  $name (แก้ไว้แล้ว)"; return }
  $i = $script:t.IndexOf($old)
  if ($i -lt 0) {
    if ($script:t.Contains($new)) { Write-Host "[ข้าม]  $name (แก้ไว้แล้ว)"; return }
    $script:fail += $name; Write-Host "[ไม่พบ] $name"; return
  }
  if ($script:t.IndexOf($old, $i + 1) -ge 0) { $script:fail += $name; Write-Host "[ซ้ำ]   $name (พบมากกว่า 1 ที่ ไม่แก้เพื่อความปลอดภัย)"; return }
  $script:t = $script:t.Substring(0, $i) + $new + $script:t.Substring($i + $old.Length)
  Write-Host "[แก้แล้ว] $name"
}

Rep '1 ROLES เพิ่ม exec' "const ROLES = ['admin', 'editor', 'viewer'];" "const ROLES = ['admin', 'editor', 'viewer', 'exec'];"
Rep '2 currentUser_ ใช้อีเมลจากรหัสเข้าใช้' `
  "function currentUser_() {$nl  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();" `
  "function currentUser_() {$nl  const email = String((typeof CALL_USER_ !== 'undefined' && CALL_USER_ && CALL_USER_.email) || Session.getActiveUser().getEmail() || '').toLowerCase();"
Rep '3 requireEditor_ ผู้บริหารแก้ไม่ได้' "if (u.role === 'viewer') throw new Error('บัญชีนี้ดูได้อย่างเดียว ให้ผู้ดูแลระบบเพิ่มสิทธิ์ในแท็บผู้ดูแลระบบ');" "if (u.role !== 'admin' && u.role !== 'editor') throw new Error('บัญชีนี้ดูได้อย่างเดียว ให้ผู้ดูแลระบบเพิ่มสิทธิ์ในแท็บผู้ดูแลระบบ');"
Rep '4 uploadAttachment' "if (u.role === 'viewer' && !staff) throw new Error('ไม่มีสิทธิ์แนบไฟล์" "if (u.role !== 'admin' && u.role !== 'editor' && !staff) throw new Error('ไม่มีสิทธิ์แนบไฟล์"
Rep '5 getVisit ผู้บริหารดูไฟล์แนบ' "const priv = u.role === 'admin' || u.role === 'editor' || !!staff;" "const priv = u.role === 'admin' || u.role === 'editor' || u.role === 'exec' || !!staff;"
Rep '6a ownerOnly_ setup' "function setup() {$nl" "function setup() {$nl  ownerOnly_();$nl"
Rep '6b ownerOnly_ connectExistingDb' "function connectExistingDb(id) {$nl" "function connectExistingDb(id) {$nl  ownerOnly_();$nl"
Rep '6c ownerOnly_ upgradeDb' "function upgradeDb() {$nl" "function upgradeDb() {$nl  ownerOnly_();$nl"
Rep '6d ownerOnly_ authorizeDrive' "function authorizeDrive() {$nl" "function authorizeDrive() {$nl  ownerOnly_();$nl"

if ($fail.Count) { Write-Host ''; Write-Host ('ไม่สำเร็จ ' + $fail.Count + ' จุด: ' + ($fail -join ' | ')); Write-Host 'ไม่ได้สร้างไฟล์ผลลัพธ์ (ไฟล์ Code.gs ของคุณอาจไม่ใช่รุ่นเดียวกับที่วางมา)'; exit 1 }
[IO.File]::WriteAllText($Out, $t, (New-Object Text.UTF8Encoding $false))
Write-Host ''; Write-Host "สำเร็จ ครบทุกจุด → $Out"

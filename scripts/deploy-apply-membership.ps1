Param(
  [string]$ProjectRef,
  [string]$SupabaseUrl,
  [SecureString]$ServiceRoleKey,
  [SecureString]$RecaptchaSecret
)

function Write-Section($text){ Write-Host "`n=== $text ===" -ForegroundColor Cyan }
function To-Plain([SecureString]$s){
  if (-not $s) { return $null }
  $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)
  try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
  finally { if ($ptr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) } }
}

# 1) Check Supabase CLI
Write-Section "Checking Supabase CLI"
$cli = Get-Command supabase -ErrorAction SilentlyContinue
if (-not $cli){
  Write-Host "Supabase CLI not found. Please install it first:" -ForegroundColor Yellow
  Write-Host " - Windows installer: https://github.com/supabase/cli/releases" -ForegroundColor Yellow
  Write-Host " - Scoop (if installed): scoop install supabase" -ForegroundColor Yellow
  Write-Host " - Chocolatey: choco install supabase" -ForegroundColor Yellow
  exit 1
}

# 2) Prompt for inputs if not provided
if (-not $ProjectRef) { $ProjectRef = Read-Host "Enter Supabase project ref (e.g., abcdefghijklmnop)" }
if (-not $SupabaseUrl) { $SupabaseUrl = Read-Host "Enter Supabase URL (e.g., https://xyzcompany.supabase.co)" }
if (-not $ServiceRoleKey) { $ServiceRoleKey = Read-Host -AsSecureString "Enter SERVICE ROLE KEY (will not echo)" }
if (-not $RecaptchaSecret) { $RecaptchaSecret = Read-Host -AsSecureString "Enter reCAPTCHA SECRET KEY (will not echo)" }

$ServiceRolePlain = To-Plain $ServiceRoleKey
$RecaptchaPlain   = To-Plain $RecaptchaSecret

if ([string]::IsNullOrWhiteSpace($ProjectRef) -or [string]::IsNullOrWhiteSpace($SupabaseUrl) -or [string]::IsNullOrWhiteSpace($ServiceRolePlain) -or [string]::IsNullOrWhiteSpace($RecaptchaPlain)){
  Write-Host "Missing required values. Aborting." -ForegroundColor Red
  exit 1
}

# 3) Login and link project
Write-Section "Logging into Supabase CLI"
Write-Host "A browser window may open for login..."
supabase login
if ($LASTEXITCODE -ne 0){ Write-Host "supabase login failed" -ForegroundColor Red; exit 1 }

Write-Section "Linking project"
supabase link --project-ref $ProjectRef
if ($LASTEXITCODE -ne 0){ Write-Host "supabase link failed" -ForegroundColor Red; exit 1 }

# 4) Set secrets
Write-Section "Setting secrets"
supabase secrets set `
  SUPABASE_URL=$SupabaseUrl `
  SUPABASE_SERVICE_ROLE_KEY=$ServiceRolePlain `
  RECAPTCHA_SECRET_KEY=$RecaptchaPlain
if ($LASTEXITCODE -ne 0){ Write-Host "supabase secrets set failed" -ForegroundColor Red; exit 1 }

# 5) Deploy function
Write-Section "Deploying Edge Function: apply_membership"
supabase functions deploy apply_membership
if ($LASTEXITCODE -ne 0){ Write-Host "Function deploy failed" -ForegroundColor Red; exit 1 }

Write-Section "Done"
Write-Host "Edge Function 'apply_membership' deployed successfully." -ForegroundColor Green
Write-Host "You can now submit the membership form and verify inserts appear under Admin → Üyeler (Onay bekleyenler) and Admin → Mesajlar (Üyelik Başvurusu)."

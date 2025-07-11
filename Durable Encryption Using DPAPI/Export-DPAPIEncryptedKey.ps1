param (
    [Parameter(Mandatory)]
    [ValidatePattern('^[a-zA-Z_][a-zA-Z0-9_]*$')]
    [string]$PlainVarName,

    [string]$OutPath,

    [string]$OutFile
)

# Resolve script directory or fallback to current directory
$ResolvedOutPath = if ($OutPath) {
    $OutPath
} elseif ($PSScriptRoot) {
    $PSScriptRoot
} else {
    Get-Location | Select-Object -ExpandProperty Path
}

# Ensure DPAPI is available at encryption time
try {
    Add-Type -AssemblyName System.Security
} catch {
    throw "Required .NET assembly 'System.Security' could not be loaded. DPAPI is unavailable."
}

# Resolve output path
if (-not $OutFile) {
    $OutFile = Join-Path -Path $ResolvedOutPath -ChildPath "$PlainVarName.ps1"
} elseif (-not [System.IO.Path]::IsPathRooted($OutFile)) {
    $OutFile = Join-Path -Path $ResolvedOutPath -ChildPath $OutFile
}

# Prompt for the API key securely
$SecureApiKey = Read-Host "Enter API Key" -AsSecureString
$BSTR = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureApiKey)
$UnsecureApiKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Encrypt using DPAPI
$apiBytes = [System.Text.Encoding]::UTF8.GetBytes($UnsecureApiKey)
$encrypted = [System.Security.Cryptography.ProtectedData]::Protect(
    $apiBytes,
    $null,
    [System.Security.Cryptography.DataProtectionScope]::CurrentUser
)
$encryptedBase64 = [Convert]::ToBase64String($encrypted)

# Create the decryption header with try/catch for Add-Type
$header = @"
# DPAPI Encrypted API Key Header

try {
    Add-Type -AssemblyName System.Security
} catch {
    throw "Required .NET assembly 'System.Security' could not be loaded. DPAPI is unavailable."
}

`$EncryptedApiKey = "$encryptedBase64"
`$ApiKeyBytes = [Convert]::FromBase64String(`$EncryptedApiKey)
`$${PlainVarName} = [System.Text.Encoding]::UTF8.GetString(
    [System.Security.Cryptography.ProtectedData]::Unprotect(
        `$ApiKeyBytes,
        `$null,
        [System.Security.Cryptography.DataProtectionScope]::CurrentUser
    )
)
"@

# Write to file
$header | Set-Content -Encoding UTF8 -Path $OutFile -Force

# Clean up
[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
Remove-Variable SecureApiKey, UnsecureApiKey, BSTR, apiBytes, encrypted, encryptedBase64, header

param (
    [Parameter(Mandatory)]
    [ValidatePattern('^[a-zA-Z_][a-zA-Z0-9_]*$')]
    [string]$PlainVarName,

    [string]$OutPath,

    [string]$OutFile,

    [ValidateSet('CurrentUser', 'LocalMachine')]
    [string]$Scope = 'CurrentUser'
)

# Resolve script directory or fallback to current directory
$ResolvedOutPath = if ($OutPath) {
    $OutPath
} elseif ($PSScriptRoot) {
    $PSScriptRoot
} else {
    Get-Location | Select-Object -ExpandProperty Path
}

# Ensure DPAPI is available
try {
    Add-Type -AssemblyName System.Security
} catch {
    throw "Required .NET assembly 'System.Security' could not be loaded. DPAPI is unavailable."
}

# Determine output file path
if (-not $OutFile) {
    $OutFile = Join-Path -Path $ResolvedOutPath -ChildPath "$PlainVarName.ps1"
} elseif (-not [System.IO.Path]::IsPathRooted($OutFile)) {
    $OutFile = Join-Path -Path $ResolvedOutPath -ChildPath $OutFile
}

# Securely prompt for API key
$SecureApiKey = Read-Host "Enter API Key" -AsSecureString
$BSTR = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureApiKey)
$UnsecureApiKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Encrypt the API key using selected DPAPI scope
$apiBytes = [System.Text.Encoding]::UTF8.GetBytes($UnsecureApiKey)
$dpapiScope = [System.Security.Cryptography.DataProtectionScope]::$Scope
$encrypted = [System.Security.Cryptography.ProtectedData]::Protect(
    $apiBytes,
    $null,
    $dpapiScope
)
$encryptedBase64 = [Convert]::ToBase64String($encrypted)

# Generate reusable decryption script with proper escaping
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
        [System.Security.Cryptography.DataProtectionScope]::$Scope
    )
)
"@

# Write header to file
$header | Set-Content -Encoding UTF8 -Path $OutFile -Force

# Clean up sensitive variables
[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
Remove-Variable SecureApiKey, UnsecureApiKey, BSTR, apiBytes, encrypted, encryptedBase64, header

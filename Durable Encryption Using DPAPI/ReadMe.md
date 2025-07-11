# DPAPI Encrypted API Key Generator for PowerShell

This project provides a dependency-free way to embed encrypted API keys in PowerShell scripts using Windows DPAPI.

## Overview

This script:

- Prompts securely for an API key (not stored in history or visible on screen)
- Encrypts the key using Windows DPAPI with CurrentUser scope
- Outputs a `.ps1` file containing decryption logic
- Allows setting a custom variable name for the decrypted key
- Ensures only the same Windows user can decrypt the API key

## Usage Example

```powershell
# Run this to create an encrypted key script (e.g., ApiToken.ps1)
.\Export-DPAPIEncryptedKey.ps1 -PlainVarName 'ApiToken'

# Use the generated script in another script
. .\ApiToken.ps1
Invoke-RestMethod -Uri "https://api.example.com" -Headers @{ Authorization = "Bearer $ApiToken" }

# Optionally clear memory
$ApiToken = $null
[System.GC]::Collect()
```

## Why Use This

- No plaintext secrets in source code or config files
- No dependencies beyond built-in .NET and PowerShell
- Avoids accidentally leaking API keys into version control, logs, or environment variables
- Suitable for environments where full secrets management solutions (e.g., Azure Key Vault, HashiCorp Vault) are not available or too complex
- Provides user-specific access control: only the same user who encrypted the key can decrypt it

## Security Model

| Scenario                               | Protection Mechanism                                             |
|----------------------------------------|------------------------------------------------------------------|
| Script theft                           | Encrypted with DPAPI; unusable without the original user context |
| CLI argument or history leakage        | Avoided by using secure prompts (`Read-Host -AsSecureString`)    |
| Decryption by other users              | Blocked by DPAPI user scope                                      |
| Decryption on another machine          | Works only if user profile and DPAPI master key roam             |
| Memory scraping after decryption       | Key is in memory temporarily; clear with `$null` and `GC`        |

**Note**: Once decrypted, the API key is a normal string in memory. Use short-lived sessions and clear the variable after use.

## Validating Use Across Servers

To decrypt the key on multiple machines:

1. Generate the encrypted file on Server A:

   ```powershell
   .\Export-DPAPIEncryptedKey.ps1 -PlainVarName 'ApiToken'
   ```

2. Copy the resulting file (e.g., `ApiToken.ps1`) to Server B.

3. Log into Server B with the **same domain user**.

4. Ensure the environment supports roaming or Credential Roaming:

   - Group Policy:  
     `User Configuration > Administrative Templates > System > Credentials Delegation`  
     Enable: `Allow DPAPI Credential Roaming`

5. Test decryption on Server B:

   ```powershell
   . .\ApiToken.ps1
   Write-Output $ApiToken
   ```

If the key is correctly decrypted, roaming is functional.

## License

MIT License

Use at your own risk. Suitability depends on your environment and threat model.

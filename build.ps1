[CmdletBinding()]
param(
    [switch]$Sign,
    [string]$CertificatePath = $env:GMAILDOWNLOADER_SIGNING_CERT,
    [string]$TimestampUrl = $env:GMAILDOWNLOADER_TIMESTAMP_URL,
    [string]$CertificatePassword = $env:GMAILDOWNLOADER_SIGNING_PASSWORD
)

$ErrorActionPreference = 'Stop'
python -m PyInstaller --noconfirm --clean GmailDownloader.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}

if ($Sign -or $CertificatePath) {
    if (-not $CertificatePath) {
        throw 'Signing requested but GMAILDOWNLOADER_SIGNING_CERT or -CertificatePath is missing.'
    }
    if (-not (Test-Path -LiteralPath $CertificatePath -PathType Leaf)) {
        throw "Signing certificate was not found: $CertificatePath"
    }
    $signtool = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if (-not $signtool) {
        throw 'Signing requested but signtool.exe is not available in PATH.'
    }
    if (-not $TimestampUrl) {
        $TimestampUrl = 'http://timestamp.digicert.com'
    }
    $signArgs = @('sign', '/fd', 'SHA256', '/f', $CertificatePath, '/tr', $TimestampUrl, '/td', 'SHA256')
    if ($CertificatePassword) {
        $signArgs += @('/p', $CertificatePassword)
    }
    $signArgs += (Join-Path $PSScriptRoot 'dist\GmailDownloader.exe')
    & $signtool.Source @signArgs
    if ($LASTEXITCODE -ne 0) {
        throw "signtool failed with exit code $LASTEXITCODE"
    }
    Write-Host 'Signed dist\GmailDownloader.exe with SHA-256 Authenticode.'
}

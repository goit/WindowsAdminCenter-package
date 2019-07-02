#
# Deployment script for installing Windows Admin Center.
#

param (
    $targetPath = $env:SystemRoot
)

. "$PSScriptRoot\func_Get-MsiInformation.ps1"

$installer = 'bin\WindowsAdminCenter1904.1.msi'
$installerInfo = Get-MsiInformation -Path $installer

$productName = $installerInfo.ProductName
$version = $installerInfo.ProductVersion

Write-Host Installing $productName `($version`)

& "msiexec.exe" /i $installer /qn /L*v wac_install.log SME_PORT=443 SSL_CERTIFICATE_OPTION=generate

$errorCode = $lastexitcode

if ($errorCode -eq 0) {
    Write-Host Successful installation.
}
else {
    Write-Host Error code: $errorCode
}

exit $errorCode

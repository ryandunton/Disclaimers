# SetupKeyVaultSecrets.ps1
# Script to set up required secrets in Azure Key Vault for managed identity approach

# Replace these values with the actual values from your application
$sharePointSiteId = "" # Example: e169d789-7cb6-4c39-bb59-1952c90ccf4b
$sharePointListId = "" # Example: b1bc0b8b-d2e3-4527-82b3-9adb4a9891cb

# Key Vault name
$keyVaultName = "" # Example: mycompanydisclaimerkv

# Login to Azure Government
# Comment out this line if you're already logged in
# Connect-AzAccount -Environment AzureUSGovernment

# Set up secrets in Key Vault
Write-Host "Setting up secrets in Key Vault $keyVaultName..."

# Set SharePointSiteId secret
Write-Host "Setting SharePointSiteId secret..."
$secureStringSiteId = ConvertTo-SecureString $sharePointSiteId -AsPlainText -Force
Set-AzKeyVaultSecret -VaultName $keyVaultName -Name "SharePointSiteId" -SecretValue $secureStringSiteId

# Set SharePointListId secret
Write-Host "Setting SharePointListId secret..."
$secureStringListId = ConvertTo-SecureString $sharePointListId -AsPlainText -Force
Set-AzKeyVaultSecret -VaultName $keyVaultName -Name "SharePointListId" -SecretValue $secureStringListId

Write-Host "All secrets have been set up in Key Vault successfully!"
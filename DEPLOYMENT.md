# Deployment Guide

This guide describes how to deploy the Disclaimers Office Add-in to development, test, and production environments.

## Deployment Model

The add-in has two deployable pieces:

1. **Web application**: `SharePointListApi`, which hosts the task pane UI and backend APIs.
2. **Office add-in manifests**:
   - `DisclaimersManifest/Disclaimers.xml` for Word, Excel, and PowerPoint
   - `DisclaimersOutlook/DisclaimersOutlookManifest/DisclaimersOutlook.xml` for Outlook

The manifests must point to the HTTPS URL where the web application is hosted.

## Cloud Endpoint Defaults

This repository is configured for Azure Government by default.

| Service | Azure Government | Commercial Azure |
| --- | --- | --- |
| Azure App Service | `.azurewebsites.us` | `.azurewebsites.net` |
| Azure Key Vault | `.vault.usgovcloudapi.net` | `.vault.azure.net` |
| Microsoft Graph | `https://graph.microsoft.us` | `https://graph.microsoft.com` |
| Microsoft identity platform | `https://login.microsoftonline.us` | `https://login.microsoftonline.com` |

Update endpoints before deploying to a different cloud.

## Prerequisites

- Azure subscription or equivalent hosting environment
- Microsoft 365 admin access
- SharePoint site if using SharePoint-backed disclaimers
- Visual Studio 2022 or another .NET publish workflow
- Azure PowerShell or Azure CLI for resource setup

## 1. Prepare Configuration

Review these files before deployment:

- `SharePointListApi/appsettings.json`
- `SharePointListApi/appsettings.Development.json`
- `SharePointListApi/Program.cs`
- `DisclaimersManifest/Disclaimers.xml`
- `DisclaimersOutlook/DisclaimersOutlookManifest/DisclaimersOutlook.xml`

Replace sample values such as `MyCompany`, `mycompanydisclaimers`, and `mycompanydisclaimerkv` with your organization-specific values if needed.

## 2. Create or Select a Key Vault

Example for Azure Government:

```powershell
$resourceGroup = "YourResourceGroup"
$location = "usgovvirginia"
$keyVaultName = "YourKeyVaultName"

New-AzKeyVault -Name $keyVaultName -ResourceGroupName $resourceGroup -Location $location
```

The SharePoint-backed API expects these secrets:

```text
SharePointSiteId
SharePointListId
```

Use `SetupKeyVaultSecrets.ps1` as a starting point, but never commit real secret values or tenant-specific IDs.

## 3. Configure SharePoint

Create a SharePoint list with the required columns:

| Column | Type | Required |
| --- | --- | --- |
| `Title` | Single line of text | Yes |
| `Text` | Multiple lines of text | Yes |
| `Ver` | Single line of text | Yes |
| `RichText` | Multiple lines of text with rich text enabled | No |
| `Priority` | Number | No |

Store the site ID and list ID in Key Vault.

## 4. Configure Authentication

### Production: Managed Identity

Production deployments should use Managed Identity where possible.

1. Enable a system-assigned managed identity on the App Service.
2. Grant the identity access to Key Vault secrets.
3. Grant the identity the Microsoft Graph permissions required to read the SharePoint list.
4. Set `ASPNETCORE_ENVIRONMENT` to `Production`.

Example Key Vault access policy:

```powershell
$managedIdentityId = (Get-AzWebApp -ResourceGroupName $resourceGroup -Name $webAppName).Identity.PrincipalId
Set-AzKeyVaultAccessPolicy -VaultName $keyVaultName -ObjectId $managedIdentityId -PermissionsToSecrets get,list
```

Microsoft Graph application permissions must be reviewed and granted by an administrator. Use the least privilege suitable for your SharePoint access model.

### Local Development: Optional Client Credential Fallback

Local development does not have an App Service managed identity. If fallback authentication is enabled, store these additional values in Key Vault:

```text
TenantId
ClientId
SharePointSecret
```

This fallback should be limited to development scenarios. Do not store client secrets in source control.

## 5. Deploy the Web Application

Example Azure App Service flow:

```powershell
$resourceGroup = "YourResourceGroup"
$appServicePlan = "YourAppServicePlan"
$webAppName = "YourWebAppName"
$location = "usgovvirginia"

New-AzAppServicePlan -ResourceGroupName $resourceGroup -Name $appServicePlan -Location $location -Tier Standard -NumberofWorkers 1
New-AzWebApp -ResourceGroupName $resourceGroup -Name $webAppName -Location $location -AppServicePlan $appServicePlan
Set-AzWebApp -ResourceGroupName $resourceGroup -Name $webAppName -AssignIdentity $true
```

Then publish `SharePointListApi` from Visual Studio, CI/CD, or your preferred .NET deployment process.

Recommended App Service settings:

| Setting | Value |
| --- | --- |
| `ASPNETCORE_ENVIRONMENT` | `Production` |
| `KeyVault__Url` | `https://YourKeyVaultName.vault.usgovcloudapi.net/` |

## 6. Update Manifest URLs

Update both manifests so all resource URLs point to your deployed web app.

Files:

- `DisclaimersManifest/Disclaimers.xml`
- `DisclaimersOutlook/DisclaimersOutlookManifest/DisclaimersOutlook.xml`

Example:

```xml
<SourceLocation DefaultValue="https://YourWebAppName.azurewebsites.us/Home.html" />
```

Also update icon and function file URLs in the resource sections.

## 7. Validate the Deployment

Before uploading manifests to Microsoft 365, verify these URLs in a browser:

```text
https://YourWebAppName.azurewebsites.us/Home.html
https://YourWebAppName.azurewebsites.us/api/disclaimer
https://YourWebAppName.azurewebsites.us/api/sharepointlist
```

Expected results:

- `Home.html` loads the task pane UI.
- `/api/disclaimer` returns demo JSON.
- `/api/sharepointlist` returns SharePoint-backed JSON or a controlled error if SharePoint configuration is incomplete.

## 8. Deploy the Add-ins in Microsoft 365

1. Open the Microsoft 365 Admin Center: `https://admin.microsoft.com/`.
2. Go to **Settings** > **Integrated apps**.
3. Choose **Upload custom apps**.
4. Upload each manifest separately.
5. Assign the add-ins to the appropriate users or groups.
6. Test in each target Office host.

## Updating an Existing Deployment

For web-only changes:

1. Deploy the updated web application.
2. Confirm static assets and APIs load from the production URL.
3. No manifest upload is required unless manifest URLs, icons, commands, permissions, or metadata changed.

For manifest changes:

1. Increment the manifest version if appropriate.
2. Re-upload the updated manifest in Microsoft 365 Admin Center.
3. Allow time for Office clients to refresh add-in metadata.

## Monitoring

Recommended production monitoring:

- Application Insights or equivalent application telemetry
- HTTP error rate alerts
- Authentication failure alerts
- Key Vault access failure alerts
- Microsoft Graph throttling or permission error alerts

## Troubleshooting

### Add-in Does Not Appear

- Confirm the manifest uploaded successfully.
- Confirm the user is assigned to the add-in.
- Clear the Office add-in cache.
- Verify the manifest XML is valid.

### Task Pane Does Not Load

- Confirm the web app is reachable over HTTPS.
- Confirm manifest URLs point to the deployed web app.
- Check browser developer tools in the Office task pane.
- Confirm TLS certificates are valid.

### SharePoint Data Does Not Load

- Confirm `SharePointSiteId` and `SharePointListId` exist in Key Vault.
- Confirm the managed identity or fallback app registration has appropriate Graph permissions.
- Confirm admin consent has been granted.
- Confirm the SharePoint list has the expected columns.

### Key Vault Access Fails

- Confirm the Key Vault URL is correct for your cloud.
- Confirm the managed identity has `get` and `list` permissions for secrets.
- Check Key Vault firewall and networking rules.
- Review application logs for Azure credential errors.

## Deployment Security Checklist

- [ ] No real secrets committed to the repository
- [ ] `.user`, publish profile, `bin`, and `obj` files excluded from source control
- [ ] HTTPS enforced
- [ ] Production uses Managed Identity where possible
- [ ] Graph permissions reviewed for least privilege
- [ ] Key Vault access restricted to required identities
- [ ] App Service logs and monitoring enabled
- [ ] Manifest URLs reviewed before upload
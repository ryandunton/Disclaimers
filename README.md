# Disclaimers Office Add-in

Disclaimers is a sample Microsoft Office add-in that helps users insert centrally managed disclaimer or boilerplate text into Office documents and messages. It supports Word, Excel, PowerPoint, and Outlook through Office add-in manifests and a small ASP.NET Core Web API.

The sample is designed for organizations that maintain approved response text, legal notices, classification markings, or other reusable snippets in a central source such as a SharePoint list.

## Attribution

This repository is a fork and public-friendly adaptation of [breakpoint7/Disclaimers](https://github.com/breakpoint7/Disclaimers), an Office Add-in sample project. The upstream repository history is preserved in git; the latest upstream commit in this branch is authored by `breakpoint7`. Changes after that point are public-release updates made for this version, including generic company naming, documentation cleanup, configuration guidance, and repository hygiene files.
## Features

- Office task pane add-in for Word, Excel, PowerPoint, and Outlook
- Separate Outlook manifest for Outlook-specific add-in requirements
- ASP.NET Core Web API backend
- Demo API endpoint with hardcoded sample disclaimers
- SharePoint list integration through Microsoft Graph
- Optional rich text disclaimer support
- Word-specific insertion into cursor position, header, footer, or end of document
- Azure Key Vault configuration support
- Azure Managed Identity support for production deployments

## Repository Structure

```text
.
├── DisclaimersManifest/
│   └── Disclaimers.xml
├── DisclaimersOutlook/
│   └── DisclaimersOutlookManifest/
│       └── DisclaimersOutlook.xml
├── SharePointListApi/
│   ├── Controllers/
│   ├── Properties/
│   ├── wwwroot/
│   ├── Program.cs
│   └── appsettings.json
├── DEPLOYMENT.md
├── DEVELOPER_GUIDE.md
├── SECURITY.md
└── SetupKeyVaultSecrets.ps1
```

## Architecture

The add-in is made of two main parts:

- **Office manifests** define which Office hosts are supported, where the task pane loads from, and how ribbon commands appear.
- **Web application** hosts the task pane UI and backend APIs used to retrieve disclaimer content.

Typical flow:

1. An Office app loads the add-in manifest.
2. The manifest opens the task pane from the hosted web application.
3. The task pane calls the backend API.
4. The backend returns disclaimer content from either demo data or a SharePoint list.
5. The user selects one or more disclaimers.
6. Office.js inserts the selected content into the active document, presentation, workbook, or message.

## Cloud Environment

This sample is configured for Azure Government by default:

- Azure Web Apps: `.azurewebsites.us`
- Azure Key Vault: `.vault.usgovcloudapi.net`
- Microsoft Graph: `https://graph.microsoft.us`
- Microsoft identity platform: `https://login.microsoftonline.us`

For commercial Azure, update the relevant endpoints to `.com`, `vault.azure.net`, `https://graph.microsoft.com`, and `https://login.microsoftonline.com`.

## Prerequisites

- Visual Studio 2022 with Office/SharePoint development workload
- .NET SDK compatible with the `SharePointListApi` project
- Microsoft 365 account with access to the target Office apps
- Azure subscription for hosted deployments
- SharePoint site if using the SharePoint-backed disclaimer source
- Microsoft 365 admin access for centralized add-in deployment

## Quick Start: Demo API

Use this path if you want to run the add-in without SharePoint integration first.

1. Open `Disclaimers.sln` in Visual Studio 2022.
2. Restore NuGet packages and build the solution.
3. Configure multiple startup projects:
   - `SharePointListApi`: Start
   - `Disclaimers`: Start
   - `DisclaimersOutlook`: None initially
4. Start debugging.
5. Confirm the web app is available at `https://localhost:7057/Home.html`.
6. Launch an Office host such as Word, Excel, or PowerPoint through the add-in project settings.
7. Use the task pane to retrieve and insert sample disclaimers.

If the Office add-in loads before the web app is ready, wait for the web app to finish starting and use the retry option in Office.

## SharePoint List Setup

To use SharePoint as the disclaimer source, create a SharePoint list with these columns:

| Column | Type | Required | Purpose |
| --- | --- | --- | --- |
| `Title` | Single line of text | Yes | Display name for the disclaimer |
| `Text` | Multiple lines of text | Yes | Plain text disclaimer content |
| `Ver` | Single line of text | Yes | Version or revision label |
| `RichText` | Multiple lines of text with rich text enabled | No | HTML-formatted disclaimer content |
| `Priority` | Number | No | Optional sort order |

After creating the list, store the SharePoint site ID and list ID in Key Vault using `SetupKeyVaultSecrets.ps1` or your normal secret-management process.

## Configuration

Primary configuration lives in:

- `SharePointListApi/appsettings.json`
- `SharePointListApi/appsettings.Development.json`
- Azure App Service application settings
- Azure Key Vault secrets

The default sample values use generic `MyCompany` names. Replace them with values for your own environment before deployment.

Required Key Vault secrets for SharePoint-backed operation:

- `SharePointSiteId`
- `SharePointListId`

Optional local-development fallback secrets if you enable client credential fallback:

- `TenantId`
- `ClientId`
- `SharePointSecret`

Production deployments should prefer Managed Identity instead of client secrets.

## Deployment

See [DEPLOYMENT.md](DEPLOYMENT.md) for full deployment instructions.

High-level production steps:

1. Deploy `SharePointListApi` to a HTTPS-capable hosting environment such as Azure App Service.
2. Configure Key Vault access and required secrets.
3. Enable Managed Identity for the web app where applicable.
4. Update both manifest files to point to the production web app URL.
5. Upload the manifests through Microsoft 365 Admin Center.
6. Assign the add-ins to the intended users or groups.

## Security

Important defaults:

- Do not commit real secrets, client credentials, tenant-specific IDs, publish profiles, or `.user` files.
- Use HTTPS for all add-in resources.
- Prefer Managed Identity in production.
- Keep SharePoint and Microsoft Graph permissions as narrow as practical.
- Review rich text content before rendering or inserting it.

## Development Guide

See [DEVELOPER_GUIDE.md](DEVELOPER_GUIDE.md) for architecture details, extension points, debugging tips, and common development tasks.

## Troubleshooting

Common checks:

- Confirm `SharePointListApi` is running before the add-in loads.
- Confirm manifest URLs match the running web app URL.
- Browse directly to `https://localhost:7057/Home.html` during local development.
- Browse directly to `https://localhost:7057/api/disclaimer` to test the demo endpoint.
- Check Office task pane developer tools for JavaScript errors.
- Clear the Office web add-in cache if stale manifests or scripts keep loading.

Office cache location on Windows:

```text
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef
```

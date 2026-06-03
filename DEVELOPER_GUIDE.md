# Developer Guide

This guide is for developers maintaining or extending the Disclaimers Office Add-in.

## Projects

| Project | Purpose |
| --- | --- |
| `Disclaimers` | Office add-in manifest project for Word, Excel, and PowerPoint |
| `DisclaimersOutlook` | Office add-in manifest project for Outlook |
| `SharePointListApi` | ASP.NET Core web app that hosts the task pane and APIs |

## Important Files

```text
SharePointListApi/
├── Controllers/
│   ├── DisclaimersController.cs
│   └── SharePointListController.cs
├── wwwroot/
│   ├── Home.html
│   ├── Home.js
│   ├── AltHome.html
│   ├── AltHome.js
│   └── Images/
├── Program.cs
├── appsettings.json
└── appsettings.Development.json

DisclaimersManifest/
└── Disclaimers.xml

DisclaimersOutlook/
└── DisclaimersOutlookManifest/
    └── DisclaimersOutlook.xml
```

## Runtime Flow

1. Office loads one of the manifest XML files.
2. The manifest opens `Home.html` or `AltHome.html` in the task pane.
3. Office.js initializes the client script.
4. Client JavaScript calls the backend API.
5. The backend returns disclaimer records.
6. The user selects disclaimers.
7. Office.js inserts the selected content into the active Office item.

## Backend Components

### `DisclaimersController.cs`

Returns demo disclaimer data. This endpoint is useful for validating that the add-in and API are wired together before configuring SharePoint.

Endpoint:

```text
GET /api/disclaimer
```

### `SharePointListController.cs`

Retrieves disclaimer data from a SharePoint list through Microsoft Graph.

Responsibilities:

- Reads SharePoint site and list IDs from Key Vault
- Acquires a Microsoft Graph access token
- Calls the SharePoint list items API
- Maps SharePoint fields to the add-in response model
- Returns disclaimer data to the task pane

Endpoint:

```text
GET /api/sharepointlist
```

## Authentication Model

Production deployments should use Managed Identity where possible.

Local development can optionally use client credential fallback if Managed Identity is unavailable. This fallback requires values such as `TenantId`, `ClientId`, and `SharePointSecret` to be stored in Key Vault, not in source code.

Do not commit tenant IDs, client IDs, client secrets, publish profiles, or `.user` files.

## Configuration

Configuration sources include:

- `appsettings.json`
- `appsettings.Development.json`
- environment variables
- Azure App Service application settings
- Azure Key Vault

The default Key Vault URL is generic:

```text
https://mycompanydisclaimerkv.vault.usgovcloudapi.net/
```

Replace it with your actual Key Vault URL for deployed environments.

## SharePoint Data Contract

The default SharePoint list mapping expects these fields:

| SharePoint field | API property | Required |
| --- | --- | --- |
| `Title` | `Description` | Yes |
| `Text` | `Text` | Yes |
| `Ver` | `Version` | Yes |
| `RichText` | `RichText` | No |
| `Priority` | Used for optional ordering | No |

If you rename list columns, update both the Graph query and the mapping in `SharePointListController.cs`.

## Client Components

### `Home.html` and `Home.js`

Main task pane experience. Handles loading disclaimers, rendering selectable items, and inserting selected content.

Primary responsibilities:

- Detect Office host
- Fetch disclaimer records
- Render checkbox list
- Support plain text and optional rich text
- Insert selected content at the cursor or end of content where supported

### `AltHome.html` and `AltHome.js`

Alternative Word-focused UI with header and footer insertion support.

## Common Development Tasks

### Add a New Disclaimer Field

1. Add the field to the SharePoint list.
2. Add the field to the Graph query in `SharePointListController.cs`.
3. Add the field to the response model.
4. Map the field from `item.fields`.
5. Update `Home.js` or `AltHome.js` to display or use it.

### Enable Priority Ordering

1. Add a `Priority` number column to the SharePoint list.
2. Index the column in SharePoint if required.
3. Set the priority flag in `SharePointListController.cs`.
4. Verify the Graph query returns items in the expected order.

### Add a New Office Host

1. Confirm Office.js supports the insertion behavior you need for that host.
2. Update host detection logic in the client script.
3. Add host-specific insertion behavior if needed.
4. Update the relevant manifest.
5. Test in the target Office application.

## Debugging

### Backend

- Run `SharePointListApi` first.
- Test API endpoints directly in a browser.
- Use Visual Studio breakpoints in controller methods.
- Review application logs for Key Vault, Graph, and authentication errors.

### Task Pane

- Use Office task pane developer tools.
- Check browser console output.
- Check network requests from the task pane.
- Confirm the web app URL in the manifest matches the running server.

### Office Add-in Cache

If Office keeps loading stale assets or manifests, clear the cache:

```text
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef
```

## Testing Checklist

- [ ] `SharePointListApi` builds successfully
- [ ] `GET /api/disclaimer` returns demo data
- [ ] `GET /api/sharepointlist` returns SharePoint data or a controlled configuration error
- [ ] Task pane loads in Word
- [ ] Task pane loads in Excel
- [ ] Task pane loads in PowerPoint
- [ ] Outlook manifest loads in Outlook
- [ ] Plain text insertion works
- [ ] Rich text insertion works where supported
- [ ] Word header/footer insertion works if using `AltHome`
- [ ] Production errors do not expose sensitive details

## Coding Guidelines

- Keep secrets out of source control.
- Keep environment-specific values in configuration.
- Prefer small, host-specific insertion helpers over large conditional blocks when adding new Office behavior.
- Validate and sanitize rich text before rendering or inserting it.
- Use clear error messages for users and detailed logs for developers.
- Avoid broad formatting-only changes in feature pull requests.

## Useful Resources

- [Office Add-ins documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Office JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office)
- [Microsoft Graph documentation](https://learn.microsoft.com/en-us/graph/)
- [SharePoint lists with Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/resources/list)
- [Azure Key Vault documentation](https://learn.microsoft.com/en-us/azure/key-vault/)
- [Managed identities for Azure resources](https://learn.microsoft.com/en-us/entra/identity/managed-identities-azure-resources/overview)
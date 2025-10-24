# Simple PowerBI Deployment Scripts

Optimized, minimal PowerShell scripts for deploying PowerBI reports to OnPrem and Cloud environments.

## Features

### Both Scripts
- **Auto-create folders/workspaces** - Creates target locations if they don't exist
- **Bulk deployment** - Deploys all `.pbix` files from specified folder
- **Report mappings** - Custom paths per report using regex patterns
- **Priority resolution** - `ReportTargetPath` > `TargetFolderPath`/`ReportName` > defaults
- **Overwrite control** - Skip existing reports or force overwrite
- **DataSource configuration** - Automatic datasource updates
- **Permissions** - Set access rights per report or globally
- **Parameters** - Update report parameters per report or globally

### OnPrem Only ([Deploy-OnPrem.ps1](Deploy-OnPrem.ps1))
- Windows Authentication support
- Automatic folder hierarchy creation
- Report Server REST API v2.0

### Cloud Only ([Deploy-Cloud.ps1](Deploy-Cloud.ps1))
- Service Principal authentication
- User authentication fallback
- Gateway binding
- Refresh schedule configuration
- Initial refresh trigger with wait option
- Workspace-level permissions

## Quick Start

### 1. OnPrem Deployment

```powershell
.\Deploy-OnPrem.ps1 -ConfigPath ".\config-sample.json"
```

### 2. Cloud Deployment

```powershell
.\Deploy-Cloud.ps1 -ConfigPath ".\config-sample.json"
```

## Configuration Structure

```json
{
  "Reports": [
    {
      "Type": "OnPrem" | "Cloud",
      "TargetServer": "http://reportserver/Reports",
      "TargetFolderPath": "/PowerBI/Reports",
      "WorkspaceName": "Production Workspace",
      "FolderPath": "./Reports",
      "Overwrite": true,

      "ServicePrincipal": true,
      "TenantId": "tenant-id",
      "ClientId": "client-id",
      "ClientSecret": "secret",

      "DataSources": [...],
      "Permissions": [...],
      "Parameters": [...],
      "RefreshSchedule": {...},
      "ReportMappings": [...]
    }
  ]
}
```

## ReportMappings Priority

### OnPrem
1. **ReportTargetPath** - Full path `/Folder/Report` (highest)
2. **TargetFolderPath** - Folder only, uses file base name
3. **TargetFolderPath** (global) - From report config

### Cloud
1. **ReportTargetPath** - Custom report name (highest)
2. **ReportName** - Explicit name
3. File base name - Automatic from `.pbix`

## Examples

### Example 1: Simple OnPrem Deployment

```json
{
  "Reports": [
    {
      "Type": "OnPrem",
      "TargetServer": "http://reportserver/Reports",
      "TargetFolderPath": "/PowerBI",
      "FolderPath": "./Reports",
      "Overwrite": true
    }
  ]
}
```

Deploys all `.pbix` files from `./Reports` to `/PowerBI/{filename}`

### Example 2: OnPrem with Custom Paths

```json
{
  "Reports": [
    {
      "Type": "OnPrem",
      "TargetServer": "http://reportserver/Reports",
      "TargetFolderPath": "/PowerBI/Reports",
      "FolderPath": "./Reports",
      "Overwrite": true,
      "ReportMappings": [
        {
          "SourcePattern": "SalesDashboard.*\\.pbix",
          "ReportTargetPath": "/PowerBI/Sales/Dashboard"
        },
        {
          "SourcePattern": "Finance.*\\.pbix",
          "TargetFolderPath": "/PowerBI/Finance"
        }
      ]
    }
  ]
}
```

- `SalesDashboard.pbix` → `/PowerBI/Sales/Dashboard` (Priority 1)
- `FinanceReport.pbix` → `/PowerBI/Finance/FinanceReport` (Priority 2)
- Other files → `/PowerBI/Reports/{filename}` (Priority 3)

### Example 3: Cloud with Service Principal

```json
{
  "Reports": [
    {
      "Type": "Cloud",
      "WorkspaceName": "Production",
      "FolderPath": "./Reports",
      "Overwrite": true,
      "ServicePrincipal": true,
      "TenantId": "your-tenant-id",
      "ClientId": "your-client-id",
      "ClientSecret": "your-secret",
      "RefreshSchedule": {
        "Enabled": true,
        "Days": ["Monday", "Wednesday", "Friday"],
        "Times": ["06:00", "18:00"],
        "TimeZone": "Eastern Standard Time",
        "NotifyOnFailure": true,
        "TriggerInitialRefresh": true,
        "WaitForRefresh": true,
        "RefreshTimeout": 30
      }
    }
  ]
}
```

### Example 4: DataSources Configuration

```json
{
  "DataSources": [
    {
      "Name": "SalesDB",
      "Type": "SQL",
      "ConnectionString": "Data Source=sql-server;Initial Catalog=SalesDB",
      "CredentialType": "Windows",
      "UseDefaultCredentials": true
    },
    {
      "Name": "CloudDB",
      "Type": "Sql",
      "Server": "azure-sql.database.windows.net",
      "Database": "CloudDB",
      "GatewayId": "gateway-id",
      "DatasourceId": "datasource-id",
      "Username": "sqluser",
      "Password": "password"
    }
  ]
}
```

### Example 5: Permissions

```json
{
  "Permissions": [
    {
      "Principal": "DOMAIN\\Users",
      "Role": "Browser"
    },
    {
      "Principal": "users@company.com",
      "AccessRight": "Viewer"
    }
  ]
}
```

### Example 6: Parameters

```json
{
  "Parameters": [
    {
      "Name": "Environment",
      "Value": "Production"
    },
    {
      "Name": "StartDate",
      "Value": "2024-01-01"
    }
  ]
}
```

### Example 7: Per-Report Overrides

```json
{
  "Reports": [
    {
      "Type": "Cloud",
      "WorkspaceName": "Production",
      "FolderPath": "./Reports",
      "Overwrite": true,
      "Permissions": [
        {
          "Principal": "allteam@company.com",
          "AccessRight": "Viewer"
        }
      ],
      "ReportMappings": [
        {
          "SourcePattern": "ExecutiveDashboard.*\\.pbix",
          "ReportTargetPath": "Executive Dashboard",
          "Permissions": [
            {
              "Principal": "executives@company.com",
              "AccessRight": "Viewer"
            }
          ]
        }
      ]
    }
  ]
}
```

Report-specific permissions override global permissions.

## Advanced Features

### Multiple Environments in One Config

```json
{
  "Reports": [
    {
      "Type": "OnPrem",
      "TargetServer": "http://dev-server/Reports",
      "TargetFolderPath": "/Dev",
      "FolderPath": "./Reports/Dev"
    },
    {
      "Type": "OnPrem",
      "TargetServer": "http://prod-server/Reports",
      "TargetFolderPath": "/Prod",
      "FolderPath": "./Reports/Prod"
    },
    {
      "Type": "Cloud",
      "WorkspaceName": "Test Workspace",
      "FolderPath": "./Reports/Test"
    },
    {
      "Type": "Cloud",
      "WorkspaceName": "Production Workspace",
      "FolderPath": "./Reports/Prod"
    }
  ]
}
```

Single config file deploys to multiple environments.

### Regex Pattern Matching

```json
{
  "ReportMappings": [
    {
      "SourcePattern": "^Sales.*\\.pbix$",
      "TargetFolderPath": "/Sales"
    },
    {
      "SourcePattern": "^Finance.*\\.pbix$",
      "TargetFolderPath": "/Finance"
    },
    {
      "SourcePattern": "^Executive.*\\.pbix$",
      "ReportTargetPath": "/Executive/Dashboard"
    }
  ]
}
```

Use regex to match multiple files with one mapping.

### Gateway Configuration (Cloud)

```json
{
  "DataSources": [
    {
      "Name": "OnPremSQL",
      "Type": "Sql",
      "Server": "onprem-server",
      "Database": "Database",
      "GatewayId": "your-gateway-id",
      "DatasourceId": "your-datasource-id",
      "Username": "DOMAIN\\user",
      "Password": "password"
    }
  ]
}
```

Automatically binds datasets to on-premises gateways.

## Workflow

### OnPrem Workflow
1. Load configuration
2. For each report configuration where `Type` = "OnPrem"
3. Get all `.pbix` files from `FolderPath`
4. For each file:
   - Match against `ReportMappings` patterns
   - Resolve target path (priority: ReportTargetPath > TargetFolderPath > global)
   - Create folder hierarchy if needed
   - Upload report (create or overwrite)
   - Configure datasources
   - Set permissions
   - Update parameters

### Cloud Workflow
1. Load configuration
2. Import PowerBI module
3. For each report configuration where `Type` = "Cloud"
4. Connect to Power BI Service (Service Principal or User)
5. Get or create workspace
6. Get all `.pbix` files from `FolderPath`
7. For each file:
   - Match against `ReportMappings` patterns
   - Resolve report name (priority: ReportTargetPath > ReportName > filename)
   - Upload report (create or overwrite)
   - Configure datasources and bind to gateway
   - Update parameters
   - Configure refresh schedule
   - Set permissions
   - Trigger initial refresh (optional)

## Requirements

### OnPrem
- PowerShell 5.1+
- Network access to Report Server
- Windows Authentication or credentials

### Cloud
- PowerShell 5.1+
- MicrosoftPowerBIMgmt module
- Service Principal or User credentials
- Power BI Pro/Premium license

## Installation

### Install PowerBI Module (Cloud only)

```powershell
Install-Module -Name MicrosoftPowerBIMgmt -Force -AllowClobber
```

## Troubleshooting

### OnPrem: "Cannot connect to server"
- Verify `TargetServer` URL is correct
- Ensure you have network access
- Check Windows Authentication permissions

### Cloud: "Workspace not found"
- Script will auto-create workspace
- Ensure Service Principal has workspace creation rights
- Verify workspace name spelling

### "Report already exists"
- Set `Overwrite: true` to replace
- Or set `Overwrite: false` to skip

### DataSource configuration fails
- OnPrem: Verify connection strings and credentials
- Cloud: Ensure Gateway ID and Datasource ID are correct
- Check datasource name matches report's datasource

### Permissions not applied
- OnPrem: Ensure principal format is correct (DOMAIN\\User or DOMAIN\\Group)
- Cloud: Use email format for users, group name for groups
- Verify account has admin rights to set permissions

## Performance Tips

1. **Use ReportMappings** - More efficient than deploying all files
2. **Disable initial refresh** - Set `TriggerInitialRefresh: false` for faster deployments
3. **Parallel execution** - Run multiple configs in parallel for different environments
4. **Pre-create folders/workspaces** - Slightly faster than auto-creation

## Security Notes

- **Never commit secrets** to source control
- Use Azure Key Vault or environment variables for credentials
- Service Principal secrets should be rotated regularly
- Use least-privilege permissions

## License

MIT License - Use freely in your projects

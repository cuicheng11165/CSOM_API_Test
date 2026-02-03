# SharePoint CSOM and Graph API Demo

This repository contains standalone C# console applications demonstrating SharePoint Online CSOM and Microsoft Graph API scenarios. The solution has been reorganized with a unified testing approach for CSOM operations and centralized configuration management.

## 🎯 Solution Overview

This solution provides production-ready examples for:
- **SharePoint CSOM API**: Comprehensive scenarios for file operations, permissions, taxonomy, views, and more
- **Microsoft Graph API**: Token generation and REST API interactions
- **Authentication**: Certificate-based app-only and delegated token flows
- **Configuration**: Centralized management of tenant credentials and tokens

## 📁 Solution Structure

### Core Projects

#### **UnifiedCsomTests** 🌟
An interactive console application with a unified menu system for testing all SharePoint CSOM API scenarios:

- **File Operations**: Multiple upload methods (bytes, stream, chunked upload), metadata management
- **Permission Management**: Group creation, role definitions, effective permissions
- **Tenant API**: Container management, hub sites, site properties
- **List Operations**: CAML queries (basic, paginated), list item operations
- **View Management**: View and view field testing, ETag retrieval
- **Exception Handling**: ExceptionHandlingScope patterns for batch operations
- **Taxonomy**: Managed metadata (groups, term sets, terms, fields)
- **Web Properties**: AllProperties management, DenyAddAndCustomizePages configuration
- **Time Zone Testing**: DateTime timezone behavior validation
- **Container Operations**: Application container management
- **Compliance**: Bulk compliance tag operations
- **Update Conflicts**: Optimistic concurrency testing

#### **Configuration**
Shared library providing centralized configuration management via `CSOM.Common.EnvConfig`:

```csharp
using CSOM.Common;

// Get site URLs
string siteUrl = EnvConfig.GetSiteUrl("/sites/mysite");
string adminUrl = EnvConfig.GetAdminCenterUrl();

// Get tokens
string csomToken = EnvConfig.GetCsomToken();
```

#### **RestApiTest**
Demonstrates SharePoint and Graph REST API usage:
- Shared link creation
- Loop container APIs
- REST-based operations

### Token Generation Projects

#### **ExportCsomToken**
Generates SharePoint CSOM bearer tokens using certificate-based authentication. Run this **first** to create `CSOMAuthorization.txt`.

#### **ExportGraphToken**
Generates Microsoft Graph application-only access tokens using certificate authentication.

#### **ExportGraphDelegateToken**
Generates Microsoft Graph delegated access tokens using the Resource Owner Password Credentials (ROPC) flow.
> ⚠️ **Note**: ROPC flow is deprecated and not recommended for production use.

## 🚀 Getting Started

### Prerequisites

- **.NET 8.0 SDK** or later
- **Azure AD App Registration** with:
  - SharePoint permissions (Sites.FullControl.All or Sites.Selected)
  - Microsoft Graph permissions (as needed)
- **Certificate (.pfx)** for app-only authentication
- **SharePoint Online** tenant

### Initial Setup

1. **Create Configuration Directory**
   
   Create a `Config/` folder at the solution root:
   ```
   C:\Drive\SharePointCSOMApiDemo\Config\
   ```

2. **Add Configuration Files**
   
   Place the following files in the `Config/` directory:

   | File | Description | Example |
   |------|-------------|---------|
   | `HostName.txt` | SharePoint hostname | `contoso.sharepoint.com` |
   | `TenantId.txt` | Azure AD tenant ID | `12345678-1234-1234-1234-123456789abc` |
   | `ClientId.txt` | App registration client ID | `87654321-4321-4321-4321-cba987654321` |
   | `Certificate.pfx` | Authentication certificate | (binary file) |
   | `CertificatePassword.txt` | Certificate password | `YourPassword123` |
   | `UserName.txt` | (Optional) User principal name | `user@contoso.com` |
   | `Password.txt` | (Optional) User password | `UserPassword123` |

3. **Generate CSOM Token**
   
   ```bash
   cd ExportCsomToken
   dotnet run
   ```
   
   This creates `Config/CSOMAuthorization.txt` with a Bearer token.

4. **Build the Solution**
   
   ```bash
   dotnet build
   ```

### Running Tests

#### Option 1: Interactive Menu (Recommended)

```bash
cd UnifiedCsomTests
dotnet run
```


#### Option 2: Direct Execution

Run individual token generators or REST API tests:

```bash
# Generate Graph token
cd ExportGraphToken
dotnet run

# Test REST APIs
cd RestApiTest
dotnet run
```

## 🔐 Authentication Pattern

All CSOM projects use a consistent authentication pattern:

```csharp
using (var context = new ClientContext(siteUrl))
{
    context.ExecutingWebRequest += (object? sender, WebRequestEventArgs e) =>
    {
        e.WebRequestExecutor.WebRequest.Headers[System.Net.HttpRequestHeader.Authorization] =
            EnvConfig.GetCsomToken();
    };
    
    // Your CSOM operations here
    context.ExecuteQuery();
}
```

## 📚 Key Features

### Centralized Configuration
- Single source of truth for all credentials
- Environment-based configuration with `EnvConfig`
- Secure file-based storage for tokens and certificates

### Token Reuse Pattern
1. Generate token once with `ExportCsomToken`
2. Token stored in `Config/CSOMAuthorization.txt`
3. All other projects read and reuse the token
4. Reduces authentication calls and improves performance

### Interactive Testing
- Menu-driven interface in `UnifiedCsomTests`
- No hardcoded URLs or credentials
- Runtime input for test parameters
- Exception handling with detailed error messages

### Comprehensive CSOM Coverage
Demonstrates nearly all major CSOM APIs:
- Client Object Model fundamentals
- File and folder operations (7 upload methods)
- Permission and security management
- Tenant administration
- Managed metadata/taxonomy
- Views and queries
- Exception handling scopes

## 🛠️ Development Notes

### Adding New Scenarios

1. Create a new scenario class in `UnifiedCsomTests/Scenarios/`:
   ```csharp
   namespace UnifiedCsomTests.Scenarios
   {
       internal static class MyScenarios
       {
           internal static void MyMethod()
           {
               var siteUrl = EnvConfig.GetSiteUrl("/sites/test");
               // Implementation
           }
       }
   }
   ```

2. Add menu option in `UnifiedCsomTests/Program.cs`

### Configuration Access

Always use `EnvConfig` for configuration:
```csharp
using CSOM.Common;

string host = EnvConfig.HostName;
string siteUrl = EnvConfig.GetSiteUrl("/sites/relative/path");
string adminUrl = EnvConfig.GetAdminCenterUrl();
string token = EnvConfig.GetCsomToken();
```

### Debugging Tips

- Check `Config/CSOMAuthorization.txt` exists and contains a Bearer token
- Verify certificate password is correct
- Ensure app registration has proper SharePoint permissions
- Review exception messages for API-specific errors

## 📖 Documentation

- [UnifiedCsomTests README](UnifiedCsomTests/README.md) - Detailed usage guide
- [UnifiedCsomTests MIGRATION](UnifiedCsomTests/MIGRATION.md) - Migration history
- [Configuration README](Configuration/README.md) - Configuration project details
- [.github/copilot-instructions.md](.github/copilot-instructions.md) - Project conventions

## 🤝 Contributing

This is a demo repository for learning and reference. Key conventions:

- Use `EnvConfig` for all configuration access
- Follow the bearer token pattern for CSOM authentication
- Add XML comments to public methods
- Keep scenarios focused and self-contained
- Include interactive prompts for runtime parameters

## 📝 License

This project is provided as-is for educational and demonstration purposes.

## 🔗 Related Resources

- [SharePoint CSOM Documentation](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-client-library-code)
- [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview)
- [MSAL.NET Authentication](https://learn.microsoft.com/en-us/entra/msal/dotnet/)

---

**Last Updated**: January 28, 2026  
**Target Framework**: .NET 8.0  
**CSOM Version**: Microsoft.SharePointOnline.CSOM 16.1.26712.12000
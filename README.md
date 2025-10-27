# VCATechManager

A PowerShell-based IT management tool for VCA hospital networks.

## Security Notice

This tool requires a configuration file (`Private/config.json`) containing sensitive internal infrastructure information. The config file is **NOT included** in the repository for security reasons.

## Initial Setup

1. **Download the tool** from the repository or provided zip file
2. **Create the configuration file** at `Private/config.json` with your organization's settings
3. **Distribute the config file securely** (see below)

## Configuration File Setup

Create `Private/config.json` with the following structure:

```json
{
    "InternalDomains": {
        "PrimaryDomain": "yourdomain.local",
        "TrustedSitesDomain": "yourcompany.com"
    },
    "SecuritySettings": {
        "RequireDomainJoin": true,
        "CredentialCacheMinutes": 10
    },
    "NetworkSettings": {
        "PrimaryDHCPServer": "your-dhcp-server",
        "DHCPServers": ["dhcp1", "dhcp2"]
    },
    "ServerNaming": {
        "DHCPServerPattern": "your-dhcp-pattern"
    },
    "FilePaths": {
        "HospitalMasterPath": "/sites/your-site/regions/Documents/HOSPITALMASTER.xlsx",
        "SharePointBaseUrl": "https://yourcompany.sharepoint.com"
    }
}
```

## Secure Distribution

**DO NOT include `Private/config.json` in the main distribution zip file.**

Instead, distribute the config file separately through secure channels:

- Encrypted email
- Secure file share
- Password-protected archive
- Direct secure transfer

## Testing Mode

If `Private/config.json` is missing, the script will run in testing mode with default placeholder values. This allows basic functionality testing but will not work with real infrastructure.

## Requirements

- PowerShell 5.1 or higher
- Domain-joined machine (if RequireDomainJoin is true)
- Active Directory access
- Appropriate network permissions

## Features

- User session management
- Device connectivity checks
- Error logging and reporting
- Infrastructure monitoring
- DHCP reservation management
- SharePoint integration

## Support

For issues or questions, contact your IT administrator.
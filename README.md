# LGPO Tool and Group Policy Files - Technical Reference

## Overview

This document explains how the Microsoft LGPO (Local Group Policy Object) tool applies GPO backup settings, the relationship between registry.pol, ADMX/ADML files, and how local policy is updated.

---

## Files in This GPO Package

### LGPO Tool
| File | Description |
|------|-------------|
| `LGPO.exe` | Microsoft's command-line utility for managing Local Group Policy (v3.0) |
| `LGPO.pdf` | Official documentation/manual |

### Chrome ADMX/ADML Template Files
| File | Description |
|------|-------------|
| `chrome.admx` | Policy definitions file - defines structure, registry paths, data types (694 policies, 39 categories) |
| `chrome.adml` | Language file - provides display names and explanatory text (1,669 strings) |

### GPO Backup: `{EAB6D3D1-B58D-4AA4-8A31-656AB8B1E52A}`

This is a **DoD Google Chrome STIG Computer v2r11** GPO backup containing:

| File | Purpose |
|------|---------|
| `Backup.xml` | GPO structure metadata, lists files to import, references Group Policy extensions |
| `bkupInfo.xml` | Backup metadata (timestamp, GPO display name, source domain) |
| `gpreport.xml` | Human-readable report of all GPO settings |
| `DomainSysvol\GPO\Machine\registry.pol` | **Binary file containing registry-based policy settings** |
| `DomainSysvol\GPO\Machine\comment.cmtx` | Policy comments/annotations |
| `DomainSysvol\GPO\Machine\Microsoft\Windows NT\SecEdit\GptTmpl.inf` | Security template settings |

---

## How LGPO Applies the GPO Backup

### Command to Apply

```cmd
LGPO.exe /g "GPO\{EAB6D3D1-B58D-4AA4-8A31-656AB8B1E52A}"
```

Or with verbose logging:
```cmd
LGPO.exe /g "GPO\{EAB6D3D1-B58D-4AA4-8A31-656AB8B1E52A}" /v > lgpo.out 2> lgpo.err
```

### What LGPO Does When You Run `/g`

1. **Searches the backup directory** for:
   - `registry.pol` files (in `Machine` or `User` subdirectories)
   - `GptTmpl.inf` security templates
   - `audit.csv` advanced auditing backups
   - `backup.xml` for CSE (Client Side Extension) references

2. **Imports registry.pol** into Local Group Policy:
   - Parses the binary registry.pol file
   - Writes settings to `C:\Windows\System32\GroupPolicy\Machine\registry.pol`
   - Each entry specifies: registry key, value name, data type, and value

3. **Applies security templates** (GptTmpl.inf):
   - Uses `secedit.exe` internally to apply security settings
   - Includes user rights assignments, security options, etc.

4. **Registers CSEs** from backup.xml:
   - Enables required Group Policy client-side extensions
   - Example: `{35378EAC-683F-11D2-A89A-00C04FBBCFA2}` for Registry processing

---

## The Registry.pol File Format

### What It Is
A **binary file** that stores registry-based Group Policy settings. It's the core mechanism for Administrative Template policies.

### Structure
Each entry in registry.pol contains:
- Registry key path (e.g., `Software\Policies\Google\Chrome`)
- Value name (e.g., `PasswordManagerEnabled`)
- Data type (DWORD, SZ, etc.)
- Value data (e.g., `0` or `1`)

### Settings in This GPO Backup

The registry.pol in this backup contains **44 Chrome policy settings**, including:

| Setting | Registry Value | Type | Value | Effect |
|---------|---------------|------|-------|--------|
| `RemoteAccessHostFirewallTraversal` | DWORD | 0 | Disable Chrome Remote Desktop firewall traversal |
| `PasswordManagerEnabled` | DWORD | 0 | Disable password manager |
| `SyncDisabled` | DWORD | 1 | Disable Chrome Sync |
| `AutoplayAllowed` | DWORD | 0 | Block autoplay (except .mil/.gov) |
| `DefaultCookiesSetting` | DWORD | 4 | Block third-party cookies |
| `SafeBrowsingProtectionLevel` | DWORD | 1 | Enable Safe Browsing |
| `ExtensionInstallBlocklist` | SZ | * | Block all extensions by default |
| `URLBlocklist` | SZ | javascript://* | Block javascript: URLs |

### Viewing registry.pol Contents

Use LGPO to parse and view:
```cmd
LGPO.exe /parse /m "GPO\{EAB6D3D1-B58D-4AA4-8A31-656AB8B1E52A}\DomainSysvol\GPO\Machine\registry.pol"
```

---

## Relationship: registry.pol, ADMX, and ADML

### How They Connect

```
                          +------------------+
                          |  Group Policy    |
                          |  Editor (gpedit) |
                          +--------+---------+
                                   |
              +--------------------+--------------------+
              |                    |                    |
              v                    v                    v
      +-------+-------+    +-------+-------+    +------+------+
      |   chrome.admx  |    |  chrome.adml  |    | registry.pol|
      +---------------+    +---------------+    +-------------+
      | - Policy names |    | - Display text|    | - Actual    |
      | - Registry keys|    | - Descriptions|    |   values    |
      | - Value types  |    | - Help strings|    | - Applied   |
      | - Categories   |    | - Translations|    |   settings  |
      +---------------+    +---------------+    +-------------+
              |                    |                    |
              +--------------------+--------------------+
                                   |
                                   v
                          +--------+--------+
                          |    Registry     |
                          | HKLM\Software\  |
                          | Policies\Google |
                          |    \Chrome      |
                          +-----------------+
```

### Detailed Relationship

1. **ADMX (chrome.admx)** - The "Schema"
   - Defines the **structure** of available policies
   - Specifies the **registry path** for each setting
   - Defines **data types** (DWORD, SZ, etc.)
   - Contains references to ADML strings: `$(string.PolicyName)`
   - Example:
     ```xml
     <policy name="PasswordManagerEnabled"
             key="Software\Policies\Google\Chrome"
             valueName="PasswordManagerEnabled">
     ```

2. **ADML (chrome.adml)** - The "Language Pack"
   - Provides **human-readable text** for the Group Policy Editor UI
   - Contains **display names** and **explanatory text**
   - Can have multiple versions for different languages
   - Example:
     ```xml
     <string id="PasswordManagerEnabled">Enable saving passwords to the password manager</string>
     <string id="PasswordManagerEnabled_Explain">If you enable this setting, users can...</string>
     ```

3. **Registry.pol** - The "Configuration Data"
   - Contains the **actual configured values**
   - Binary format, not human-readable
   - Created/modified when you configure policies in gpedit.msc
   - Applied by the Group Policy engine at startup/login/refresh

### Key Insight

> **ADMX/ADML files do NOT store policy values.** They only define what policies exist and how they appear in the UI. The actual configured settings are stored in registry.pol.

---

## How Local Policy Gets Updated

### When LGPO Applies Settings

1. **LGPO reads** the backup's registry.pol
2. **Writes to** `C:\Windows\System32\GroupPolicy\Machine\registry.pol`
3. **Triggers** Group Policy refresh (or you can run `gpupdate /force`)

### What Happens at Policy Refresh

1. **Group Policy Engine** reads the local registry.pol
2. **Applies settings** to the registry under `HKLM\Software\Policies\Google\Chrome`
3. **Chrome reads** these registry keys on startup
4. **Settings take effect** in the browser

### Where Settings Appear in Group Policy Editor

After applying with LGPO:

1. Open `gpedit.msc`
2. Navigate to: `Computer Configuration > Administrative Templates > Google > Google Chrome`
3. **IMPORTANT**: The ADMX/ADML files must be installed for settings to display properly

### Installing ADMX/ADML for Group Policy Editor

For the Group Policy Editor to show Chrome policies:

```
Copy chrome.admx to:  C:\Windows\PolicyDefinitions\
Copy chrome.adml to:  C:\Windows\PolicyDefinitions\en-US\
```

Without these files:
- Settings **still apply** (registry values exist)
- But gpedit.msc shows them as "Extra Registry Settings" instead of named policies

---

## Complete Workflow Diagram

```
[GPO Backup]                          [Local Machine]
     |                                      |
     v                                      v
+----+----+                         +-------+-------+
| Backup  |                         | PolicyDefs   |
|  .xml   |--- Metadata ----------->| (empty until |
+---------+                         |  ADMX added) |
                                    +-------+-------+
+----+----+                                 |
|registry |                                 |
|  .pol   |--- LGPO.exe /g ----+            |
+---------+    imports         |            |
                               v            |
                      +--------+--------+   |
                      | C:\Windows\     |   |
                      | System32\       |   |
                      | GroupPolicy\    |   |
                      | Machine\        |   |
                      | registry.pol    |   |
                      +--------+--------+   |
                               |            |
                               v            |
                      +--------+--------+   |
                      | gpupdate /force |   |
                      | (policy refresh)|   |
                      +--------+--------+   |
                               |            |
                               v            v
                      +--------+------------+--------+
                      |        Registry              |
                      | HKLM\Software\Policies\      |
                      |        Google\Chrome         |
                      +--------+---------------------+
                               |
                               v
                      +--------+--------+
                      |   Google Chrome |
                      | reads policies  |
                      | on startup      |
                      +-----------------+
```

---

## LGPO Command Reference

### Apply a GPO Backup
```cmd
LGPO.exe /g <path-to-backup-folder>
```

### Parse/View registry.pol Contents
```cmd
LGPO.exe /parse /m <path>\registry.pol    # Machine policy
LGPO.exe /parse /u <path>\registry.pol    # User policy
```

### Apply Individual Files
```cmd
LGPO.exe /m <path>\registry.pol           # Machine registry.pol
LGPO.exe /u <path>\registry.pol           # User registry.pol
LGPO.exe /s <path>\GptTmpl.inf            # Security template
```

### Export Current Local Policy
```cmd
LGPO.exe /b <output-path> /n "My Policy Backup"
```

---

## Important Notes

1. **Administrative Rights Required**: LGPO.exe requires elevation

2. **LGPO Does NOT Clear Existing Settings**: It adds/modifies settings but doesn't remove pre-existing ones unless explicitly specified (DELETEALLVALUES, DELETE actions)

3. **Chrome Must Be Restarted**: After policy changes, Chrome needs to restart to read new registry values

4. **Verify Application**: Check registry at `HKLM\Software\Policies\Google\Chrome` or use `chrome://policy` in the browser

5. **Group Policy Preferences NOT Supported**: LGPO doesn't handle GPP (Preferences)

---

## Troubleshooting

### Settings Don't Appear in gpedit.msc
- Ensure chrome.admx/adml are copied to `C:\Windows\PolicyDefinitions\`

### Settings Not Taking Effect in Chrome
1. Verify registry values exist: `reg query "HKLM\Software\Policies\Google\Chrome"`
2. Run `gpupdate /force`
3. Restart Chrome
4. Check `chrome://policy` for applied policies

### LGPO Errors
- Run with `/v` flag for verbose output
- Redirect stderr to see error messages: `2> error.log`

---

## References

- [LGPO Documentation (included PDF)](LGPO.pdf)
- [Microsoft Security Compliance Toolkit](https://www.microsoft.com/download/details.aspx?id=55319)
- [Registry Policy File Format](https://docs.microsoft.com/previous-versions/windows/desktop/Policy/registry-policy-file-format)
- [Chrome Enterprise Policy Documentation](https://cloud.google.com/docs/chrome-enterprise/policies)

---

## GPO Policy Manager GUI Tool

A PowerShell WPF application is included for managing GPO policy settings with a graphical interface.

### Features

- **Load ADMX/ADML Templates**: Import multiple template files to understand policy definitions
- **Load GPO Backups**: Parse GPO backup folders to see configured settings
- **Policy Matching**: Automatically matches GPO settings with ADMX definitions to show display names and descriptions
- **Select/Deselect**: Choose which policies to include in your export
- **Search & Filter**: Find policies by name, registry key, or category
- **Export**: Generate a new registry.pol file with only selected policies

### Usage

```powershell
# Run the tool (requires PowerShell 5.1+)
.\GPO-Policy-Manager.ps1

# Or right-click the file and select "Run with PowerShell"
```

### Workflow

1. **Launch** the tool (it will auto-detect LGPO.exe if in the same folder)
2. **Load ADMX/ADML** files (e.g., chrome.admx) to get policy definitions
3. **Load GPO Backup** folder to see the current settings
4. **Review** the policies in the grid - display names and categories will show if ADMX was loaded
5. **Select/Deselect** policies using checkboxes
6. **Export** selected policies to a new registry.pol file
7. **Apply manually** using: `LGPO.exe /m "path\to\registry.pol"`

### Screenshot Layout

```
+----------------------------------------------------------+
| GPO Policy Manager                              [_][□][X] |
+----------------------------------------------------------+
| [Load ADMX/ADML...] [Load GPO Backup...] [Export Selected]|
+----------------------------------------------------------+
| Loaded Templates: chrome.admx (694 policies)             |
| Loaded GPOs: DoD Chrome STIG v2r11 (44 settings)         |
+----------------------------------------------------------+
| [Search: _______________] [Category: All        v]       |
+----------------------------------------------------------+
| [x] | Policy Name    | Display Name | Value | Type       |
|-----|----------------|--------------|-------|------------|
| [x] | PasswordMgr... | Enable sav...| 0     | DWORD      |
| [x] | SyncDisabled   | Disable sy...| 1     | DWORD      |
| [ ] | AutoplayAll... | Allow auto...| 0     | DWORD      |
+----------------------------------------------------------+
| Selected: 42/44 | [Select All] [Deselect All]            |
+----------------------------------------------------------+
```

### Requirements

- Windows PowerShell 5.1 or later
- LGPO.exe (included or specify location)
- .NET Framework 4.5+ (included with Windows 10/11)

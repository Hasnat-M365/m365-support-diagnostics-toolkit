# Microsoft 365 Support Diagnostics Toolkit

A practical PowerShell-based Microsoft 365 troubleshooting project designed to automate common user diagnostic tasks using Microsoft Graph.

The project is being developed as a hands-on Microsoft 365 administration and support toolkit, with a focus on identity, licensing, access, and troubleshooting workflows.

## 🎯 Project Goal

The goal of this project is to build reusable PowerShell diagnostic tools for common Microsoft 365 support scenarios.

The toolkit is designed to demonstrate practical skills in:

- Microsoft Entra ID administration
- Microsoft 365 user troubleshooting
- Microsoft 365 licensing investigation
- Group membership analysis
- Microsoft Graph PowerShell
- PowerShell automation
- Structured troubleshooting and root-cause analysis

## ✅ Current Features

The first diagnostic script, `Get-M365UserDiagnostic.ps1`, can currently:

- Look up a Microsoft 365 user by User Principal Name (UPN)
- Display basic user information
- Check whether the user account is enabled
- Check the user's usage location
- Review assigned Microsoft 365 licenses
- Retrieve the user's group memberships

## 🧰 Technologies

The project currently uses:

- PowerShell 7
- Microsoft Graph PowerShell SDK
- Microsoft Entra ID
- Microsoft 365
- Git
- GitHub

Additional Microsoft 365 technologies will be introduced as the toolkit develops.

## 📂 Repository Structure

```text
m365-support-diagnostics-toolkit/
├── docs/
├── sample-output/
├── scripts/
│   ├── Get-M365UserDiagnostic.ps1
│   └── README.md
├── LICENSE
└── README.md
```

## 🚀 Running the Diagnostic Script

### Prerequisites

- PowerShell 7
- Microsoft Graph PowerShell SDK
- A Microsoft 365 account with appropriate Microsoft Graph permissions

### Run the script

From PowerShell:

```powershell
.\scripts\Get-M365UserDiagnostic.ps1
```

Enter the User Principal Name when prompted.

Example:

```text
user@contoso.com
```

The script retrieves diagnostic information about the specified Microsoft 365 user.

## 📊 Sample Output

A sanitized example of the diagnostic output is available in:

[`sample-output/Example-UserDiagnostic.txt`](sample-output/Example-UserDiagnostic.txt)

Example:

```text
DIAGNOSTIC SUMMARY
------------------
[PASS] Account enabled
[PASS] Usage location configured
[PASS] 2 license(s) assigned
[PASS] 2 group membership(s) found
[PASS] Additional authentication method(s) registered
```

All sample users, identifiers, groups, and tenant information in this repository are fictional or sanitized.

## 🧪 Lab Approach

Development and testing are performed using non-production scenarios with fictional or sanitized user information.

Troubleshooting scenarios follow a structured workflow:

1. Identify the reported issue
2. Gather diagnostic information
3. Review account and configuration data
4. Develop troubleshooting hypotheses
5. Identify the root cause
6. Apply the appropriate remediation
7. Validate the result
8. Document the troubleshooting process

## 🗺️ Roadmap

Planned improvements include:

- MFA-related checks
- Sign-in troubleshooting
- Improved license analysis
- Exchange Online mailbox diagnostics
- Additional error handling
- Diagnostic summaries
- Exportable reports
- Additional Microsoft 365 troubleshooting scripts

## 🔐 Security and Privacy

This repository is intended for learning, lab testing, and administrative troubleshooting.

Do not commit:

- Real user credentials
- Access tokens
- Client secrets
- Production tenant information
- Sensitive or personally identifiable information

Use fictional or sanitized information in examples and sample output.

## 📄 License

This project is provided under the license included in the repository.
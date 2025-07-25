# Get-UserMessagesReadStatus
This PowerShell script uses the Microsoft Graph SDK with application permissions to retrieve messages from a user's mailbox within a specified date range and export them to a CSV file, including key information like IsRead status

---

## Features

- Authenticates using Azure AD App Registration (client credentials flow)
- Retrieves messages using Microsoft Graph API
- Filters messages by `receivedDateTime`
- Exports message metadata to a CSV file
- Includes fields like subject, sender, recipients, read status, and more
---

## Prerequisites

- PowerShell 7+
- Microsoft Graph PowerShell SDK installed:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```
- An Azure AD App Registration
---

## Azure App Registration Setup

1. Register a New App
- Go to Azure Portal – App registrations
- Click New registration
- Name your app (e.g., GraphMailExport)
- Set Supported account types to: Accounts in this organizational directory only
- Click Register
  
2. Create a Client Secret
- In the app's blade, go to Certificates & secrets
- Click New client secret
- Add a description and expiration
- Copy the Value (you’ll use this as SecuredPassword in the script)
  
3. Assign API Permissions
- Go to API permissions
- Click Add a permission
- Choose Microsoft Graph
- Select Application permissions
- Search for and add:
      - Mail.Read (or Mail.ReadBasic)
      - User.Read.All (to resolve user IDs)
- Click Grant admin consent for your tenant
---

## Script Parameters

| Parameter | Description |
| ------------- | ------------- |
| ApplicationId | 	Azure AD App (client) ID |
| SecuredPassword | Client secret value |
| tentantId | Azure AD tenant ID | 
| UserPrincipalName | Email address of the target mailbox | 
| StartDate | Start of the message data range (UTC) | 
| EndDate | End of the message date range (UTC) | 
| CsvOutputPath | PAth to save the exported CSV | 
---
## Example Usage

```powershell
.\Get-UserMessages.ps1 `
    -ApplicationId "your-app-id" `
    -SecuredPassword "your-client-secret" `
    -tenantID "your-tenant-id" `
    -UserPrincipalName "user@contoso.com" `
    -StartDate "2025-06-01T00:00:00Z" `
    -EndDate "2025-07-01T00:00:00Z" `
    -CsvOutputPath "C:\Reports\UserMessages.csv"
```
---

## Output CSV Fields

- Subject
- ReceivedDateTime
- From
- ToRecipients
- CcRecipients
- IsRead
- Importance
- BodyPreview
- ConversationId
- MessageId
- WebLink
---

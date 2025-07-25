<#
.SYNOPSIS
    Retrieves messages from a user's mailbox using Microsoft Graph SDK with app permissions.

.PARAMETER UserPrincipalName
    The UPN (email address) of the user whose mailbox messages you want to retrieve.

.PARAMETER StartDate
    The start date for filtering messages.

.PARAMETER EndDate
    The end date for filtering messages.

.PARAMETER CsvOutputPath
    The file path where the output CSV will be saved.
#>

# =========================
# Script Parameters
# =========================
param (
    [string]$ApplicationId = "YourAppId",  # Azure AD App (Client) ID
    [string]$SecuredPassword = "YourClientSecret",  # Client Secret (secure this in production!)
    [string]$tenantID = "YourTenantId",  # Azure AD Tenant ID
    [string]$UserPrincipalName = "targetUser@contoso.com",  # Target user's email address
    [datetime]$StartDate = "2025-06-01T00:00:00Z",  # Start of date range
    [datetime]$EndDate = "2025-07-16T00:00:00Z",    # End of date range
    [string]$CsvOutputPath = "C:\temp\UserMessages.csv"  # Output CSV file path
)

# =========================
# Connect to Microsoft Graph using app credentials
# =========================
function Connect-ToGraph {
    # Convert the plain text client secret to a secure string
    $SecuredPasswordPassword = ConvertTo-SecureString -String $SecuredPassword -AsPlainText -Force

    # Create a PSCredential object using the App ID and secure secret
    $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationId, $SecuredPasswordPassword

    # Connect to Microsoft Graph using the credential and tenant ID
    Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential
}

# =========================
# Get the Object ID of the user from their UPN
# =========================
function Get-UserObjectId {
    param ([string]$UserPrincipalName)

    # Retrieve the user's Graph object ID
    return (Get-MgUser -UserId $UserPrincipalName).Id
}

# =========================
# Retrieve messages from the user's mailbox within the specified date range
# =========================
function Get-UserMessages {
    param (
        [string]$UserId,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    # Construct the filter string for the Graph query
    $filter = "receivedDateTime ge $($StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")) and receivedDateTime le $($EndDate.ToString("yyyy-MM-ddTHH:mm:ssZ"))"

    # Retrieve messages using the Graph SDK with filtering and pagination
    $messages = Get-MgUserMessage -UserId $UserId -Filter $filter -Top 50 -All

    # Format the output as a collection of custom objects for CSV export
    $output = foreach ($msg in $messages) {
        [PSCustomObject]@{
            Subject          = $msg.Subject
            ReceivedDateTime = $msg.ReceivedDateTime
            From             = $msg.From.EmailAddress.Address
            ToRecipients     = ($msg.ToRecipients | ForEach-Object { $_.EmailAddress.Address }) -join "; "
            CcRecipients     = ($msg.CcRecipients | ForEach-Object { $_.EmailAddress.Address }) -join "; "
            IsRead           = $msg.IsRead
            Importance       = $msg.Importance
            BodyPreview      = $msg.BodyPreview
            ConversationId   = $msg.ConversationId
            MessageId        = $msg.Id
            WebLink          = $msg.WebLink
        }
    }

    return $output
}

# =========================
# Main Execution Function
# =========================
function Main {
    # Authenticate to Microsoft Graph
    Connect-ToGraph

    # Get the user's object ID
    $UserId = Get-UserObjectId -UserPrincipalName $UserPrincipalName

    # Retrieve messages within the specified date range
    $messages = Get-UserMessages -UserId $UserId -StartDate $StartDate -EndDate $EndDate

    # If no CSV path is provided, generate a default one with timestamp
    if ([string]::IsNullOrWhiteSpace($CsvOutputPath)) {
        $CsvOutputPath = ".\UserMessages_$($UserPrincipalName.Replace('@','_'))_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
    }

    # Export the messages to a CSV file
    $messages | Export-Csv -Path $CsvOutputPath -NoTypeInformation -Encoding UTF8

    # Output summary to console
    Write-Output "Exported $($messages.Count) messages to $CsvOutputPath"
}

# =========================
# Call Main with Parameters
# =========================
Main -UserPrincipalName $UserPrincipalName `
     -StartDate $StartDate `
     -EndDate $EndDate `
     -CsvOutputPath $CsvOutputPath

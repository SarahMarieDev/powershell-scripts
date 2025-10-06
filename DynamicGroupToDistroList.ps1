# ------------------------------
# Configuration
# ------------------------------
$dynamicGroupId   = "<AZURE-GROUP-ID>"     # Azure AD group (dynamic)
$distributionList = "<DISTRO-GROUP-NAME-OR-ALIAS>"     # DL name or alias

# ------------------------------
# Connect to Microsoft Graph (to get dynamic group members)
# ------------------------------
Write-Output "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "GroupMember.Read.All"
Select-MgProfile beta  # or v1.0 if you prefer

Write-Output "Fetching members from dynamic group..."
$members = @()
$members += Get-MgGroupMember -GroupId $dynamicGroupId -All

Write-Output "Retrieved $($members.Count) members from dynamic group."

# ------------------------------
# Connect to Exchange Online
# ------------------------------
Write-Output "Connecting to Exchange Online..."
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -UserPrincipalName (Read-Host "Enter your admin UPN")

# ------------------------------
# Add members to distribution list
# ------------------------------
foreach ($member in $members) {
    try {
        # Try to get the user's email address
        $email = $member.AdditionalProperties.mail
        if (-not $email) {
            Write-Output "$(Get-Date) - Skipping member with no email: $($member.Id) ; $($member.AdditionalProperties.displayName)"
            continue
        }

        # Check if already a member
        $alreadyMember = Get-DistributionGroupMember -Identity $distributionList -ResultSize Unlimited |
                         Where-Object { $_.PrimarySmtpAddress -eq $email }

        if ($alreadyMember) {
            Write-Output "$(Get-Date) - Already a member: $email"
        }
        else {
            Add-DistributionGroupMember -Identity $distributionList -Member $email -ErrorAction Stop
            Write-Output "$(Get-Date) - Added: $email"
        }
    }
    catch {
        Write-Output "$(Get-Date) - Error adding $email : $($_.Exception.Message)"
    }
}

# ------------------------------
# Cleanup
# ------------------------------
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
Write-Output "Done."

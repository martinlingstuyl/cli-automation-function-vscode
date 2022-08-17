# Input bindings are passed in via param block.
param($Timer)

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

Write-Host "Logging into Microsoft 365 using Managed Identity"

m365 login --authType identity

$status = m365 status

Write-Host "Current Status: $status"

m365 spo set --url "https://blimped.sharepoint.com"

$sites = m365 spo site list --type All | ConvertFrom-Json 

Write-Host "Number of sites: $($sites.Count)"
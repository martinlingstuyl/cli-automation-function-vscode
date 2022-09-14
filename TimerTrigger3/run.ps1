# Input bindings are passed in via param block.
param($Timer)

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

Write-Host "Logging into Microsoft 365 using Managed Identity"

m365 login --authType identity

m365 spo set --url "https://contoso.sharepoint.com"

$sites = m365 spo site list --type All | ConvertFrom-Json

$guestList = [System.Collections.ArrayList]::new()

foreach ($site in $sites) {
    $users = m365 request --url "$($site.Url)/_api/web/siteusers?`$filter=IsShareByEmailGuestUser eq true&`$expand=Groups&`$select=Title,LoginName,Email,Groups/LoginName" | ConvertFrom-Json  

    foreach($user in $users.value) {
        foreach($group in $user.Groups | Where-Object { $_.LoginName -cnotmatch "Limited Access System Group" -and $_.LoginName -cnotmatch "SharingLinks"}) {
            $obj = [PSCustomObject][ordered]@{
                Title = $user.Title;
                Email = $user.Email;
                LoginName = $user.LoginName;
                SiteUrl = $site.Url;
                Group = $group.LoginName;
            }
            $guestList.Add($obj) | Out-Null
        }
    }
}

$path = Join-Path -Path "$PSScriptRoot" -ChildPath ".." | Join-Path -ChildPath ".." -Resolve

m365 spo file get --webUrl "/sites/mysite" --url "/sites/mysite/shared documents/guest-list.csv" --path "$path/previous-guest-list.csv" --asFile | out-null

If ((Test-Path -Path "$path/previous-guest-list.csv") -eq $true) {
    $previousList = Import-Csv -Path "$path/previous-guest-list.csv" -ErrorAction Ignore
}

$postUpdates = $null -eq $previousList -or (Compare-Object -ReferenceObject $guestList -DifferenceObject $previousList).length -gt 0

if ($postUpdates -eq $true) {
    $guestList | Export-Csv -Path "$path/guest-list.csv" -NoTypeInformation
    
    m365 spo file add --webUrl "/sites/mysite" --folder "shared documents" --path "$path/guest-list.csv"

    m365 outlook mail send --to "martin@contoso.com" --sender "<user-id-of-my-account>" --subject "Guest access change notification" --bodyContents `@TimerTrigger3/email-body.html --bodyContentType HTML
}

If ((Test-Path -Path "$path/previous-guest-list.csv") -eq $true) {
    Remove-Item -Path "$path/previous-guest-list.csv" -Force
}

If ((Test-Path -Path "$path/guest-list.csv") -eq $true) {
    Remove-Item -Path "$path/guest-list.csv" -Force
}
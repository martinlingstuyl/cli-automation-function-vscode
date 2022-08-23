# Input bindings are passed in via param block.
param($Timer)

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

Write-Host "Logging into Microsoft 365 using Managed Identity"

m365 login --authType identity

m365 spo set --url "https://contoso.sharepoint.com"

$webUrl = "/sites/contoso"
$incomingWebhookURL = "<enter an incoming webhook url>"

$list = m365 spo list get --webUrl $webUrl --title HealthIssuesSent | ConvertFrom-Json

if ($null -eq $list) {
    $list = m365 spo list add --webUrl $webUrl --title HealthIssuesSent --baseTemplate GenericList | ConvertFrom-Json
}

$cachedIssues = m365 spo listitem list --webUrl $webUrl --listId $list.Id --fields "Id,Title" | ConvertFrom-Json

Write-Host "Loading health issues for SharePoint" -ForegroundColor Green

$issues = m365 tenant serviceannouncement healthissue list --service "SharePoint Online" --query "[?!isResolved]" | ConvertFrom-Json


Write-Host "Health issues found for SharePoint: $($issues.Length)"

foreach ($issue in $issues) {
    $savedIssue = $cachedIssues | Where-Object { $_.Title -eq $issue.id }

    if ($null -eq $savedIssue) {
        Write-Host "New issue found for SharePoint: $($issue.id)" -ForegroundColor Green

        Write-Host "Sending notification to Teams $pwd"
        m365 adaptivecard send --card `@TimerTrigger2/adaptive-card.json --url $incomingWebhookURL --cardData "{ \`"title\`": \`"A health incident occurred on SharePoint\`", \`"description\`": \`"$($issue.Title)\`", \`"issueId\`": \`"$($issue.id)\`", \`"issueTimestamp\`": \`"$($issue.startDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ"))\`", \`"viewUrl\`": \`"https://admin.microsoft.com/Adminportal/Home#/servicehealth/:/alerts/$($issue.id)\`", \`"properties\`":[{\`"key\`":\`"Classification\`",\`"value\`":\`"$($issue.classification)\`"},{\`"key\`":\`"Feature Group\`",\`"value\`":\`"$($issue.featureGroup)\`"},{\`"key\`":\`"Feature\`",\`"value\`":\`"$($issue.feature)\`"}] }"

        Write-Host "Saving issue to List to avoid repetition"
        m365 spo listitem add --webUrl $webUrl --listId $list.Id --Title $issue.id | out-null
    } 
}

# Remove resolved items
foreach ($cachedIssue in $cachedIssues) {

    $isResolved = @($issues | Where-Object { $_.id -eq $cachedIssue.Title }).Count -eq 0

    if ($isResolved -eq $true) {
        Write-Host "Removing resolved issue from list: $($issue.id)" -ForegroundColor Green

        m365 spo listitem remove --webUrl $webUrl --listId $list.Id --id $cachedIssue.Id --confirm | out-null
    }
}


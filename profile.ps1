# Azure Functions profile.ps1
#
# This profile.ps1 will get executed every "cold start" of your Function App.
# "cold start" occurs when:
#
# * A Function App starts up for the very first time
# * A Function App starts up after being de-allocated due to inactivity
#
# You can define helper functions, run commands, or specify environment variables
# NOTE: any variables defined that are not environment variables will get reset after the first execution

Write-Host "Setting up m365 in PATH variable"
Write-Host "Current Path: $($Env:Path)" 

$functionPath = "$PWD\node_modules\.bin" 
if ($Env:PATH.Contains($functionPath) -eq $false) {
    [System.Environment]::SetEnvironmentVariable('PATH',$Env:PATH + ";$functionPath")
}

Write-Host "Changed Path: $($Env:Path)" 
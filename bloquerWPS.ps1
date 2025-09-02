param(
    # Specifies the name of the rule set to apply. Corresponds to an XML file in the 'rules' folder.
    [string]$RuleSet = "bloquearWPS"
)

# This script manages AppLocker rules to block WPS Office.

# Establecemos el valor para el inicio del servicio AppIDSvc 2: automatic, 3: manual
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\AppIDSvc" -Name "Start" -Value 2

# Start the service if it's not already running.
Start-Service -Name AppIDSvc

# Backup existing AppLocker rules
$timestamp = Get-Date -Format "yyMMddHHmmss"
$backupFile = "rules-backup-$timestamp.xml"
Get-AppLockerPolicy -Local -XML > $backupFile
# Clear all existing AppLocker rules
# A new, empty AppLocker policy object is created.

#$emptyPolicy = New-Object Microsoft.Security.ApplicationId.Policy.PolicyObjectModel

# The Set-AppLockerPolicy cmdlet sets the AppLocker policy.
# By providing an empty policy object, we effectively remove all existing rules.

#Set-AppLockerPolicy -PolicyObject $emptyPolicy

# Apply new rules from the specified XML file
# Construct the path to the rule file based on the script's location and the RuleSet parameter.
$ruleFile = Join-Path -Path $PSScriptRoot -ChildPath "rules/$($RuleSet).xml"
Write-Host $ruleFile

# Verify that the selected rule file exists before attempting to apply it.
if (-not (Test-Path -Path $ruleFile)) {
    Write-Host "Error: The rule file for '$($RuleSet)' was not found at '$ruleFile'." -ForegroundColor Red
    # Exit the script if the file doesn't exist to prevent unintended states.
    exit 1
}

# Apply the new AppLocker policy from the specified XML file.
# The Set-AppLockerPolicy cmdlet sets the policy, overwriting the (now empty) current policy.
Set-AppLockerPolicy -XMLPolicy $ruleFile

Write-Host "Successfully applied AppLocker rules from '$ruleFile'." -ForegroundColor Green

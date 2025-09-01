param(
    # Specifies the name of the rule set to apply. Corresponds to an XML file in the 'rules' folder.
    [string]$RuleSet = "default"
)

# This script manages AppLocker rules to block WPS Office.

# Configure and start the Application Identity service (AppIDSvc).
# This service is essential for AppLocker to function.
# Set the service to start automatically with the system.
Set-Service -Name AppIDSvc -StartupType Automatic
# Start the service if it's not already running.
Start-Service -Name AppIDSvc

# Backup existing AppLocker rules
# The Get-AppLockerPolicy cmdlet gets the local AppLocker policy.
# The Export-Clixml cmdlet creates an XML-based representation of the object or objects and stores it in a file.
Get-AppLockerPolicy -Local | Export-Clixml -Path "rules-backup.xml"

# Clear all existing AppLocker rules
# A new, empty AppLocker policy object is created.
$emptyPolicy = New-Object Microsoft.Security.ApplicationId.Policy.PolicyObjectModel
# The Set-AppLockerPolicy cmdlet sets the AppLocker policy.
# By providing an empty policy object, we effectively remove all existing rules.
Set-AppLockerPolicy -PolicyObject $emptyPolicy

# Apply new rules from the specified XML file
# Construct the path to the rule file based on the script's location and the RuleSet parameter.
$ruleFile = Join-Path -Path $PSScriptRoot -ChildPath "rules/$($RuleSet).xml"

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

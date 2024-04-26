Import-Module PnP.PowerShell

$migrationConfiguration = Get-Content .\config.json -Raw | ConvertFrom-Json

if([string]::IsNullOrEmpty($migrationConfiguration.DestinationCredentials.UserName)) {
    Write-Host "Connecting to SPO via interactive method" -ForegroundColor Green
    Connect-PnPOnline $migrationConfiguration.DestinationSite -Interactive
} else {
    Write-Host "Connecting to SPO via using config file credentials" -ForegroundColor Green
    $securePassword = ConvertTo-SecureString -AsPlainText $migrationConfiguration.DestinationCredentials.Password -Force
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $migrationConfiguration.DestinationCredentials.UserName $securePassword
    Connect-PnPOnline $migrationConfiguration.DestinationSite -Credentials $cred
}


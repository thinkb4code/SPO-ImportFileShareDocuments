# Import SharePoint PnP Scripts.
Import-Module PnP.PowerShell

# Load Configuration file.
$migrationConfiguration = Get-Content .\config.json -Raw | ConvertFrom-Json

# Migration configuration
$libraryPath = $migrationConfiguration.DestinationLibraryTitle
$manifestFile = Import-Csv $migrationConfiguration.ManifestFile
$fileCount = 1

# Connect to SharePoint Online using method defined in configuration file.
if([string]::IsNullOrEmpty($migrationConfiguration.DestinationCredentials.UserName)) {
    Write-Host "Connecting to SPO via interactive method" -ForegroundColor Green
    Connect-PnPOnline $migrationConfiguration.DestinationSite -Interactive
} else {
    Write-Host "Connecting to SPO via using config file credentials" -ForegroundColor Green
    $securePassword = ConvertTo-SecureString -AsPlainText $migrationConfiguration.DestinationCredentials.Password -Force
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $migrationConfiguration.DestinationCredentials.UserName $securePassword
    Connect-PnPOnline $migrationConfiguration.DestinationSite -Credentials $cred
}



foreach($csvRow in $manifestFile){
    $uploadPath = "$libraryPath"
    
    # Set up folder migration details
    if($migrationConfiguration.UploadFolderHierarchy) {
        $folderPath = @($migrationConfiguration.FolderHierarchyFieldsSequence | % {$csvRow.$_}) -join "/"
        $uploadPath = "$($libraryPath)/$($folderPath)"
    }

    Write-Host "Uploading $fileCount of $($manifestFile.Count) to: $uploadPath"

    # Read file from network drive as a stream

$csvRow."$($migrationConfiguration.FileName)"

    #Add-PnPFile -Folder $libraryPath -FileName $fileName -Stream $stream -Values $metaData -Publish -PublishComment "Automated upload using PnP PowerShell migration script."
    $fileCount++
}

try {
    

    
}
catch {
    <#Do this if a terminating exception happens#>
    #Error Log
}
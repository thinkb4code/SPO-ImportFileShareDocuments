# Import SharePoint PnP Scripts.
Import-Module PnP.PowerShell

# Load Configuration file.
$migrationConfiguration = Get-Content .\config.json -Raw | ConvertFrom-Json

# Migration configuration
$libraryPath = $migrationConfiguration.DestinationLibraryTitle
$manifestFile = Import-Csv $migrationConfiguration.ManifestFile
$fileCount = 1
$publishComment = "Automated upload using PnP PowerShell migration script."

# Migration Log Output File
$outputFile = ".\Logs\MigrationLog_$((Get-Date).ToString("yyyyMMddHHmmss")).csv"

if(!(Test-Path -Path ".\Logs")){
    New-Item -Path ".\Logs" -ItemType Directory | Out-Null
}

#CSV Header
"Operation,Status,Source,Destination,Details" | Out-File -FilePath $outputFile

# Check for site Add and Customize Pages permission enabled for site. 
function Check-PnPAdminSiteForScriptPermissions {
    if([string]::IsNullOrEmpty($migrationConfiguration.DestinationCredentials.UserName)) {
        Write-Host "Connecting to SPO Admin Center via interactive method" -ForegroundColor Green
        Connect-PnPOnline $migrationConfiguration.TenantAdminUrl -Interactive
    }else {
        Write-Host "Connecting to SPO Admin Center via using config file credentials" -ForegroundColor Green
        $securePasswordAdmin = ConvertTo-SecureString -AsPlainText $migrationConfiguration.DestinationCredentials.Password -Force
        $credAdmin = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $migrationConfiguration.DestinationCredentials.UserName $securePasswordAdmin
        Connect-PnPOnline $migrationConfiguration.TenantAdminUrl -Credentials $credAdmin
    }

    $site = Get-PnPTenantSite $migrationConfiguration.DestinationSite -Detailed
    if($site.DenyAddAndCustomizePages -eq [Microsoft.Online.SharePoint.TenantAdministration.DenyAddAndCustomizePagesStatus]::Enabled){
        # Scripting permissions need to be granted.
        Write-Host "Disable site Deny Add & Customized Pages permissions at tenant level"
        Set-PnPTenantSite $migrationConfiguration.DestinationSite -DenyAddAndCustomizePages:$false
    }
}

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

# Uncomment following line if Add-PnPFile result in Access Denied error
# Check-PnPAdminSiteForScriptPermissions

foreach($csvRow in $manifestFile){
    $uploadPath = "$libraryPath"
    $fileName = $csvRow."$($migrationConfiguration.FileNameMappingInCSV)"
    $networkFileLocation = $csvRow."$($migrationConfiguration.NetworkFileLocationMappingInCSV)"
    $metadata = $null
    $metadataValidation = $null
    
    # Set up folder migration details
    if($migrationConfiguration.UploadFolderHierarchy) {
        $folderPath = @($migrationConfiguration.FolderHierarchyFieldsSequence | % {$csvRow.$_}) -join "/"
        $uploadPath = "$($libraryPath)/$($folderPath)"
    }

    Write-Host "Uploading $fileCount of $($manifestFile.Count) to: $uploadPath/$fileName"

    # Setup metadata of uploaded file
    if($migrationConfiguration.MetaDataMapping.Count -gt 0){
        $metadata = @{}
        $metadataValidation = New-Object System.Collections.ArrayList

        foreach($mapping in $migrationConfiguration.MetaDataMapping){
            $columnData = $csvRow."$($mapping.SourceColumn)"
            # Date validation in mm/dd/yyyy format
            switch -CaseSensitive ($mapping.Type) {
                "Date" {
                    if($columnData -match '\d{1,2}\/\d{1,2}\/(19\d{2}|20[0-2][0-4])'){
                        $metadata.Add($mapping.DestinationColumn, $columnData)
                    }else {
                        $metadataValidation.Add("Invalid Date")
                    }
                    break
                }
                Default {
                    $metadata.Add($mapping.DestinationColumn, $columnData)
                }
            }
        }
    }

    # Read file from network drive as a stream
    $fileContent = Get-Content -Path $networkFileLocation -AsByteStream -Raw
    $filestream = [System.IO.MemoryStream]::new($fileContent)

    try {
        if($metadata.Count -gt 0){
            $newFileInSPO = Add-PnPFile -Folder $uploadPath -FileName $fileName -Stream $filestream -Values $metadata -PublishComment $publishComment
        }else {
            $newFileInSPO = Add-PnPFile -Folder $uploadPath -FileName $fileName -Stream $filestream -PublishComment $publishComment
        }
        
        "Upload,Success,$networkFileLocation,$($newFileInSPO.ServerRelativeUrl)," | Out-File -FilePath $outputFile -Append
    }
    catch {
        <#Do this if a terminating exception happens#>
        "Upload,Failed,$networkFileLocation,#NA,$($_.Exception.Message)" | Out-File -FilePath $outputFile -Append
    } finally {
        <#Do this after the try block regardless of whether an exception occurred or not#>
        foreach($validation in $metadataValidation){
            "Validation,Failed,$networkFileLocation,$($newFileInSPO.ServerRelativeUrl),$validation" | Out-File -FilePath $outputFile -Append
        }
    }
    
    $fileCount++
}

#Disconnect-PnPOnline
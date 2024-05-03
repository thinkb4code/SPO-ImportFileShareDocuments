# Import SharePoint PnP Scripts.
Import-Module PnP.PowerShell

# Load Configuration file.
$migrationConfiguration = Get-Content .\config.json -Raw | ConvertFrom-Json

# Check for site Add and Customize Pages permission enabled for site. 
# This feature uses SPO Tenant Admin permission. 
function Update-RunScriptPermission {
    param(
        [string]$TenantAdminSite,
        [string]$DestinationSite,
        [string]$UserName,
        [string]$Password
    )

    if([string]::IsNullOrEmpty($UserName)) {
        Write-Host "Connecting to SPO Admin Center via interactive method" -ForegroundColor Green
        Connect-PnPOnline $TenantAdminSite -Interactive
    }else {
        Write-Host "Connecting to SPO Admin Center via using config file credentials" -ForegroundColor Green
        $securePasswordAdmin = ConvertTo-SecureString -AsPlainText $Password -Force
        $credAdmin = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName $securePasswordAdmin
        Connect-PnPOnline $TenantAdminSite -Credentials $credAdmin
    }

    $site = Get-PnPTenantSite $DestinationSite -Detailed
    if($site.DenyAddAndCustomizePages -eq [Microsoft.Online.SharePoint.TenantAdministration.DenyAddAndCustomizePagesStatus]::Enabled){
        # Scripting permissions need to be granted.
        Write-Host "Disable site Deny Add & Customized Pages permissions at tenant level" -ForegroundColor Yellow
        Set-PnPTenantSite $DestinationSite -DenyAddAndCustomizePages:$false
    }
}

# Generate the document metadata for uploading file to SPO.
function New-DocumentMetaData {
    param (
        $FieldMaps,
        $ValueMaps
    )
    
    $metadata = @{}
    $metadataValidation = New-Object System.Collections.ArrayList

    # Loop all the metadata fields configured in JSON Config and Generate Upload Metadata

    foreach($FieldMap in $FieldMaps){
        $columnData = $ValueMaps."$($FieldMap.SourceField)"

        if([string]::IsNullOrEmpty($columnData) -and ($FieldMap.Required -eq 1)){
            $metadataValidation.Add("Required column missing value. Field Name: $($FieldMap.SourceField)") | Out-Null
        }elseif(![string]::IsNullOrEmpty($columnData)){
            switch -CaseSensitive ($FieldMap.TargetType) {
                "Text" {
                    if($columnData.Length -gt 255) {
                        $metadataValidation.Add("Text Column Length Exceeded for: $($FieldMap.SourceField)") | Out-Null
                    }elseif($FieldMap.TargetField -eq "Name"){
                        # Skip adding value for Name column
                    }else {
                        $metadata.Add($FieldMap.TargetField, $columnData)
                    }
                    Break
                }
                "DateTime" {
                    try {
                        $sourceDate = Get-Date -Date $columnData
                        $metadata.Add($FieldMap.TargetField, $sourceDate.ToString("MM/dd/yyyy HH:mm:ss"))
                    }
                    catch {
                        $metadataValidation.Add("Invalid data. Field Name: $($FieldMap.SourceField)") | Out-Null
                    }
                    Break
                }
                "Choice" {
                    $metadata.Add($FieldMap.TargetField, $columnData)
                    Break
                }
                "Lookup" {
                    $camlQuery = "<View>
                        <Query>
                            <Where>
                                <Eq>
                                    <FieldRef Name='$($FieldMap.LookupColumn)'/>
                                    <Value Type='Text'>$($columnData)</Value>
                                </Eq>
                            </Where>
                            <ViewFields>
                                <FieldRef  Name='ID' />
                            </ViewFields>
                        </Query>
                    </View>"

                    $item = Get-PnPListItem -List "Insurance Details" -Query $camlQuery

                    If($item.Id){
                        $metadata.Add($FieldMap.TargetField, $item.Id)
                    }else {
                        $metadataValidation.Add("Lookup value '$($columnData)' not found in '$($FieldMap.LookupList)' list.")
                    }
                }
                "User" {
                    $metadata.Add($FieldMap.TargetField, $columnData)
                    Break
                }
                Default {
                    $metadata.Add($FieldMap.TargetField, $columnData)
                }
            }
        }else {
            $metadataValidation.Add("Skipping blank metadata value upload: Field Name: $($FieldMap.SourceField)") | Out-Null
        }
    }

    return @{"MetaData" = $metadata; "Validation" = $metadataValidation}
}

foreach($task in $migrationConfiguration.Tasks){
    Write-Host "$($task.Name) import process started at $(Get-Date -Format "HH:mm:ss")" -ForegroundColor Yellow

    # Migration configuration
    $libraryPath = $task.TargetLibrary
    $manifestFile = Import-Csv -Path $task.ManifestFile
    $fileCount = 1
    $publishComment = "Automated upload using PnP PowerShell migration script."

    # Migration Log Output File
    $outputFile = ".\Logs\$($task.Name)_$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
    $stackTraceFile = ".\Logs\StackTrace_$($task.Name)_$((Get-Date).ToString("yyyyMMddHHmmss")).txt"

    # Verify and create Log folder if not exists already
    if(!(Test-Path -Path ".\Logs")){
        New-Item -Path ".\Logs" -ItemType Directory | Out-Null
    }

    # Create CSV Log file with Header
    "Operation,Status,Source,Destination,Details" | Out-File -FilePath $outputFile
    "" | Out-File -FilePath $stackTraceFile

    # Uncomment following line if Add-PnPFile result in Access Denied error, provide the Tenant Admin Site URL in following too
    # Update-RunScriptPermission -TenantAdminSite "" -DestinationSite $task.TargetUrl -UserName $task.TargetCredentials.UserName -Password $task.TargetCredentials.Password

    # Connect to SharePoint Online using method defined in configuration file.
    if([string]::IsNullOrEmpty($task.TargetCredentials.UserName)) {
        Write-Host "Connecting to SPO via interactive method" -ForegroundColor Green
        Connect-PnPOnline $task.TargetUrl -Interactive
    } else {
        Write-Host "Connecting to SPO via using config file credentials" -ForegroundColor Green
        $securePassword = ConvertTo-SecureString -AsPlainText $task.TargetCredentials.Password -Force
        $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $task.TargetCredentials.UserName $securePassword
        Connect-PnPOnline $task.TargetUrl -Credentials $cred
    }

    # For each file listed in Manifest CSV file, validate the file upload context and finally push to SPO 
    foreach($csvRow in $manifestFile){
        # Setup metadata of uploaded file
        $newDocMetadata = New-DocumentMetaData -FieldMaps $task.FieldMap -ValueMaps $csvRow
        
        try {
            $uploadPath = "$libraryPath"
            $fileName = $csvRow."$($task.FileNameInManifestCSV)"
            $networkFileLocation = -join($task.SourceUnc, "\", $csvRow."$($task.FileNameInManifestCSV)")
            
            # Check if folder hierarchy is enabled and set up folder hierarchy using configuration from Manifest file
            if($task.ParentFolderHierarchyUpload) {
                $folderPath = @($task.ParentFolder | ForEach-Object {$csvRow.$_}) -join "/"
                $uploadPath = "$($libraryPath)/$($folderPath)"
                if($uploadPath.LastIndexOf("/") -eq $uploadPath.Length - 1){
                    $uploadPath = $uploadPath.Substring(0, $uploadPath.Length - 1)
                }
            }
            
            # Read file from network drive as a stream
            $fileContent = Get-Content -Path $networkFileLocation -AsByteStream -Raw
            $filestream = [System.IO.MemoryStream]::new($fileContent)

            # Upload File to SPO
            if($newDocMetadata.MetaData.Count -gt 0){
                Write-Host "Uploading $fileCount of $($manifestFile.Count) to: $uploadPath/$fileName with metadata" -ForegroundColor Green
                $newFileInSPO = Add-PnPFile -Folder $uploadPath -FileName $fileName -Stream $filestream -Values $newDocMetadata.MetaData -PublishComment $publishComment
            }else {
                Write-Host "Uploading $fileCount of $($manifestFile.Count) to: $uploadPath/$fileName without metadata" -ForegroundColor Yellow
                $newFileInSPO = Add-PnPFile -Folder $uploadPath -FileName $fileName -Stream $filestream -PublishComment $publishComment
            }
            
            "Upload,Success,$networkFileLocation,$($newFileInSPO.ServerRelativeUrl)," | Out-File -FilePath $outputFile -Append
        }
        catch {
            <#Do this if a terminating exception happens#>
            Write-Host "Uploading $fileCount of $($manifestFile.Count) to: $uploadPath/$fileName with metadata" -ForegroundColor Red
            "Upload,Failed,$networkFileLocation,#NA,$($_.Exception.Message)" | Out-File -FilePath $outputFile -Append
            $fileName | Out-File -FilePath $stackTraceFile -Append
            $_ | Out-File -FilePath $stackTraceFile -Append
            $_.Exception.StackTrace | Out-File $stackTraceFile -Append
            "`r`n" | Out-File $stackTraceFile -Append
            "`r`n" | Out-File $stackTraceFile -Append
            
        } finally {
            <#Do this after the try block regardless of whether an exception occurred or not#>
            foreach($validation in $newDocMetadata.Validation){
                if($validation.StartsWith("Skipping")){
                    "Validation,Skip,$networkFileLocation,$($newFileInSPO.ServerRelativeUrl),$validation" | Out-File -FilePath $outputFile -Append
                } else {
                    "Validation,Failed,$networkFileLocation,$($newFileInSPO.ServerRelativeUrl),$validation" | Out-File -FilePath $outputFile -Append
                }
            }
        }
        
        $fileCount++
    }

    Disconnect-PnPOnline
}

$FileName = "SubscriptionLicenseCount.xlsx"
$FilePath = Join-Path -Path $PWD.Path -ChildPath $FileName

Write-Progress -Activity "Preparation" -Status "Checking for existing file and cleaning up" -PercentComplete 0
if (Test-Path $FilePath -PathType Leaf) {
    Remove-Item $FilePath -Force
}

$ApplicationClientId = '' 
$ApplicationClientSecret = ''
$TenantId = ''

Write-Progress -Activity "Authentication" -Status "Connecting to Microsoft Graph" -PercentComplete 10
$SecureClientSecret = ConvertTo-SecureString -String $ApplicationClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationClientId, $SecureClientSecret
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

Write-Progress -Activity "Fetching Team Information" -Status "Retrieving group and drive details" -PercentComplete 30
$Team = "GIT Infra Team"
$Group = Get-MgGroup -Filter "displayName eq '$Team'"
$GroupDrive = Get-MgGroupDrive -GroupId $Group.Id | Select-Object -First 1
$DriveItem = Get-MgDriveItem -DriveId $GroupDrive.Id -Filter "file ne null" | Where-Object { $_.Name -eq 'Microsoft Licensing' }
$DriveChild = Get-MgDriveItemChild -DriveId $GroupDrive.Id -DriveItemId $DriveItem.Id -Filter "name eq '$($FileName)'"
Get-MgDriveItemContent -DriveId $GroupDrive.Id -DriveItemId $DriveChild.Id -OutFile $FilePath

Write-Progress -Activity "Fetching Users" -Status "Getting Azure user details" -PercentComplete 50
$azureUsers = Get-MgUser -All -Filter "(assignedLicenses/any(x:x/skuId eq 18181a46-0d4e-45cd-891e-60aabd171b4e) or assignedLicenses/any(x:x/skuId eq 6fd2c87f-b296-42f0-b197-1e91e994b900) or assignedLicenses/any(x:x/skuId eq 05e9a617-0261-4cee-bb44-138d3ef5d965) or assignedLicenses/any(x:x/skuId eq 078d2b04-f1bd-4111-bbd4-b4b1b354cef4)) and UserType eq 'Member'" -ConsistencyLevel eventual -CountVariable licensedUserCount -Property DisplayName, UserPrincipalName, CompanyName, Id, JobTitle | Select-Object DisplayName, UserPrincipalName, CompanyName, Id, JobTitle

Write-Progress -Activity "Processing Users" -Status "Processing each user for license details" -PercentComplete 70
$userLicenses = @()
$totalUsers = $azureUsers.Count
$currentUserIndex = 0

foreach ($user in $azureUsers) {
    $currentUserIndex++
    $progressPercent = [math]::Round(($currentUserIndex / $totalUsers) * 100, 0)
    Write-Progress -Activity "Processing Users" -Status "Processing $currentUserIndex of $totalUsers users" -PercentComplete $progressPercent

    if ($null -eq $user.Id) {
        Write-Output "Skipping user $($user.UserPrincipalName) because Id is null"
        continue
    }

    $licenses = Get-MgUserLicenseDetail -UserId $user.Id
    $licenseSkus = $licenses | Select-Object -ExpandProperty SkuPartNumber -Unique
    $mailboxSettings = Get-MgUserMailboxSetting -UserId $user.Id -ErrorAction SilentlyContinue

    $userInfo = [PSCustomObject]@{
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        MailboxType       = $mailboxSettings.UserPurpose
        JobTitle          = $user.JobTitle
        CompanyName       = $user.CompanyName
        Licenses          = ($licenseSkus -join ", ")

        "O365 E1"         = if ($licenseSkus -contains "STANDARDPACK") { 1 } else { 0 }
        "O365 E3"         = if ($licenseSkus -contains "ENTERPRISEPACK") { 1 } else { 0 }
        "M365 E3"         = if ($licenseSkus -contains "SPE_E3") { 1 } else { 0 }
        "POWERBI PRO"     = if ($licenseSkus -contains "POWER_BI_PRO") { 1 } else { 0 }
        "VISIO 2"         = if ($licenseSkus -contains "VISIO_PLAN2_DEPT") { 1 } else { 0 }
        "PROJECT PLAN 3"  = if ($licenseSkus -contains "PROJECTPROFESSIONAL") { 1 } else { 0 }
        "AAD_PREMIUM"     = if ($licenseSkus -contains "AAD_PREMIUM") { 1 } else { 0 }
    }

    $userLicenses += $userInfo
}

Write-Progress -Activity "Exporting Data" -Status "Exporting user licenses to Excel" -PercentComplete 90
$userLicenses | Export-Excel -Path $FilePath -WorksheetName "Data" -AutoSize -ClearSheet -BoldTopRow
Set-MgDriveItemContent -DriveId $GroupDrive.Id -DriveItemId $DriveChild.Id -InFile $FilePath

Write-Progress -Activity "Finalizing" -Status "Disconnecting from Microsoft Graph" -PercentComplete 100
Disconnect-MgGraph

Write-Output "Script execution completed successfully."
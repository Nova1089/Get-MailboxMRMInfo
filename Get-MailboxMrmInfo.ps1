<#
Version 1.02

This script displays mailbox info regarding messaging records management (MRM), which used to be known as email life cycle (ELC).
(This has to do with archive, retention, and deletion policies.)

View info about things such as storage limits, archive mailbox, last ELC run, folder sizes, retention tags applied to folders, and MRM error logs.

Most of this script is leveraging knowledge and cmdlets given in this article:
https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/retention/troubleshoot-mrm-email-archive-deletion

This article can also help to understand the info obtained by this script:
https://www.itprotoday.com/email-and-calendaring/behind-scenes-managed-folder-assistant
#>

# functions
function Initialize-ColorScheme
{
    $script:successColor = "Green"
    $script:infoColor = "DarkCyan"
    $script:failColor = "Red"
    # warning color is yellow, but that is built into Write-Warning
}

function Show-Introduction
{
    Write-Host ("This script displays mailbox info regarding messaging records management (MRM), which used to be known as email life cycle (ELC). `n" +
                "(This has to do with archive, retention, and deletion policies.) `n`n" +

                "Most of this script is leveraging knowledge and cmdlets given in this article: `n" +
                "https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/retention/troubleshoot-mrm-email-archive-deletion `n`n" +

                "This article can also help to understand the info obtained by this script: `n" +
                "https://www.itprotoday.com/email-and-calendaring/behind-scenes-managed-folder-assistant `n") -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule $moduleName
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor $infoColor
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required." -ForegroundColor $infoColor
        $confirmInstall = Read-Host -Prompt "Would you like to install the module? (y/n)"
    }
    while ($confirmInstall -inotmatch "^\s*y\s*$") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Write-Host ("Please run script with admin privileges.`n" +
        "1. Open Powershell as admin.`n" +
        "2. CD into script directory.`n" +
        "3. Run .\scriptname`n") -ForegroundColor $failColor
        Read-Host "Press Enter to exit"
        exit
    }
}

function TryConnect-ExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue

    while ($null -eq $connectionStatus)
    {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor $infoColor
        Connect-ExchangeOnline -ErrorAction SilentlyContinue
        $connectionStatus = Get-ConnectionInformation

        if ($null -eq $connectionStatus)
        {
            Write-Warning "Failed to connect to Exchange Online."
            Read-Host "Press Enter to try again"
        }
    }
}

function TryConnect-MsolService
{
    Get-MsolDomain -ErrorVariable errorConnecting -ErrorAction SilentlyContinue | Out-Null

    if ($errorConnecting)
    {
        Read-Host "You must also connect to MsolService, press Enter to continue"
    }

    while ($errorConnecting)
    {
        Write-Host "Connecting to MsolService..." -ForegroundColor $infoColor
        Connect-MsolService -ErrorAction SilentlyContinue
        Get-MSolDomain -ErrorVariable errorConnecting -ErrorAction SilentlyContinue | Out-Null   

        if ($errorConnecting)
        {
            Read-Host -Prompt "Failed to connect to MsolService. Press Enter to try again"
        }
    }
}

function Prompt-Mailbox
{
    Write-Host "Enter mailbox display name or UPN:"
    do
    {
        $mailboxId = Read-Host
        if ([string]::IsNullOrWhiteSpace($mailboxId))
        {
            Write-Warning "Please enter mailbox display name or UPN."
            continue
        }

        $mailbox = Get-Mailbox -Identity $mailboxId

        if ($mailbox)
        {
            Write-Host "Mailbox found: $($mailbox.UserPrincipalName)" -ForegroundColor $successColor
        }
        else
        {
            Write-Warning "Mailbox not found: $mailboxId.`nPlease try again."
        }
        
    }
    while ($null -eq $mailbox)
    
    return $mailbox
}

function Show-MailboxInfo($mailbox)
{
    Write-Host "Getting mailbox info..." -ForegroundColor $infoColor

    $mailboxStats = $mailbox | Get-EXOMailboxStatistics
    $elcLogs = Get-ElcLogs $mailbox.UserPrincipalName

    if ($mailbox.ArchiveStatus -eq "Active")
    {
        $archiveMailboxStats = $mailbox | Get-EXOMailboxStatistics -Archive
        $archiveStorageConsumed = $archiveMailboxStats.TotalItemSize
    }
    else
    {
        $archiveStorageConsumed = ""
    }
    
    [PSCustomObject]@{
        UserPrincipalName = $mailbox.UserPrincipalName
        DisplayName = $mailbox.DisplayName
        Type = $mailbox.RecipientTypeDetails
        StorageConsumed = $mailboxStats.TotalItemSize
        StorageLimit = $mailbox.ProhibitSendReceiveQuota
        ArchiveStatus = $mailbox.ArchiveStatus
        AutoExpandingArchiveEnabled = $mailbox.AutoExpandingArchiveEnabled
        ArchiveStorageConsumed = $archiveStorageConsumed
        ArchiveStorageQuota = $mailbox.ArchiveQuota
        RetentionPolicy = $mailbox.RetentionPolicy
        LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
        RetentionHoldEnabled = $mailbox.RetentionHoldEnabled
        ElcProcessingDisabled = $mailbox.ElcProcessingDisabled
        ElcLastSuccessTimestamp = $elcLogs.ELCLastSuccessTimestamp
        "ELcLastRunTotalProcessingTime (Minutes)" = ([int]$elcLogs.ElcLastRunTotalProcessingTime / 60000)
        "ElcLastRunSubAssistantProcessingTime (Minutes)" = ([int]$elcLogs.ElcLastRunSubAssistantProcessingTime / 60000)
        ELcLastRunUpdatedFolderCount = $elcLogs.ElcLastRunUpdatedFolderCount
        ElcLastRunTaggedFolderCount = $elcLogs.ElcLastRunTaggedFolderCount
        ElcLastRunUpdatedItemCount = $elcLogs.ElcLastRunUpdatedItemCount
        ElcLastRunTaggedWithArchiveItemCount = $elcLogs.ElcLastRunTaggedWithArchiveItemCount
        ElcLastRunTaggedWithExpiryItemCount = $elcLogs.ElcLastRunTaggedWithExpiryItemCount
        ElcLastRunDeletedFromRootItemCount = $elcLogs.ElcLastRunDeletedFromRootItemCount
        ElcLastRunDeletedFromDumpsterItemCount = $elcLogs.ElcLastRunDeletedFromDumpsterItemCount
        ElcLastRunArchivedFromRootItemCount = $elcLogs.ElcLastRunArchivedFromRootItemCount
        ElcLastRunArchivedFromDumpsterItemCount = $elcLogs.ElcLastRunArchivedFromDumpsterItemCount
    }
}

function Get-ElcLogs($mailboxId)
{
    # ELC stands for email life cycle
    $logProps = Export-MailboxDiagnosticLogs $mailboxId -ExtendedProperties
    $xmlProps = [xml]($logProps.MailboxLog)
    $logs = $xmlProps.Properties.MailboxTable.Property | Where-Object { $_.Name -like "ELC*" }
    return Convert-XmlToObject $logs
}

function Convert-XmlToObject($xmlElementArray)
{
    $outputObject = New-Object -TypeName PSCustomObject

    foreach ($element in $xmlElementArray)
    {
        Add-Member -InputObject $outputObject -Name $element.Name -Value $element.Value -MemberType NoteProperty
    }

    return $outputObject
}

function Prompt-YesOrNo($question)
{
    Write-Host "$question`n[Y] Yes  [N] No"

    do
    {
        $response = Read-Host
        $validResponse = $response -imatch '^\s*[yn]\s*$' # regex matches y or n but allows spaces
        if (-not($validResponse)) 
        {
            Write-Warning "Please enter y or n."
        }
    }
    while (-not($validResponse))

    if ($response -imatch '^\s*y\s*$') # regex matches a y but allows spaces
    {
        return $true
    }
    return $false
}

function Show-MailboxFolderStatsSimplified($stats)
{
    $stats |
    Where-Object { $_.ContainerClass -ilike "*IPF.Note*" } |
    Format-Table -Property FolderPath, OldestItemReceivedDate, FolderType, FolderSize
}

function Export-MailboxFolderStatsDetailed($stats, $path)
{
    $folderStats = $stats | Where-Object { $_.ContainerClass -ilike "*IPF.Note*" }

    foreach ($folder in $folderStats)
    {
        [PSCustomObject]@{
            Name = $folder.Name
            FolderPath = $folder.FolderPath
            FolderType = $folder.FolderType
            TargetQuota = $folder.TargetQuota
            FolderSize = $folder.FolderSize
            FolderAndSubfolderSize = $folder.FolderAndSubfolderSize
            OldestItemReceivedDate = $folder.OldestItemReceivedDate
            OldestItemLastModifiedDate = $folder.OldestItemLastModifiedDate
            DeletePolicy = $folder.DeletePolicy
            ArchivePolicy = $folder.ArchivePolicy
            CompliancePolicy = $folder.CompliancePolicy
            RetentionFlags = $folder.RetentionFlags
        } | Export-CSV -Path $path -Append -NoTypeInformation
    }

    Write-Host "Finished exporting to $path" -ForegroundColor $successColor
}

function Show-MRMErrorLogs($mailbox, [switch]$raw)
{    
    # Although this says export, it's not exporting data to a file.
    $logs = Export-MailboxDiagnosticLogs -Identity $mailbox.UserPrincipalName -ComponentName MRM
    if ($null -eq $logs.MailboxLog)
    {
        Write-Host "There are no error logs to show." -ForegroundColor $infoColor
        return
    }

    if ($raw) # output the logs raw and unsimplified
    {
        $logs | Out-Host
        return
    }

    # regex to parse the convoluted log messages and extract the gist
    $regex = '(\d{1,2}\/\d{1,2}\/\d{4}.*)'
    $regex += '|(Exception:|InnerException:|\(Inner Exception #)[\S\s]*?hr=0x\d*, ec=\d*\)'
    $regex += '|(Exception:[\S\s]*?(?=\n\s*at\s))'

    $regexMatchInfo = Select-String -InputObject $logs.MailboxLog -Pattern $regex -AllMatches
    if ($regexMatchInfo)
    {        
        foreach ($match in $regexMatchInfo.Matches)
        {
            Write-Host $match.Value
            Write-Host "`n"
        }   
    }
    else
    {
        Write-Warning "Found no relevant logs when parsing. Logs will be displayed in their raw form."
        $logs | Out-Host
    }
}

function New-DesktopPath($fileName, $fileExt)
{
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $timeStamp = (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
    return "$desktopPath\$fileName $timeStamp.$fileExt"
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module("ExchangeOnlineManagement")
TryConnect-ExchangeOnline

$mailbox = Prompt-Mailbox
Show-MailboxInfo $mailbox

$viewFolderStats = Prompt-YesOrNo "Would you like to view mailbox folder stats?"
if ($viewFolderStats)
{
    $stats = Get-MailboxFolderStatistics -Identity $mailbox.UserPrincipalName -IncludeOldestAndNewestItems
    Show-MailboxFolderStatsSimplified $stats

    $exportFolderStats = Prompt-YesOrNo "Would you like to export mailbox folder stats detailed?"
    if ($exportFolderStats)
    {
        $path = New-DesktopPath -fileName "$($mailbox.DisplayName) Mailbox Folder Stats" -fileExt "csv"
        Export-MailboxFolderStatsDetailed -stats $stats -path $path
    }
}

$viewMrmErrorLogs = Prompt-YesOrNo "Would you like to see messaging records managemenet (MRM) error logs?"
if ($viewMrmErrorLogs)
{
    Show-MRMErrorLogs $mailbox

    $viewMrmErrorLogsRaw = Prompt-YesOrNo "Would you like to see the MRM logs raw and unsimplified?"
    if ($viewMrmErrorLogsRaw)
    {
        Show-MRMErrorLogs $mailbox -raw
    }
}

$startMfa = Prompt-YesOrNo "Would you like to start the managed folder assistant for this mailbox?"
if ($startMfa)
{
    Write-Host "Executing command: Start-ManagedFolderAssistant -Identity $($mailbox.UserPrincipalName)" -ForegroundColor $infoColor
    Start-ManagedFolderAssistant -Identity $mailbox.UserPrincipalName -ErrorVariable $mfaError

    if ($mfaError)
    {
        Write-Warning "There was an error starting the managed folder assistant:`n$mfaError"
    }
    else
    {
        Write-Host "Managed folder assistant was started successfully." -ForegroundColor $successColor
    }
}

Read-Host "Press Enter to exit"
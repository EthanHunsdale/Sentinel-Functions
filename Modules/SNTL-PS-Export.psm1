Function Export-StaffMailboxPermissions {
    <#
  .SYNOPSIS
  Create report of all mailbox permissions
  .DESCRIPTION
  Get all mailbox permissions, including folder permissions for all or a selected group of users
  .EXAMPLE
  Generate the mailbox report with Shared mailboxes, store the csv file in the script root location.
  Export-StaffMailboxPermissions -adminUPN john@contoso.com
  .EXAMPLE
  Get only the shared mailboxes
  Export-StaffMailboxPermissions -adminUPN john@contoso.com -sharedMailboxes only
  .EXAMPLE
  Get only the user mailboxes
  Export-StaffMailboxPermissions -adminUPN john@contoso.com -sharedMailboxes no
  .EXAMPLE
  Get the mailbox permissions without the folder (inbox and calendar) permissions
  Export-StaffMailboxPermissions -adminUPN john@contoso.com -folderPermissions:$false
  .EXAMPLE
  Get the mailbox permissions for a selection of users
  Export-StaffMailboxPermissions -adminUPN john@contoso.com -UserPrincipalName jane@contoso.com,alex@contoso.com

  .NOTES
  Version:         1.0
  Original Author: R. Mens - LazyAdmin.nl
  Modified By:     E Hunsdale
#>

    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = "Enter the Exchange Online or Global admin username"
        )]
        [String]$adminUPN,

        [Parameter(
            Mandatory = $false,
            HelpMessage = "Enter a single UserPrincipalName or a comma separted list of UserPrincipalNames"
        )]
        [String[]]$UserPrincipalName,

        [Parameter(
            Mandatory = $false,
            HelpMessage = "Get (only) Shared Mailboxes or not. Default include them"
        )]
        [ValidateSet("no", "only", "include")]
        [String]$sharedMailboxes = "include",

        [Parameter(
            Mandatory = $false,
            HelpMessage = "Enter path to save the CSV file"
        )]
        [String]$CSVpath = ""
    )

    Function Connect-ToEXO {
        <#
      .SYNOPSIS
          Connects to EXO when no connection exists. Checks for EXO v2 module
    #>
    
        process {
            # Check if EXO is installed and connect if no connection exists
            If ($null -eq (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
                Write-Host "Exchange Online PowerShell v2 module is requied, do you want to install it?" -ForegroundColor Yellow
        
                $install = Read-Host Do you want to install module? [Y] Yes [N] No 
                If ($install -match "[yY]") { 
                    Write-Host "Installing Exchange Online PowerShell v2 module" -ForegroundColor Cyan
                    Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
                } 
                else {
                    Write-Error "Please install EXO v2 module."
                }
            }

            If ($null -ne (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
                # Check if there is a active EXO sessions
                $psSessions = Get-PSSession | Select-Object -Property State, Name
                If (((@($psSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0) -ne $true) {
                    Connect-ExchangeOnline -UserPrincipalName $adminUPN
                }
            }
            Else {
                Write-Error "Please install EXO v2 module."
            }
        }
    }

    Function Get-DisplayName {
        <#
      .SYNOPSIS
        Get the full displayname (if requested) or return only the userprincipalname
    #>
        Param(
            [Parameter(
                Mandatory = $true
            )]
            $identity
        )

        Try {
            Return (Get-EXOMailbox -Identity $identity -ErrorAction Stop).DisplayName
        }
        Catch {
            Return $identity
        }
    }

    Function Get-SingleUser {
        <#
      .SYNOPSIS
        Get only the requested mailbox
    #>
        Param(
            [Parameter(
                Mandatory = $true
            )]
            $identity
        )
  
        Get-EXOMailbox -Identity $identity -Properties GrantSendOnBehalfTo, ForwardingSMTPAddress | 
        Select-Object UserPrincipalName, DisplayName, PrimarySMTPAddress, RecipientType, RecipientTypeDetails, GrantSendOnBehalfTo, ForwardingSMTPAddress
    }

    Function Get-AllMailboxes {
        <#
      .SYNOPSIS
          Get all the mailboxes for the report
    #>
        Process {
            Switch ($sharedMailboxes) {
                "include" { $mailboxTypes = "UserMailbox,SharedMailbox" }
                "only" { $mailboxTypes = "SharedMailbox" }
                "no" { $mailboxTypes = "UserMailbox" }
            }
  
            Get-EXOMailbox -ResultSize unlimited -RecipientTypeDetails $mailboxTypes -Properties GrantSendOnBehalfTo, ForwardingSMTPAddress | 
            Select-Object UserPrincipalName, DisplayName, PrimarySMTPAddress, RecipientType, RecipientTypeDetails, GrantSendOnBehalfTo, ForwardingSMTPAddress
        }
    }

    Function Get-SendOnBehalf {
        <#
      .SYNOPSIS
          Get Display name for each Send on Behalf entity
    #>
        Param(
            [Parameter(
                Mandatory = $true
            )]
            $mailbox
        )
  
        # Get Send on Behalf
        $SendOnBehalfAccess = @();
        If ($null -ne $mailbox.GrantSendOnBehalfTo) {
      
            # Get a proper displayname of each user
            $mailbox.GrantSendOnBehalfTo | ForEach-Object {
                $sendOnBehalfAccess += Get-DisplayName -identity $_
            }
        }
        Return $SendOnBehalfAccess
    }

    Function Get-SendAsPermissions {
        <#
      .SYNOPSIS
          Get all users with Send as Permissions
    #>
        Param(
            [Parameter(
                Mandatory = $true
            )]
            $identity
        )
        Write-Host $identity;
        $users = Get-EXORecipientPermission -Identity $identity | Where-Object { -not ($_.Trustee -match "NT AUTHORITY") -and ($_.IsInherited -eq $false) }
  
        $sendAsUsers = @();
    
        # Get a proper displayname of each user
        $users | ForEach-Object {
            $sendAsUsers += Get-DisplayName -identity $_.Trustee
        }
        Return $sendAsUsers
    }

    Function Get-FullAccessPermissions {
        <#
      .SYNOPSIS
          Get all users with Read and manage (full access) permissions
    #>
        Param(
            [Parameter(
                Mandatory = $true
            )]
            $identity
        )
    
        $users = Get-EXOMailboxPermission -Identity $identity | Where-Object { -not ($_.User -match "NT AUTHORITY") -and ($_.IsInherited -eq $false) }
  
        $fullAccessUsers = @();
    
        # Get a proper displayname of each user
        $users | ForEach-Object {
            $fullAccessUsers += Get-DisplayName -identity $_.User
        }
        Return $fullAccessUsers
    }

    Function Get-FolderPermissions {
        <#
      .SYNOPSIS
        Get Inbox folder permisions
    #>
        Param(
            [Parameter(Mandatory = $true)] $identity,
            [Parameter(Mandatory = $true)] $folder
        )
    
        $return = @{
            users      = @()
            permission = @()
            delegated  = @()
        }
  
        Try {
            $ErrorActionPreference = "Stop"; #Make all errors terminating
            $users = Get-EXOMailboxFolderPermission -Identity "$($identity):\$($folder)" | Where-Object { -not ($_.AccessRights -match "None") }
        }
        Catch {
            Return $return
        }
        Finally {
            $ErrorActionPreference = "Continue"; #Reset the error action pref to default
        }
  
        $folderUsers = @();
        $folderAccessRights = @();
        $folderDelegated = @();
    
        # Get a proper displayname of each user
        $users | ForEach-Object {
            $folderUsers += Get-DisplayName -identity $_.User
            $folderAccessRights += $_.AccessRights
            $folderDelegated += $_.SharingPermissionFlags
        }
  
        $return.users = $folderUsers
        $return.permission = $folderAccessRights
        $return.delegated = $folderDelegated
  
        Return $return
    }
 
    Function Get-AllMailboxPermissions {
        <#
      .SYNOPSIS
        Get all the permissions of each mailbox
          
        Permission are spread into 4 parts.
        - Read and Manage permission
        - Send as Permission
        - Send on behalf of permission
        - Folder permissions (inbox and calendar set by the user self)
    #>
        Process {
  
            If ($UserPrincipalName) {
        
                Write-Host "Collecting mailboxes" -ForegroundColor Cyan
                $mailboxes = @()
  
                # Get the requested mailboxes
                ForEach ($user in $UserPrincipalName) {
                    Write-Host "- Get mailbox $user" -ForegroundColor Cyan
                    $mailboxes += Get-SingleUser -identity $user
                }
            }
            Else {
                Write-Host "Collecting mailboxes" -ForegroundColor Cyan
                $mailboxes = Get-AllMailboxes
            }
      
            $i = 0
            Write-Host "Collecting permissions" -ForegroundColor Cyan
            $mailboxesqty = $mailboxes.Count
            $mailboxes | ForEach-Object {
       
                # Get Send on Behalf Permissions
                $sendOnbehalfUsers = Get-SendOnBehalf -mailbox $_
        
                # Get Fullaccess Permissions
                $fullAccessUsers = Get-FullAccessPermissions -identity $_.UserPrincipalName
  
                # Get Send as Permissions
                $sendAsUsers = Get-SendAsPermissions -identity $_.UserPrincipalName
  
                # Count number or records
                $sob = if ($sendOnbehalfUsers -is [array]) { $sendOnbehalfUsers.Count } else { if ($null -ne $sendOnbehalfUsers) { 1 }else { 0 } }
                $fa = if ($fullAccessUsers -is [array]) { $fullAccessUsers.Count } else { if ($null -ne $fullAccessUsers) { 1 }else { 0 } }
                $sa = if ($sendAsUsers -is [array]) { $sendAsUsers.Count } else { if ($null -ne $sendAsUsers) { 1 }else { 0 } }
        
                If ($folderPermissions.IsPresent) {
          
                    # Get Inbox folder permission
                    $inboxFolder = Get-FolderPermissions -identity $_.UserPrincipalName -folder $inboxFolderName
                    $ib = $inboxFolder.users.Count
  
                    # Get Calendar permissions
                    $calendarFolder = Get-FolderPermissions -identity $_.UserPrincipalName -folder $calendarFolderName
                    $ca = $calendarFolder.users.Count
                }
                Else {
                    $inboxFolder = @{
                        users      = @()
                        permission = @()
                        delegated  = @()
                    }
                    $calendarFolder = @{
                        users      = @()
                        permission = @()
                        delegated  = @()
                    }
                    $ib = 0
                    $ca = 0
                }
       
                $mostRecords = Find-LargestValue -sob $sob -fa $fa -sa $sa -ib $ib -ca $ca
  
                $x = 0
                If ($mostRecords -gt 0) {
            
                    Do {
                        [PSCustomObject]@{
                            "Display Name"            = $_.DisplayName
                            "Emailaddress"            = $_.PrimarySMTPAddress
                            "Mailbox type"            = $_.RecipientTypeDetails
                            "Read and manage"         = @($fullAccessUsers)[$x]
                            "Send as"                 = @($sendAsUsers)[$x]
                            "Send on behalf"          = @($sendOnbehalfUsers)[$x]
                            "Inbox folder"            = @($inboxFolder.users)[$x]
                            "Inbox folder Permission" = @($inboxFolder.permission)[$x]
                            "Inbox folder Delegated"  = @($inboxFolder.delegated)[$x]
                            "Calendar"                = @($calendarFolder.users)[$x]
                            "Calendar Permission"     = @($calendarFolder.permission)[$x]
                            "Calendar Delegated"      = @($calendarFolder.delegated)[$x]
                        }
                        $x++;
  
                        $currentUser = $_.DisplayName
                        If ($mailboxes.Count -gt 1) {
                            Write-Progress -Activity "Collecting mailbox permissions" -Status "Current count: $i of $mailboxesqty" -PercentComplete (($i / $mailboxes.Count) * 100) -CurrentOperation "Processing mailbox: $currentUser"
                        }
                    }
                    While ($x -ne $mostRecords)
                }
                $i++;
            }
        }
    }

    # Connect to Exchange Online
    Connect-ToEXO

    If ($CSVpath) {
        # Get mailbox status
        Get-AllMailboxPermissions | Export-CSV -Path $CSVpath -NoTypeInformation -Encoding UTF8
        If ((Get-Item $CSVpath).Length -gt 0) {
            Write-Host "Report finished and saved in $CSVpath" -ForegroundColor Green
        } 
        Else {
            Write-Host "Failed to create report" -ForegroundColor Red
        }
    }
    Else {
        Get-AllMailboxPermissions
    }

    # Close Exchange Online Connection
    $close = Read-Host Close Exchange Online connection? [Y] Yes [N] No 

    If ($close -match "[yY]") {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    }
}


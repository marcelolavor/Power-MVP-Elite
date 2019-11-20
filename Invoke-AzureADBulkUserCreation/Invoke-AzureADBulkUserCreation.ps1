<#
//-----------------------------------------------------------------------

//     Copyright (c) {charbelnemnom.com}. All rights reserved.

//-----------------------------------------------------------------------

.SYNOPSIS
Create Azure AD User Account.

.DESCRIPTION
Azure AD Bulk user creation and assign the new users to an Azure AD group.

.PROGRAM-INFORMATION
File Name : Invoke-AzureADBulkUserCreation.ps1
Author    : Charbel Nemnom
Date      : 27-February-2018
Requires  : PowerShell Version 3.0 or above
Module    : AzureAD Version 2.0.0.155 or above
Product   : Azure Active Directory

.PREVIOUS-RELEASE
Version   : 1.7
Update    : 30-July-2019
Author    : Charbel Nemnom

.RELEASE-NOTES
Version   : 1.8
Update    : 19-November-2019
Author    : Marcelo Lavor

.NEW FEATURES
- Config file definition and load for global variables
- Corporative domains registry addition
- Default parameters stabilished
- Azure Multitenant feature addition
- User invitation feature implementation
- User and group creation sync with timer-sleep 
- User domain analysis for New or Invited user decision
- InviteForwardURL configuration
- Groups members list report in verbose mode
- File report with added users
- Broad verbose implementation for issues tracking
- Several comments addition for better code reading
- Initial changes to become a module acting as az cmdlet

.REMARK: 
- Does not run under PowerShell Version 6.x

.OPEN BUGS
- Still does not add users to groups right after have been created due to some latency or thread visibility issues

.LINK
To provide feedback or for further assistance please visit:
https://charbelnemnom.com

.EXAMPLE-1
./Invoke-AzureADBulkUserCreation -FilePath <FilePath> -Credential <Username\Password> -Verbose
This example will import all users from a CSV File and then create the corresponding account in Azure Active Directory.
The user will be asked to change his password at first log on.

.EXAMPLE-2
./Invoke-AzureADBulkUserCreation -FilePath <FilePath> -Credential <Username\Password> -AadGroupName <AzureAD-GroupName> -Verbose
This example will import all users from a CSV File and then create the corresponding account in Azure Active Directory.
The user will be a member of the specified Azure AD Group Name.
The user will be asked to change his password at first log on.

#>

[CmdletBinding()]
Param(
    [Parameter(Position = 0, Mandatory = $false, HelpMessage = 'Specify the path of the CSV file')]
    [Alias('CSVFile')]
    [string]$FilePath="bulk-data.csv",
    [Parameter(Position = 1, Mandatory = $false, HelpMessage = 'Specify Credentials')]
    [Alias('User')]
    [string]$CredUser,
    #MFA Account for Azure AD Account
    [Parameter(Position = 2, Mandatory = $false, HelpMessage = 'Specify if account is MFA enabled')]
    [Alias('2FA')]
    [Switch]$MFA,
    [Parameter(Position = 2, Mandatory = $false, HelpMessage = 'Specify Azure AD Group Name')]
    [Alias('AADGN')]
    [string]$AadGroupName
)

# LOAD CONFIG FILE
$scriptFiles = Get-ChildItem "$PSScriptRoot\config\*.ps1" -Recurse

foreach ($script in $scriptFiles)
{
    try
    {
        . $script.FullName
        Write-Verbose "Config file $($script.FullName) loaded."
    }
    catch [System.Exception]
    {
        throw
    }
}

# GLOBAL VARIABLES SET
$DomainName = $global:config.varTenant
If (!$CredUser) {
    $CredUser = $global:config.varCredential
    Write-Verbose "Using default credential $CredUser"
}

#Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\PowerShell\1\ShellIds" -Name ConsolePrompting -Value $false
$Credential = Get-Credential -Credential $CredUser

# INSTALL AZURE-AD LIB
Function Install-AzureAD {
    Set-PSRepository -Name PSGallery -Installation Trusted -Verbose:$false
    Install-Module -Name AzureAD -AllowClobber -Verbose:$false
}

# TEST TENANT
function Test-DomainExistsInAad {

      param(
             [Parameter(mandatory=$true)]
             [string]$DomainName
       )

      $response = Invoke-WebRequest -Uri "https://login.microsoftonline.com/getuserrealm.srf?login=user@$DomainName&xml=1"

     if($response -and $response.StatusCode -eq 200) {

           $namespaceType = ([xml]($response.Content)).RealmInfo.NameSpaceType

           if ($namespaceType -eq "Unknown")
           {
                return $false
           }
           else{
                Write-Verbose "User domain registred as $namespaceType"
                return $true
           }

    } else {

        Write-Error -Message 'Domain could not be verified. Please check your connectivity to login.microsoftonline.com'

    }

}

# IMPORT MODULES
Try {
    Import-Module -Name AzureAD -ErrorAction Stop -Verbose:$false | Out-Null
}
Catch {
    Write-Verbose "Azure AD PowerShell Module not found..."
    Write-Verbose "Installing Azure AD PowerShell Module..."
    Install-AzureAD
}

# CONNECT AZURE
Write-Verbose $sep

Try {
    Write-Verbose "Connecting to Azure AD..."
    if ($MFA) {
        Connect-AzureAD -ErrorAction Stop | Out-Null
    }
    Else {
        Connect-AzureAD -TenantId $DomainName -Credential $Credential -ErrorAction Stop | Out-Null
    }
}
Catch {
    Write-Verbose "Cannot connect to Azure AD. Please check your credentials. Exiting!"
    Break
}

# INSERT DATA
Write-Verbose $sep

# IMPORT CSV DATA
Try {
    $CSVData = @(Import-CSV -Path $FilePath -ErrorAction Stop)
    Write-Verbose "Successfully imported entries from $FilePath"
    Write-Verbose "Total no. of entries in CSV are : $($CSVData.count)"
} 
Catch {
    Write-Verbose "Failed to read from the CSV file $FilePath Exiting!"
    Break
}

# CREATE FILE FOR ADDED-USERS LIST REPORT
$stream = [System.IO.StreamWriter] "$PSScriptRoot\added-users.csv"
$stream.WriteLine("OID, EMAIL")

# LOAD CSV DATA
$CheckedDomains = @{}
Foreach ($Entry in $CSVData) {
    # Verify that mandatory properties are defined for each object
    $DisplayName = $Entry.DisplayName
    $MailNickName = $Entry.MailNickName
    $UserPrincipalName = $Entry.UserPrincipalName
    $Password = $Entry.PasswordProfile
    
    If (!$DisplayName) {
        Write-Warning '$DisplayName is not provided. Continuing to the next record'
        Continue
    }

    If (!$MailNickName) {
        Write-Warning '$MailNickName is not provided. Continuing to the next record'
        Continue
    }

    If (!$UserPrincipalName) {
        Write-Warning '$UserPrincipalName is not provided. Continuing to the next record'
        Continue
    }

    If (!$Password) {
        Write-Warning "Password is not provided for $DisplayName in the CSV file!"
        $Password = Read-Host -Prompt "Enter desired Password" -AsSecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $Password
        $PasswordProfile.EnforceChangePasswordPolicy = 1
        $PasswordProfile.ForceChangePasswordNextLogin = 1
    }
    Else {
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $Password
        $PasswordProfile.EnforceChangePasswordPolicy = 1
        $PasswordProfile.ForceChangePasswordNextLogin = 1
    }   

    #Verify that the domain is registered in AAD
    $domain = $UserPrincipalName.SubString($UserPrincipalName.IndexOf("@") + 1)

    $domainExists = $False
    if (!$CheckedDomains.ContainsKey($domain))
    {
        $CheckedDomains.Add($domain, (Test-DomainExistsInAad($domain)))
    }
    $domainExists = $CheckedDomains[$domain];

    if(!$domainExists)
    {
        Write-Warning "Domain for user $UserPrincipalName is not registered in Azure AD. Continuing to next user."
        Continue
    }
    Write-Verbose "Domain $domain is valid."
    
    #See if the user exists.
    Write-Verbose $sep

    Try{
        # See if the user does not belongs to corporative domain to handle internal $UserPrincipalName with #EXT#
        If (!$global:config.varCorpDomain.ContainsKey($domain)) {
          # Get "#EXT#" UserPrincipalName
          $UserEmail = $UserPrincipalName
          $UserPrincipalName = $UserPrincipalName -replace "@", '_'
          $UserPrincipalName = "$UserPrincipalName#EXT#@$DomainName"
        }
        Write-Verbose "Look for userPrincipalName $UserPrincipalName"
        $ADuser = Get-AzureADUser -Filter "userPrincipalName eq '$UserPrincipalName'"
    }
    Catch{}

    #If so then move along, otherwise create the user.
    If ($ADuser)
    {
        Write-Verbose "User found and will be added to group if specified."
    }
    Else
    {
        Write-Verbose "DisplayName: $($DisplayName)"
        Write-Verbose "GivenName: $($Entry.GivenName)"
        Write-Verbose "Surname: $($Entry.Surname)"
        Write-Verbose "MailNickName: $MailNickName"
        Write-Verbose "UserPrincipalName: $UserPrincipalName"
        Write-Verbose "City: $($Entry.City)"
        Write-Verbose "State: $($Entry.State)"
        Write-Verbose "Country: $($Entry.Country)"
        Write-Verbose "Department: $($Entry.Department)"
        Write-Verbose "JobTitle: $($Entry.JobTitle)"
        Write-Verbose "Mobile: $($Entry.Mobile)"
        Write-Verbose "UsageLocation: $($Entry.UsageLocation)"

        if (!$global:config.varCorpDomain.ContainsKey($domain))
        {
            # INVITE USER
            Write-Verbose "Inviting user..."
            $varInviteRedirectUrl = $global:config.varAppURIbyGroups[$Entry.GroupNames]
            Write-Verbose "Invite App URL: $varInviteRedirectUrl"
            
            Try {    
                New-AzureADMSInvitation `
                    -InvitedUserDisplayName $DisplayName `
                    -InvitedUserEmailAddress $UserEmail `
                    -SendInvitationMessage $true `
                    -InviteRedirectUrl $varInviteRedirectUrl | Out-Null
                    }
            Catch {
                Write-Error "$DisplayName : Error occurred while creating Azure AD Account. $_"
                Continue
            }
        }
        else
        {
            # ADD NEW USER
            Write-Verbose "Creating new user..."
            Try {    
                New-AzureADUser `
                    -AccountEnabled $true `
                    -DisplayName $DisplayName `
                    -PasswordProfile $PasswordProfile `
                    -GivenName $Entry.GivenName `
                    -Surname $Entry.Surname `
                    -MailNickName $MailNickName `
                    -UserPrincipalName $UserPrincipalName `
                    -City $Entry.City `
                    -State $Entry.State `
                    -Country $Entry.Country `
                    -Department $Entry.Department `
                    -JobTitle $Entry.JobTitle `
                    -Mobile $Entry.Mobile `
                    -UsageLocation $Entry.UsageLocation | Out-Null
                    }
            Catch {
                Write-Error "$DisplayName : Error occurred while creating Azure AD Account. $_"
                Continue
            }
        }
        
        #Make sure the user exists now.
        
        $i=1
        DO
        {
            Try
            {
                #$ADuser = Get-AzureADUser -Filter "userPrincipalName eq '$UserPrincipalName'"
                $ADuser = Get-AzureADUser -ObjectId '$UserPrincipalName'
            }
            Catch
            {
                Write-Verbose "Waiting for 5sec to get $UserPrincipalName"
                $i++
                Start-Sleep -Seconds 5
                continue
            }
        } while ((!$ADuser) -and ($i -le 10))

        if ($ADuser)
        {
            Write-Verbose "$($ADuser.DisplayName): AAD Account is created successfully with ObjectId $($ADuser.ObjectID)"             
        } else {
            Write-Warning "$DisplayName : Newly created account could not be found. Continuing to next user. $_"
            Continue
        }
    }

    #Add the user to a group, creating it if necessary.
    Write-Verbose $sep

    If ($Entry.GroupNames) {
        $GroupNames = ($Entry.GroupNames).Split(";")

        Foreach ($GroupName in $GroupNames)
        {
            Try {   
                $AadGroup = Get-AzureADGroup -SearchString "$GroupName"
            }
            Catch {                
            }

            If (!$AadGroup)
            {
                Try {   
                $AadGroup = New-AzureADGroup -DisplayName "$GroupName" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
                }
                Catch {                
                    Write-Warning "Failed to create group $GroupName. Continuing to the next group."
                    Continue
                }
            }

            #Determine if user is already part of the group
            $GroupMembers = (Get-AzureADGroupMember -ObjectId $AadGroup.ObjectID | Select ObjectId)            
            If ($GroupMembers -Match $ADuser.ObjectID){
                Write-Verbose "$UserPrincipalName is already a member of Azure AD Group $GroupName"
            }
            Else
            {

                Try {   
                        Add-AzureADGroupMember -ObjectId $AadGroup.ObjectID -RefObjectId $ADuser.ObjectID 
                        Write-Verbose "User $DisplayName assigned to Azure AD Group $GroupName"
                        
                        # WRITE TO REPORT ADDED-USER LIST
                        $stream.WriteLine("$($MemberName.ObjectId), $($MemberName.Mail)")
                    }
                    Catch {                
                        Write-Warning "Failed to add $DisplayName to Azure AD Group $GroupName. Continuing to the next group."
                        Continue
                    }
            }
        }

        # GENERATE FINAL GROUP REPORT
        Foreach ($GroupName in $GroupNames)
        {
            Try {   
                $AadGroup = Get-AzureADGroup -SearchString "$GroupName"
            }
            Catch {
            }
            $GroupMembers = (Get-AzureADGroupMember -ObjectId $AadGroup.ObjectID | Select ObjectId, mail)            

            Write-Verbose $sep
            Write-Verbose "'$($AadGroup.DisplayName)' group members"
            Foreach ($MemberName in $GroupMembers)
            {
                Write-Verbose "$($MemberName.ObjectId), $($MemberName.Mail)"
            }
            Write-Verbose $sep
        }
    }
}

# CLOSE FILE OF ADDED-USERS LIST REPORT
if ($stream) { $stream.close()}
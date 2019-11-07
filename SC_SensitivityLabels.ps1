<# 
#Version : 2.1
# 12.Sep.2019 = Andrew Auret
# Converted to Pester
## $currentDate =  (Get-Date -Format "yyyy-MM-dd_HHmm").ToString()
## Invoke-Pester -Script @{Path = '.\O365ValidationEXO-OD4Bver2.1.ps1'; Parameters = @{Service = 'Exchange';MFA = $true}} -Show All -OutputFormat NUnitXml -OutputFile .\$currentDate.xml  
## Invoke-Command -ScriptBlock{.\ReportUnit.exe "$($currentDate).xml"}
## Update Parameters in script block to include MFA ($True) or Basic Auth ($false)
# Update SPODomain to reflect correct Tenant (line 46)

#Version : 2.0
# 02.Sep.2019 = Gabriel Antohi
# Added a run parameter to skip the EXO-SPO Connect sessions: type Yes
    # To Connect using O365 Managemtn shells press any key
# Extract the exact UPN and the Primary Email Address for the user for comparison (as Identity can be one or another in get-recipient)
# Separate the verification for Forwarding SMTP Address into incorrect and not set
# Added OD4B validation against the UPN not $CurrentUser variable and catch any "Cannot find Site" error to detect non-existent OD4B site 
# Added the variable $IdentityColumn = "Target Mailbox", as in the GMAIL batch file, for easy manipulation of batches
# Catch the errors and write it in the output file
# Make decision about the input file: .CSV or .TSV - based on the file name extension 

O365ValidationEXO-OD4B.ps1 : This script performs Pre-flight Validation for GDRIVE and GMAIL Migrations

WCC Tenant URL = https://warwickshiregovuk-admin.sharepoint.com/
WCC Domain = warwickshire.gov.uk
SPO Domain = warwickshiregovuk
FWD Address domain = @pilot.warwickshire.gov.uk
#$TenantUrl = Read-Host "Enter the SharePoint Online Tenant Admin Url"

For this script to run properly you need AAD, EXO, SPO Management Shells 

Sample input file:
Target Mailbox
_________________________
User1@warwickshire.gov.uk

#>
#--------------------------------------------------------------------------------------------------------
[OutputType()]
[CmdletBinding(DefaultParameterSetName)]
Param (
    [Parameter(Mandatory = $True, Position = 1)]
    [ValidateSet('AzureAD', 'Exchange',  'SharePoint')]
    [string[]]$Service,
    [Parameter(Mandatory = $False, Position = 2)]
    [string[]]$SPODomain = "warwickshiregovuk"    , 
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = 'Credential')]
    [PSCredential]$Credential,
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = 'MFA')]
    [Switch]$MFA
)

#region Functions

function Initialize-Modules() {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [String] $Name
    )

    Write-Verbose "Checking for $Name module..."

    If ($Name -eq "CreateExoPSSession"){
            $Module = Get-Module -Name $Name
    }
    else {
        $Module = Get-Module -Name $Name -ListAvailable
    }

    if ($null -eq $Module) {
        Write-Verbose ""
        Write-Verbose "$($Name) Powershell module not present..." 
        Write-Verbose "Installing $($Name)" 
        Write-Verbose ""
        If ($Name -eq "CreateExoPSSession"){
            Import-Module (Get-ChildItem -Path $scriptFolder\Exchange_Online_Module\ -Filter '*ExoPowershellModule.dll' -Recurse | ForEach-Object{(Get-ChildItem -Path $_.Directory -Filter CreateExoPSSession.ps1)} | Sort-Object LastWriteTime | Select-Object -Last 1).FullName  -Verbose
        }
        else {
                Write-Verbose ""
                Write-Verbose "$($Name) Powershell module not present..." 
                Write-Verbose "Installing $($Name)" 
                Write-Verbose ""
                Install-Module -Name $Name -Scope CurrentUser -Force -Verbose
        }
    }
    Else{
        Write-Verbose "$($Name) Powershell module is installed"
    }
}
Function Get-FileName($InitialDirectory, $Title)
{   
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = $Title
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.Filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
    $OpenFileDialog.ShowHelp = $true
}

function WCCBatch($Users)
{
Write-Verbose "Starting Validation of all Users from Input File..." 

Write-Verbose "Getting a list of all the users from the domain 1st "

$Users | ForEach-Object {
    Try
    {	 
        $currentUser = $_.$IdentityColumn
        Describe "Checks for User:$($currentUser)"{
            #

            $expression = "EmailAddresses -eq '$currentUser'"
            $UserRecipientAttributes = Get-Recipient -Filter $expression | Select-Object Name, RecipientType
            $UserAttributes = Get-MailBox -Filter $expression  |  Select-Object PrimarySmtpAddress, UserPrincipalName, ForwardingSMTPAddress
            $userTargetAddress = $UserAttributes.ForwardingSMTPAddress
            $UPNAddress = $UserAttributes.UserPrincipalName
            $emailAddress = $UserAttributes.PrimarySmtpAddress
            $PersonalName = $UPNAddress -replace "[^a-zA-Z0-9,-]", "_" # for OD4B
            #
            #
            It "Verifying that $($currentUser) exists in the $($SPODomain) O365 Tenant"{
                $UserRecipientAttributes.Name | Should Not Be $Null
            }
            #
            It "Verifying that $($currentUser) has a Mailbox provisioned on the $($SPODomain) Office 365 Tenant"{
                $UserRecipientAttributes.RecipientType | Should Be "UserMailBox"
            }


            It "Verifying that $($currentUser) has a ForwardingSMTPAddress set [Not Null]"{
                $userTargetAddress | Should Not Be $Null
            }
            #
            It "Verifying that $($currentUser) has the ForwardingSMTPAddress set correctly [Should match $($ForwardingDomain)]"{
                If ($Null -eq $userTargetAddress){
                    Set-ItResult -Skipped -Because 'The ForwardingSMTPAddress is $Null or Empty'
                }
                ($userTargetAddress -like "*@"+$ForwardingDomain) | should Be $True
            }
            #
            It "Verifying that $($currentUser) Email Address Matches the UPN Address"{
                ($emailAddress -eq $UPNAddress) | should Be $True
            }
            #
            It "Verifying that $($currentUser) has a OD4B site on $($SPODomain) O365 Tenant"{
                $OD4BSiteURL = "https://$SPOdomain-my.sharepoint.com/personal/$PersonalName"
                ([bool] (Get-SPOSite -Identity $OD4BSiteURL -ErrorAction SilentlyContinue | Select-Object owner)) | should Be $True
            }
        }


    }
    Catch [Exception]
    {
        $Exception = $_.Exception
        if ($Exception -like "*Cannot get site*")
        {
            Write-Warning "User doens't have an OD4B site on O365 WCC Tenant."
        }
        else
        {
            Write-Host "There was an error runing the script in Office 365 WCC Tenant." -ForegroundColor Red
        }
    }
}

Write-Host "Script Execution Completed Successfully." -ForegroundColor Green
exit
}
#endregion
#--------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------
$VerbosePreference = "Continue" 

$users= $null
$OutputFile = $null
$InputFile = $null
$InputFileType = $null
$scriptPath = $myInvocation.MyCommand.Path
$scriptFolder = Split-Path $scriptPath;
$IdentityColumn = "Target Mailbox" #align this value with what you have in the CSV batch file for GMAIL
$currentDate =  (Get-Date -Format "yyyy-MM-dd_HHmm").ToString()
#$OutputFile = $scriptFolder + "\" + $currentDate + '_Result.csv'

#$SPODomain = "warwickshiregovuk"
$ForwardingDomain = "pilot.warwickshire.gov.uk"
$TenantUrl = "https://$SPODomain-admin.sharepoint.com/"
#
#

If ($MFA.IsPresent){
    Initialize-Modules("CreateExoPSSession")
}
Initialize-Modules("Microsoft.Online.SharePoint.PowerShell")
Initialize-Modules("Pester")
#
If (!($MFA.IsPresent))
{
    Write-Verbose "Gathering Credentials for non-MFA sign on"
    $Credential = Get-Credential -Message "Please enter your Office 365 credentials"
}
#
If (!(Get-PSSession).ConfigurationName -eq "Microsoft.Exchange"){
    If ($MFA.IsPresent){
        . $scriptFolder"\Exchange_Online_Module\CreateExoPSSession.ps1"
        Connect-IPPSSession -ConnectionUri https://outlook.office365.com/powershell-liveid 
        Push-Location $scriptFolder
        Connect-SPOService -Url $TenantUrl
    }
    else {
        $newPSSessionSplat = @{
            ConfigurationName = 'Microsoft.Exchange'
            ConnectionUri   = "https://outlook.office365.com/powershell-liveid"
            Authentication    = 'Basic'
            Credential       = $Credential
            AllowRedirection  = $true
        }
        $Session = New-PSSession @newPSSessionSplat
        Write-Verbose "Connecting to Exchange Online"
        Import-PSSession $Session -AllowClobber
        Connect-SPOService -Url $TenantUrl -Credential $Credential
    }
}

$InputFile = Get-FileName -initialDirectory $scriptFolder -Title "Please Select the Input File to GMail and GDrive Pre-Flight Tool"
if($InputFile -eq "")
{
    Write-Host ""
    Write-Host -ForegroundColor Red "No File Found. Please select File again. Quitting ..... "
    Exit
}
elseif ($InputFile -like "*.csv")
{
    $Users = Import-Csv $InputFile | Select-Object $IdentityColumn
}
else
{
    $Users = Import-Csv $InputFile -Delimiter "`t" | Select-Object $IdentityColumn
}

Write-Host "Input File : " $InputFile
Write-Host "Result File : " $OutputFile

WCCBatch($Users)



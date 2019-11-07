###########################################################################################################
# Script Disclaimer
###########################################################################################################
# This script is not supported under any Microsoft standard support program or service.
# This script is provided AS IS without warranty of any kind.
# Microsoft disclaims all implied warranties including, without limitation, any implied warranties of
# merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
# performance of this script and documentation remains with you. In no event shall Microsoft, its authors,
# or anyone else involved in the creation, production, or delivery of this script be liable for any damages
# whatsoever (including, without limitation, damages for loss of business profits, business interruption,
# loss of business information, or other pecuniary loss) arising out of the use of or inability to use
# this script or documentation, even if Microsoft has been advised of the possibility of such damages.

<#
.SYNOPSIS
  Control Script for querying Security and Compliance Center - Sensitivy Labels and Policies 
.DESCRIPTION
  Demo Control Script for querying Security and Compliance Center - Sensitivy Labels and Policies
.PARAMETER  Platform
  Description: Name of the OS being targeted
  Possible values: AndroidE, iOS or Win10	
.EXAMPLE
  .\Intune-MAM.ps1 -Platform AndroidE -GraphApiVersion Beta -Mode get -Path '.\MAM\AppProtection\Prod\Production Android Browser.json'
  Gets the current policy settings for a specific Policy 
.INPUTS
   <none>
.OUTPUTS
   <none>
.NOTES
    Script Name     : SC_SensitivityLabels.ps1
    Requires        : Powershell Version 5.1
    Tested          : Powershell Version 5.1
    Author          : Andrew Auret
    Email           : 
    Version         : 1.0
    Date            : 2019-11-07 (ISO 8601 standard date notation: YYYY-MM-DD)    
#>

#######################################################################################################################
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



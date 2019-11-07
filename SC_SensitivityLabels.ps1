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
#--------------------------------------------------------------------------------------------------------#
[OutputType()]
[CmdletBinding(DefaultParameterSetName)]
Param (
    [Parameter(Mandatory = $False, Position = 0, ParameterSetName = 'Credential')]
    [PSCredential]$Credential,

    [Parameter(Mandatory = $False, Position = 0, ParameterSetName = 'MFA')]
    [Switch]$MFA,

    [Parameter(Mandatory=$true)]
    [ValidateSet("get","test")]
    [string]$Mode,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Label","LabelPolicy")]
    [string]$Type

) 


#--------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------
Process{
    $VerbosePreference = "Continue" 
    $scriptPath = $myInvocation.MyCommand.Path
    $scriptFolder = Split-Path $scriptPath
    #
    
    Write-Verbose "Checking for Common_Functions module..."
    
    $CommonModule = Get-Module -Name "Common_Functions" -ListAvailable
    
    if ($null -eq $CommonModule) {
        Write-Verbose ""
        Write-Verbose "Common_Functions Powershell module not installed..." 
        Write-Verbose "Installing Common_Functions module" 
        Write-Verbose ""
        Import-Module $scriptFolder\Common_Functions.psd1 -Force -Verbose
    }
    Else{
        Write-Verbose "Common_Functions Powershell module is installed"
    }
    #

    If ($MFA.IsPresent){
        Initialize-Modules("CreateExoPSSession")
    }
    If (!($MFA.IsPresent))
    {
        If (!($Credential.UserName)){
            Write-Verbose "Gathering Credentials for non-MFA sign on"
            $Credential = Get-Credential -Message "Please enter your Office 365 credentials"
        }
    }
    #
    If (!(Get-PSSession).ConfigurationName -eq "Microsoft.Exchange"){
        If ($MFA.IsPresent){
            . $scriptFolder"\Exchange_Online_Module\CreateExoPSSession.ps1"
            Connect-IPPSSession
            Push-Location $scriptFolder
        }
        else {
            $newPSSessionSplat = @{
                ConfigurationName = 'Microsoft.Exchange'
                ConnectionUri   = "https://ps.compliance.protection.outlook.com/powershell-liveid/"
                Authentication    = 'Basic'
                Credential       = $Credential
                AllowRedirection  = $true
            }
            try{
                $Session = New-PSSession @newPSSessionSplat -ErrorAction Stop
                Write-Verbose "Connecting to Security & Compliance Center"
                Import-PSSession $Session -AllowClobber -ErrorAction Stop
            }
            catch{
                Write-Output $_.Exception.Message
                Write-Error "Please connect using Multi-Factor authentication instead using the -MFA switch"
                break            
            }
        }
    }
    #
    switch ($Mode) {
        "get" {
            switch ($Type) {
                "Label" {
                    Write-Output "Getting the content of the current Sensitivity Labels"
                    $labels = Get-Label
                    foreach($label in $labels){
                        $labelpolicyRule = Get-Labelpolicyrule | Where-Object {$_.LabelName -eq $label.Name}
                        Write-Output "Name         : $($label.Name)"
                        Write-Output "Created by   : $($label.CreatedBy)"
                        Write-Output "Last modified: $($label.LastModifiedBy)"
                        Write-Output "Display Name : $($label.DisplayName)"
                        Write-Output "Tooltip      : $($label.Tooltip)"
                        Write-Output "Description  : $($label.Comment)"
                        $labelEncryptAction = ($labelpolicyRule | Where-Object {$_.LabelActionName -eq 'encrypt'}).LabelActionName
                        If ($labelEncryptAction -eq "encrypt"){
                            Write-Output "Encryption   : On"
                        }
                        else {
                            Write-Output "Encryption   : Off"
                        }
                        Write-Output ""
                    }

                  }
                "LabelPolicy" {

                }
            }
        }
        "test" {
            switch ($Type) {
                "Label" {

                  }
                "LabelPolicy" {
                    
                }
            }
        }
    }
}
#
    


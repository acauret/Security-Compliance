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
  Control Script for creating and querying Security and Compliance Center - Retention Labels and Policies 
.DESCRIPTION
  Demo Control Script for creating and querying Security and Compliance Center - Retention Labels and Policies
.PARAMETER  Mode
  Description: Determines the mode of operation of the script
  Possible values: get, create
.PARAMETER  Type
Description: Determines what aspect of Sensitivity labels to query
Possible values: Label, LabelPolicy
.PARAMETER  Credential
  Description: For use with connecting the Security and Compliance center using Basic Auth
  Possible values: User Principal Name
.PARAMETER  MFA
Description: Switch to specifiy using MFA with modern auth setting
Possible values: -MFA
.EXAMPLE
  .\SC_Retention.ps1 -Mode get -Type Label -MFA
  Connects to the S&C Center to query the current label settings using Modern auth
.EXAMPLE
.\SC_Retention.ps1 -Mode get -Type LabelPolicy -MFA
Connects to the S&C Center to query the current labelPolicy settings using Modern auth
.EXAMPLE
.\SC_Retention.ps1 -Mode get -Type LabelPolicy
Connects to the S&C Center to query the current labelPolicy settings using Basic auth
.EXAMPLE
.\SC_Retention.ps1 -Mode get -Type Label
Connects to the S&C Center to query the current label settings using Basic auth
.EXAMPLE
get-help .\SC_Retention.ps1 -Detailed
Displays the help file
.INPUTS
   <none>
.OUTPUTS
   <none>
.NOTES
    Script Name     : SC_Retention.ps1
    Requires        : Powershell Version 5.1, Windows Remote Management (WinRM) on your computer needs to allow basic authentication
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
    [Parameter(Mandatory=$true, Position = 0)]
    [ValidateSet("get", "create")]
    [string]$Mode,

    [Parameter(Mandatory=$true, Position = 1)]
    [ValidateSet("Label","LabelPolicy")]
    [string]$Type,

    [Parameter(Mandatory = $False)]
    [PSCredential]$Credential,

    [Parameter(Mandatory = $False)]
    [Switch]$MFA,
    [switch]$ResultCSV
) 

DynamicParam
{
    if (($Mode.Equals("create"))-and ($Type.Equals("Label")))
    {
      $LabelListCSV = 'LabelListCSV'
      $attributes = New-Object -Type `
        System.Management.Automation.ParameterAttribute
      $attributes.ParameterSetName = "DefaultSet"
      $attributes.Mandatory = $true
      $attributeCollection = New-Object `
        -Type System.Collections.ObjectModel.Collection[System.Attribute]

      # Add the attributes to the attributes collection
      $attributeCollection.Add($attributes)

      $dynParam1 = New-Object -Type `
        System.Management.Automation.RuntimeDefinedParameter($LabelListCSV, [string],
          $attributeCollection)
  
      $paramDictionary = New-Object `
        -Type System.Management.Automation.RuntimeDefinedParameterDictionary
      $paramDictionary.Add($LabelListCSV, $dynParam1)
      return $paramDictionary
    }
    if (($Mode.Equals("create"))-and ($Type.Equals("LabelPolicy")))
    {
      $PolicyListCSV = 'PolicyListCSV'
      $attributes = New-Object -Type `
        System.Management.Automation.ParameterAttribute
      $attributes.ParameterSetName = "DefaultSet"
      $attributes.Mandatory = $true
      $attributeCollection = New-Object `
        -Type System.Collections.ObjectModel.Collection[System.Attribute]

      # Add the attributes to the attributes collection
      $attributeCollection.Add($attributes)

      $dynParam1 = New-Object -Type `
        System.Management.Automation.RuntimeDefinedParameter($PolicyListCSV, [string],
          $attributeCollection)
  
  
      $paramDictionary = New-Object `
        -Type System.Management.Automation.RuntimeDefinedParameterDictionary
      $paramDictionary.Add($PolicyListCSV, $dynParam1)
      return $paramDictionary
    }    
}


#--------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------
Process{
    #$VerbosePreference = "Continue" 

    #
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
    Push-Location $scriptFolder
    #
    if ($PSBoundParameters.ContainsKey(("LabelListCSV"))){
        $LabelListCSV = $PSBoundParameters[$LabelListCSV]
    }
    #
    if ($PSBoundParameters.ContainsKey(("PolicyListCSV"))){
        $PolicyListCSV = $PSBoundParameters[$PolicyListCSV]
    }
    # Prepare LogFile
    Create-Log -LogFolderRoot $scriptFolder -LogFunction "Publish_Compliance_Tag" | Out-Null

    switch ($Mode) {
        "create" {
            switch ($Type) {
                "Label" {
                    # Create compliance tag
                    CreateComplianceTag -FilePath $LabelListCSV
                    # Export to result csv
                    if ($ResultCSV)
                    {
                        Create-ResultCSV -ResultFolderRoot $scriptFolder -ResultFunction "Tag_Creation" | Out-Null
                        $global:tagRetFile = $retfilePath
                        ExportCreatedComplianceTag -LabelFilePath $LabelListCSV
                    }
                }
                "LabelPolicy" {
                    # Create retention policy and publish compliance tag with the policy
                    CreateRetentionCompliancePolicy -FilePath $PolicyListCSV
                    # Export to result csv
                    if ($ResultCSV)
                    {
                        #ExportCreatedComplianceTag -LabelFilePath $LabelListCSV
                        $global:tagPubRetFile = $retfilePath
                        Create-ResultCSV -ResultFolderRoot $scriptFolder -ResultFunction "Tag_Publish" | Out-Null
                        ExportPublishedComplianceTagAndPolicy -PolicyFilePath $PolicyListCSV 
                    }
                }
            }
        }
        # TBC 
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
    


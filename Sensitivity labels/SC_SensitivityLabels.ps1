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
.PARAMETER  Mode
  Description: Determines the mode of operation of the script
  Possible values: get, set
.PARAMETER  Type
Description: Determines what aspect of Sensitivity labels to query
Possible values:  Label, LabelPolicy
.PARAMETER  MFA
Description: Switch to specifiy using MFA with modern auth setting
Possible values: -MFA
.EXAMPLE
  .\SC_SensitivityLabels.ps1 -MFA -Mode get -Type Label
  Connects to the S&C Center to query the current label settings using Modern auth
.EXAMPLE
.\SC_SensitivityLabels.ps1 -MFA -Mode get -Type LabelPolicy 
Connects to the S&C Center to query the current labelPolicy settings using Modern auth
.EXAMPLE
.\SC_SensitivityLabels.ps1 -Mode get -Type LabelPolicy
Connects to the S&C Center to query the current labelPolicy settings using Basic auth
.EXAMPLE
.\SC_SensitivityLabels.ps1 -Mode get -Type Label
Connects to the S&C Center to query the current label settings using Basic auth
.EXAMPLE
get-help .\SC_SensitivityLabels.ps1 -Detailed
Displays the help file
.INPUTS
   <none>
.OUTPUTS
   <none>
.NOTES
    Script Name     : SC_SensitivityLabels.ps1
    Requires        : Powershell Version 5.1, Windows Remote Management (WinRM) on your computer needs to allow basic authentication
    Tested          : Powershell Version 5.1
    Author          : Andrew Auret
    Email           : 
    Version         : 1.3
    Date            : 2019-11-07 (ISO 8601 standard date notation: YYYY-MM-DD)
    
    
#>

#######################################################################################################################
#--------------------------------------------------------------------------------------------------------#
[OutputType()]
[CmdletBinding(DefaultParameterSetName,PositionalBinding=$true)]
Param (
    [Parameter(Mandatory = $False, Position = 0)]
    [Switch]$MFA,

    [Parameter(Mandatory=$true, Position = 1)]
        [ValidateSet("get","set")]
        [string]$Mode
) 

DynamicParam
{
    if ($Mode.Equals("get"))
    {
      $Type = 'Type'
      $attributes = New-Object -Type `
        System.Management.Automation.ParameterAttribute
      $attributes.ParameterSetName = "__AllParameterSets"
      $attributes.Mandatory = $true
      $attributes.Position = 2
      $attributeCollection = New-Object `
        -Type System.Collections.ObjectModel.Collection[System.Attribute]

      # Add the attributes to the attributes collection
      $attributeCollection.Add($attributes)
      $attributeCollection.Add((New-Object System.Management.Automation.ValidateSetAttribute(("Label","LabelPolicy")))) 

      $dynParam1 = New-Object -Type `
        System.Management.Automation.RuntimeDefinedParameter($Type, [string],
          $attributeCollection)
  
      $paramDictionary = New-Object `
        -Type System.Management.Automation.RuntimeDefinedParameterDictionary
      $paramDictionary.Add($Type, $dynParam1)
      return $paramDictionary
    }
    if ($Mode.Equals("set"))
    {
        $Type = 'Type'
        $attributes = New-Object -Type `
            System.Management.Automation.ParameterAttribute
        $attributes.ParameterSetName = "__AllParameterSets"
        $attributes.Mandatory = $true
        $attributes.Position = 2
        $attributeCollection = New-Object `
          -Type System.Collections.ObjectModel.Collection[System.Attribute]
  
        # Add the attributes to the attributes collection
        $attributeCollection.Add($attributes)
        $attributeCollection.Add((New-Object System.Management.Automation.ValidateSetAttribute(("LabelPolicy")))) 
  
        $dynParam1 = New-Object -Type `
          System.Management.Automation.RuntimeDefinedParameter($Type, [string],
            $attributeCollection)

        $PolicyName = 'PolicyName'
        $attributes2 = New-Object -Type `
            System.Management.Automation.ParameterAttribute
        $attributes2.ParameterSetName = "__AllParameterSets"
        $attributes2.Mandatory = $true
        $attributes2.Position = 2
        $attributeCollection2 = New-Object `
          -Type System.Collections.ObjectModel.Collection[System.Attribute]
  
        # Add the attributes to the attributes collection
        $attributeCollection2.Add($attributes2)
          
        $dynParam2 = New-Object -Type `
          System.Management.Automation.RuntimeDefinedParameter($PolicyName, [string],
            $attributeCollection2)
    
        $paramDictionary = New-Object `
          -Type System.Management.Automation.RuntimeDefinedParameterDictionary
        $paramDictionary.Add($Type, $dynParam1)
        $paramDictionary.Add($PolicyName, $dynParam2)
        return $paramDictionary
    }
  
}

#--------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------
Process{
    #$VerbosePreference = "Continue" 
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
        Write-Verbose "Gathering Credentials for non-MFA sign on"
        $Credential = Get-Credential -Message "Please enter your Office 365 credentials"
    }
    #
    If ((Get-PSSession).ConfigurationName -eq "Microsoft.Exchange"){
        If ((Get-PSSession).State -eq "Broken"){
            Get-PSSession | Where-Object{$_.ConfigurationName -eq "Microsoft.Exchange"}  | Remove-PSSession
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
    if ($PSBoundParameters.ContainsKey(("PolicyName"))){
        $LabelPolicyName = $PSBoundParameters[$PolicyName]
    }
    #
    switch ($Mode) {
        "get" {
            switch ($PSBoundParameters.$Type) {
                "Label" {
                    Write-Verbose  "Getting the content of the current Sensitivity Labels"
                    $labels = Get-Label
                    foreach($label in $labels){
                        try {
                            $labelpolicyRule = Get-Labelpolicyrule | Where-Object {$_.LabelName -eq $label.Name} -ErrorAction SilentlyContinue
                        }
                        catch {
                            Write-Warning "The cmdlet 'Get-Labelpolicyrule' is not available for this tenant - Encryption and Content Marking will not be displayed correctly"
                        }

                        Write-Output "Name           : $($label.Name)"
                        Write-Output "Created by     : $($label.CreatedBy)"
                        Write-Output "Last modified  : $($label.LastModifiedBy)"
                        Write-Output "Display Name   : $($label.DisplayName)"
                        Write-Output "Tooltip        : $($label.Tooltip)"
                        Write-Output "Description    : $($label.Comment)"
                        Write-Output "ImmutableId    : $($label.ImmutableId)"
                        #
                        $labelEncryptAction = ($labelpolicyRule | Where-Object {$_.LabelActionName -eq 'encrypt'}).LabelActionName
                        If ($labelEncryptAction -eq "encrypt"){
                            Write-Output "Encryption     : On"
                        }
                        else {
                            Write-Output "Encryption     : Off"
                        }
                        #
                        $ContentMarking = ($labelpolicyRule | Where-Object {$_.LabelActionName -like 'applycontentmarking*'}).LabelActionName
                        If (!($null -eq $ContentMarking)){
                            Write-Output "Content marking: $($ContentMarking)"
                        }
                        else {
                            Write-Output "Content marking: Not set"
                        }
                        #
                        $Settings = $label | Select-Object -ExpandProperty Settings
                        foreach($list in $Settings.GetEnumerator()){
                            if ($list.Contains("color")){
                                $list = $list -replace "\W",''
                                $color = $list -replace "color","#"
                                Write-Output "Label colour   : $($color)"
                            }
                        }
                        Write-Output ""
                    }

                  }
                "LabelPolicy" {
                    Write-Output "Getting the content of the current Sensitivity Label policies"
                    $labelpolicies = Get-LabelPolicy
                    foreach($labelpolicy in $labelpolicies){
                        Write-Output "Name           : $($labelpolicy.Name)"
                        Write-Output "Labels         : $($labelpolicy.Labels)"
                        Write-Output "-- Settings --"
                        $Settings = $labelpolicies | Select-Object -ExpandProperty Settings
                        foreach($list in $Settings){
                            #Users must provide justification to remove a label or lower classification label
                            if ($list.Contains("requiredowngradejustification")){
                                $list = $list -replace "\W",''
                                $rdgj = $list -replace "requiredowngradejustification",""
                                Write-Output "Require a justification for changing a label              : $($rdgj)"
                            }
                            #Label is mandatory
                            if ($list.Contains("mandatory")){
                                $list = $list -replace "\W",''
                                $mandatory = $list -replace "mandatory",""
                                Write-Output "Require users to apply a label to their email or documents: $($mandatory)"
                            }
                            #For email messages with attachments, apply a label that matches the highest classification of those attachments
                            if ($list.Contains("attachmentaction")){
                                $list = $list -replace "\W",''
                                $attachmentaction = $list -replace "attachmentaction",""
                                Write-Output "For email messages with attachments, apply a label that matches the highest classification of those attachments: $($attachmentaction)"
                            }
                            #hidebarbydefault
                            if ($list.Contains("hidebarbydefault")){
                                $list = $list -replace "\W",''
                                $hidebarbydefault = $list -replace "hidebarbydefault",""
                                Write-Output "HideBarByDefault                                          : $($hidebarbydefault)"
                            }
                            #outlookjustifyuntrustedcollaborationlabel
                            if ($list.Contains("outlookjustifyuntrustedcollaborationlabel")){
                                #$list = $list -replace "\D",''
                                $list = $list -replace "outlookjustifyuntrustedcollaborationlabel,",""
                                Write-Output "outlookjustifyuntrustedcollaborationlabel                 : $($list)"
                            }
                        }
                        Write-Output ""
                    }
                }
            }
        }
        "set" {
            switch ($PSBoundParameters.$Type) {
                "Label" {
                    # TBC
                }
                "LabelPolicy" {
                    Write-Host "Getting the content of the current Sensitivity Label policies" -ForegroundColor Blue
                    try {
                        $labelpolicy = Get-LabelPolicy -Identity $LabelPolicyName -ErrorAction Stop
                    }
                    catch {
                        Write-Warning "The LabelPolicy:[$($LabelPolicyName)] does not exist"
                        break
                    }
                    #
                    Initialize-Modules("MSOnline")
                    #Check if we are authenticated already before prompting
                    If (!(MSOLConnected)){
                        Write-Host "Prompting for Authentication since we still need seperate sessions for Exchange Online Remote PowerShell Module and Azure Active Directory PowerShell for Graph module" -ForegroundColor Green
                        Connect-MsolService
                    }
                    #
                    Write-Host "Checking the Groups that have been assigned to the Policy to ensure that they are Office 365 or MailEnabledSecurity groups" -ForegroundColor Blue
                    foreach($Group in $labelpolicy.ModernGroupLocation.Name){
                        $MSOLGroup = Get-MsolGroup -SearchString $Group | Where-Object{$_.GroupType -eq "DistributionList"}
                        If ($MSOLGroup){
                           Write-Warning "Group: $($Group) is a $($MSOLGroup.GroupType) and could cause replication issues"
                        }
                    }
                    #
                    Write-Host "Getting LabelPolicy - 'Advanced settings' - Before Change" -ForegroundColor Blue
                    Get-LabelPolicy -Identity $LabelPolicyName | Select-Object Settings -ExpandProperty Settings
                    #
                    Write-Output ""
                    $CSV = Import-CSVtoHash .\AdvancedSettings.csv
                    If ($CSV.contains("HideBarByDefault")){
                        Write-Host "Setting LabelPolicy - 'HideBarByDefault'" -ForegroundColor Green
                        Set-LabelPolicy -Identity $LabelPolicyName -AdvancedSettings @{HideBarByDefault="$(($CSV.HideBarByDefault).toLower())"}
                    }
                    #
                    If ($CSV.contains("AttachmentAction")){
                        Write-Host "Setting LabelPolicy - 'AttachmentAction'" -ForegroundColor Green
                        Set-LabelPolicy -Identity $LabelPolicyName -AdvancedSettings @{AttachmentAction="$($CSV.AttachmentAction)"}
                    }
                    #
                    If ($CSV.contains("OutlookJustifyUntrustedCollaborationLabel")){
                        Write-Host "Setting LabelPolicy - 'OutlookJustifyUntrustedCollaborationLabel'" -ForegroundColor Green
                        Set-LabelPolicy -Identity $LabelPolicyName -AdvancedSettings @{OutlookJustifyUntrustedCollaborationLabel="$($CSV.OutlookJustifyUntrustedCollaborationLabel)"}
                    }
                    #
                    If ($CSV.contains("OutlookJustifyTrustedDomains")){
                        Write-Host  "Setting LabelPolicy - 'OutlookJustifyTrustedDomains'" -ForegroundColor Green
                        Set-LabelPolicy -Identity $LabelPolicyName -AdvancedSettings @{OutlookJustifyTrustedDomains="$($CSV.OutlookJustifyTrustedDomains)"}
                    }
                    #
                    Write-Output ""
                    Write-Host  "Getting LabelPolicy - 'Advanced settings' - After Change" -ForegroundColor Blue
                    Get-LabelPolicy -Identity $LabelPolicyName | Select-Object Settings -ExpandProperty Settings
                    #
                }
            }
        }
    }
}
#
    


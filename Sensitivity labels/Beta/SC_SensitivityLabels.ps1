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
  Possible values: get
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
  .\SC_SensitivityLabels.ps1 -Mode get -Type Label -MFA
  Connects to the S&C Center to query the current label settings using Modern auth
.EXAMPLE
.\SC_SensitivityLabels.ps1 -Mode get -Type LabelPolicy -MFA
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
    Version         : 1.1
    Date            : 2019-11-07 (ISO 8601 standard date notation: YYYY-MM-DD)
    
    
#>

#######################################################################################################################
#--------------------------------------------------------------------------------------------------------#
[OutputType()]
[CmdletBinding(DefaultParameterSetName)]
Param (
    [Parameter(Mandatory=$true, Position = 0)]
    [ValidateSet("get")]
    [string]$Mode,

    [Parameter(Mandatory=$true, Position = 1)]
    [ValidateSet("Label","LabelPolicy")]
    [string]$Type
) 
#--------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------
Process{
    #$VerbosePreference = "Continue" 
    $scriptPath = $myInvocation.MyCommand.Path
    $scriptFolder = Split-Path $scriptPath
    #
    
    Write-Verbose "Checking for Graph module..."

    $GraphModule = Get-Module -Name "Graph"

    if ($null -eq $IntuneModule) {
        Write-Verbose ""
        Write-Verbose "Graph Powershell module not installed..." 
        Write-Verbose "Installing Graph module" 
        Write-Verbose ""
        Import-Module $scriptFolder\Graph.psm1 -Force -Verbose
    }
    Else{
        Write-Verbose "Graph Powershell module is installed"
    }
    #
    # If the module count is greater than 1 then find the latest version

    if($GraphModule.count -gt 1){

    $Latest_Version = ($GraphModule | Select-Object version | Sort-Object)[-1]

    $GraphModule = $GraphModule | Where-Object { $_.version -eq $Latest_Version.version}

        # Checking if there are multiple versions of the same module found

        if($GraphModule.count -gt 1){
            $GraphModule = $GraphModule | Select-Object -Unique
        }
    }
    break
    #
    switch ($Mode) {
        "get" {
            switch ($Type) {
                "Label" {
                    Write-Verbose  "Getting the content of the current Sensitivity Labels"
                    $labels = Get-Label
                    foreach($label in $labels){
                        $labelpolicyRule = Get-Labelpolicyrule | Where-Object {$_.LabelName -eq $label.Name} -ErrorAction SilentlyContinue
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
    


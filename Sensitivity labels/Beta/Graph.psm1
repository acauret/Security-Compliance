<#
.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.
.VERSION
    1.0             : Initial version
#>
#region Authentication
####################################################
function Get-AuthToken {

<#
.SYNOPSIS
This function is used to authenticate with the Graph API REST interface
.DESCRIPTION
The function authenticate with the Graph API Interface with the tenant name
.EXAMPLE
Get-AuthToken
Authenticates you with the Graph API interface
.NOTES
NAME: Get-AuthToken
#>

[cmdletbinding()]

param
(
    [Parameter(Mandatory=$true)]
    [string]$User
)

$userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User

$tenant = $userUpn.Host

Write-verbose "Checking for AzureAD module..."

    $AadModule = Get-Module -Name "AzureAD" -ListAvailable


    if ($null -eq $AadModule) {
        Write-Output ""
        Write-OutPut "AzureAD Powershell module not installed..." 
        Write-OutPut "Install by running 'Install-Module AzureAD' from an elevated PowerShell prompt" 
        Write-OutPut "Script can't continue..."
        Write-OutPut ""
        break
    }

# Getting path to ActiveDirectory Assemblies
# If the module count is greater than 1 find the latest version

    if($AadModule.count -gt 1){

        $Latest_Version = ($AadModule | Select-Object version | Sort-Object)[-1]

        $aadModule = $AadModule | Where-Object { $_.version -eq $Latest_Version.version }

            # Checking if there are multiple versions of the same module found

            if($AadModule.count -gt 1){

            $aadModule = $AadModule | Select-Object -Unique

            }

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    }

    else {

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    }

[System.Reflection.Assembly]::LoadFrom($adal) | Out-Null

[System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null

$clientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"
#
$redirectUri = "urn:ietf:wg:oauth:2.0:oob"
#
$resourceAppIdURI = "https://graph.microsoft.com"
#
$authority = "https://login.windows.net/$Tenant"
#
    try {

    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority

    # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
    # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession

    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"

    $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")

    $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$clientId,$redirectUri,$platformParameters,$userId).Result

        # If the accesstoken is valid then create the authentication header

        if($authResult.AccessToken){

        # Creating header for Authorization token

        $authHeader = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer " + $authResult.AccessToken
            'ExpiresOn'=$authResult.ExpiresOn
            }

        return $authHeader

        }

        else {
            Write-OutPut "Authorization Access Token is null, please re-run authentication..."
            Write-Output ""
            break
        }

    }

    catch {
        Write-Output $_.Exception.Message
        Write-Output $_.Exception.ItemName
        Write-Output ""
        break
    }

}
##################################################################################
Function CheckAuthorisation(){

<#
.SYNOPSIS
This function is used to check authorisation token
.DESCRIPTION
.EXAMPLE
.NOTES
#>

    try {
            # Checking if authToken exists before running authentication
            if($Global:authToken){
              # Setting DateTime to Universal time to work in all timezones
              $DateTime = (Get-Date).ToUniversalTime()
              # If the authToken exists checking when it expires
              $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes
              if($TokenExpires -le 0){
                Write-Output "Azure Authentication Token expired $($TokenExpires) minutes ago" 
                Write-Output ""
                # Defining User Principal Name if not present
                if($null -eq $User -or $User -eq ""){
                    $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
                    Write-Output ""
                }
                $Global:authToken = Get-AuthToken -User $User
              }
            }
            # Authentication doesn't exist, calling Get-AuthToken function
            else {
                write-Output "Azure Authentication Token does not exist." 
                write-Output ""
                if($Null -eq $User -or $User -eq ""){
                    $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
                    write-Output ""
                }
                # Getting the authorization token
                $Global:authToken = Get-AuthToken -User $User
            }
        }
    catch {
        $_.Exception
        break
    }
}
##################################################################################
#endregion
function Trace-Error{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Message,

        [switch]
        $NonTerminating
    )

    $Message += $(Get-PSCallStack | Select-Object -skip 1 | Out-String)

    if ($NonTerminating) {
        Write-Error $Message
    } 
    else {
        throw $Message
    }
}
##################################################################################
Function Trace-Execution {
    [CmdletBinding()]
    param (
        [string]
        $Message
    )

    Write-Verbose $Message -Verbose
}
##################################################################################
Function Submit-JSON {

[CmdletBinding()]

    param(
        [Parameter(mandatory=$True)]
        [String] $path,

        [Parameter(mandatory=$false)]
        [String] $JSON,

        [Parameter(Mandatory=$False)]
        [string]$graphApiVersion = "v1.0",
        [Switch] $WaitForUpdate,
        $Silent = $True,

        [Parameter(mandatory=$True)]
        [ValidateSet("Get","Patch","Post")]
        [String] $method      

    )


    $uriRoot = "https://graph.microsoft.com"

    $uri = $uriRoot+"/"+$graphApiVersion+"/"+$($path)
    
    if (!$Silent) {
        Trace-Execution "Submit-JSON $method [$path]"
    }

    try {
        $NotFinished = $true
        do {
            switch ($method){
              "Get"{
                $result = Invoke-RestMethod -Uri $uri -Headers $authToken -Method $method 
              }
              "Post"{
                $result = Invoke-RestMethod -Uri $uri -Headers $authToken -Method $method -Body $JSON -ContentType "application/json"
              }
              "Patch"{
                $result = Invoke-RestMethod -Uri $uri -Headers $authToken -Method $method -Body $JSON -ContentType "application/json"
              }
            }          
            if($null -eq $result) {
                return $null    
            }
            
            #$toplevel = convertfrom-json $result.Content # Remove invoke-webrequest and use Invoke-RestMethod instead
            $toplevel = $result
            if ($null -eq $toplevel.value)
            {
                $obj = $toplevel
            } 
            else 
            {
                $obj = $toplevel.value
            }

            if ($WaitForUpdate.IsPresent) {
                if ($obj.properties.provisioningState -eq "Updating")
                {
                    Trace-Execution "Submit-JSON: the object's provisioningState is Updating. Wait 1 second and check again."
                    Start-Sleep 1 #then retry
                }
                else
                {
                    $NotFinished = $false
                }
            }
            else
            {
                $notFinished = $false
            }
      } while ($NotFinished)

      if ($obj.properties.provisioningState -eq "Failed") {
         Trace-Error ("Provisioning failed: {0}`nReturned Object: {1}`nObject properties: {2}" -f @($uri, $obj, $obj.properties))
      }
      return $obj
    }
    catch
    {
        Trace-Execution "GET Exception: $_"
        Trace-Execution "GET Exception: $($_.Exception.Response)"
        Trace-Execution "GET Exception: $($_.Exception.Response.GetResponseStream())"
        Trace-Execution  "---------"
        Trace-Execution  "URI: $uri"
        Trace-Execution  "Method: $method"
        Trace-Execution  "JSON: $JSON"
        break
        return $null
    }
}

##################################################################################
function Test-Json {
    [CmdletBinding()]
    param (
      [Parameter(Mandatory=$true)]
      [AllowNull()]
      [AllowEmptyString()]
      [AllowEmptyCollection()]
      [string]$JSON
    )
  
    try {
        
      if($JSON -eq "" -or  $null -eq $JSON){
          write-error -Message "Test-Json: No JSON specified, please specify valid JSON ..."
      }
      else {
          ConvertFrom-Json $JSON -ErrorAction Stop
          $Result = $True
      }
    }
    catch {
      Trace-Execution "GET Exception: $($_.Exception)"
      $result = $false
    }
    return $result
  }

##################################################################################
Function Get-sensitivityLabels(){

    <#
    .SYNOPSIS
    This function is used to get simple sensitivityLabels information from the Graph API REST interface
    .DESCRIPTION
    The function connects to the Graph API Interface and gets any information sensitivityLabels
    .EXAMPLE
    Get-IntuneManagedAppPolicy
    Returns any sensitivityLabels
    .NOTES
    NAME: Get-sensitivityLabels
    #>
    
    [CmdletBinding()]
    
    param
    (
        [Parameter(Mandatory=$False)]
        [string]$graphApiVersion = "v1.0"

    )
        #
        CheckAuthorisation
        #
        $Resource = "me/informationProtection/policy/labels"


        Submit-JSON -method Get -path $Resource -graphApiVersion $graphApiVersion
    }
##################################################################################
Function Get-sensitivityLabels_adv(){

    <#
    .SYNOPSIS
    This function is used to get advanced sensitivityLabels information from the Graph API REST interface
    .DESCRIPTION
    The function connects to the Graph API Interface and gets any information sensitivityLabels
    .EXAMPLE
    Get-IntuneManagedAppPolicy
    Returns any sensitivityLabels
    .NOTES
    NAME: Get-sensitivityLabels_adv
    #>
    
    [CmdletBinding()]
    
    param
    (
        [Parameter(Mandatory=$False)]
        [string]$graphApiVersion = "v1.0"

    )
        #
        CheckAuthorisation
        #
        $Resource = "me/informationProtection/sensitivityLabels"


        Submit-JSON -method Get -path $Resource -graphApiVersion $graphApiVersion
    }
##################################################################################
Export-ModuleMember -Function Test-Json,`
                              Get-sensitivityLabels, `
                              Get-sensitivityLabels_adv

                              
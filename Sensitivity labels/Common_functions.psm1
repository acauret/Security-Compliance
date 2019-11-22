#region Functions

function Initialize-Modules() 
{
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
            Import-Module (Get-ChildItem -Path $scriptFolder\Exchange_Online_Module\ -Filter '*ExoPowershellModule.dll' -Recurse | ForEach-Object{(Get-ChildItem -Path $_.Directory -Filter CreateExoPSSession.ps1)} | Sort-Object LastWriteTime | Select-Object -Last 1).FullName
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


function MSOLConnected {
    Get-MsolGroup -ErrorAction SilentlyContinue
    $result = $?
    return $result
}

Function Import-CSVtoHash {
  
    [cmdletbinding()]
    
    Param (
    [Parameter(Position=0,Mandatory=$True,HelpMessage="Enter a filename and path for the CSV file")]
    [ValidateNotNullorEmpty()]
    [ValidateScript({Test-Path -Path $_})]
    [string]$Path
    )
    
    Write-Verbose "Importing data from $Path"
    
    Import-Csv -Path $path | ForEach-Object -begin {
         #define an empty hash table
         $hash=@{}
        } -process {
           <#
           if there is a type column, then add the entry as that type
           otherwise we'll treat it as a string
           #>
           if ($_.Type) {
             
             $type=[type]"$($_.type)"
           }
           else {
             $type=[type]"string"
           }
           Write-Verbose "Adding $($_.key)"
           Write-Verbose "Setting type to $type"
           
           $hash.Add($_.Key,($($_.Value) -as $type))
    
        } -end {
          #write hash to the pipeline
          Write-Output $hash
        }
    
    write-verbose "Import complete"
    
} #end function
#endregion
Export-ModuleMember -Function Initialize-Modules, `
                              MSOLConnected,`
                              Import-CSVtoHash

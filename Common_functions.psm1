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

#endregion
Export-ModuleMember -Function Initialize-Modules

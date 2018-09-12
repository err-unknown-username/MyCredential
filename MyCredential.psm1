﻿ 

function PromptForConfirmation 
{
    param(
        [OutputType([boolean])]
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Title,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
    )
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Description."
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Description."
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
 
    $result = $host.ui.PromptForChoice($title, $message, $options, 1)

    $result -eq 0 
}

function Add-MyCredential
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Username,
        [Parameter(Mandatory=$false)][string]$PasswordsFolder = "C:\Users\DavidWhitehead\OneDrive\OneDrive - RecordPoint Software\Documents\WindowsPowerShell\Passwords"
    )

    if( -not (Test-Path $PasswordsFolder) )
    {
        if( PromptForConfirmation -Title "Create Folder" -Message "Passwords folder not found.`r`nDid you intent to create a new password folder?`r`n`r`n$($PasswordsFolder)" ) 
        {
            $void = New-Item -Path $PasswordsFolder -ItemType Directory
        }
    }

    $passwordFile = Join-Path $PasswordsFolder $( $username.Replace('\','##') + ".pwd" )
    $cred = Get-Credential -Message "Enter Password" -UserName $Username
    $cred.Password | ConvertFrom-SecureString | Set-Content $passwordFile
}
Export-ModuleMember -Function Add-MyCredential


function Get-MyCredential
{
    [CmdletBinding()]
    Param (
        [OutputType([System.Management.Automation.PSCredential])]
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$Username,
        [Parameter(Mandatory=$false)][string]$PasswordsFolder =  (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'MyCredential')
    )

    if( ($Username -eq $null) -or ($Username -eq [string]::Empty) )
    {
        Write-Verbose 'Username parameter missing, select from registered usernames:' -Verbose
        $options = @{}
        $i=0;  dir $PasswordsFolder | %{ $options[$i]=($_.Name).Replace('##','\').Replace(".pwd",""); $i++} 
        $options.GetEnumerator() | Sort-Object -Property Value | %{ 
            Write-Verbose "`t$($_.Key)`t:`t$($_.Value)" -Verbose 
        }
        Write-Verbose 'Select one: ' -Verbose
        $option = Read-Host

        if( $options.ContainsKey([int]$option) )
        {
            $Username = $options[[int]$option]
        }
        else
        {
            Write-Warning "Invalid Option $option"
            return
        }
    }

    $passwordFile = Join-Path $PasswordsFolder $( $Username.Replace('\','##') + ".pwd" )
    if( -not (Test-Path $passwordFile) )
    {
        throw "Password file not found $passwordFile"
    }


    $password     = Get-Content $passwordFile | ConvertTo-SecureString
    $cred         = New-Object System.Management.Automation.PSCredential $Username, $password

    $cred
}
Export-ModuleMember -Function Get-MyCredential


function Get-MySPOnlineCredential
{
    [CmdletBinding()]
    Param (
        [OutputType([System.Management.Automation.PSCredential])]
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$Username,
        [Parameter(Mandatory=$false)][string]$PasswordsFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'MyCredential')
    )

    if( (Get-Module Microsoft.SharePoint.Client) -eq $null )
    {
        Write-Verbose 'Importing SharePoint Client Module' -Verbose
        Import-Module 'C:\nuget\Microsoft.SharePointOnline.CSOM.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll'
    }

    if( ($Username -eq $null) -or ($Username -eq [string]::Empty) )
    {
        Write-Verbose 'Username parameter missing, select from registered usernames:' -Verbose
        $options = @{}
        $i=0;  dir $PasswordsFolder | %{ $options[$i]=($_.Name).Replace(".pwd",""); $i++} 
        $options.GetEnumerator() | Sort-Object -Property Value | %{ 
            Write-Verbose "`t$($_.Key)`t:`t$($_.Value)" -Verbose 
        }
        Write-Verbose 'Select one: ' -Verbose
        $option = Read-Host

        if( $options.ContainsKey([int]$option) )
        {
            $Username = $options[[int]$option]
        }
        else
        {
            Write-Warning "Invalid Option $option"
            return
        }
    }

    $passwordFile = Join-Path $PasswordsFolder $( $Username.Replace('\','##') + ".pwd" )
    if( -not (Test-Path $passwordFile) )
    {
        throw "Password file not found $passwordFile"
    }

    $password     = Get-Content $passwordFile | ConvertTo-SecureString
    $cred         = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials $Username, $password

    $cred
}
Export-ModuleMember -Function Get-MySPOnlineCredential

# This script adds items to the list picking up the data from the json file.

if (-not (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline)) 
{
    Install-Module SharePointPnPPowerShellOnline
}

#Import-Module SharePointPnPPowerShellOnline

# Gets or Sets the JSON data.
$jsonconfig = $null

# Gets or Sets the tenant admin credentials.
$credentials = $null

# Windows Credentials Manager credential label.
$winCredentialsManagerLabel = "ShareOnline"

# Imports the json configuration template.
try 
{ 
    $jsonconfig = Get-Content .\PollData.json|ConvertFrom-Json
} 
catch 
{
    # Prompts for json configuration template path.
    $jsonconfigPath = Read-Host -Prompt "Please enter the json configuration template full path"
    $jsonconfig = Get-Content -Path $jsonconfigPath|ConvertFrom-Json
}

# Gets stored credentials from the Windows Credential Manager or show prompt.
# How to use windows credential manager:
# https://github.com/SharePoint/PnP-PowerShell/wiki/How-to-use-the-Windows-Credential-Manager-to-ease-authentication-with-PnP-PowerShell
if((Get-PnPStoredCredential -Name $winCredentialsManagerLabel) -ne $null)
{
    $credentials = $winCredentialsManagerLabel
}
else
{
    # Prompts for credentials, if not found in the Windows Credential Manager.
    $email = Read-Host -Prompt "Please enter tenant admin email"
    $pass = Read-host -AsSecureString "Please enter tenant admin password"
    $credentials = New-Object –TypeName "System.Management.Automation.PSCredential" –ArgumentList $email, $pass
}

if($credentials -eq $null -or $jsonconfig -eq $null) 
{
    Write-Host "Error: Not enough details." -ForegroundColor DarkRed
    exit 1
}


$listItems=$jsonconfig.PollQuestions
Connect-PnPOnline $jsonconfig.TenantUrl -Credentials $credentials
foreach($item in $listItems){
    $options=$null
    $item.Options|%{if($options -ne $null){$options=$options+"`n"+$_}else{$options=$options+$_}}
    $valuestr=@{"Title"=$item.Question;"Question"=$item.Question;"Options"=$options;"Published_x0020_Date"=$item.PublishedDate;"Expiry_x0020_Date"=$item.ExpiryDate}
    Add-PnPListItem -List $jsonconfig.ListName -Values $valuestr
}

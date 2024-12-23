# Export DL Groups - Script by Khalil
# Exporting a list of Members within Nested DL's

function WriteR($Message) { Write-Host -ForegroundColor Red $Message }
function WriteY($Message) { Write-Host -ForegroundColor Yellow $Message }

# Module Check
try {
    $Modules = ("Microsoft.Graph.Groups", "Microsoft.Graph.Authentication")
    foreach ($Module in $Modules) {
        $Command = Get-Module -ListAvailable -Name $Module
        if ($Command.Name -eq $Module) {
            WriteY "$Module Module already installed"
        } else {
            WriteY "Installing $Module Module"
            Install-Module -Name $Module -Force -ErrorAction Stop | Out-Null
        }
    }
} catch {
    WriteR "Couldn't install Powershell Modules"
    exit
}

# Connecting to Microsoft Graph
try {
    Connect-MgGraph -NoWelcome | Out-Null
} catch {
    WriteR "Couldn't connect to Microsoft Graph"
    exit
}

# Variables are outside the function because it'll use this function multiple times
$AllMembers = @()
$DLGroups = @()
function CheckMembers($list) {
    foreach ($object in $list) {
        $InfoObject = $object | Select-Object -ExpandProperty AdditionalProperties
        
        if ($InfoObject."@odata.type" -eq "#microsoft.graph.group") {
            $nestedGroup = Get-MgGroupMember -GroupId $object.Id
            $DLGroups += $InfoObject.displayName
            CheckMembers($nestedGroup)
        } elseif ($InfoObject."@odata.type" -eq "#microsoft.graph.user") {
            $AllMembers += $InfoObject.mail
        }
        $InfoObject = $null
        $nestedGroup = $null
    }
}

$Group = Read-Host "What is the name of the Distribution List?"

# Getting the DL Members by the Display Name
try {
    $MainGroup = Get-MgGroupMember -Filter "displayName eq '$Group'"
} catch {
    WriteR "Couldn't get the Group Members with the DL Name '$Group'"
    exit
}

CheckMembers($MainGroup)

Export-Csv -InputObject $AllMembers -Path "C:\Users\$([Environment]::Username)\members.csv"
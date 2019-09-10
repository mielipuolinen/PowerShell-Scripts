#Requires -Version 5.0

<#
.SYNOPSIS
Export company's endpoints in excel format from BitDefender GravityZone

.DESCRIPTION
Export customer company's endpoint list in excel format from BitDefender GravityZone by using JSON-RPC 2.0 API.

.PARAMETER CompanyName
Search string for company. The script expects only one result to be found with this parameter.

.PARAMETER ExportFormat
CSV or Excel.

.PARAMETER ExportFilename
A prefix for export file.
"${ExportFilename}_YYMMDD-HHMM.ext"

.EXAMPLE
PS> BitDefenderGravityZoneAPI_CompanyEndpoints.ps1 -CompanyName "X" -ExportFormat "Excel" -ExportFilename "CompanyEndpoints"
Customer XYZ found, exported endpoints to CompanyEndpoints__X_YYMMDD-HHMM.xlsx

.EXAMPLE
PS> BitDefenderGravityZoneAPI_CompanyEndpoints.ps1 -CompanyName "X" -ExportFormat "CSV" -ExportFilename "CompanyEndpoints"
Customer XYZ found, exported endpoints to CompanyEndpoints_X_YYMMDD-HHMM.csv

.INPUTS
Pipelining not yet available.

.OUTPUTS

.LINK
https://download.bitdefender.com/business/API/Bitdefender_GravityZone_Cloud_APIGuide_forPartners_enUS.pdf

.LINK
https://github.com/dfinke/ImportExcel

.LINK
http://www.jsonrpc.org/specification

.NOTES
NOTE: This script is still in development. This needs some refactoring but it works.

PowerShell module ImportExcel required, install it by using command:
Install-Module ImportExcel

This script expects a configuration file which has $API variable with base64 encoded API key. The file needs to be located in same path with this script.
"${PSScriptRoot}\BitDefenderGravityZoneAPI.config.ps1":
$API = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"

API access required to Companies API and Network API. This access is declared during API key generation.

Author: Niko Mielikäinen
Git: https://github.com/mielipuolinen/PowerShell-Scripts
#>

[CmdletBinding()]
Param(
    [Parameter(ValueFromPipeline=$true)]
    [String]$CompanyName = "",
    [String]$ExportFormat = "CSV",
    [String]$ExportFilename = "CompanyEndpoints_${CompanyName}"
)

#Set-StrictMode -Version Latest

#HTTP POST
#Content-Type application/json
#HTTP Statuses; 200 Successful (incl. wrong params), 401 Unauthorized (API Key?), 403 Forbiddden (API not enabled?), 405 Method Not Allowed (Not HTTP POST?), 429 Too Many Requests (>10req/s)
#JSON-RPC Erros: -32700 Parse Error, -32600 Invalid Request, -32601 Method not found, -32602 Invalid params, -32000 Server error

try{
    . "${PSScriptRoot}\BitDefenderGravityZoneAPI.config.ps1"
}catch{
    Write-Host "ERROR - Couldn't load ${PSScriptRoot}\BitDefenderGravityZoneAPI.config.ps1"
    Exit
}

if(!(Get-Variable API -ErrorAction Ignore) -or $API -eq "" -or $API.length -lt 40 ){
    Write-Host "ERROR - Couldn't load API key."
    Exit
}

$BaseAPI = "https://cloudgz.gravityzone.bitdefender.com/api/v1.0/jsonrpc/"
$CompaniesAPI = "${BaseAPI}companies"
#$LicensingAPI = "${BaseAPI}licensing"
#$AccountsAPI = "${BaseAPI}accounts"
$NetworkAPI = "${BaseAPI}network"
#$PackagesAPI = "${BaseAPI}packages"
#$PoliciesAPI = "${BaseAPI}policies"
#$IntegrationsAPI = "${BaseAPI}integrations"
#$ReportsAPI = "${BaseAPI}reports"
#$PushAPI = "${BaseAPI}push"
$headers = @{"Authorization" = "Basic ${API}"}

$exportCSV = "${PSScriptRoot}\${ExportFilename}_$(Get-Date -UFormat "%y%m%d%-%H%M").csv"
$exportXLSX = "${PSScriptRoot}\${ExportFilename}_$(Get-Date -UFormat "%y%m%d%-%H%M").xlsx"

#########################################################################################################
# FIND COMPANY BY NAME AND GET parentID
#########################################################################################################


$JSON = @'
{
    "params": {
        "nameFilter": ""
    },
    "method": "findCompaniesByName",
    "id": "1",
    "jsonrpc": "2.0"
}
'@ | ForEach-Object {$_.replace("`"nameFilter`": `"`"","`"nameFilter`": `"*${CompanyName}`"")} 

$response = Invoke-RestMethod -Uri $CompaniesAPI -Method POST -ContentType "application/json" -Headers $headers -Body $JSON
Start-Sleep 0.2

if($response.result.Count -eq 1){
    Write-Verbose "OK - Found $($response.result[0].name) : $($response.result[0].id)"
    $parentId = ([string]($response.result[0].id)).Trim()
    $companyName = $response.result[0].name
}else{
    Write-Host "ERROR - Found $($response.result.Count) results, exiting"
    $response.result | ForEach-Object {Write-Verbose "$_"}
    exit
}


#########################################################################################################
# GET PAGE COUNT
#########################################################################################################


$JSON = @'
{
    "params": {
        "parentId": "",
        "page": 1,
        "perPage": 30,
        "filters": {
            "depth": {
                "allItemsRecursively": true
            }
        }
    },
    "method": "getEndpointsList",
    "id": "1",
    "jsonrpc": "2.0"
}
'@ | ForEach-Object {$_.replace("`"parentId`": `"`"","`"parentId`": `""+$parentId+"`"")}

$response = Invoke-RestMethod -Uri $NetworkAPI -Method POST -ContentType "application/json" -Headers $headers -Body $JSON
Start-Sleep 0.2

if($response.result){
    Write-Verbose "OK - $($response.result.pagesCount) pages"
    $pageCount = ([int]([string]($response.result.pagesCount)).Trim())
}else{
    Write-Host "ERROR - Failed results, exiting"
    exit
}


#########################################################################################################
# QUERY CLIENTS
#########################################################################################################

if($clients){ Remove-Variable clients }

for($i=1; $i -le $pageCount; $i++){

$JSON = @'
{
"params": {
    "parentId": "",
    "page": "",
    "perPage": 30,
    "filters": {
        "depth": {
            "allItemsRecursively": true
        }
    }
},
"method": "getEndpointsList",
"id": "1",
"jsonrpc": "2.0"
}
'@ | ForEach-Object {$_.replace("`"parentId`": `"`"","`"parentId`": `""+$parentId+"`"")} | % {$_.replace("`"page`": `"`"","`"page`": "+$i)}

    $response = Invoke-RestMethod -Uri $NetworkAPI -Method POST -ContentType "application/json" -Headers $headers -Body $JSON
    Start-Sleep 0.2

    if($response.result){
        Write-Verbose "OK - Querying page n:o $($i), $($response.result.items.Count) clients"
        $clients += $response.result.items
    }else{
        Write-Host "Error - Failed to query a page n:o $($i), exiting"
        exit
    }
}

#########################################################################################################
# DUPLICATE CHECK
#########################################################################################################

$duplicates = @()

ForEach ($client in $clients){
    $i = 0;
    ForEach ($client2 in $clients){
        if ($client.name -eq $client2.name){ $i++; }
    }

    if (($i -gt 1) -and (!$duplicates.Contains($client.name))){
        #Write-Host "WARNING - Duplicate: $($client.name)"
        $duplicates += $client.name
    }
}

Write-Host "WARNING - Found duplicate endpoint(s): ${duplicates}"

#########################################################################################################
# EXPORT
#########################################################################################################

$clients | Select-Object fqdn,isManaged,managedWithBest,operatingSystemVersion,ip | Export-Csv -Path $exportCSV -NoTypeInformation
Write-Verbose "OK - Exported CSV to $($exportCSV)"
Write-Verbose "DONE - $($clients.Count) clients for $($companyName)"
$endpointCount = $clients.Count

if($ExportFormat -eq "Excel"){
    $clients | Select-Object @{Label="Hostname";Expression={($_.fqdn)}},@{Label="Managed";Expression={($_.isManaged)}},@{Label="Managed with BEST";Expression={($_.managedWithBest)}},@{Label="OS";Expression={($_.operatingSystemVersion)}},@{Label="IP";Expression={($_.ip)}} | Export-Excel -Path $exportXLSX -Title "BDGZ $($companyName) Endpoints" -WorkSheetname "BDGZ Endpoints" -AutoSize -TableName "Endpoints"
    Remove-item $exportCSV
    
    Write-Host "Company ${companyName} found, exported ${endpointCount} endpoints to ${exportXLSX}"
}else{
    Write-Host "Company ${companyName} found, exported ${endpointCount} endpoints to ${exportCSV}"
}
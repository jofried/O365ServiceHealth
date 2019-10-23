##############################################################################################
#
# This script is not officially supported by Microsoft, use it at your own risk.
# Microsoft has no liability, obligations, warranty, or responsibility regarding
# any result produced by use of this file.
#
##############################################################################################
# The sample scripts are not supported under any Microsoft standard support
# program or service. The sample scripts are provided AS IS without warranty
# of any kind. Microsoft further disclaims all implied warranties including, without
# limitation, any implied warranties of merchantability or of fitness for a particular
# purpose. The entire risk arising out of the use or performance of the sample scripts
# and documentation remains with you. In no event shall Microsoft, its authors, or
# anyone else involved in the creation, production, or delivery of the scripts be liable
# for any damages whatsoever (including, without limitation, damages for loss of business
# profits, business interruption, loss of business information, or other pecuniary loss)
# arising out of the use of or inability to use the sample scripts or documentation,
# even if Microsoft has been advised of the possibility of such damages
##############################################################################################


##############################################################################################
# Description:
# This script can be used to access the Office 365 Service Communications Management API.
# Once configured using the app registration from Azure AD, you can view Office 365 Service
# Health Information. This script will output the incident information immediately for you 
# in a grid view, as well as saving all of the Incident, Advisory, and Message Center Posts
# to a CSV file under C:\temp\O365ServiceHealth_<datetime>
##############################################################################################


# App Registration Variables (Must be configured to match the app registration created in Azure AD)
##############################################################################################
$tenantId = "Please insert Tenant ID"
$client_id = "Please insert Client ID"
$client_secret = 'Please insert Client Secret'
##############################################################################################



#OAuth Connection URI
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

#OAuth Token Body
$tokenBody = @{
    client_id     = $client_id
    scope         = "https://manage.office.com/.default"
    client_secret = $client_secret
    grant_type    = "client_credentials"
}

#OAuth Token Request
$tokenRequest = try {

    Invoke-RestMethod -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $tokenBody -ErrorAction Stop

}
catch [System.Net.WebException] {

    Write-Warning "Error: $($_.Exception.Message)"
    
}

$token = $tokenRequest.access_token

# Get O365 Service Health
$o365ServiceHealth = try {

    Invoke-RestMethod -Method Get -Uri "https://manage.office.com/api/v1.0/$tenantid/ServiceComms/Messages" -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop

}
catch [System.Net.WebException] {

    Write-Warning "Error: $($_.Exception.Message)"
    
} 


#Export ALL Data to CSV (This includes Incidents, Advisories, and Message Center post)
$o365ServiceHealthdata = $o365ServiceHealth.Value
$O365healthData = $o365ServiceHealthdata | ForEach-Object { 
    [PSCustomObject]@{
        "Id" = $_.Id
        "Classification" = $_.Classification
        "Status" = $_.Status
        "LastUpdatedTime" = $_.LastUpdatedTime
        "Title" = $_.Title
        "ImpactDescription" = $_.ImpactDescription
        "Messages" = ($_.Messages.MessageText | Out-String).trim()      
        "WorkloadDisplayName" = $_.WorkloadDisplayName
        "ActionType" = $_.ActionType
        "FeatureDisplayName" = $_.FeatureDisplayName        
        "PostIncidentDocumentUrl" = $_.PostIncidentDocumentUrl
        "Severity" = $_.Severity
        "StartTime" = $_.StartTime
        "EndTime" = $_.EndTime
        "AffectedWorkloadDisplayNames" = ($_.AffectedWorkloadDisplayNames |out-string).trim()
        "AffectedTenantCount" = $_.AffectedTenantCount
        "AffectedUserCount" = $_.AffectedUserCount
        "UserFunctionalImpact" = $_.UserFunctionalImpact
        "AdditionalDetails" = ($_.AdditionalDetails |out-string).trim()
        "MessageType" = $_.MessageType
    }
}

$path = Test-Path -Path "C:\temp\"
If ($path -eq "True")
    {
        $O365healthData | export-csv C:\temp\O365ServiceHealth_$((Get-Date).ToString('MM-dd-yyyy_hh-mm-ss')).csv -NoTypeInformation
    }
    Else
    {
        Write-Warning "Filepath 'C:\temp\' does not exist. Please create the filepath 'C:\temp\'"
    }



# List Incidents only in gridView for a quick view of information
$O365healthData | where{$_.Classification -eq "Incident"} |select-object Classification, ID, Status, LastUpdatedTime, Title, ImpactDescription, StartTime, Workload, Severity, EndTime, Messages, PostIncidentDocumentUrl |Out-GridView
    

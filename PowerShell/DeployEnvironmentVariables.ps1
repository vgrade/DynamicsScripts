Write-Host "---start---"
function UpdateEnvVariableCurrentValue {
  param(
      [string]$EnvironmentVariableDefinitionSchemaName,
      [string]$EnvironmentVariableValue
  )
  
  $rec = Get-CrmRecords -EntityLogicalName environmentvariabledefinition -FilterAttribute "schemaname" -FilterOperator eq -FilterValue $EnvironmentVariableDefinitionSchemaName -Fields "*" -TopCount 1
  #Write-Host "log:retrieved parent record from crm"
  $rec
  $parent = New-Object Microsoft.Xrm.Sdk.EntityReference
              $parent.Id = $rec.CrmRecords[0].environmentvariabledefinitionid
              $parent.LogicalName = "environmentvariabledefinition"
  #Write-Host "log:Created parent object"
  #Write-Host "log:Creating record in crm"
  $newrecordId = New-CrmRecord  -EntityLogicalName environmentvariablevalue -Fields @{"value"=$EnvironmentVariableValue;"environmentvariabledefinitionid"=$parent}
  #Write-Host "log:Created record in crm, details below"
  #$newrecordId
  #Write-Host "---end---"
}

Write-Host "log: retrieving pipeline variables"
$allEnvVars = Get-ChildItem env:
$envVarsToUpdate = $allEnvVars.GetEnumerator() | Where-Object { $_.Name -like "abc_*" } 
Write-Host "log: found " $envVarsToUpdate.count " pipelines variables which need to be updated in dynamics."

if ($envVarsToUpdate.count -eq 0) {
  Write-Host "0 environment variables to update in dyanmics"
}
else {
  
  Write-Host "log:Intsalling Microsoft.Xrm.Tooling.CrmConnector.PowerShell"

  Install-Module -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell -Force
  
  Write-Host "log:Installed Microsoft.Xrm.Tooling.CrmConnector.PowerShell"
  
  Write-Host "log:Intsalling Microsoft.Xrm.Data.PowerShell"
  
  Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser -AllowClobber -Force
  
  Write-Host "log:Installed Microsoft.Xrm.Data.PowerShell"
  
  [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12  
  Write-Host "log:Connecting to crm"
  Write-Host $env:destinationEnvironmenturl
  Write-Host $env:app_dynamics365_crmdeploy_clientid
  $conn = Get-CrmConnection -ConnectionString "AuthType=ClientSecret;
  url=$env:destinationEnvironmenturl;
  ClientId=$env:app_dynamics365_crmdeploy_clientid;
  ClientSecret=$env:ClientSecretValue"

   Write-Host "log:Connected to crm org " $conn.ConnectedOrgFriendlyName

  $envVarsToUpdate | ForEach-Object {
    Write-Host "Updating env variable in dynamics name "$_.Name " value " $_.Value
    UpdateEnvVariableCurrentValue -EnvironmentVariableDefinitionSchemaName $_.Name -EnvironmentVariableValue $_.Value
    Write-Host "Updated env variable in dynamics name "$_.Name " value " $_.Value
   }

}

Write-Host "---finish---"

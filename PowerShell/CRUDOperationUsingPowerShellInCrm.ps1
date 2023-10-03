
#---------------------------- Install Required Module----------------------------#
$ModuleName = "Microsoft.Xrm.Data.PowerShell"
Install-Module -Name $ModuleName -Scope CurrentUser -AllowClobber -Force
Import-Module -Name $ModuleName -Force
#---------------------------- Install Required Module----------------------------#

#---------------------------- Create CRM Connection----------------------------#
$conn = ''
# Chnage the org url in below connection string to your crm organization 
$conn = Get-CrmConnection -ConnectionString "authtype=OAuth;url=https://org4f487103.crm.dynamics.com;appid=51f81489-12ee-4a9e-aaae-a2591f45987d;redirecturi=app://58145B91-0C36-4500-8554-080854F2AC97;"
Write-Host "log:Connected to crm org " $conn.ConnectedOrgFriendlyName
$conn
#---------------------------- Create CRM Connection----------------------------#

#---------------------------- CREATE ----------------------------#
# Create a account record 
$CreateAccount_Fieldvalues = @{
  "name"                  = "XYZ Ltd"; 
}
$newAccountGuid = New-CrmRecord  -EntityLogicalName account -Fields  $CreateAccount_Fieldvalues

# Create a contact  record with text, optionset, boolean, datetime and lookup type fields
# Make newly created account as parentcustomer of contact
$parentcustomerid = [Microsoft.Xrm.Sdk.EntityReference]::new("account", [Guid]::new($newAccountGuid))
$preferredcontactmethodcode = [Microsoft.Xrm.Sdk.OptionSetValue]::new(2) #Email

$dateString = "10/3/2023 12:00:00 AM"
$format = "M/d/yyyy h:mm:ss tt"
$birthDate = [System.DateTime]::ParseExact($dateString, $format, [System.Globalization.CultureInfo]::InvariantCulture)


$CreateContact_Fieldvalues = @{
  "firstname"                  = "Tony1"; 
  "lastname"                   = "Baker";
  "preferredcontactmethodcode" = $preferredcontactmethodcode;
  "creditonhold"               = $true
  "birthdate"                  = $birthDate;
  "parentcustomerid"           = $parentcustomerid;
}

$newContactGuid = New-CrmRecord  -EntityLogicalName contact -Fields  $CreateContact_Fieldvalues
#---------------------------- CREATE ----------------------------#

#---------------------------- READ ----------------------------#
#  Retrieve a contact record from crm using guid
$RetrievedContactsByGuid = Get-CrmRecords -EntityLogicalName contact -FilterAttribute contactid -FilterOperator eq -FilterValue $newContactGuid "*" 
$RetrievedContactsByGuid.CrmRecords

#  Retrieve contact records from crm using filter on a single column filter
$RetrievedContactsByName = Get-CrmRecords -EntityLogicalName contact -FilterAttribute firstname -FilterOperator eq -FilterValue "Tony" "*" -TopCount 100
$RetrievedContactsByName.CrmRecords

#  Retrieve contact records from crm using fetch xml
$ContactFetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
 <entity name="contact">
   <attribute name="fullname" />
   <attribute name="telephone1" />
   <attribute name="contactid" />
   <order attribute="fullname" descending="false" />
   <filter type="and">
     <condition attribute="firstname" operator="eq" value="Tony" />
     <condition attribute="lastname" operator="eq" value="Baker" />
   </filter>
 </entity>
</fetch>
"@

$RetrievedContactsByFetch = Get-CrmRecordsByFetch -Fetch  $ContactFetchXml
$RetrievedContactsByFetch.CrmRecords

#---------------------------- READ ----------------------------#


#---------------------------- UPDATE ----------------------------#
# Update the contact record with new value of first name and preferredcontactmethodcode
$preferredcontactmethodcode = [Microsoft.Xrm.Sdk.OptionSetValue]::new(3) #Phone
$Update_Fieldvalues = @{
  "firstname"                  = "Tonyyyyy"; 
  "preferredcontactmethodcode" = $preferredcontactmethodcode
}

Set-CrmRecord -EntityLogicalName contact -Id $newContactGuid -Fields $Update_Fieldvalues

#---------------------------- UPDATE ----------------------------#

#---------------------------- DELETE ----------------------------#
# Delete a contact record by guid
Remove-CrmRecord -EntityLogicalName contact -Id $newContactGuid
#---------------------------- DELETE ----------------------------#

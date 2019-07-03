#region variables
$debuggingon = $false

#Define variables:
$exportcsv = "$PSScriptRoot\csvgalexport.csv"
$delimiter = ";"
[string]$Indentification = "TealContactAdder"
[string]$onlinetarget = "www.teal-consulting.de"
[string]$tcptestprotocol = "http"
[string]$ExportCompanyName = "Teal Technology Consutling GmbH"
[string]$matchmail = "*@teal-consulting.de"
[string]$matchcontactfolder = "\\\\[a-z|A-Z]*.[a-z|A-Z]*\@teal-consulting.de\\[Contacts|Kontakte]"

$propertiestoupdate=@'
FirstName
LastName
JobTitle
BusinessTelephoneNumber
MobileTelephoneNumber
CompanyName
BusinessAddressStreet
BusinessAddressCity
'@ -split "`r`n"

#endregion variables

#region prepare script execution
write-host "Script start ..."

$testtcpresult = Test-NetConnection -ComputerName $onlinetarget -CommonTCPPort $tcptestprotocol

if (!($testtcpresult.tcptestsucceeded)){
    Write-host  "System is offline - please establish an internet connection!"
    exit 99
    }

$usertestsucceeded = $false
[array]$testuserresult = @()

#endregion prepare script execution

#region contact export and import.
#Connection to outlook:

try {
    write-host "Connect to outlook ..."
    [Microsoft.Office.Interop.Outlook.Application] $ConnectionToOutlook = New-Object -ComObject Outlook.Application  -ea 1
    $contactstoexport = $ConnectionToOutlook.Session.GetGlobalAddressList().AddressEntries
    write-host "Outlook connection created."
    }
catch {
    write-host "Outlook connection not created."
    }

#Connect to the contacts default folder in outlook. (Default folder id for Contacts of default account is 10)
try {
    write-host "Connect to the contacts default folder in outlook. ..."
    $defaultcontactsfolder = $ConnectionToOutlook.session.GetDefaultFolder(10)
    $namespace = $ConnectionToOutlook.GetNamespace("MAPI")
    $outlookfolders = $namespace.folders
    $outlooksubfolders = ""
    $folder = ""
    write-host "Connected to the contacts default folder in outlook. ..."
    write-host "Contactsfolder: ... $($defaultcontactsfolder.FullFolderPath)"
    }
catch{
    write-host "Could not connect to the contacts default folder in outlook. ..."
    }

function Add-Contact 
{
param ($folder,$contact)
    $newcontact = $folder.Items.Add()
    $newcontact.FirstName = $contact.FirstName
    $newcontact.lastname = $contact.lastname
    $newcontact.JobTitle = $contact.JobTitle
    $newcontact.MobileTelephoneNumber = $contact.MobileTelephoneNumber
    $newcontact.OfficeLocation = $contact.OfficeLocation
    $newcontact.Email1Address = $contact.Email1Address
    $newcontact.Email1AddressType = $contact.Email1AddressType
    $newcontact.User1 = $contact.Email1Address
    $newcontact.BusinessAddressState = $contact.StateOrProvince
    $newcontact.BusinessAddressStreet = $contact.BusinessAddressStreet
    $newcontact.ManagerName = $contact.Manager
    $newcontact.AssistantName = $contact.AssistantName
    $newcontact.BusinessTelephoneNumber = $contact.BusinessTelephoneNumber
    $newcontact.BusinessAddressCity = $contact.BusinessAddressCity
    $newcontact.CompanyName = $contact.CompanyName
    $newcontact.Department = $contact.Department
    $newcontact.Save()
    Start-Sleep -Seconds 10
}

function Update-Contact 
{
param ($item,$contact,$properties)
        
    foreach ($property in $properties){
        if ($property -like "BusinessTelephoneNumber"){
            if (((($item.$($property).ToString()) -replace "\s","") -replace "\+","") -notmatch ((($contact.$($property).ToString()) -replace "\s","") -replace "\+","")){
                if ($debuggingon){write-host "$Property - OLDvalue:$item.$($property) NEWValue: $contact.$($property)"}
                $item.$($property) = $contact.$($property)
                $item.Save()
                Start-Sleep -Seconds 10
                }
            continue
            }
        if ($property -like "MobileTelephoneNumber"){
            if (((($item.$($property).ToString()) -replace "\s","") -replace "\+","") -notmatch ((($contact.$($property).ToString()) -replace "\s","") -replace "\+","")){
                if ($debuggingon){write-host "$Property - OLDvalue:$item.$($property) NEWValue: $contact.$($property)"}
                $item.$($property) = $contact.$($property)
                $item.Save()
                Start-Sleep -Seconds 10
                }
            continue
            }
        if ($item.$($property) -notlike $contact.$($property)){
            if ($debuggingon){write-host "$Property - OLDvalue:$item.$($property) NEWValue: $contact.$($property)"}
            $item.$($property) = $contact.$($property)
            $item.Save()
            Start-Sleep -Seconds 10
            continue
            }
        write-host "Nothing to Update: ... $property"
        }
}

#Test if old export exists, if so it get deleted.
if (Test-path $exportcsv) {
    Remove-Item -Path $exportcsv
    }


#Export of contacts with all possible values:

write-host "Export contacts ..."

foreach ($contacttoexport in $contactstoexport){
    #Verify contact is not empty. If empty process next contact.
    if ($contacttoexport -eq $null ) { continue }

    #Verify contact is an User contact. If not process next contact.
    if ($contacttoexport.AddressEntryUserType -ne "0") { continue }

    #Create objects for export.
    
    try {
        $firstname = $contacttoexport.GetExchangeUser().firstname
        $lastname = $contacttoexport.GetExchangeUser().lastname
        $JobTitle = $contacttoexport.GetExchangeUser().JobTitle
        $MobileTelephoneNumber = $contacttoexport.GetExchangeUser().MobileTelephoneNumber
        $OfficeLocation = $contacttoexport.GetExchangeUser().OfficeLocation
        $Postalcode = $contacttoexport.GetExchangeUser().PostalCode
        $Email1Address = $contacttoexport.GetExchangeUser().PrimarySmtpAddress
        $Email1AddressType = "SMTP"
        $StateOrProvince = $contacttoexport.GetExchangeUser().StateOrProvince
        $StreetAddress = $contacttoexport.GetExchangeUser().StreetAddress
        $Manager = $contacttoexport.GetExchangeUser().Manager
        $Name = $contacttoexport.GetExchangeUser().Name
        $Alias = $contacttoexport.GetExchangeUser().Alias
        $AssistantName = $contacttoexport.GetExchangeUser().AssistantName
        $BusinessTelephoneNumber = $contacttoexport.GetExchangeUser().BusinessTelephoneNumber
        $City = $contacttoexport.GetExchangeUser().City
        $CompanyName = $ExportCompanyName
        $Department = $contacttoexport.GetExchangeUser().Department
        }

    catch {
        write-host "ERROR: Failed to create export object ..."
        }

    #Identify if the contact is a contact to export, defined by the Busniess Telephone number.
    #If the contact have no number, next contact is processed.
    if ($MobileTelephoneNumber.Length -le "0" ) { continue }

    #Create object for CSV export. And export the object to csv. As append is used, it need to be a new file.
    try {
       write-host "Export contacts to csv (only contacts which are from type Address Entry User) ... "
        [PSCustomobject]@{
            FirstName = $firstname
            LastName = $lastname
            Name = $Name
            Alias = $Alias
            JobTitel = $JobTitel
            Email1Address = $Email1Address
            Email1AddressType = $Email1AddressType
            MobileTelephoneNumber = $MobileTelephoneNumber
            BusinessTelephoneNumber = $BusinessTelephoneNumber
            BusinessAddressStreet = $StreetAddress
            BusinessAddressCity = $City
            BusinessAddressPostalCode = $Postalcode
            StateOrProvince = $StateOrProvince
            Department = $Department
            CompanyName = $CompanyName
            OfficeLocation = $OfficeLocation
            } | Export-Csv $exportcsv -encoding Default -Delimiter $delimiter -NoTypeInformation -Append
        write-host "Export object to csv ..."
        }
    catch {
        write-host "ERROR:  Failed to export object to csv ..."
        }
    }

#Import exportet contact to outlook default contact folder.
if ($defaultcontactsfolder.FolderPath -match $matchcontactfolder){
    write-host "Get contact folder ..."
    $DefaultAddressBookID = $defaultcontactsfolder.EntryID
    $DefaultAddressBookNamespace = $namespace.GetFolderFromID($DefaultAddressBookID)
    $folder = $DefaultAddressBookNamespace
    if ($debuggingon) {write-host "Contact folder: ... $($DefaultAddressBookNamespace.FullFolderPath)"}
    else {write-host "Contact folder: ... "}
    }

#get existing contacts.
Start-Sleep -Seconds 10
$existingcontacts = @()
write-host "Start exporting all existing $matchmail contacts from local addressbook. ..."
foreach ($item in $folder.items | Where-Object {$_.User1 -like $matchmail}){
    Start-Sleep -Seconds 2
    $existingcontacts += $item.User1
    }

$existingcontactcount = 1
foreach ($existingcontact in $existingcontacts) {
    if ($debuggingon) {write-host "Existing contatc $existingcontactcount of $($existingcontacts.count): $($existingcontact)"}
    $existingcontactcount ++
    }

# Add Contacts to Contact Folder.
if (Test-path $exportcsv) {
    $ContactsImport = Import-csv $exportcsv -Delimiter $delimiter

    write-host "Start creating or updating contacts ..."
    if ($folder){
        
        foreach ($contacttoimport in $contactsimport) {
            $contactexists = $false
            $updatecontact = @()
            if ($existingcontacts -contains $contacttoimport.Email1Address) {
                $contactexists = $true
                $updatecontact = $folder.items | Where-Object {$_.User1 -match $contacttoimport.Email1Address}
                }

            if ($contactexists) {
                try {
                    if ($debuggingon) {write-host "Update contact ... $($contacttoimport.Email1Address)"}
                    else {write-host "Update contact ... "}
                    Update-Contact $updatecontact $contacttoimport $propertiestoupdate
                    write-host "Update contact ..."
                    }
                    catch{
                        write-host "Failed to Update contact ..."
                        }
                 }
            if (!($contactexists)) {
                try {
                    if ($debuggingon) {write-host "Create contact ... $($contacttoimport.Email1Address)"}
                    else {write-host "Create contact ... "}
                    Add-Contact $folder $contacttoimport 
                    write-host "Create contact ..."
                    }
                catch{
                    write-host "Failed to Create contact ..."
                    }
                }
        }
    }
}

if (Test-path $exportcsv) {
    if (!($debuggingon)) {Remove-Item -Path $exportcsv}
    }

write-host "Finished script deleting csv ..."

exit 0
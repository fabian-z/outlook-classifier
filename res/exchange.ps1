# Prequisites: ApplicationImpersionation for user account!
# New-ManagementRoleAssignment -name:impersonationAssignmentName -Role:ApplicationImpersonation -User:serviceAccount 

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

Import-Module "./Microsoft.Exchange.WebServices.dll"
$ewsService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)

# Configuration required here:

$autodiscoverUrl = "zaremba@concept.lab"

$categoriesToEnforce = @{
    "TLP White" = [PSCustomObject]@{
        Subject       = "[Classified White ⚪]"
        Color         = "-1"
        FoundCategory = $false
        FoundRule     = $false
    }
    "TLP Green" = [PSCustomObject]@{
        Subject       = "[Classified Green 🟢]"
        Color         = "4"
        FoundCategory = $false
        FoundRule     = $false
    }
    "TLP Amber" = [PSCustomObject]@{
        Subject       = "[Classified Amber 🟠]"
        Color         = "1"
        FoundCategory = $false
        FoundRule     = $false
    }
    "TLP Red"   = [PSCustomObject]@{
        Subject       = "[Classified Red 🔴]"
        Color         = "0"
        FoundCategory = $false
        FoundRule     = $false
    }
}

## Start of script

#$username=Read-Host -Prompt "Enter UserName"
#$password=Read-host -Prompt "Enter Password" -AsSecureString
#$domain=Read-Host -Prompt "Enter Domain"
#$creds = New-Object System.Management.Automation.PSCredential($username,$password)

$creds = Get-Credential

$ewsService.Credentials = $creds.GetNetworkCredential()
#$ewsService.UseDefaultCredentials = $true

$ewsService.AutodiscoverUrl($autodiscoverUrl)

# Check for URL
if (!$ewsService.Url) {
    Write-Host "EWS Service Url is blank"
    exit
}

Write-Host "Service URL = $($ewsService.Url)"

Get-Mailbox -RecipientTypeDetails UserMailbox | ForEach-Object { 

    $targetMailboxObject = $_
    $targetMailbox = $_.PrimarySmtpAddress
    $ewsService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $targetMailbox);

    ## Reset tracking values
    foreach ($c in $categoriesToEnforce.GetEnumerator()) {
        $c.Value.FoundRule = $false
        $c.Value.FoundCategory = $false
    }

    ## Process categories
    $mb = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($targetMailbox)	

    $folderCalendar = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar
    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($folderCalendar, $mb)

    Write-Host $folderId
    Write-Host "Processing categories for mailbox $mailbox"

    # If you just wanted YOUR (cred) mailbox you wouldn't need to fetch the folderID
    #$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ewsService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
	
    # Use the folderID of a mailbox folder the provided creds have access to.

    $xmlProperty = [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::XmlData
    $userConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($ewsService, "CategoryList", $folderId, $xmlProperty)

    #$userConfig.Delete()
    #exit

    $xmlData = $userConfig.XmlData
    $xmlString = ""

    if (($xmlData[0]) -eq 0xEF -and ($xmlData[1]) -eq 0xBB -and ($xmlData[2] -eq 0xBF)) {
        $cleanXmlData = $xmlData[3..$xmlData.Length]
        $xmlString = [System.Text.Encoding]::UTF8.GetString($cleanXmlData)
    }
    else {
        $xmlString = [System.Text.Encoding]::UTF8.GetString($xmlData)
    }

    #Write-Host $xmlString
    $xml = New-Object -TypeName XML
    $xml.LoadXml($xmlString)

    $xml.categories.category | ForEach-Object {
        $categoryFound = $_.name
        if ($categoriesToEnforce.ContainsKey($categoryFound)) {
            Write-Host "Found $($categoryFound)"
            $categoriesToEnforce[$categoryFound].FoundCategory = $true
        }
    }

    foreach ($c in $categoriesToEnforce.GetEnumerator() ) {
        if (-not $c.Value.FoundCategory) {
            Write-Host "Missing $($c.Name) : $($c.Value)"

            $timestamp = [datetime]::Now.ToUniversalTime().ToString("yyyy-mm-ddThh:mm:ss.fffffffZ")

            $newCat = $xml.CreateNode("element", "category", $xml.categories.NamespaceURI)
            $newCat.SetAttribute("renameOnFirstUse", "0")
            $newCat.SetAttribute("name", $c.Name)
            $newCat.SetAttribute("color", $c.Value.Color)
            $newCat.SetAttribute("keyboardShortcut", "0")
            $newCat.SetAttribute("lastTimeUsedNotes", $timestamp)
            $newCat.SetAttribute("lastTimeUsedJournal", $timestamp)
            $newCat.SetAttribute("lastTimeUsedContacts", $timestamp)
            $newCat.SetAttribute("lastTimeUsedTasks", $timestamp)
            $newCat.SetAttribute("lastTimeUsedCalendar", $timestamp)
            $newCat.SetAttribute("lastTimeUsedMail", $timestamp)
            $newCat.SetAttribute("lastTimeUsed", $timestamp)
            $newCat.SetAttribute("guid", "{$([guid]::NewGuid())}")

            $xml.categories.AppendChild($newCat)

        }
    }

    $newXmlString = [System.Text.Encoding]::UTF8.GetBytes($xml.OuterXml)
    $userConfig.XmlData = $newXmlString
    $userConfig.Update()

    Write-Host "Updated $($targetMailbox) categories"

    ## Process rules

    Get-InboxRule -Mailbox $targetMailboxObject | ForEach-Object {
        $ruleFound = $_.Name
        if ($categoriesToEnforce.ContainsKey($ruleFound)) {
            Write-Host "Found $($ruleFound)"
            $categoriesToEnforce[$ruleFound].FoundRule = $true
        }
    }

    foreach ($c in $categoriesToEnforce.GetEnumerator() ) {
        if (-not $c.Value.FoundRule) {
            Write-Host "Missing $($c.Name) : $($c.Value)"

            New-InboxRule $c.Name -Mailbox $targetMailboxObject -SubjectContainsWords $c.Value.Subject -ApplyCategory $c.Name
        }
    }
}

Write-Host "Updated $($targetMailbox) rules"


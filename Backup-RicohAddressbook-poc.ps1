function ConvertTo-Base64
{
param($String)
[System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($String))
}

function Connect-MFP{
param(
    $Hostname,
    $Authentication,
    $Username,
    $Password,
    $SecurePassword, 
    $Protocol
    )

#$url = "$Protocol" + "://" + $Hostname + "/DH/udirectory"
$url = $Protocol + "://" + $Hostname + "/DH/udirectory"
Write-host "Connect-MFP $url "
#$url = "http://$Hostname/DH/udirectory"

$login = [xml]@'
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
 <s:Body>
  <m:startSession xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
   <stringIn></stringIn>
   <timeLimit>30</timeLimit>
   <lockMode>X</lockMode>
  </m:startSession>
 </s:Body>
</s:Envelope>
'@
if($SecurePassword -eq $NULL){$pass = ConvertTo-Base64 $Password}else{$pass = $SecurePassword; $enc = "gwpwes003"}
$login.Envelope.Body.startSession.stringIn = "SCHEME="+(ConvertTo-Base64 $Authentication)+";UID:UserName="+(ConvertTo-Base64 $Username)+";PWD:Password=$pass;PES:Encoding=$enc"
try{
[xml]$xml = Invoke-WebRequest $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#startSession"} -Body $login
if($xml.Envelope.Body.startSessionResponse.returnValue -eq "OK"){
    $sessionID = $xml.Envelope.Body.startSessionResponse.stringOut}
    Return $sessionID
}catch{
    $SessionID = $False
    Return $SessionID
}
}

function Search-MFPAB
{
param(
    [string]$Hostname, 
    [string]$SessionID
    )
    
#$url = "http://$Hostname/DH/udirectory"
$url = $Protocol + "://" + $Hostname + "/DH/udirectory"
Write-host "Search-MFPAB $url"
$search = [xml]@'
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
 <s:Body>
  <m:searchObjects xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
    <sessionId></sessionId>
   <selectProps xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.w3.org/2001/XMLSchema" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" soap-enc:arrayType="itt:string[1]">
    <item>id</item>
   </selectProps>
    <fromClass>entry</fromClass>
    <parentObjectId></parentObjectId>
    <resultSetId></resultSetId>
   <whereAnd xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" soap-enc:arrayType="itt:queryTerm[1]">
    <item>
     <operator></operator>
     <propName>all</propName>
     <propVal></propVal>
     <propVal2></propVal2>
    </item>
   </whereAnd>
   <whereOr xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" soap-enc:arrayType="itt:queryTerm[1]">
    <item>
     <operator></operator>
     <propName></propName>
     <propVal></propVal>
     <propVal2></propVal2>
    </item>
   </whereOr>
   <orderBy xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" soap-enc:arrayType="itt:queryOrderBy[1]">
    <item>
     <propName></propName>
     <isDescending>false</isDescending>
    </item>
   </orderBy>
    <rowOffset>0</rowOffset>
    <rowCount>50</rowCount>
    <lastObjectId></lastObjectId>
   <queryOptions xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" soap-enc:arrayType="itt:property[1]">
    <item>
     <propName></propName>
     <propVal></propVal>
    </item>
   </queryOptions>
  </m:searchObjects>
 </s:Body>
</s:Envelope>
'@
$search.Envelope.Body.searchObjects.sessionId = $SessionID
try{
    [xml]$xml = Invoke-WebRequest $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#searchObjects"} -Body $search 
    #$xml.Save("c:\temp\dump.xml")
    $AddressBook = $xml.SelectNodes("//rowList/item") | ForEach-Object {$_.item.propVal} | Where-Object {$_.length -lt "10"} | ForEach-Object {[int]$_} | Sort-Object
    $ReturnValue = $AddressBook
    Return $ReturnValue
}catch{
    $ReturnValue = "Error"
    Return $ReturnValue
}Finally{

}
}

function Get-MFPAB
{
param($Hostname,$Authentication="BASIC",$Username="admin",$Password,$SecurePassword,$SessionID, $Protocol)
#Connect-MFP $Hostname $Authentication $Username $Password $SecurePassword
write-host "Get-MFPAB"
#$url = "http://$Hostname/DH/udirectory"
$url = $Protocol + "://" + $Hostname + "/DH/udirectory"
write-host "Get-MFPAB $url"
$get = [xml]@'
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
 <s:Body>
  <m:getObjectsProps xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
   <sessionId></sessionId>
  <objectIdList xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.w3.org/2001/XMLSchema" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="">
  </objectIdList>
  <selectProps xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.w3.org/2001/XMLSchema" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="itt:string[71]">
  <item>entryType</item>
<item>id</item>
<item>name</item>
<item>longName</item>
<item>phoneticName</item>
<item>index</item>
<item>passwordEncoding</item>
<item>isDestination</item>
<item>isSender</item>
<item>auth:</item>
<item>auth:name</item>
<item>auth:password</item>
<item>password:</item>
<item>password:password</item>
<item>password:usedForMailSender</item>
<item>password:usedForRemoteFolder</item>
<item>password:passwordEncoding</item>
<item>mail:</item>
<item>mail:address</item>
<item>mail:parameter</item>
<item>mail:isDirectSMTP</item>
<item>fax:</item>
<item>fax:number</item>
<item>fax:lineType</item>
<item>fax:isAbroad</item>
<item>fax:parameter</item>
<item>faxAux:</item>
<item>faxAux:ttiNo</item>
<item>faxAux:label1</item>
<item>faxAux:label2String</item>
<item>faxAux:messageNo</item>
<item>remoteFolder:</item>
<item>remoteFolder:type</item>
<item>remoteFolder:serverName</item>
<item>remoteFolder:path</item>
<item>remoteFolder:accountName</item>
<item>remoteFolder:password</item>
<item>remoteFolder:port</item>
<item>remoteFolder:characterEncoding</item>
<item>remoteFolder:passwordEncoding</item>
<item>remoteFolder:select</item>
<item>ldap:</item>
<item>ldap:accountName</item>
<item>ldap:password</item>
<item>ldap:passwordEncoding</item>
<item>ldap:select</item>
<item>smtp:</item>
<item>smtp:accountName</item>
<item>smtp:password</item>
<item>smtp:passwordEncoding</item>
<item>smtp:select</item>
<item>ifax:</item>
<item>ifax:address</item>
<item>ifax:parameter</item>
<item>ifax:isDirectSMTP</item>
<item>tagId</item>
<item>auth:lockOut</item>
<item>browser:</item>
<item>browser:homepageUrl</item>
<item>browser:fontSize</item>
<item>browser:charCode</item>
<item>browser:isHistoryKeeping</item>
<item>browser:historyKeepingDate</item>
<item>browser:isUsingProxySrv</item>
<item>browser:proxySrvUrl</item>
<item>browser:proxySrvPort</item>
<item>browser:proxySrvAccountName</item>
<item>browser:proxySrvPassword</item>
<item>browser:proxyPasswordEncoding</item>
<item>browser:proxyException</item>
<item>browser:isURLBarDisplay</item>
<item>browser:isPanelLocked</item>
  </selectProps>
   <options xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="itt:property[1]">
    <item>
     <propName></propName>
     <propVal></propVal>
    </item>
   </options>
  </m:getObjectsProps>
 </s:Body>
</s:Envelope>
'@
$get.Envelope.Body.getObjectsProps.sessionId = $SessionID
Search-MFPAB -Hostname $Hostname -SessionID $SessionID | ForEach-Object{
$x = $get.CreateElement("item")
$x.set_InnerText("entry:$_")
$o = $get.Envelope.Body.getObjectsProps.objectIdList.AppendChild($x)
}
$get.Envelope.Body.getObjectsProps.objectIdList.arrayType = "itt:string["+$get.Envelope.Body.getObjectsProps.objectIdList.item.count+"]"
try{
[xml]$xml = Invoke-WebRequest $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#getObjectsProps"} -Body $get
$xml.Save("c:\temp\dump.xml")
$xml.SelectNodes("//returnValue/item") | ForEach-Object{
New-Object PSObject -Property @{
    entryType                    =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "entryType"}).propVal
    id                           =   [int](Foreach-Object{$_.item} | Where-Object{$_.propName -eq "id"}).propVal 
    name                         =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "name"}).propVal
    longName                     =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "longname"}).propVal
    phoneticName                 =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "longname"}).propVal
    index                        =   [int](Foreach-Object{$_.item} | Where-Object{$_.propName -eq "phoneticName"}).propVal
    passwordEncoding             =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "passwordEncoding"}).propVal
    isDestination                =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "isDestination"}).propVal
    isSender                     =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "isSender"}).propVal
    Auth                         =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "auth:"}).propVal
    AuthName                     =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "auth:name"}).propVal
    AuthPassword                 =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "auth:password"}).propVal
    Password                     =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "password"}).propVal
    PasswordUsedForMailSender    =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "password:usedForMailSender"}).propVal
    PasswordUsedForRemoteFolder  =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "password:usedForRemoteFolder"}).propVal
    PasswordPasswordEncoding     =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "password:passwordEncoding"}).propVal
    Mail                         =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "mail"}).propVal
    MailAddress                  =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "mail:address"}).propVal
    MailParameter                =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "mail:parameter"}).propVal
    MailisDirectSMTP             =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "mail:isDirectSMTP"}).propVal
    Fax                          =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "fax:"}).propVal
    FaxNumber                    =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "fax:number"}).propVal
    FaxLineType                  =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "fax:lineType"}).propVal
    FaxisAbroad                  =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "fax:isAbroad"}).propVal
    FaxParameter                 =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "fax:parameter"}).propVal
    FaxAux                       =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "faxAux:"}).propVal
    FaxAuxTtiNo                  =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "faxAux:ttiNo"}).propVal
    FaxAuxLabel1                 =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "faxAux:label1"}).propVal
    FaxAuxLabel2String           =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "faxAux:label2String"}).propVal
    FaxAuxMessageNo              =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "faxAux:messageNo"}).propVal
    RemoteFolder                 =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "remoteFolder:"}).propVal
    RemoteFolderType             =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "remoteFolder:type"}).propVal
    RemoteFolderServerName       =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "remoteFolder:serverName"}).propVal
    RemoteFolderPath             =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "remoteFolder:path"}).propVal
    RemoteFolderAccountName      =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "remoteFolder:accountName"}).propVal
    RemoteFolderPassword         =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "remoteFolder:password"}).propVal
    RemoteFolderPort             =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "remoteFolder:port"}).propVal
    RemoteFolderCharacterEncoding  =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "remoteFolder:characterEncoding"}).propVal
    RemoteFolderPasswordEncoding   =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "remoteFolder:passwordEncoding"}).propVal
    RemoteFolderSelect           =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "remoteFolder:select"}).propVal
    ldap                       =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "ldap:"}).propVal
    LdapAccountName            =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "ldap:accountName"}).propVal
    LdapPassword               =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "ldap:password"}).propVal
    LdapPasswordEncoding       =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "ldap:passwordEncoding"}).propVal
    LdapSelect                 =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "ldap:select"}).propVal
    Smtp                       =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "smtp:"}).propVal
    SmtpAccountName            =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "smtp:accountName"}).propVal
    SmtpPassword               =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "smtp:password"}).propVal
    SmtpPasswordEncoding       =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "smtp:passwordEncoding"}).propVal
    SmtpSelect                 =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "smtp:select"}).propVal
    Ifax                       =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "ifax:"}).propVal 
    IfaxAddress                =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "ifax:address"}).propVal
    IfaxParameter              =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "ifax:parameter"}).propVal
    IfaxisDirectSMTP           =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "ifax:isDirectSMTP"}).propVal
    tagId                      =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "tagId"}).propVal
    authlockOut                =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "auth:lockOut"}).propVal
    Browser                    =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:"}).propVal
    BrowserHomepageUrl         =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:homepageUrl"}).propVal
    BrowserfontSize            =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:fontSize"}).propVal
    BrowserCharCode            =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:charCode"}).propVal
    BrowserisHistoryKeeping    =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:isHistoryKeeping"}).propVal
    BrowserhistoryKeepingDate  =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:historyKeepingDate"}).propVal
    BrowsersUsingProxySrv      =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:isUsingProxySrv"}).propVal
    BrowserProxySrvUrl         =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:isUsingProxySrv"}).propVal
    BrowserProxySrvPort        =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:proxySrvUrl"}).propVal
    BrowserProxySrvAccountName =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:proxySrvAccountName"}).propVal
    BrowserProxySrvPassword    =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:proxySrvPassword"}).propVal
    BrowserProxyPasswordEncoding   =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:proxyPasswordEncoding"}).propVal
    BrowserProxyException      =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:proxyException"}).propVal
    BrowserisURLBarDisplay     =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:isURLBarDisplay"}).propVal
    BrowserisPanelLocked       =   (Foreach-Object{$_.item} | Where-Object{$_.propName -eq "browser:isPanelLocked"}).propVal
}} 
}catch{
    write-host "error"
}Finally{
}
Disconnect-MFP -Hostname $Hostname -SessionID $SessionID $Protocol
}

function Add-MFPAB
{

<#
param(
		$Hostname,
		$Protocol,
		$SessionID,
   		[string]$password,
		[string]$UserCode,
		[string]$entryType="User",
		[string]$id, 
		[string]$name,
		[string]$longName, 
		[string]$phoneticName, 
		[string]$index, 
		[string]$passwordEncoding, 
		[string]$isDestination="true", 
		[string]$isSender="false", 
		[string]$Auth, 
		[string]$AuthName, 
		[string]$AuthPassword,    
		[string]$PasswordUsedForMailSender, 
		[string]$PasswordUsedForRemoteFolder, 
		[string]$PasswordPasswordEncoding, 
		[string]$Mail="True", 
		[string]$MailAddress, 
		[string]$MailParameter, 
		[string]$MailisDirectSMTP, 
		[string]$Fax,   
		[string]$FaxNumber, 
		[string]$FaxLineType, 
		[string]$FaxisAbroad,  
		[string]$FaxParameter,  
		[string]$FaxAux,   
		[string]$FaxAuxTtiNo,  
		[string]$FaxAuxLabel1,
		[string]$FaxAuxLabel2String,  
		[string]$FaxAuxMessageNo,   
		[string]$RemoteFolder, 
		[string]$RemoteFolderType,   
		[string]$RemoteFolderServerName,   
		[string]$RemoteFolderPath,   
		[string]$RemoteFolderAccountName,  
		[string]$RemoteFolderPassword,
		[string]$RemoteFolderPort,   
		[string]$RemoteFolderCharacterEncoding,  
		[string]$RemoteFolderPasswordEncoding, 
		[string]$RemoteFolderSelect,  
		[string]$ldap,  
		[string]$LdapAccountName,  
		[string]$LdapPassword,  
		[string]$LdapPasswordEncoding,  
		[string]$LdapSelect,  
		[string]$Smtp, 
		[string]$SmtpAccountName,   
		[string]$SmtpPassword,  
		[string]$SmtpPasswordEncoding,  
		[string]$SmtpSelect, 
		[string]$Ifax,   
		[string]$IfaxAddress,   
		[string]$IfaxParameter, 
		[string]$IfaxisDirectSMTP,  
		[string]$tagId,  
		[string]$authlockOut, 
		[string]$Browser,   
		[string]$BrowserHomepageUrl,  
		[string]$BrowserfontSize,   
		[string]$BrowserCharCode,  
		[string]$BrowserisHistoryKeeping, 
		[string]$BrowserhistoryKeepingDate,
		[string]$BrowsersUsingProxySrv,
		[string]$BrowserProxySrvUrl,   
		[string]$BrowserProxySrvPort,   
		[string]$BrowserProxySrvAccountName, 
		[string]$BrowserProxySrvPassword,
		[string]$BrowserProxyPasswordEncoding,   
		[string]$BrowserProxyException,  
		[string]$BrowserisURLBarDisplay, 
		[string]$BrowserisPanelLocked 
)#> 

param(
$Hostname,
$Authentication="BASIC",
$Username="admin",
$Password,
$SecurePassword,
$EntryType="user",
$Index,
$Name,
$LongName,
$UserCode,
$isDestination="true",
$isSender="false",
$Mail="true",
$MailAddress,
$Protocol,
$SessionID
)

#Connect-MFP $Hostname $Authentication $Username $Password $SecurePassword
#$url = "http://$Hostname/DH/udirectory"
$url = $Protocol + "://" + $Hostname + "/DH/udirectory"
Write-host "Add-MFPAB $url" 
$add = [xml]@'
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
 <s:Body>
  <m:putObjects xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
   <sessionId></sessionId>
   <objectClass>entry</objectClass>
   <parentObjectId></parentObjectId>
   <propListList xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="">
	<item xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="itt:property[7]">

	</item>	
   </propListList>
  </m:putObjects>
 </s:Body>
</s:Envelope>
'@
$add.Envelope.Body.putObjects.sessionId = $SessionID

#UserCode
if($UserCode -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("UserCode")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($UserCode)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#entrytype
if($EntryType -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("entryType")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($EntryType)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}
#id
if($id -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("id")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($id)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}
#name
if($name -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("name")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($name)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#longName
if($longName -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("longName")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($longName)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}
#phoneticName
if($phoneticName -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("phoneticName")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($phoneticName)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}
#index
if($index -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("index")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($index)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#passwordEncoding
if($passwordEncoding -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("passwordEncoding")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($passwordEncoding)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#isDestination
if($isDestination -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("isDestination")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($isDestination)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}
#isSender
if($isSender -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("isSender")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($isSender)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#auth:
if($Auth -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("auth:")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($Auth)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#Auth:Name
if($AuthName -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("auth:name")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($AuthName)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#Auth:Password
if($AuthPassword -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("auth:password")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($AuthPassword)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#password:
if($password -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("password:")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($password)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#password:password
if($PasswordPassword -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("password:password")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($passwordpassword)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#password:usedForMailSender
if($passwordusedForMailSender -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("password:usedForMailSender")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($passwordusedForMailSender)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#password:usedForMailSender
if($passwordusedForRemoteFolder -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("password:usedForRemoteFolder")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($passwordusedForRemoteFolder)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#password:usedForMailSender
if($passwordpasswordEncoding -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("password:passwordEncoding")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($PasswordPasswordEncoding)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#password:usedForMailSender
if($passwordpasswordEncoding -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("password:passwordEncoding")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($PasswordPasswordEncoding)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#mail:
if($mail -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("mail:")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($mail)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#mail:address
if($MailAddress -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("mail:address")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($MailAddress)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#Mail:Parameter
if($MailParameter -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("mail:parameter")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($MailParameter)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#Mail:Parameter
if($MailParameter -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("mail:parameter")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($MailParameter)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}
#mail:isDirectSMTP
if($mailisDirectSMTP -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("mail:isDirectSMTP")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($MailisDirectSMTP)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#fax:
if($fax -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("fax:")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($fax)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#faxNumber
if($FaxNumber -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("fax:Name")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($faxNumber)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#fax:lineType
if($FaxNumber -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("fax:lineType")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($faxNlineType)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#fax:lineType
if($faxlineType -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("fax:lineType")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($faxlineType)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#fax:isAbroad
if($faxisAbroad -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("fax:isAbroad")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($faxisAbroad)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#fax:parameter
if($faxparameter -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("fax:parameter")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($faxparameter)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#faxAux:
if($faxAux -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("faxAux:")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($faxAux)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#faxAux:ttiNo
if($faxAuxttiNo -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("faxAux:ttiNo")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($FaxAuxTtiNo)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#faxAux:label1
if($faxAux:label1 -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("faxAux:label1")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($faxAuxlabel1)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#faxAux:label1
if($faxAux:label1 -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("faxAux:label1")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($faxAuxlabel1)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#faxAux:label2String
if($faxAux:label1 -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("faxAux:label2String")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($faxAuxlabel2String)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#faxAux:messageNo
if($faxAux:messageNo -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("faxAux:messageNo")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($faxAuxmessageNo)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#remoteFolder:
if($remoteFolder -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFolder)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#remoteFoldertype:
if($remoteFoldertype -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:type")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFoldertype)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#remoteFolderserverName:
if($remoteFolderserverName -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:serverName")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFoldertypeserverName)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#remoteFolderpath:
if($remoteFolderpath -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:path")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFolderpath)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#remoteFolderaccountName:
if($remoteFolder -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:accountName")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFolderaccountName)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#remoteFolderpassword:
if($remoteFolderPassword -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:password")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFolderpassword)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#remoteFolderport:
if($remoteFolderport -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:port")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFolderport)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#remoteFoldercharacterEncoding:
if($remoteFoldercharacterEncoding -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:characterEncoding")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFoldercharacterEncoding)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#remoteFolderpasswordEncoding
if($remoteFolderpasswordEncoding -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:passwordEncoding")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFolderpasswordEncoding)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#remoteFolderpasswordEncoding
if($remoteFolderpasswordEncoding -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:passwordEncoding")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFolderpasswordEncoding)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#remoteFolderselect
if($remoteFolderselect -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("remoteFolder:select")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($remoteFolderselect)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#ldap:
if($ldap -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("ldap:")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($ldap)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}
#ldap:accountName
if($ldapAccountName -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("ldap:accountName")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($LdapAccountName)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#ldap:password
if($ldappassword -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("ldap:password")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($ldappassword)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#ldap:passwordEncoding
if($ldappasswordEncoding -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("ldap:passwordEncoding")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($ldappasswordEncoding)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#ldap:select
if($ldapselect -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("ldap:select")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($ldapselect)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#ldap:select
if($ldapselect -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("ldap:select")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($ldapselect)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#smtp
if($smtp -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("smtp:")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($smtp)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#smtp:accountName
if($smtpaccountName -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("smtp:accountName")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($smtpaccountName)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#smtp:password
if($smtppassword -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("smtp:password")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($smtppassword)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#smtp:passwordEncoding
if($smtppasswordEncoding -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("smtp:passwordEncoding")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($smtppasswordEncoding)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#smtp:select
if($smtpselect -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("smtp:select")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($smtpselect)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#ifax:
if($ifax -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("ifax:")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($ifax)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#ifax:address
if($ifaxaddress -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("ifax:address")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($ifaxaddress)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#ifax:parameter
if($ifaxparameter -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("ifax:parameter")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($ifaxparameter)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#ifax:isDirectSMTP
if($ifaxisDirectSMTP -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("ifax:isDirectSMTP")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($ifaxisDirectSMTP)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#tagId
if($tagId -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("tagId")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($tagId)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#auth:lockOut
if($authlockOut -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("auth:lockOut")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($authlockOut)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:
if($browser -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browser)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#browser:homepageUrl
if($browserhomepageUrl -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:homepageUrl")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserhomepageUrl)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#browser:fontSize
if($browserfontSize -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:fontSize")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserfontSize)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:charCode
if($browsercharCode -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:charCode")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browsercharCode)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:isHistoryKeeping
if($browserisHistoryKeeping -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:isHistoryKeeping")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserisHistoryKeeping)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:historyKeepingDate
if($browserhistoryKeepingDate -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:historyKeepingDate")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserhistoryKeepingDate)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:isUsingProxySrv
if($browserhistoryKeepingDate -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:isUsingProxySrv")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserisUsingProxySrv)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:iproxySrvUrl
if($browserhistoryKeepingDate -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:proxySrvUrl")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserproxySrvUrl)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:proxySrvPort
if($browserhistoryKeepingDate -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:proxySrvPort")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserproxySrvPort)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:proxySrvPort
if($browserproxySrvPort -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:proxySrvPort")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserproxySrvPort)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:proxySrvAccountName
if($browserproxySrvAccountName -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:proxySrvAccountName")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserproxySrvAccountName)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#browser:proxySrvPassword
if($browserproxySrvPassword -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:proxySrvPassword")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserproxySrvPassword)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


#browser:proxyPasswordEncoding
if($browserproxyPasswordEncoding -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:proxyPasswordEncoding")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserproxyPasswordEncoding)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:proxyException
if($browserproxyException -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:proxyException")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserproxyException)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:isURLBarDisplay
if($browserisURLBarDisplay -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:isURLBarDisplay")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserisURLBarDisplay)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}

#browser:isPanelLocked
if($browserisPanelLocked -ne $NULL){
	$a = $add.CreateElement("item")
	$a.set_InnerText("")
	$b = $add.CreateElement("propName")
	$b.set_InnerText("browser:isPanelLocked")
	$o = $a.AppendChild($b)
	$c = $add.CreateElement("propVal")
	$c.set_InnerText($browserisPanelLocked)
	$o = $a.AppendChild($c)
	$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}


$add.Envelope.Body.putObjects.propListList.arrayType = "itt:string[]["+$add.Envelope.Body.putObjects.propListList.item.item.count+"]"

$add.Save([console]::Out)

[xml]$xml = Invoke-WebRequest $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjects"} -Body $add
Disconnect-MFP -Hostname $Hostname -SessionID $SessionID -Protocol $Protocol 
}

function Remove-MFPAB
{
param(
    $Hostname,
    $Authentication="BASIC",
    $Username="admin",
    $Password,
    $SecurePassword,
    $ID, 
    $SessionID)
#Connect-MFP $Hostname $Authentication $Username $Password $SecurePassword

#$url = "http://$Hostname/DH/udirectory"
$url = $Protocol + "://" + $Hostname + "/DH/udirectory"
Write-Host "Remove-MFPAB $url"
$remove = [xml]@'
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
 <s:Body>
  <m:deleteObjects xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
   <sessionId></sessionId>
  <objectIdList xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.w3.org/2001/XMLSchema" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="">
  </objectIdList>
   <options xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="itt:property[1]">
    <item>
     <propName></propName>
     <propVal></propVal>
    </item>
   </options>
  </m:deleteObjects>
 </s:Body>
</s:Envelope>
'@
$remove.Envelope.Body.deleteObjects.sessionId = $SessionID
$ID | ForEach-Object{
$x = $remove.CreateElement("item")
$x.set_InnerText("entry:$_")
$o = $remove.Envelope.Body.deleteObjects.objectIdList.AppendChild($x)
}
$remove.Envelope.Body.deleteObjects.objectIdList.arrayType = "itt:string["+$remove.Envelope.Body.deleteObjects.objectIdList.item.count+"]"
[xml]$xml = Invoke-WebRequest $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#deleteObjects"} -Body $remove
Disconnect-MFP $Hostname $Protocol
}

function Disconnect-MFP
{
param(
    [string]$Hostname, 
    [string]$SessionID, 
    [string]$Protocol )
Write-host "Disconnect-MFP"
#$url = "http://$Hostname/DH/udirectory"
$url = $Protocol + "://" + $Hostname + "/DH/udirectory"
Write-host "Disconnect-MFP $url"
$logout = [xml]@'
<?xml version="1.0" encoding="utf-8" ?>
 <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
  <s:Body>
   <m:terminateSession xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
    <sessionId></sessionId>
   </m:terminateSession>
  </s:Body>
 </s:Envelope>
'@
$logout.Envelope.Body.terminateSession.sessionId = $SessionID
[xml]$xml = Invoke-WebRequest $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#terminateSession"} -Body $logout
}

Add-Type @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            ServicePointManager.ServerCertificateValidationCallback += 
                delegate
                (
                    Object obj, 
                    X509Certificate certificate, 
                    X509Chain chain, 
                    SslPolicyErrors errors
                )
                {
                    return true;
                };
        }
    }
"@
 
[ServerCertificateValidationCallback]::Ignore();


Function Get-HostPTR{
# Parameter help description
Param(
    [Parameter(Mandatory=$true)]
    [string]$IPaddress
)

try{
    $Hostname =  [System.Net.Dns]::GetHostEntry("10.128.10.245").HostName
    Return $Hostname 
}Catch{

}
}

Function Test-TCPConnection {
    Param(
        [Parameter(Mandatory=$True)]
        [string]$Hostname,
        [Parameter(Mandatory=$True)]
        [string]$Port
    )

$TCPConnectionObject = New-Object Net.Sockets.TcpClient
 
Try{
    $TCPConnectionObject.Connect($Hostname,$Port)
    $ReturnValue = $True
    Return $ReturnValue
}Catch{
    $ReturnValue = $False
    Return $ReturnValue
}Finally{
    $TCPConnectionObject.Dispose()  
}
}#Function

Function Test-PrinterAvailability {
Param(
    [Parameter(Mandatory=$True)]
    [string]$Hostname
)
$PortList = @()
If (Test-Connection -ComputerName "$Hostname" -Quiet -Count 1){
$OpenPort80 = Test-TCPConnection -Hostname $Hostname -Port 80
if ($OpenPort80 -eq $true){
    $PortList += "HTTP"
}

$OpenPort443 = Test-TCPConnection -Hostname $Hostname -Port 443
if ($OpenPort443 -eq $true){
    $PortList += "HTTPS"
}
    $ReturnValue = $PortList
    Return $ReturnValue    
}else{
    $ReturnValue = "False"
    Return $ReturnValue
}
}


#Connect-MFP -Hostname isstffc1 -Protocol "HTTP" -Username "admin" -Password "" $Authentication "BASIC" $SecurePassword ""

$AddressBookDir = "c:\temp\"

$PrinterList =  Import-Csv -Path "C:\TEMP\13-02-2018 Device Population.csv"

foreach ($Printer in $PrinterList){

    Write-Host $Printer.IPAddress $Printer.ModelName

    $Hostname = $Printer.IPAddress

    $PrinterAvalible = Test-PrinterAvailability -Hostname $Hostname
    if ($PrinterAvalible -eq $false){
        Write-host "Printer is offline"
    }else{
        $SessionID = ""
        Foreach ($Port in $PrinterAvalible){
            If ($port.Contains("HTTPS")){

                $Hostname = $Printer.IPAddress
                $Authentication="BASIC"
                $Username="admin"
                $Password=""
                $SecurePassword=""
                $SessionID = Connect-MFP -Hostname $Hostname -Authentication $Authentication -Username $Username -Password $Password -SecurePassword $SecurePassword -Protocol "HTTPS"
                if ($SessionID -eq $false){
                    Write-host "Can't Authenticate"
                }else{
                    Write-host "SessionID = $SessionID"
                    $AddressBook = Get-MFPAB -Hostname $Hostname -Authentication $Authentication -Username $Username -Password $Password -SecurePassword $SecurePassword -SessionID $SessionID -Protocol "HTTPS"
                    If ($AddressBook -like "error") {
                        Write-Host "something broke"    
                    }else{
                        $AddressBook | Format-Table -AutoSize  
                        $output = "$AddressBookDir$Hostname.csv"
                        $AddressBook | Export-Csv -Path $output
                    }
                }   

             }elseif ($port.Contains("HTTP")) {
                
                $Hostname = $Printer.IPAddress
                $Authentication="BASIC"
                $Username="admin"
                $Password=""
                $SecurePassword=""
                $SessionID = Connect-MFP -Hostname $Hostname -Authentication $Authentication -Username $Username -Password $Password -SecurePassword $SecurePassword -Protocol "HTTP"
                if ($SessionID -eq $false){
                    Write-host "Can't Authenticate"
                }else{
                    Write-host "SessionID = $SessionID"
                    $AddressBook = Get-MFPAB -Hostname $Hostname -Authentication $Authentication -Username $Username -Password $Password -SecurePassword $SecurePassword -SessionID $SessionID -Protocol "HTTP"
                    If ($AddressBook -like "error") {
                         Write-Host "something broke"    
                    }else{
                                          
                        $AddressBook | Format-Table -AutoSize  
                        $output = "$AddressBookDir$Hostname.csv"
                        $AddressBook | Export-Csv -Path $output
                    }
                    } #if else  
             } 

           
    }#Foreach
            $SessionID =""
        }
}#>

#$Hostname = "10.128.22.239"
#$Authentication="BASIC"
#$Username="admin"
#$Password=""
#$SecurePassword=""
#$SessionID = Connect-MFP -Hostname $Hostname -Authentication $Authentication -Username $Username -Password $Password -SecurePassword $SecurePassword -Protocol "HTTP"

#Add-MFPAB -Hostname $Hostname -Password "false" -Protocol "HTTP" -SessionID $SessionID -entryType "user" -name "Test23" -longName "More test23" -isDestination "true" -isSender "true" -Mail "True" -Smtp "false" -UserCode "9924" -MailAddress "gregory.machin@waikatodhb.health.nz" -AuthName "9999"

#Add-MFPAB  -hostname $Hostname -Protocol "HTTP" -SessionID $SessionID  -Name "Test14" -LongName "Test Forteen"  -isDestination "true" -isSender "false" -Mail "true" -MailAddress "gregory.machin@waikatodhb.health.nz" 
#-UserCode "9988"
#$SessionID =""


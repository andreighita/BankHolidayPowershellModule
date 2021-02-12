Function Get-DesktopOAuthCredential
{
    <#
    .SYNOPSIS
    Acquires an OAuth authentication token using MSAL.Net

    .Description

    #>

[CmdletBinding()]
param(
[string]$ClientId = "f570b6de-4b04-46d0-a141-7367145db206",
[string]$TenantId,
[string]$RedirectUri = "https://localhost",
[string]$MSALNetDllPath
);

        if ($MSALNetDllPath -eq $null -or $MSALNetDllPath -eq [string]::Empty)
        {
            Import-MSALNet
        }
        else
        {
            # Using Microsoft.Identity.Client 4.22.0
            Write-Verbose "Loading the Microsoft.Identity.Client.dll assembly from location $MSALNetDllPath"
            [Void] [Reflection.Assembly]::LoadFile($MSALNetDllPath)
        }
        # Configure the MSAL client to get tokens
        Write-Verbose "Creating PublicClientApplicationOptions object with specified ClientId, TenantId, and RedirectUri"
        $pcaOptions = New-Object Microsoft.Identity.Client.PublicClientApplicationOptions
        $pcaOptions.ClientId = $ClientId
        $pcaOptions.TenantId = $TenantId
        $pcaOptions.RedirectUri = $RedirectUri
        Write-Debug ($pcaOptions | Out-String)
        #if ($PSBoundParameters['Debug']) { $pcaOptions }
        Write-Verbose "Creating PublicClientApplication object"
        $pca = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::CreateWithApplicationOptions($pcaOptions).Build()
        Write-Debug ($pca | Out-String)

        # The permission scope required for EWS access
        [System.Collections.Generic.IEnumerable`1[System.String]]$ewsScopes = [string[]]@("https://outlook.office365.com/EWS.AccessAsUser.All")
    
        Write-Verbose "Acquiring token"
        $task = $pca.AcquireTokenInteractive([System.Collections.Generic.IEnumerable[string]]$ewsScopes).ExecuteAsync().GetAwaiter().GetResult()
        Write-Debug ($task | Out-String)

        return $task.AccessToken
}


function TrustAllCertificates()
{
    <#
    .SYNOPSIS
    Set certificate trust policy to trust self-signed certificates (for test servers).
    #>

    ## Code From http://poshcode.org/624
    ## Create a compilation environment
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
    $Compiler=$Provider.CreateCompiler()
    $Params=New-Object System.CodeDom.Compiler.CompilerParameters
    $Params.GenerateExecutable=$False
    $Params.GenerateInMemory=$True
    $Params.IncludeDebugInformation=$False
    $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

    $TASource=@'
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
        public class TrustAll : System.Net.ICertificatePolicy {
            public TrustAll() {}
            public bool CheckValidationResult(System.Net.ServicePoint sp,
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                                                System.Net.WebRequest req, int problem)
            { return true; }
        }
        }
'@
    $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
    $TAAssembly=$TAResults.CompiledAssembly

    ## We now create an instance of the TrustAll and attach it to the ServicePointManager
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
    ## END Code From http://poshcode.org/624
    write-verbose "Trusting all certificates!"
}

Function Import-EwsManagedAPI
{
param (
[string]$Path
);

    if ($Path -eq [String]::Empty -or $Path -eq $null)
    {
        $Path = $null
        foreach ($path in ($env:PSModulePath -split ";"))
        {
            $results = Get-ChildItem -Path $path -Recurse | where { $_.FullName -like "*BankHoliday\Assembly\Microsoft.Exchange.WebServices.dll" }
            if ($results.count -gt 0)
            {
                $Path = $results[0].FullName
                break;
            }
        }
        if ($Path -eq $null)
        {
            $Path = (Get-Item "C:\Program Files\Microsoft\Exchange\Web Services\" -ErrorAction SilentlyContinue).FullName
        
            if ($Path -eq $null)
            {
                $Path = (Get-Item "C:\Program Files (x86)\Microsoft\Exchange\Web Services\" -ErrorAction SilentlyContinue).FullName
            }
    
            if ($Path -ne $null)
            {
                $Path = (Get-childItem -Path $Path | where {$_.Attributes -like "*Directory*"} | Sort-Object -Descending)[0].FullName
                $Path = (Get-ChildItem -Path "$($Path)\Microsoft.Exchange.WebServices.dll" -ErrorAction SilentlyContinue)[0].FullName
            }

            if ($Path -eq $null)
            {
                throw "Unable to locate the Microsoft.Exchange.WebServices.dll. Please install the EWS Managed API (https://www.microsoft.com/en-ie/download/details.aspx?id=42951) or specify a correct path using the -Path parameter"
            }
        }

    }
    try
    {
    # Loading ManagedAPI dll
    [Void] [Reflection.Assembly]::LoadFile($Path)
    }
    catch
    {
        write-error "Unable to load the EWS Managed API assembly. Please specify a correct path using the -Path parameter"
        Write-error "Exception details: $($_.Exception.Message)"
        throw "Unable to load EWS Managed API";
        return;
    }
}

Function Import-MSALNet
{
param (
[string]$Path
);

    if ($Path -eq [String]::Empty -or $Path -eq $null)
    {
        $Path = $null
        foreach ($path in ($env:PSModulePath -split ";"))
        {
            $results = Get-ChildItem -Path $path -Recurse | where { $_.FullName -like "*BankHoliday\Assembly\Microsoft.Identity.Client.dll" }
            if ($results.count -gt 0)
            {
                $Path = $results[0].FullName
                break;
            }
        }
    }
    try
    {
    # Loading MSALNet dll
    [Void] [Reflection.Assembly]::LoadFile($Path)
    }
    catch
    {
        write-error "Unable to load the MSAL.Net assembly. Please specify a correct path using the -Path parameter"
        Write-error "Exception details: $($_.Exception.Message)"
        throw "Unable to load MSAL.Net";
    }
}

function Get-HolFolderPath
{
    $FolderPath = $null
    foreach ($path in ($env:PSModulePath -split ";"))
    {
        $results = Get-ChildItem -Path $path -Recurse | where { $_.FullName -like "*BankHoliday\HOL" -and $_.Attributes -eq "Directory" }
        if ($results.count -gt 0)
        {
            $FolderPath = $results[0].FullName
        }
    }

    return $FolderPath 
}

Function New-ExchangeService
{
param(
    $ExchangeVersion = "Exchange2013_SP1",
    [Parameter()]
    [ValidateSet("Dateline Standard Time","UTC-11","Aleutian Standard Time","Hawaiian Standard Time","Marquesas Standard Time","Alaskan Standard Time","UTC-09","Pacific Standard Time (Mexico)","UTC-08","Pacific Standard Time","US Mountain Standard Time","Mountain Standard Time (Mexico)","Mountain Standard Time","Yukon Standard Time","Central America Standard Time","Central Standard Time","Easter Island Standard Time","Central Standard Time (Mexico)","Canada Central Standard Time","SA Pacific Standard Time","Eastern Standard Time (Mexico)","Eastern Standard Time","Haiti Standard Time","Cuba Standard Time","US Eastern Standard Time","Turks And Caicos Standard Time","Paraguay Standard Time","Atlantic Standard Time","Venezuela Standard Time","Central Brazilian Standard Time","SA Western Standard Time","Pacific SA Standard Time","Newfoundland Standard Time","Tocantins Standard Time","E. South America Standard Time","SA Eastern Standard Time","Argentina Standard Time","Greenland Standard Time","Montevideo Standard Time","Magallanes Standard Time","Saint Pierre Standard Time","Bahia Standard Time","UTC-02","Mid-Atlantic Standard Time","Azores Standard Time","Cape Verde Standard Time","UTC","GMT Standard Time","Greenwich Standard Time","Sao Tome Standard Time","Morocco Standard Time","W. Europe Standard Time","Central Europe Standard Time","Romance Standard Time","Central European Standard Time","W. Central Africa Standard Time","Jordan Standard Time","GTB Standard Time","Middle East Standard Time","Egypt Standard Time","E. Europe Standard Time","Syria Standard Time","West Bank Standard Time","South Africa Standard Time","FLE Standard Time","Israel Standard Time","Kaliningrad Standard Time","Sudan Standard Time","Libya Standard Time","Namibia Standard Time","Arabic Standard Time","Turkey Standard Time","Arab Standard Time","Belarus Standard Time","Russian Standard Time","E. Africa Standard Time","Iran Standard Time","Arabian Standard Time","Astrakhan Standard Time","Azerbaijan Standard Time","Russia Time Zone 3","Mauritius Standard Time","Saratov Standard Time","Georgian Standard Time","Volgograd Standard Time","Caucasus Standard Time","Afghanistan Standard Time","West Asia Standard Time","Ekaterinburg Standard Time","Pakistan Standard Time","Qyzylorda Standard Time","India Standard Time","Sri Lanka Standard Time","Nepal Standard Time","Central Asia Standard Time","Bangladesh Standard Time","Omsk Standard Time","Myanmar Standard Time","SE Asia Standard Time","Altai Standard Time","W. Mongolia Standard Time","North Asia Standard Time","N. Central Asia Standard Time","Tomsk Standard Time","China Standard Time","North Asia East Standard Time","Singapore Standard Time","W. Australia Standard Time","Taipei Standard Time","Ulaanbaatar Standard Time","Aus Central W. Standard Time","Transbaikal Standard Time","Tokyo Standard Time","North Korea Standard Time","Korea Standard Time","Yakutsk Standard Time","Cen. Australia Standard Time","AUS Central Standard Time","E. Australia Standard Time","AUS Eastern Standard Time","West Pacific Standard Time","Tasmania Standard Time","Vladivostok Standard Time","Lord Howe Standard Time","Bougainville Standard Time","Russia Time Zone 10","Magadan Standard Time","Norfolk Standard Time","Sakhalin Standard Time","Central Pacific Standard Time","Russia Time Zone 11","New Zealand Standard Time","UTC+12","Fiji Standard Time","Kamchatka Standard Time","Chatham Islands Standard Time","UTC+13","Tonga Standard Time","Samoa Standard Time","Line Islands Standard Time")]
    $TimeZoneName,
    [Parameter(ParameterSetName = "DefaultCredentials")]
    [switch]$UseDefaultCredentials,
    [Parameter(ParameterSetName = "Credentials")]
    [System.Management.Automation.PSCredential]$Credentials,
    [Parameter(ParameterSetName = "ModernAuthentication")]
    [string]$OAuthCredentials,
    [Parameter(ParameterSetName = "ModernAuthentication")]
    [switch]$UseModernAuthentication,
    [Parameter()]    
    [string]$EwsUrl,
    [Parameter(ParameterSetName = "ModernAuthentication")]
    [string]$ClientId,
    [Parameter(ParameterSetName = "ModernAuthentication")]
    [string]$TenantId,
    [Parameter(ParameterSetName = "ModernAuthentication")]
    [string]$RedirectUri,
    [string]$EWSManagedAPIPath,
    [string]$MSALNetDllPath
    
);
    
    if (([System.AppDomain]::CurrentDomain.GetAssemblies() | where { $_.FullName -like "*Microsoft.Exchange.WebServices*" }).Count -eq 0) { Import-EwsManagedAPI -Path $EWSManagedAPIPath }
    # Setting up the service
    [Microsoft.Exchange.WebServices.Data.ExchangeService] $Service = $null;
    switch ($ExchangeVersion)
    {
        "Exchange2007_SP1" 
        {
            if ($TimeZoneName -ne $null -and $TimeZoneName -ne [string]::Empty)
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1, [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneName));
            }
            else
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1);
            }
        }
        "Exchange2010" 
        {
            if ($TimeZoneName -ne $null -and $TimeZoneName -ne [string]::Empty)
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010, [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneName));
            }
            else
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010);
            }
        }
        "Exchange2010_SP1" 
        {
            if ($TimeZoneName -ne $null -and $TimeZoneName -ne [string]::Empty)
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1, [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneName));
            }
            else
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1);
            }
        }
        "Exchange2010_SP2" 
        {
            if ($TimeZoneName -ne $null -and $TimeZoneName -ne [string]::Empty)
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2, [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneName));
            }
            else
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2);
            }
        }
        "Exchange2013" 
        {
            if ($TimeZoneName -ne $null -and $TimeZoneName -ne [string]::Empty)
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013, [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneName));
            }
            else
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013);
            }
        }
        "Exchange2013_SP1" 
        {
            if ($TimeZoneName -ne $null -and $TimeZoneName -ne [string]::Empty)
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1, [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneName));
            }
            else
            {
                $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1);
            }
        }
    }

    # Setting up credentials
    if ($UseModernAuthentication.IsPresent)
    {
        if ($OAuthCredentials -eq $null -or $OAuthCredentials -eq [string]::Empty)
        {
            $OAuthCredentials = Get-DesktopOAuthCredential -ClientId $ClientId -TenantId $TenantId -RedirectUri $RedirectUri -MSALNetDllPath $MSALNetDllPath
        }        
        $Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($OAuthCredentials);
    }
    else 
    {
        if ($UseDefaultCredentials.IsPresent)
        {
            $Service.UseDefaultCredentials = $true;
        }
        else
        {  
            if ($Credentials -eq $null)
            {
                $Credentials = Get-Credential;
                $Service.Credentials = New-Object System.Net.NetworkCredential($credentials.UserName, $credentials.Password);
            }
            else
            {
                $Service.Credentials = $Credentials.GetNetworkCredential()
            }
            write-host "Connecting as: " $credentials.UserName.ToString() -ForegroundColor Yellow
        }
    }


    return $service
}

Function Get-EwsCalendarFolder
{
param(
    [Parameter()]
    [Microsoft.Exchange.WebServices.Data.ExchangeService]$ExchangeService,
    [Parameter()]
    [bool]$Impersonate,
    [Parameter()]
    [string]$EmailAddress,
    [Parameter(ParameterSetName="Autodiscover")]
    [bool]$UseAutodiscover,
    [Parameter(ParameterSetName="ServiceUri")]
    [string]$EwsUrl,
    [string]$CalendarFolderId = $null
);


    $ServiceUrl = $EwsUrl
    if ($UseAutodiscover)
    {    
        $ServiceUrl = Get-EwsUrl -EmailAddress $EmailAddress
        if ($ServiceUrl -ne $null)
        {
            $ExchangeService.Url = new-Object Uri($ServiceUrl)
        }
        else
        {
            $ExchangeService.AutodiscoverUrl($EmailAddress)
        }
    }
    else
    {
        if ($ServiceUrl -ne $null -and $ServiceUrl -ne [string]::Empty)
        {
            $ExchangeService.Url = new-Object Uri($ServiceUrl)
        }
        else
        {
            return $null
        }

    }

    # If we are using impersonation then setup impersonation
    if ($Impersonate)
    {
        $ExchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
    }

    try
    {
        # Binding to the Calendar folder
        if ($CalendarFolderId -eq $null -or $CalendarFolderId -eq [string]::Empty)
        {
            $CalendarFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService, (new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $EmailAddress)))
            return $CalendarFolder
        }
        $CalendarFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService, (new-object Microsoft.Exchange.WebServices.Data.FolderId$CalendarFolderId))
        return $CalendarFolder
    }
    catch
    {
        Write-Host "Unable to bind to Calendar folder. Please check your credentials and try again.";
        Write-host "Exception details: " $_.Exception.Message -ForegroundColor Red
        return;
    }
}

function Get-EwsUrl{
param(
[ValidateNotNullOrEmpty()]
[string]$EmailAddress
)
    $AutodiscoverV2 = $null;

    try
    {
        $AutodiscoverV2 = Invoke-RestMethod -Method GET -Uri ("https://outlook.office365.com/autodiscover/autodiscover.json?Email={0}&Protocol=EWS" -f $EmailAddress) -ErrorAction Continue
        if ($AutodiscoverV2 -ne $null)
        {
            if($AutodiscoverV2.Url -ne $null)
            {
                return $AutodiscoverV2.Url
            }
        }
    }
    catch { }
}


class HolidayEntry {
[string]$EmailAddress
[string]$CalendarFolderId
[string]$ItemId
[string]$CountryName
[string]$Date
[string]$HolidayName
[string]$BusyStatus
}

class Holiday {
[string]$CountryName
[string]$Date
[string]$HolidayName
[string]$Calendar
[string]$BusyStatus
}


function Get-BankHoliday
{
<#
.SYNOPSIS
Retrieves information about bank holidays in a given mailbox or in multipe mailboxes. 

.DESCRIPTION
Uses Exchange Web Services to connect to one or multiple mailboxes and retrieves information about bank holidays in a given mailbox or in multipe mailboxes. The script requires the Ews Managed API assembly be referenced, as well as the MSAL.Net assembly.
The function allows the automatic discovery of service endpoints, as well as the use of modern authentication to connect to Exchange Online hosted mailboxes or on premises hosted mailboxes. 

.PARAMETER EmailAddress
Single email address, comma separated list of mailbox email addresses, or an array of strings containing the email addresses to connect to. See the examples section for examples.

.PARAMETER Impersonate
Use this switch parameter to have the authenticated account use Exchange Impersonation when connecting to mailboxes.

.PARAMETER UseAutodiscover
Use this switch parameter use automatic discovery of the Exchange Web Services Service Url using Autodiscover V1 and V2.

.PARAMETER UseModernAuthentication
Use this switch parameter use Modern Authentication (OAuth credentials) to authenticate when connecting to Exchange mailboxes. Use this switch parameter for Exchange Online or Hybrid Exchange Server deployments.

.PARAMETER ClientId
The Application (client) ID value of the Azure Portal app registration that exposes the EWS.AccessAsUser.All permission. 

.PARAMETER TenantId
The unique GUID tenant identifier or the fully qualified domain name. For example "105a90c6-224d-4609-9fbd-af1f927a554f" or "contoso.com".

.PARAMETER RedirectUri
The redirect URI defined in the Azure Portal app registration. 

.PARAMETER EWSManagedAPIPath
The path to the Microsoft.Exchange.WebServices.dll file. 
Note: If you've imported all module files, this parameter is not necessary as the script will try and locate the required assemblies automatically in the well known module locations.

.PARAMETER MSALNetDllPath
The path to the Microsoft.Identity.Client.dll file. This is required if -UseModernAuthentication is used. 
Note: If you've imported all module files, this parameter is not necessary as the script will try and locate the required assemblies automatically in the well known module locations.

.PARAMETER Culture
The culture name. This is used to search bank holidays using the country names corresponding to specific cultures. 
Each culture contains the country names and bank holiday names for that locale. The default culture is "English (United Kingdom)"
For example, English (United States) will contain all the country and bank holiday names in Enlish. 
German (Germany) will contain all the names in German.
The following are the possible culture values: 
"Arabic (Saudi Arabia)","Bulgarian (Bulgaria)","Chinese (Simplified, PRC)","Chinese (Traditional, Taiwan)","Croatian (Croatia)","Czech (Czech Republic)","Danish (Denmark)","Dutch (Netherlands)","English (United States)","Estonian (Estonia)","Finnish (Finland)","French (France)","German (Germany)","Greek (Greece)","Hebrew (Israel)","Hindi (India)","Hungarian (Hungary)","Indonesian (Indonesia)","Italian (Italy)","Japanese (Japan)","Kazakh (Kazakhstan)","Korean (Korea)","Latvian (Latvia)","Lithuanian (Lithuania)","Malay (Malaysia)","Norwegian, Bokmål (Norway)","Polish (Poland)","Portuguese (Brazil)","Portuguese (Portugal)","Romanian (Romania)","Russian (Russia)","Serbian (Latin, Serbia and Montenegro (Former))","Serbian (Latin, Serbia)","Slovak (Slovakia)","Slovenian (Slovenia)","Spanish (Spain)","Swedish (Sweden)","Thai (Thailand)","Turkish (Turkey)","Ukrainian (Ukraine)","Vietnamese (Vietnam)"

.PARAMETER CountryName
The name of the country the bank holiday entries correspond to, to be returned by the search.  
To get the correct country name to use, please use the Get-CountryNames function.
To specify a custom country name, please specify the custom value for the CountryName parameter and specify the -CountryNameOverride switch.
	
.PARAMETER CountryNameOverride
Use this switch to be able to specify a non-standard CountryName value. Otherwise, the country name value specified will be validated against the list of countries for the specified culture.

.PARAMETER Year
A specific year to retrieve bank holidays for.

.PARAMETER StartDate
Use this parameter to specify the date after which to look for bank holidays in the selected mailbox(es). Cannot be used if -Year is present. Can be used together with $EndDate.
Note: The value needs to be a string in the following format: "yyyy/MM/dd" for example "2021/05/12" for May 12, 2021.

.PARAMETER EndDate
Use this parameter to specify the date before which to look for bank holidays in the selected mailbox(es). Cannot be used if -Year is present. Can be used together with $StartDate.
Note: The value needs to be a string in the following format: "yyyy/MM/dd" for example "2021/05/12" for May 12, 2021.

.PARAMETER Subject
Use this parameter to specify Subject of the bank holiday entries to retrieve for each mailbox specified.

.PARAMETER ListAll
Use this switch to look up bank holiday entries for all countries corresponding to a specific culture. Cannot be used if -CountryName is present.

.PARAMETER Summary
Displayes a table view of the resulting set of appointments returned by the search.

.INPUTS

None. You cannot pipe objects to Get-BankHoliday.

.OUTPUTS

A collection of HolidayEntry objects.
The defining class signature is the following:
class HolidayEntry {
[string]$EmailAddress
[string]$CalendarFolderId
[string]$ItemId
[string]$CountryName
[string]$Date
[string]$HolidayName
[string]$BusyStatus
}

.EXAMPLE
Retreieve all bank holiday entries for the United Kingdom ocurring in 2021.
Get-BankHoliday -EmailAddress user1@contoso.com -Impersonate -UseAutodiscover -UseModernAuthentication -TenantId contoso.com -Year 2021 -CountryName "United Kingdom" -Summary
Holiday entries for mailbox user1@contoso.com:
17 entries returned


EmailAddress      CountryName    Date                HolidayName                                    BusyStatus
------------      -----------    ----                -----------                                    ----------
user1@contoso.com United Kingdom 08/02/2021 00:00:00 August Bank Holiday (Scotland)                 OOF
user1@contoso.com United Kingdom 11/30/2021 00:00:00 St. Andrew's Day (Scotland)                    OOF
user1@contoso.com United Kingdom 01/04/2021 00:00:00 New Year's Day (2nd Day) (Scotland) (Observed) OOF
user1@contoso.com United Kingdom 01/02/2021 00:00:00 New Year's Day (2nd Day) (Scotland)            OOF
user1@contoso.com United Kingdom 12/27/2021 00:00:00 Christmas Bank Holiday                         OOF
user1@contoso.com United Kingdom 12/28/2021 00:00:00 Boxing Day Bank Holiday                        OOF
user1@contoso.com United Kingdom 03/17/2021 00:00:00 St. Patrick's Day (N. Ireland)                 OOF
user1@contoso.com United Kingdom 05/31/2021 00:00:00 Spring Bank Holiday                            OOF
user1@contoso.com United Kingdom 01/01/2021 00:00:00 New Year's Day                                 OOF
user1@contoso.com United Kingdom 05/03/2021 00:00:00 May Day Bank Holiday                           OOF
user1@contoso.com United Kingdom 08/30/2021 00:00:00 Late Summer Holiday                            OOF
user1@contoso.com United Kingdom 04/02/2021 00:00:00 Good Friday                                    OOF
user1@contoso.com United Kingdom 04/05/2021 00:00:00 Easter Monday                                  OOF
user1@contoso.com United Kingdom 04/04/2021 00:00:00 Easter Day                                     OOF
user1@contoso.com United Kingdom 12/25/2021 00:00:00 Christmas Day                                  OOF
user1@contoso.com United Kingdom 12/26/2021 00:00:00 Boxing Day                                     OOF
user1@contoso.com United Kingdom 07/12/2021 00:00:00 Battle of the Boyne (N. Ireland)               OOF

.EXAMPLE
Retrieve all bank holidays in year 2021 for all countries corresponding to the "English (United States)" culture for two mailboxes:
Get-BankHoliday -EmailAddress user1@contoso.com, user2@contoso.com -Impersonate -UseAutodiscover -UseModernAuthentication -TenantId contoso.com -Year 2021 -Summary -ListAll
Holiday entries for mailbox user1@contoso.com:

15 entries returned for country France
17 entries returned for country United Kingdom

Holiday entries for mailbox user2@contoso.com:
17 entries returned for country United Kingdom

EmailAddress      CountryName    Date                HolidayName                                    BusyStatus
------------      -----------    ----                -----------                                    ----------
user1@contoso.com France         12/24/2021 00:00:00 Christmas Eve                                  OOF
user1@contoso.com France         05/23/2021 00:00:00 Pentecost                                      OOF
user1@contoso.com France         04/04/2021 00:00:00 Easter Sunday                                  OOF
user1@contoso.com France         04/02/2021 00:00:00 Good Friday                                    OOF
user1@contoso.com France         07/14/2021 00:00:00 Bastille Day                                   OOF
user1@contoso.com France         05/01/2021 00:00:00 Labor Day                                      OOF
user1@contoso.com France         05/24/2021 00:00:00 Pentecost Monday                               OOF
user1@contoso.com France         05/08/2021 00:00:00 Victory Day                                    OOF
user1@contoso.com France         01/01/2021 00:00:00 New Year's Day                                 OOF
user1@contoso.com France         04/05/2021 00:00:00 Easter Monday                                  OOF
user1@contoso.com France         12/25/2021 00:00:00 Christmas Day                                  OOF
user1@contoso.com France         08/15/2021 00:00:00 Assumption                                     OOF
user1@contoso.com France         05/13/2021 00:00:00 Ascension                                      OOF
user1@contoso.com France         11/11/2021 00:00:00 Armistice Day 1918                             OOF
user1@contoso.com France         11/01/2021 00:00:00 All Saints' Day                                OOF
user1@contoso.com United Kingdom 08/02/2021 00:00:00 August Bank Holiday (Scotland)                 OOF
user1@contoso.com United Kingdom 11/30/2021 00:00:00 St. Andrew's Day (Scotland)                    OOF
user1@contoso.com United Kingdom 01/04/2021 00:00:00 New Year's Day (2nd Day) (Scotland) (Observed) OOF
user1@contoso.com United Kingdom 01/02/2021 00:00:00 New Year's Day (2nd Day) (Scotland)            OOF
user1@contoso.com United Kingdom 12/27/2021 00:00:00 Christmas Bank Holiday                         OOF
user1@contoso.com United Kingdom 12/28/2021 00:00:00 Boxing Day Bank Holiday                        OOF
user1@contoso.com United Kingdom 03/17/2021 00:00:00 St. Patrick's Day (N. Ireland)                 OOF
user1@contoso.com United Kingdom 05/31/2021 00:00:00 Spring Bank Holiday                            OOF
user1@contoso.com United Kingdom 01/01/2021 00:00:00 New Year's Day                                 OOF
user1@contoso.com United Kingdom 05/03/2021 00:00:00 May Day Bank Holiday                           OOF
user1@contoso.com United Kingdom 08/30/2021 00:00:00 Late Summer Holiday                            OOF
user1@contoso.com United Kingdom 04/02/2021 00:00:00 Good Friday                                    OOF
user1@contoso.com United Kingdom 04/05/2021 00:00:00 Easter Monday                                  OOF
user1@contoso.com United Kingdom 04/04/2021 00:00:00 Easter Day                                     OOF
user1@contoso.com United Kingdom 12/25/2021 00:00:00 Christmas Day                                  OOF
user1@contoso.com United Kingdom 12/26/2021 00:00:00 Boxing Day                                     OOF
user1@contoso.com United Kingdom 07/12/2021 00:00:00 Battle of the Boyne (N. Ireland)               OOF
user2@contoso.com United Kingdom 08/02/2021 00:00:00 August Bank Holiday (Scotland)                 Busy
user2@contoso.com United Kingdom 11/30/2021 00:00:00 St. Andrew's Day (Scotland)                    Busy
user2@contoso.com United Kingdom 01/04/2021 00:00:00 New Year's Day (2nd Day) (Scotland) (Observed) Busy
user2@contoso.com United Kingdom 01/02/2021 00:00:00 New Year's Day (2nd Day) (Scotland)            Busy
user2@contoso.com United Kingdom 12/27/2021 00:00:00 Christmas Bank Holiday                         Busy
user2@contoso.com United Kingdom 12/28/2021 00:00:00 Boxing Day Bank Holiday                        Busy
user2@contoso.com United Kingdom 05/31/2021 00:00:00 Spring Bank Holiday                            Busy
user2@contoso.com United Kingdom 01/01/2021 00:00:00 New Year's Day                                 Busy
user2@contoso.com United Kingdom 05/03/2021 00:00:00 May Day Bank Holiday                           Busy
user2@contoso.com United Kingdom 08/30/2021 00:00:00 Late Summer Holiday                            Busy
user2@contoso.com United Kingdom 04/02/2021 00:00:00 Good Friday                                    Busy
user2@contoso.com United Kingdom 04/05/2021 00:00:00 Easter Monday                                  Busy
user2@contoso.com United Kingdom 04/04/2021 00:00:00 Easter Day                                     Busy
user2@contoso.com United Kingdom 12/25/2021 00:00:00 Christmas Day                                  Busy
user2@contoso.com United Kingdom 12/26/2021 00:00:00 Boxing Day                                     Busy
user2@contoso.com United Kingdom 07/12/2021 00:00:00 Battle of the Boyne (N. Ireland)               Busy
user2@contoso.com United Kingdom 03/17/2021 00:00:00 St. Patrick's Day (N. Ireland)                 Free


.EXAMPLE

Listing bank holiday entries for Germany between dates 1st January 2021 and 31st December 2023
Get-BankHoliday -EmailAddress user1@contoso.com -Impersonate -UseAutodiscover -UseModernAuthentication -TenantId contoso.com -CountryName "Germany" -StartDate "2021/01/01" -EndDate "2023/12/31" -Summary
Holiday entries for mailbox user1@contoso.com:
18 entries returned for country Germany


EmailAddress      CountryName Date                HolidayName                  BusyStatus
------------      ----------- ----                -----------                  ----------
user1@contoso.com Germany     12/24/2023 00:00:00 Christmas Eve                OOF
user1@contoso.com Germany     05/01/2023 00:00:00 Labor Day                    OOF
user1@contoso.com Germany     05/28/2023 00:00:00 Whit Sunday                  OOF
user1@contoso.com Germany     05/29/2023 00:00:00 Whit Monday                  OOF
user1@contoso.com Germany     10/31/2023 00:00:00 Reformation Day              OOF
user1@contoso.com Germany     01/01/2023 00:00:00 New Year's Day               OOF
user1@contoso.com Germany     04/07/2023 00:00:00 Good Friday                  OOF
user1@contoso.com Germany     01/06/2023 00:00:00 Epiphany                     OOF
user1@contoso.com Germany     04/10/2023 00:00:00 Easter Monday                OOF
user1@contoso.com Germany     04/09/2023 00:00:00 Easter Day                   OOF
user1@contoso.com Germany     11/22/2023 00:00:00 Day of Prayer and Repentance OOF
user1@contoso.com Germany     10/03/2023 00:00:00 Day of German Unity          OOF
user1@contoso.com Germany     06/08/2023 00:00:00 Corpus Christi               OOF
user1@contoso.com Germany     12/25/2023 00:00:00 Christmas Day                OOF
user1@contoso.com Germany     12/26/2023 00:00:00 Christmas (2nd Day)          OOF
user1@contoso.com Germany     08/15/2023 00:00:00 Assumption                   OOF
user1@contoso.com Germany     05/18/2023 00:00:00 Ascension                    OOF
user1@contoso.com Germany     11/01/2023 00:00:00 All Saints' Day              OOF

.LINK

https://github.com/andreighita/EWSHoliday

.NOTES
File name      : BankHoliday.psm1
Author         : Andrei Ghita (catagh@microsoft.com)
Created        : 2021-02-10
Last reviewer  : Andrei Ghita (catagh@microsoft.com)
Last revision  : 2021-02-10
Disclaimer     : THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

#>
[CmdletBinding(DefaultParameterSetName = "Year Search")]
param(
    [Parameter(Mandatory=$true, HelpMessage = "Single email address, comma separated list of mailbox email addresses, or an array of strings containing the email addresses to connect to.")] [ValidateNotNullOrEmpty()]
    [string[]]$EmailAddress,
    
    [Parameter()]
    [switch]$Impersonate,
    
    [Parameter()]
    [switch]$UseAutodiscover,

    [Parameter()]
    [switch]$UseModernAuthentication,

    [Parameter()]
    [string]$ClientId = "f570b6de-4b04-46d0-a141-7367145db206",

    [Parameter()]
    [string]$TenantId,

    [Parameter()]
    [string]$RedirectUri = "https://localhost",

    [Parameter()] [ValidateScript({ return [System.IO.File]::Exists($_) })]
    [string]$EWSManagedAPIPath,

    [Parameter()] [ValidateScript({ return [System.IO.File]::Exists($_) })]
    [string]$MSALNetDllPath,

    [Parameter(ParameterSetName = 'List All')]
    [Parameter(ParameterSetName = 'Year Search')]
    [int]$Year = -1,
    
    [Parameter(ParameterSetName = 'List All')]
    [Parameter(ParameterSetName = 'Year Search')]
    [Parameter(ParameterSetName = 'Range Search')]
    [ValidateSet("Arabic (Saudi Arabia)","Bulgarian (Bulgaria)","Chinese (Simplified, PRC)","Chinese (Traditional, Taiwan)","Croatian (Croatia)","Czech (Czech Republic)","Danish (Denmark)","Dutch (Netherlands)","English (United States)","Estonian (Estonia)","Finnish (Finland)","French (France)","German (Germany)","Greek (Greece)","Hebrew (Israel)","Hindi (India)","Hungarian (Hungary)","Indonesian (Indonesia)","Italian (Italy)","Japanese (Japan)","Kazakh (Kazakhstan)","Korean (Korea)","Latvian (Latvia)","Lithuanian (Lithuania)","Malay (Malaysia)","Norwegian, Bokmål (Norway)","Polish (Poland)","Portuguese (Brazil)","Portuguese (Portugal)","Romanian (Romania)","Russian (Russia)","Serbian (Latin, Serbia and Montenegro (Former))","Serbian (Latin, Serbia)","Slovak (Slovakia)","Slovenian (Slovenia)","Spanish (Spain)","Swedish (Sweden)","Thai (Thailand)","Turkish (Turkey)","Ukrainian (Ukraine)","Vietnamese (Vietnam)")]
    [string]$Culture = 'English (United States)',
    
    
    [Parameter(ParameterSetName = 'Range Search')]
    [Parameter(ParameterSetName = 'Year Search')]
    [string[]]$CountryName,
   
    [Parameter(ParameterSetName = 'Range Search')]
    [Parameter(ParameterSetName = 'Year Search')]
    [switch]$CountryNameOverride,
    
    [Parameter(ParameterSetName = 'List All')]
    [switch]$ListAll,

    [Parameter()]
    [switch]$Summary,

    [Parameter(ParameterSetName = 'List All')]
    [Parameter(ParameterSetName = 'Range Search')]
    [string]$StartDate,
    
    [Parameter(ParameterSetName = 'List All')]
    [Parameter(ParameterSetName = 'Range Search')]
    [string]$EndDate,
    
    [Parameter()]
    [string]$Subject,

    [string]$Keyword
)
    

    $CountryNames = @()
    if ($ListAll.IsPresent)
    {
        $CountryNames+= Get-CountryNames -Culture $Culture
    }
    else
    {
        if ($CountryNameOverride.IsPresent)
        {
            $CountryNames+= $CountryName
        }
        else
        {
            foreach ($countryNameEntry in $CountryName)
            {
                if ((Get-CountryNames -Culture $Culture).Contains($countryNameEntry))
                {
                    $CountryNames  += (Get-CountryNames -Culture $Culture | Where-Object { $_ -like "$countryNameEntry*" })
                }
                else
                {
                    write-error "You must specify a valid country name. Please use one of the values below or to specify a custom value please use the -CountryNameOverride parameter."
                    (Get-CountryNames -Culture $Culture) -join ", "
                    return
                }
            }
        }
    }

    $ExchangeService = $null;
    $ExchangeService = New-ExchangeService -UseModernAuthentication -ClientId $ClientId -TenantId $TenantId -RedirectUri $RedirectUri -MSALNetDllPath $MSALNetDllPath 

    $EwsCalendarFolder = $null

    $EmailPace = 1 / $EmailAddress.count * 100
    $EmailProgress = 0
    $HolidayEntries = @()
    

    foreach ($EmailAddressEntry in $EmailAddress)
    {
        Write-Host "Holiday entries for mailbox $($EmailAddressEntry):"
        Write-Progress -Activity "Processing mailbox $($EmailAddressEntry)" -Status "Running" -PercentComplete $EmailProgress -Id 1

        $EwsUrl = Get-EwsUrl -EmailAddress $EmailAddressEntry
        $EwsCalendarFolder = Get-EwsCalendarFolder -ExchangeService $ExchangeService -EmailAddress $EmailAddressEntry -UseAutodiscover $UseAutodiscover.IsPresent -Impersonate $Impersonate.IsPresent

        if ($EwsCalendarFolder -ne $null)
        {
            # Extended properties
            $PS_PUBLIC_STRINGS_Guid = new-object Guid("{00020329-0000-0000-C000-000000000046}")
            $PSETID_Appointment_Guid = new-object Guid("{00062002-0000-0000-C000-000000000046}")
            $Keywords = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($PS_PUBLIC_STRINGS_Guid, "Keywords", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::StringArray)  
            $PidLidLocation = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($PSETID_Appointment_Guid, 0x8208, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
            [string]$DateTimeFormat = "d";

            
            $ActivityName = $null;
            if ($ListAll.IsPresent)
            {
                $ActivityName = "Listing bank holidays for all countries on mailbox $($mailbox)"
            }
            else
            {
                $ActivityName = "Listing bank holidays for $($countrynames -join ", ") on mailbox $($mailbox)"
            }

            $Pace = 1 / $countrynames.Count * 100
            $Progress = 0
            foreach ($CountryName in $countrynames)
            {
                Write-Progress -Activity $ActivityName -Status "Running" -PercentComplete $progress -Id 2
                $SearchFilterCollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
                if ($year -ne -1)
                {
                    $Start = [System.DateTime]::ParseExact("$Year/01/01", "yyyy/MM/dd", $null)
                    $End = [System.DateTime]::ParseExact("$Year/12/31", "yyyy/MM/dd", $null)

                    $SearchFilter1 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, $Start)
                    $SearchFilter2 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End, $End)
                    $SearchFilterCollection.Add($SearchFilter1)
                    $SearchFilterCollection.Add($SearchFilter2)
                }
                else
                {
                    if($StartDate -ne $null -and $StartDate -ne [string]::Empty )
                    {
                        $SearchFilter3 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, [System.DateTime]::ParseExact($StartDate, "yyyy/MM/dd", $null ))
                        $SearchFilterCollection.Add($SearchFilter3)                   
                    }
                    if($EndDate -ne $null -and $EndDate -ne [string]::Empty )
                    {
                        $SearchFilter4 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End, [System.DateTime]::ParseExact($EndDate, "yyyy/MM/dd", $null ))
                        $SearchFilterCollection.Add($SearchFilter4)  
                    }
                }
                
                if($Subject -ne $null -and $Subject -ne [string]::Empty )
                {
                    $SearchFilter5 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject, $Subject)
                    $SearchFilterCollection.Add($SearchFilter5)                   
                }

                if($CountryName -ne $null -and $CountryName -ne [string]::Empty)
                {
                    $SearchFilter6 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Location, $CountryName)
                    $SearchFilterCollection.Add($SearchFilter6)                   
                }
                
                if($Keyword -ne $null -and $Keyword -ne [string]::Empty)
                {
                    $SearchFilter7 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Exists($Keywords)
                    $SearchFilterCollection.Add($SearchFilter7)           
                    $SearchFilter8 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring($Keywords, $Keyword)
                    $SearchFilterCollection.Add($SearchFilter8)      
                }

                $itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(10000)
                $itemView.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Location, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::LegacyFreeBusyStatus, $Keywords)
        
                $findResults = $ExchangeService.FindItems($EwsCalendarFolder.Id,$SearchFilterCollection,$itemView)
        
                foreach($item in $findResults.Items)
                { 
                    $Item.Load();

                    $HolidayEntry =[HolidayEntry]::new()
                    $HolidayEntry.EmailAddress = $EmailAddressEntry
                    $HolidayEntry.CalendarFolderId = $EwsCalendarFolder.Id
                    $HolidayEntry.ItemId = $Item.Id
                    $HolidayEntry.CountryName = $CountryName
                    $HolidayEntry.Date = $Item.Start 
                    $HolidayEntry.HolidayName = $Item.subject
                    $HolidayEntry.BusyStatus = $item.LegacyFreeBusyStatus
                    $holidayEntries += $HolidayEntry
                     
                }
            
                $Progress = $Progress + $Pace
                
                
                if ($findResults.Items.count -gt 0)
                {
                    write-host "$($findResults.Items.count) entries returned for country $CountryName"
                }
                else
                {
                   # write-host "No entries returned for country $CountryName"
                }
            }
            Write-Progress -Activity $ActivityName -Status "Completed" -PercentComplete 100

        }
        else
        {
            Write-Warning "Unable to connect to calendar folder for mailbox $($EmailAddressEntry)"
        }
        $EmailProgress = $EmailProgress + $EmailPace
        Write-Host "Found $(($HolidayEntries | where { $_.EmailAddress -eq $EmailAddressEntry }).Count) entries"

    }
     Write-Progress -Activity "Mailbox processing" -Status "Completed" -PercentComplete 100 -Id 1
    
    if ($Summary.IsPresent)
    {
        $HolidayEntries | Select-Object EmailAddress, CountryName, Date, HolidayName, BusyStatus | ft
    }
    else
    {
        $HolidayEntries
    }    
    return
}

function Add-BankHoliday
{
<#
.SYNOPSIS
Adds one or multiple bank holiday calendar entries to a given mailbox or to multipe mailboxes. 

.DESCRIPTION
Uses Exchange Web Services to connect to one or multiple mailboxes and creates one or multiple bank holiday entries in each mailbox. The script requires the Ews Managed API assembly be referenced, as well as the MSAL.Net assembly.
The function allows the automatic discovery of service endpoints, as well as the use of modern authentication to connect to Exchange Online hosted mailboxes or on premises hosted mailboxes. 
The module includes bank holiday entries for 41 different locales, based on Outlook's .hol import files. The pre-canned files for each locale can be used to import official bank holiday entries. The function also allows using a custom, user defined, CSV file for import, or simply adding one single bank holiday entry et a time. 

.PARAMETER EmailAddress
Single email address, comma separated list of mailbox email addresses, or an array of strings containing the email addresses to connect to. See the examples section for examples.

.PARAMETER Impersonate
Use this switch parameter to have the authenticated account use Exchange Impersonation when connecting to mailboxes.

.PARAMETER UseAutodiscover
Use this switch parameter use automatic discovery of the Exchange Web Services Service Url using Autodiscover V1 and V2.

.PARAMETER UseModernAuthentication
Use this switch parameter use Modern Authentication (OAuth credentials) to authenticate when connecting to Exchange mailboxes. Use this switch parameter for Exchange Online or Hybrid Exchange Server deployments.

.PARAMETER ClientId
The Application (client) ID value of the Azure Portal app registration that exposes the EWS.AccessAsUser.All permission. 

.PARAMETER TenantId
The unique GUID tenant identifier or the fully qualified domain name. For example "105a90c6-224d-4609-9fbd-af1f927a554f" or "contoso.com".

.PARAMETER RedirectUri
The redirect URI defined in the Azure Portal app registration. 

.PARAMETER EWSManagedAPIPath
The path to the Microsoft.Exchange.WebServices.dll file. 
Note: If you've imported all module files, this parameter is not necessary as the script will try and locate the required assemblies automatically in the well known module locations.

.PARAMETER MSALNetDllPath
The path to the Microsoft.Identity.Client.dll file. This is required if -UseModernAuthentication is used. 
Note: If you've imported all module files, this parameter is not necessary as the script will try and locate the required assemblies automatically in the well known module locations.

.PARAMETER Culture
The culture name for the pre-canned bank holiday csv files. 
Each culture contains the country names and bank holiday names for that locale. 
For example, English (United States) will contain all the country and bank holiday names in Enlish. 
German (Germany) will contain all the names in German.
The following are the possible culture values: 
"Arabic (Saudi Arabia)","Bulgarian (Bulgaria)","Chinese (Simplified, PRC)","Chinese (Traditional, Taiwan)","Croatian (Croatia)","Czech (Czech Republic)","Danish (Denmark)","Dutch (Netherlands)","English (United States)","Estonian (Estonia)","Finnish (Finland)","French (France)","German (Germany)","Greek (Greece)","Hebrew (Israel)","Hindi (India)","Hungarian (Hungary)","Indonesian (Indonesia)","Italian (Italy)","Japanese (Japan)","Kazakh (Kazakhstan)","Korean (Korea)","Latvian (Latvia)","Lithuanian (Lithuania)","Malay (Malaysia)","Norwegian, Bokmål (Norway)","Polish (Poland)","Portuguese (Brazil)","Portuguese (Portugal)","Romanian (Romania)","Russian (Russia)","Serbian (Latin, Serbia and Montenegro (Former))","Serbian (Latin, Serbia)","Slovak (Slovakia)","Slovenian (Slovenia)","Spanish (Spain)","Swedish (Sweden)","Thai (Thailand)","Turkish (Turkey)","Ukrainian (Ukraine)","Vietnamese (Vietnam)"

.PARAMETER CountryName
The name of the country the location of the bank holiday entry to be set to. 
If using the -ImportHolFile parameter, the country name needs to match the name defined in the CSV file corresponding to the selected culture. 
To get the correct country name to use, please use the Get-CountryNames function.
To specify a custom country name, please specify the custom value for the CountryName parameter or add the custom country name to your custom, user defined, CSV import file.
This parameter cannot be used in conjunction with the -ImportFromCustomCsv when importing a custom, user defined, CSV file.
	
.PARAMETER ImportFromCustomCsv
Use this switch to import entries from a custom, user defined, CSV file. 
When using this switch parameter, you must specify the path to the user defined CSV file, using the -CustomCsvFilePath parameter.

.PARAMETER CustomCsvFilePath
Path to the custom, user defined, CSV file containing bank holidays to import. 
The file format must follow the following structure: "Country"Delimiter"HolidayName"Delimiter"Date"Delimiter"Calendar"Delimiter"BusyStatus"
For example, if the delimiter is comma (,): "Country","HolidayName","Date","Calendar","BusyStatus"
For example, if the delimiter is semicolon: "Country";"HolidayName";"Date";"Calendar";"BusyStatus"
Here is an example of a custom, user defined file that is used to import two bank holidays
"Country";"HolidayName";"Date";"Calendar";"BusyStatus"
"United Kingdom";"Queen's Diamond Jubilee";"2012/06/05";"Gregorian";"OOF"
"United Kingdom";"Queen's Platinum Jubilee";"2022/06/02";"Gregorian";"OOF"

.PARAMETER CustomCsvDelimiter
Delimiter separating the values in the file the -CustomCsvFilePath points to. 

.PARAMETER ImportHolFile
Use this switch to import bank holiday entries from the pre-canned bank holiday entries. 
When using this switch parameter, you must specify the path to the pre-canned HOL file, using the -HolFolderPath parameter. 

.PARAMETER HolFolderPath
Path to the custom, pre-canned, HOL file containing bank holidays to import.These are generated based on Outlook's .hol holiday import files.
Note: If you've imported all module files, this parameter is not necessary as the script will try and locate the required files automatically in the well known module locations.

.PARAMETER Year
A specific year to import bank holidays for when importing holidays from the pre-canned HOL files, using ImportHolFile.

.PARAMETER Date
Use this parameter to specify the date of a single custom entry bank holiday. The parameter must be passed as a string value using quotes and the following format: "yyyy/mm/dd". For example: "2021/05/28" for the 28th may 2021.

.PARAMETER HolidayName
Use this parameter to specify Subject of a single custom entry bank holiday.

.PARAMETER Calendar
Use this parameter to specify the calendar format of a single custom entry bank holiday. 
For example a date time format where the year is 14XX, the calendar type is Hijri. 
For a date format where the year is 57XX, the calendar type is Hebrew.
For a date format where the year is 20XX, the calendar type is Gregorian.
The possible values are Gregorian, Hebrew or Hijri. 

.PARAMETER BusyStatus
Use this parameter to specify the busy status of a single custom entry bank holiday. 
The possible values are "Busy", "Free", and "OOF"

.INPUTS

None. You cannot pipe objects to Add-BankHoliday.

.OUTPUTS

None.

.EXAMPLE

Using a custom pre-canned import file:
Add-BankHoliday -EmailAddress user1@contoso.com. user2@contoso.com -UseModernAuthentication -Impersonate -UseAutodiscover -TenantId contoso.com -MSALNetDllPath "C:\Assemblies\Microsoft.Identity.Client.dll" -EWSManagedAPIPath "C:\Assemblies\Microsoft.Exchange.WebServices.dll" -ImportHolFile -HolFolderPath 'C:\BankHolidays' -Culture "English (United States)" -Country "United Kingdom"

Entry: 2012/8/6 August Bank Holiday (Scotland) created
Entry: 2013/8/5 August Bank Holiday (Scotland) created
Entry: 2014/8/4 August Bank Holiday (Scotland) created
Entry: 2015/8/3 August Bank Holiday (Scotland) created
Entry: 2016/8/1 August Bank Holiday (Scotland) created
Entry: 2017/8/7 August Bank Holiday (Scotland) created
Entry: 2018/8/6 August Bank Holiday (Scotland) created
Entry: 2019/8/5 August Bank Holiday (Scotland) created
Entry: 2020/8/3 August Bank Holiday (Scotland) created
Entry: 2021/8/2 August Bank Holiday (Scotland) created
Entry: 2022/8/1 August Bank Holiday (Scotland) created
Entry: 2023/8/7 August Bank Holiday (Scotland) created
Entry: 2024/8/5 August Bank Holiday (Scotland) created
Entry: 2025/8/4 August Bank Holiday (Scotland) created
Entry: 2026/8/3 August Bank Holiday (Scotland) created
Import complete

.EXAMPLE

Adding bank holidays for Germany, for year 2023, from the pre-canned HOL file for the English (United States) culture.
Add-BankHoliday -EmailAddress user1@contoso.com -Impersonate -UseAutodiscover -UseModernAuthentication -TenantId contoso.com -Culture 'English (United States)' -CountryName Germany -Year 2023 -ImportHolFile
Processing mailbox user1@contoso.com...
Entry: 2023/11/1 All Saints' Day created
Entry: 2023/5/18 Ascension created
Entry: 2023/8/15 Assumption created
Entry: 2023/12/26 Christmas (2nd Day) created
Entry: 2023/12/25 Christmas Day created
Entry: 2023/6/8 Corpus Christi created
Entry: 2023/10/3 Day of German Unity created
Entry: 2023/11/22 Day of Prayer and Repentance created
Entry: 2023/4/9 Easter Day created
Entry: 2023/4/10 Easter Monday created
Entry: 2023/1/6 Epiphany created
Entry: 2023/4/7 Good Friday created
Entry: 2023/1/1 New Year's Day created
Entry: 2023/10/31 Reformation Day created
Entry: 2023/5/29 Whit Monday created
Entry: 2023/5/28 Whit Sunday created
Entry: 2023/5/1 Labor Day created
Entry: 2023/12/24 Christmas Eve created
Entry: 2023/12/31 Sylvester created
Import complete

.EXAMPLE

Using a custom, user defined, import file:
Add-BankHoliday -EmailAddress user1@contoso.com. user2@contoso.com -UseModernAuthentication -Impersonate -UseAutodiscover -TenantId contoso.com -MSALNetDllPath "C:\Assemblies\Microsoft.Identity.Client.dll" -EWSManagedAPIPath "C:\Assemblies\Microsoft.Exchange.WebServices.dll" -ImportFromCustomCsv -CustomCsvFilePath "C:\BankHolidays\BankHolidays.csv" -CustomCsvDelimiter ";"

Example custom, user defined, CSV file:

"Country";"HolidayName";"Date";"Calendar";"BusyStatus"
"United Kingdom";"Queen's Diamond Jubilee";"2012/06/05";"Gregorian";"OOF"
"United Kingdom";"Queen's Platinum Jubilee";"2022/06/02";"Gregorian";"OOF"

.EXAMPLE

Ading a single custom bank holiday entry:
Add-BankHoliday -EmailAddress user1@contoso.com. user2@contoso.com -UseModernAuthentication -Impersonate -UseAutodiscover -TenantId contoso.com -MSALNetDllPath "C:\Assemblies\Microsoft.Identity.Client.dll" -EWSManagedAPIPath "C:\Assemblies\Microsoft.Exchange.WebServices.dll" -CountryName Elysium -Date "2022/02/15" -HolidayName "Elysium Day" -BusyStatus "OOF" -Calendar "Gregorian"

Entry: 2022/02/15 Elysium Day already exists
Import complete

.LINK

https://github.com/andreighita/EWSHoliday

.NOTES
File name      : BankHoliday.psm1
Author         : Andrei Ghita (catagh@microsoft.com)
Created        : 2021-02-09
Last reviewer  : Andrei Ghita (catagh@microsoft.com)
Last revision  : 2021-02-09
Disclaimer     : THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

#>

[CmdletBinding(DefaultParameterSetName = 'HOL')]
param(

    <# Connection related parameters #>

    [Parameter(Mandatory=$true, HelpMessage = "Single email address, comma separated list of mailbox email addresses, or an array of strings containing the email addresses to connect to.")] [ValidateNotNullOrEmpty()]
    [string[]]$EmailAddress,
    
    [Parameter()]
    [switch]$Impersonate,
    
    [Parameter()]
    [switch]$UseAutodiscover,

    [Parameter()]
    [switch]$UseModernAuthentication,

    [Parameter()]
    [string]$ClientId = "f570b6de-4b04-46d0-a141-7367145db206",

    [Parameter()]
    [string]$TenantId,

    [Parameter()]
    [string]$RedirectUri = "https://localhost",

    [Parameter()] [ValidateScript({ return [System.IO.File]::Exists($_) })]
    [string]$EWSManagedAPIPath,

    [Parameter()] [ValidateScript({ return [System.IO.File]::Exists($_) })]
    [string]$MSALNetDllPath,

    <# pre-canned HOL import parameters #>
    [Parameter(ParameterSetName = "HOL")]
    [ValidateSet("Arabic (Saudi Arabia)","Bulgarian (Bulgaria)","Chinese (Simplified, PRC)","Chinese (Traditional, Taiwan)","Croatian (Croatia)","Czech (Czech Republic)","Danish (Denmark)","Dutch (Netherlands)","English (United States)","Estonian (Estonia)","Finnish (Finland)","French (France)","German (Germany)","Greek (Greece)","Hebrew (Israel)","Hindi (India)","Hungarian (Hungary)","Indonesian (Indonesia)","Italian (Italy)","Japanese (Japan)","Kazakh (Kazakhstan)","Korean (Korea)","Latvian (Latvia)","Lithuanian (Lithuania)","Malay (Malaysia)","Norwegian, Bokmål (Norway)","Polish (Poland)","Portuguese (Brazil)","Portuguese (Portugal)","Romanian (Romania)","Russian (Russia)","Serbian (Latin, Serbia and Montenegro (Former))","Serbian (Latin, Serbia)","Slovak (Slovakia)","Slovenian (Slovenia)","Spanish (Spain)","Swedish (Sweden)","Thai (Thailand)","Turkish (Turkey)","Ukrainian (Ukraine)","Vietnamese (Vietnam)")]
    [string]$Culture = 'English (United States)',
    
    [Parameter(ParameterSetName = "HOL")]
    [Parameter(ParameterSetName = "Custom Entry")]
    [ValidateNotNullOrEmpty()]
    [string]$CountryName,

    <# Custom CSV file import parameters #>

    [Parameter(ParameterSetName = "Custom CSV")]
    [switch]$ImportFromCustomCsv,
    
    [Parameter(ParameterSetName = "Custom CSV")] [ValidateScript({ return [System.IO.File]::Exists($_) })]
    [string]$CustomCsvFilePath,

    [Parameter(ParameterSetName = "Custom CSV")]
    [string]$CustomCsvDelimiter = ",",
    
    [Parameter(ParameterSetName = "HOL")]
    [switch]$ImportHolFile,
    
    [Parameter(ParameterSetName = "HOL")] [ValidateScript({ return [System.IO.Directory]::Exists($_) })]
    [string]$HolFolderPath,
    
    [Parameter(ParameterSetName = "HOL")]
    [int]$Year = -1,

    <# Custom Entry Bank holiday parameters #>
    
    [Parameter(ParameterSetName = "Custom Entry")] [ValidateNotNullOrEmpty()]
    [string]$Date,
    
    [Parameter(ParameterSetName = "Custom Entry")] [ValidateNotNullOrEmpty()]
    [string]$HolidayName,
    
    [Parameter(ParameterSetName = "Custom Entry")] [ValidateSet("Gregorian", "Hebrew", "Hijri")]
    [string]$Calendar = "Gregorian",
    
    [Parameter(ParameterSetName = "Custom Entry")] [ValidateSet("Free", "Busy", "OOF")]
    [string]$BusyStatus = "OOF"

);

    $CountryNames = @()

    $HolidayEntries = @()

    if ($CountryName -ne [string]::Empty -and $CountryName -ne $null)
    {
        if ($ImportHolFile.IsPresent)
        {
            foreach ($countryNameEntry in $CountryName)
            {
                if ((Get-CountryNames -Culture $Culture).Contains($countryNameEntry))
                {
                    $CountryNames  = (Get-CountryNames -Culture $Culture | Where-Object { $_ -like "$countryNameEntry*" })
                }
                else
                {
                    write-error "You must specify a valid country name. Please use one of the values below:"
                    (Get-CountryNames -Culture $Culture) -join ", "
                    return;
                }
            }
        }
        else
        {
            if ($ImportFromCustomCsv.IsPresent)
            {
                Write-Error "You cannot use the CountryName parameter in conjunction with ImportFromCustomCsv."
                return;
            }
            else
            {
                if ($CountryName.Count -gt 1)
                {
                    Write-Error "You cannot specify multiple values for the CountryName parameter when creating a single custom entry bank holiday."
                    return;
                }
                else
                {
                    $CountryNames += $CountryName
                }
            }
        }

        
        if ($ImportHolFile.IsPresent)
        {
            if ($HolFolderPath -eq $null -or $HolFolderPath -eq [string]::Empty)
            {
                $HolFolderPath = Get-HolFolderPath
                if ($HolFolderPath -eq $null -or $HolFolderPath -eq [string]::Empty)
                {
                    throw "When specifying the -ImportHolFile parameter you must also specify the path of the folder containing the holiday import CSV files using the -HolFolderPath parameter."
                }
            }

            foreach ($CountryNameEntry in $CountryName)
            {
                if ($Year -ne -1)
                { 
                    $HolidayEntries += Get-HolFileEntry -Culture $Culture -CountryName $CountryName -HolFolderPath $HolFolderPath | where { $_.Date -like "$Year*" }
                }
                else
                {
                    $HolidayEntries += Get-HolFileEntry -Culture $Culture -CountryName $CountryName -HolFolderPath $HolFolderPath
                }
            }
        }
        else
        {
                $HolidayEntry = [Holiday]::new()
                $HolidayEntry.Calendar = $Calendar
                $HolidayEntry.CountryName = $CountryName
                $HolidayEntry.Date = $Date
                $HolidayEntry.HolidayName = $HolidayName
                $HolidayEntry.BusyStatus = $BusyStatus
                $HolidayEntries += $HolidayEntry
        }
    }
    else
    {
        if ($ImportHolFile.IsPresent)
        {
            if ($HolFolderPath -eq $null -or $HolFolderPath -eq [string]::Empty)
            {
                $HolFolderPath = Get-HolFolderPath
                if ($HolFolderPath -eq $null -or $HolFolderPath -eq [string]::Empty)
                {
                    throw "When specifying the -ImportHolFile parameter you must also specify the path of the folder containing the holiday import CSV files using the -HolFolderPath parameter." 
                }
            }
            
            if ($Year -ne -1)
            { 
                $HolidayEntries += Get-HolFileEntry -Culture $Culture -HolFolderPath $HolFolderPath | where { $_.Date -like "$Year*" }
            }
            else
            {
                $HolidayEntries += Get-HolFileEntry -Culture $Culture -HolFolderPath $HolFolderPath
            }
        }
        else 
        {
            if($ImportFromCustomCsv.IsPresent)
            {
                if ($CustomCsvFilePath -eq $null -or $CustomCsvFilePath -eq [string]::Empty)
                {
                    throw "When specifying the -ImportFromCustomCsv parameter you must also specify the of the custom CSV file containing the CSV to be inported using the -CustomCsvFilePath parameter."
                    
                }
                $HolidayEntries += Get-CustomCSVHolidayEntry -CSVFilePath $CustomCsvFilePath -Delimiter $CustomCsvDelimiter
            }
            else
            {
                throw "You must specify a country name using the -CountryName parameter. Alternatively, you can use the -ImportHolFile or -ImportFromCustomCsv parameters to import entries from a CSV file"
                
           }
        }
    }

    $ExchangeService = $null;
    $ExchangeService = New-ExchangeService -UseModernAuthentication -ClientId $ClientId -TenantId $TenantId -RedirectUri $RedirectUri -MSALNetDllPath $MSALNetDllPath

    $EwsCalendarFolder = $null

    $EmailPace = 1 / $EmailAddress.count * 100
    $EmailProgress = 0
    
    foreach ($EmailAddressEntry in $EmailAddress)
    {
        Write-Host "Processing mailbox $($EmailAddressEntry)..."
        Write-Progress -Activity "Processing mailbox $($EmailAddressEntry)" -Status "Running" -PercentComplete $EmailProgress -Id 1

        $EwsUrl = Get-EwsUrl -EmailAddress $EmailAddressEntry
        $EwsCalendarFolder = Get-EwsCalendarFolder -ExchangeService $ExchangeService -EmailAddress $EmailAddressEntry -UseAutodiscover $UseAutodiscover.IsPresent -Impersonate $Impersonate.IsPresent

        if ($EwsCalendarFolder -ne $null)
        {

            # Extended properties
            $PS_PUBLIC_STRINGS_Guid = new-object Guid("{00020329-0000-0000-C000-000000000046}")
            $PSETID_Appointment_Guid = new-object Guid("{00062002-0000-0000-C000-000000000046}")
            $Keywords = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($PS_PUBLIC_STRINGS_Guid, "Keywords", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::StringArray)  
            [string]$DateTimeFormat = "d";

            $holidayPace = 1 / $HolidayEntries.Count * 100
            $HolidayProgress = 0


            foreach ($HolidayEntry in $HolidayEntries)
            {
                $found = $false
                Write-Progress -Activity "Processing holiday $($HolidayEntry.CountryName) - $($HolidayEntry.HolidayName) : $($HolidayEntry.Date)" -Status "Running" -PercentComplete $holidayProgress -Id 2

                write-host "Entry:" $HolidayEntry.Date $HolidayEntry.HolidayName -NoNewline
                $SearchFilterCollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
                $SearchFilter1 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject,$HolidayEntry.HolidayName)
                $Start = new-object System.DateTime
                $Start = [System.DateTime]::Parse($HolidayEntry.Date)
                $SearchFilter2 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, $Start)
                $SearchFilter3 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Location, $HolidayEntry.CountryName)
                $SearchFilterCollection.Add($SearchFilter1)
                $SearchFilterCollection.Add($SearchFilter2)
                $SearchFilterCollection.Add($SearchFilter3)
         
                $itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(20)
                $itemView.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
            
                $findResults = $ExchangeService.FindItems($EwsCalendarFolder.Id,$SearchFilterCollection,$itemView)
                if ($findResults.Items.Count -eq 0)
                {
                    $StartTime = $null

                    if ($HolidayEntry.Calendar -eq "Hebrew")
                    {
                        [System.Globalization.HebrewCalendar]$HC = New-Object System.Globalization.HebrewCalendar
                        $HD = [System.DateTime]::Parse($HolidayEntry.Date);
                        $StartTime = $HC.ToDateTime($HD.Year, $HD.Month, $hd.Day, 0, 0, 0, 0)
                    }
                    else
                    {
                        if ($HolidayEntry.Calendar -eq "Hijri")
                        {
                            [System.Globalization.HijriCalendar]$HC = New-Object System.Globalization.HijriCalendar
                            $HD = [System.DateTime]::Parse($HolidayEntry.Date);
                            $StartTime = $HC.ToDateTime($HD.Year, $HD.Month, $hd.Day, 0, 0, 0, 0)
                        }
                        else
                        {

                        }

                    }
                    
                    # Creating new entry
                    $Appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment -ArgumentList $ExchangeService  
                    # Set Subject  
                    $Appointment.Subject = $HolidayEntry.HolidayName
                    # Set Start Time  
                    if ($ExchangeVersion -eq "Exchange2007_SP1")
                    {
                        $Appointment.StartTimeZone = $ExchangeService.TimeZone;
                    }
                    $Appointment.Start = [System.DateTime]::Parse($HolidayEntry.Date)
                    # Set Start Time  
                    $Appointment.End = [System.DateTime]::Parse($HolidayEntry.Date).AddDays(1)
				    # No reminder
				    $Appointment.IsReminderSet = $false
                    # Mark as all day event 
                    $Appointment.IsAllDayEvent = $true;
                    switch ($HolidayEntry.BusyStatus)
                    {
                        "Free"
                        {
                            $Appointment.LegacyFreeBusyStatus = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::Free;
                        }
                        "Busy"
                        {
                            $Appointment.LegacyFreeBusyStatus = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::Busy;
                        }
                        "OOF"
                        {
                            $Appointment.LegacyFreeBusyStatus = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::OOF;
                        }
                    }
                    # Setting extended properties
                    [string[]]$KeyWordsValue = @("Holiday")
                    $Appointment.SetExtendedProperty($Keywords, $KeyWordsValue);
                    $Appointment.Location = $HolidayEntry.CountryName

                    # Saving the entry
                    $Appointment.Save($EwsCalendarFolder.Id, [Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)  
                    write-host " created"    
                }
                else
                {
                    Write-Host " already exists"

                }
                $HolidayProgress += $holidayPace
            }
            write-host "Import complete" -ForegroundColor Green
        }
        else
        {
            Write-Warning "Unable to connect to calendar folder for mailbox $($EmailAddressEntry)"
        }
        $EmailProgress = $EmailProgress + $EmailPace
        Write-Host ""
    }
    return
}

function Remove-BankHoliday
{
<#
.SYNOPSIS
Removes bank holiday entries in a given mailbox or in multipe mailboxes. 

.DESCRIPTION
Uses Exchange Web Services to connect to one or multiple mailboxes and removes bank holidays in a given mailbox or in multipe mailboxes. The script requires the Ews Managed API assembly be referenced, as well as the MSAL.Net assembly.
The function allows the automatic discovery of service endpoints, as well as the use of modern authentication to connect to Exchange Online hosted mailboxes or on premises hosted mailboxes. 

.PARAMETER EmailAddress
Single email address, comma separated list of mailbox email addresses, or an array of strings containing the email addresses to connect to. See the examples section for examples.

.PARAMETER Impersonate
Use this switch parameter to have the authenticated account use Exchange Impersonation when connecting to mailboxes.

.PARAMETER UseAutodiscover
Use this switch parameter use automatic discovery of the Exchange Web Services Service Url using Autodiscover V1 and V2.

.PARAMETER UseModernAuthentication
Use this switch parameter use Modern Authentication (OAuth credentials) to authenticate when connecting to Exchange mailboxes. Use this switch parameter for Exchange Online or Hybrid Exchange Server deployments.

.PARAMETER ClientId
The Application (client) ID value of the Azure Portal app registration that exposes the EWS.AccessAsUser.All permission. 

.PARAMETER TenantId
The unique GUID tenant identifier or the fully qualified domain name. For example "105a90c6-224d-4609-9fbd-af1f927a554f" or "contoso.com".

.PARAMETER RedirectUri
The redirect URI defined in the Azure Portal app registration. 

.PARAMETER EWSManagedAPIPath
The path to the Microsoft.Exchange.WebServices.dll file. 
Note: If you've imported all module files, this parameter is not necessary as the script will try and locate the required assemblies automatically in the well known module locations.

.PARAMETER MSALNetDllPath
The path to the Microsoft.Identity.Client.dll file. This is required if -UseModernAuthentication is used. 
Note: If you've imported all module files, this parameter is not necessary as the script will try and locate the required assemblies automatically in the well known module locations.

.PARAMETER ItemId
The ItemId value of a single bank holiday entry to be removed.

.PARAMETER InputObject
A single HolidayEntry onbject or a collection of HolidayEntry objects to be removed.


.INPUTS

HolidayEntry. You can pipe HolidayEntry objects in to Remove-BankHoliday.

.OUTPUTS

None. The Remove-BankHoliday function does not return any values.

.EXAMPLE

Retreieve all bank holiday entries for the France in mailboxes user1@contoso.com and user2@contoso.com and store the result in a variable called $Holidays.Then pass the $holidays variable value as -InputObject to the Remove-BankHoliday function for removal.

$Holidays = Get-BankHoliday -EmailAddress user1@contoso.com, user2@contoso.com -Impersonate -UseAutodiscover -UseModernAuthentication -TenantId contoso.com -Culture 'English (United States)' -CountryName "France"

Holiday entries for mailbox user1@contoso.com:
240 entries returned for country France
Found 240 entries
Holiday entries for mailbox user2@contoso.com:
240 entries returned for country France
Found 240 entries

Remove-BankHoliday -Impersonate -UseAutodiscover -UseModernAuthentication -TenantId contoso.com -InputObject $holidays

Removing holiday entries for mailbox user1@contoso.com...240 items removed
Removing holiday entries for mailbox user2@contoso.com...240 items removed
Processing complete.

.LINK

https://github.com/andreighita/EWSHoliday

.NOTES
File name      : BankHoliday.psm1
Author         : Andrei Ghita (catagh@microsoft.com)
Created        : 2021-02-10
Last reviewer  : Andrei Ghita (catagh@microsoft.com)
Last revision  : 2021-02-10
Disclaimer     : THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

#>
[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Single email address, comma separated list of mailbox email addresses, or an array of strings containing the email addresses to connect to.")] [ValidateNotNullOrEmpty()]
    [Parameter(ParameterSetName = "Item Id")]
    [string]$EmailAddress,
    
    [Parameter()]
    [switch]$Impersonate,
    
    [Parameter()]
    [switch]$UseAutodiscover,

    [Parameter()]
    [switch]$UseModernAuthentication,

    [Parameter()]
    [string]$ClientId = "f570b6de-4b04-46d0-a141-7367145db206",

    [Parameter()]
    [string]$TenantId,

    [Parameter()]
    [string]$RedirectUri = "https://localhost",

    [Parameter()] [ValidateScript({ return [System.IO.File]::Exists($_) })]
    [string]$EWSManagedAPIPath,

    [Parameter()] [ValidateScript({ return [System.IO.File]::Exists($_) })]
    [string]$MSALNetDllPath,

    [ValidateNotNullOrEmpty()]
    [Parameter(ParameterSetName = "Item Id")]
    [string]$ItemId,

    [Parameter(ParameterSetName = "Input Object", ValueFromPipeLine=$true)]
    [ValidateNotNullOrEmpty()]
    [HolidayEntry[]]$InputObject = $null
)
    

    $ExchangeService = $null;
    $ExchangeService = New-ExchangeService -UseModernAuthentication -ClientId $ClientId -TenantId $TenantId -RedirectUri $RedirectUri -MSALNetDllPath $MSALNetDllPath 

    $EwsCalendarFolder = $null

    $EmailPace = 1 / $EmailAddress.count * 100
    $EmailProgress = 0
    $HolidayEntries = @()
    
    if ($InputObject -eq $null)
    {
        [HolidayEntry]$HolidayEntry = [HolidayEntry]::new()
        $HolidayEntry.EmailAddress = $EmailAddress;
        $HolidayEntry.CalendarFolderId = $CalendarFolderId;
        $HolidayEntry.ItemId = $ItemId;
        $HolidayEntries.Add($HolidayEntry);
    }
    else
    {
        $HolidayEntries = $InputObject
    }

    [string[]]$EmailAddresses = ($HolidayEntries | Select-Object emailaddress -unique).EmailAddress
    
    foreach ($EmailAddressEntry in $EmailAddresses)
    {
        Write-Host "Removing holiday entries for mailbox $($EmailAddressEntry)..." -NoNewline
        Write-Progress -Activity "Processing mailbox $($EmailAddressEntry)" -Status "Running" -PercentComplete $EmailProgress -Id 1

        $ItemIds = ($HolidayEntries | where {$_.emailaddress -eq $EmailAddressEntry}).ItemId

        $ServiceUrl = $EwsUrl
        if ($UseAutodiscover)
        {    
            $ServiceUrl = Get-EwsUrl -EmailAddress $EmailAddressEntry
            if ($ServiceUrl -ne $null)
            {
                $ExchangeService.Url = new-Object Uri($ServiceUrl)
            }
            else
            {
                $ExchangeService.AutodiscoverUrl($EmailAddressEntry)
            }
        }
        else
        {
            if ($ServiceUrl -ne $null -and $ServiceUrl -ne [string]::Empty)
            {
                $ExchangeService.Url = new-Object Uri($ServiceUrl)
            }
            else
            {
                throw "Unable to discover the service endpoint"
            }

        }

        # If we are using impersonation then setup impersonation
        if ($Impersonate)
        {
            $ExchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddressEntry);
        }

        $itempace = 1 / $ItemIds.count / 100
        $itemprogress = 0;
        $count = 0;
        foreach ($itemIdEntry in $ItemIds)
        {
            write-Progress -Activity "Removing items from mailbox $EmailAddressEntry" -Status "Running" -Id 2 -PercentComplete $itemprogress
            $item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($ExchangeService, (New-Object Microsoft.Exchange.WebServices.Data.ItemId($itemIdEntry)))
            $item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems);
            $itemProgress += $itempace
            $count++
        }
        $EmailProgress = $EmailProgress + $EmailPace
        Write-Host "$count items removed"
    }
    
    Write-Host "Processing complete."
    return
}

function Get-CustomCSVHolidayEntry
{
param(    
    [string]$CSVFilePath,
    [string]$Delimiter = '#'
    );

    $HeaderValidationArray = @("Country","HolidayName","Date")
    $Holidays = @()
    [System.Collections.ArrayList]$Content = ((get-content -Path $CSVFilePath)[0]) -split $Delimiter
    if  ($Content[0].Equals('"Country"') -and $Content[1].Equals('"HolidayName"') -and $Content[2].Equals('"Date"') -and $Content[3].Equals('"Calendar"') -and $Content[4].Equals('"BusyStatus"'))
    {
        $CsvEntries = Import-Csv $CSVFilePath -Delimiter $Delimiter
        if ($CsvEntries.Count -gt 0)
        {
            foreach ($CsvEntry in $CsvEntries)
            { 
                $holiday = [Holiday]::new()
                $holiday.CountryName = $CsvEntry.Country
                $Holiday.Date = $CsvEntry.Date
                $Holiday.HolidayName = $CsvEntry.HolidayName
                $holiday.Calendar = $CsvEntry.Calendar
                $holiday.BusyStatus = $CsvEntry.BusyStatus
                $holidays += $holiday
            }
        }
        return $holidays 
    }
}

function Get-HolFileEntry
{
param(    
    [ValidateSet("Arabic (Saudi Arabia)","Bulgarian (Bulgaria)","Chinese (Simplified, PRC)","Chinese (Traditional, Taiwan)","Croatian (Croatia)","Czech (Czech Republic)","Danish (Denmark)","Dutch (Netherlands)","English (United States)","Estonian (Estonia)","Finnish (Finland)","French (France)","German (Germany)","Greek (Greece)","Hebrew (Israel)","Hindi (India)","Hungarian (Hungary)","Indonesian (Indonesia)","Italian (Italy)","Japanese (Japan)","Kazakh (Kazakhstan)","Korean (Korea)","Latvian (Latvia)","Lithuanian (Lithuania)","Malay (Malaysia)","Norwegian, Bokmål (Norway)","Polish (Poland)","Portuguese (Brazil)","Portuguese (Portugal)","Romanian (Romania)","Russian (Russia)","Serbian (Latin, Serbia and Montenegro (Former))","Serbian (Latin, Serbia)","Slovak (Slovakia)","Slovenian (Slovenia)","Spanish (Spain)","Swedish (Sweden)","Thai (Thailand)","Turkish (Turkey)","Ukrainian (Ukraine)","Vietnamese (Vietnam)")]
    [string]$Culture,
    [string]$CountryName,
    [string]$HolFolderPath
    );
    $Holidays = @()
    $FilePath = "{0}\{1}.hol" -f $HolFolderPath, $Culture
    
    if ($CountryName -ne $null -and $CountryName -ne [string]::Empty)
    {
        $CsvEntries = Import-Csv $FilePath -Delimiter "#" | where { $_.Country -eq $CountryName }
    }
    else
    {
        $CsvEntries = Import-Csv $FilePath -Delimiter "#"
    }

    if ($CsvEntries.Count -gt 0)
    {
        foreach ($CsvEntry in $CsvEntries)
        { 
            $holiday = [Holiday]::new()
            $holiday.CountryName = $CsvEntry.Country
            $Holiday.Date = $CsvEntry.Date
            $Holiday.HolidayName = $CsvEntry.HolidayName
            $holiday.Calendar = $CsvEntry.Calendar
            $holiday.BusyStatus = $CsvEntry.BusyStatus
            $holidays += $holiday
        }
    }
    return $holidays 
}


function Get-CountryNames
{
param(
[ValidateSet("Arabic (Saudi Arabia)","Bulgarian (Bulgaria)","Chinese (Simplified, PRC)","Chinese (Traditional, Taiwan)","Croatian (Croatia)","Czech (Czech Republic)","Danish (Denmark)","Dutch (Netherlands)","English (United States)","Estonian (Estonia)","Finnish (Finland)","French (France)","German (Germany)","Greek (Greece)","Hebrew (Israel)","Hindi (India)","Hungarian (Hungary)","Indonesian (Indonesia)","Italian (Italy)","Japanese (Japan)","Kazakh (Kazakhstan)","Korean (Korea)","Latvian (Latvia)","Lithuanian (Lithuania)","Malay (Malaysia)","Norwegian, Bokmål (Norway)","Polish (Poland)","Portuguese (Brazil)","Portuguese (Portugal)","Romanian (Romania)","Russian (Russia)","Serbian (Latin, Serbia and Montenegro (Former))","Serbian (Latin, Serbia)","Slovak (Slovakia)","Slovenian (Slovenia)","Spanish (Spain)","Swedish (Sweden)","Thai (Thailand)","Turkish (Turkey)","Ukrainian (Ukraine)","Vietnamese (Vietnam)")]
$Culture
)

$cultures = @{}
$Cultures.Add("Arabic (Saudi Arabia)", @("أذربيجان","أرمينيا","أسبانيا","أستراليا","إستونيا","الاتحاد الأوروبي","الأرجنتين","الأردن","الأعياد الدينية الإسلامية (السنّة)","الأعياد الدينية الإسلامية (الشيعة)","الأعياد الدينية المسيحية","الإكوادور","الإمارات العربية المتحدة","ألبانيا","البحرين","البرازيل","البرتغال","البوسنة والهرسك","التشيك","الجزائر","الدنمارك","السعودية","السلفادور","السويد","الصين","العراق","الفلبين","الكرسي الرسولي (دولة الفاتيكان)","الكونغو (جمهورية الكونغو الديمقراطية)","الكويت","ألمانيا","المغرب","المكسيك","المملكة المتحدة","النرويج","النمسا","الهند","الولايات المتحدة","اليابان","اليمن","اليونان","أندورا","أنغولا","أوروجواي","أوكرانيا","إيرلندا","أيسلندا","إيطاليا","باراجواي","باكستان","بروناي","بلجيكا","بلغاريا","بنما","بورتو ريكو","بولندا","بوليفيا","بيرو","بيلاروس","تايلاند","تركيا","ترينداد وتوباغو","تشيلي","تونس","جامايكا","جمهورية الدومينيكان","جنوب أفريقيا","جورجيا","روسيا","رومانيا","سان مارينو","سلوفاكيا","سلوفينيا","سنغافورة","سوريا","سويسرا","شمال مقدونيا","صربيا","عُمان","غواتيمالا","غينيا الاستوائية","فرنسا","فنزويلا","فنلندا","فيتنام","قبرص","قطر","كازاخستان","كرواتيا","كندا","كوريا","كوستاريكا","كولومبيا","كينيا","لاتفيا","لبنان","لتوانيا","لكسمبورغ","ليشتنشتاين","مالطا","ماليزيا","مصر","منطقة ماكاو الإدارية الخاصة","منطقة هونغ كونغ الإدارية الخاصة","مولدافا","موناكو","مونتينيغرو","نيجيريا","نيكاراجوا","نيوزيلندا","هندوراس","هنغاريا","هولندا"))
$Cultures.Add("Bulgarian (Bulgaria)", @("Австралия","Австрия","Азeрбайджан","Албания","Алжир","Ангола","Андора","Аржентина","Армения","Бахрейн","Беларус","Белгия","Боливия","Босна и Херцеговина","Бразилия","Бруней","България","Великобритания","Венецуела","Виетнам","Гватемала","Германия","Грузия","Гърция","Дания","Доминиканска република","Еврейски религиозни празници","Европейски съюз","Египет","Еквадор","Екваториална Гвинея","Естония","Йемен","Израел","Индия","Йордания","Ирак","Ирландия","Исландия","Ислямски (сунитски) религиозни празници","Ислямски (шиитски) религиозни празници","Испания","Италия","Казахстан","Канада","Катар","Кения","Кипър","Китай","Колумбия","Конго (Демократична република)","Корея","Коста Рика","Кувейт","Латвия","Ливан","Литва","Лихтенщайн","Люксембург","Малайзия","Малта","Мароко","Мексико","Молдова","Монако","Нигерия","Нидерландия","Никарагуа","Нова Зеландия","Норвегия","Обединени арабски емирства","Оман","Пакистан","Панама","Парагвай","Перу","Полша","Португалия","Пуерто Рико","Румъния","Русия","Салвадор","Сан Марино","САР Макао","САР Хонконг","Саудитска Арабия","Светия престол (Ватикан)","Северна Македония","Сингапур","Сирия","Словакия","Словения","Съединени щати","Сърбия","Тайланд","Тринидад и Тобаго","Тунис","Турция","Украйна","Унгария","Уругвай","Филипини","Финландия","Франция","Хондурас","Християнски религиозни празници","Хърватска","Черна гора","Чехия","Чили","Швейцария","Швеция","Южна Африка","Ямайка","Япония"))
$Cultures.Add("Chinese (Simplified, PRC)", @("中国","丹麦","乌克兰","乌拉圭","也门","亚美尼亚","以色列","伊拉克","伊斯兰(什叶派)宗教假日","伊斯兰(逊尼派)宗教假日","俄罗斯","保加利亚","克罗地亚","冰岛","列支敦士登","刚果民主共和国","加拿大","匈牙利","北马其顿","南非","卡塔尔","卢森堡","印度","危地马拉","厄瓜多尔","叙利亚","哈萨克斯坦","哥伦比亚","哥斯达黎加","土耳其","圣座(梵蒂冈)","圣马力诺","埃及","基督教节日","塞尔维亚","塞浦路斯","墨西哥","多米尼加共和国","奥地利","委内瑞拉","安哥拉","安道尔","尼加拉瓜","尼日利亚","巴基斯坦","巴拉圭","巴拿马","巴林","巴西","希腊","德国","意大利","拉脱维亚","挪威","捷克","摩尔多瓦","摩洛哥","摩纳哥","文莱","斯洛伐克","斯洛文尼亚","新加坡","新西兰","日本","智利","格鲁吉亚","欧盟","比利时","沙特阿拉伯","法国","波兰","波多黎各","波斯尼亚和黑塞哥维那","泰国","洪都拉斯","澳大利亚","澳门特别行政区","爱尔兰","爱沙尼亚","牙买加","特立尼达和多巴哥","犹太教节日","玻利维亚","瑞典","瑞士","白俄罗斯","科威特","秘鲁","突尼斯","立陶宛","约旦","罗马尼亚","美国","肯尼亚","芬兰","英国","荷兰","菲律宾","萨尔瓦多","葡萄牙","西班牙","赤道几内亚","越南","阿塞拜疆","阿尔及利亚","阿尔巴尼亚","阿拉伯联合酋长国","阿曼","阿根廷","韩国","香港特别行政区","马来西亚","马耳他","黎巴嫩","黑山"))
$Cultures.Add("Chinese (Traditional, Taiwan)", @("丹麥","亞塞拜然","亞美尼亞","以色列","伊拉克","俄羅斯","保加利亞","克羅埃西亞","冰島","列支敦斯登","剛果民主共和國","加拿大","匈牙利","北馬其頓","千里達及托巴哥","南非","卡達","印度","厄瓜多","哈薩克","哥倫比亞","哥斯大黎加","喬治亞","回教 (什葉派) 假日","回教 (遜尼派) 假日","土耳其","埃及","基督教假日","塞爾維亞","墨西哥","多明尼加","奈及利亞","奧地利","委內瑞拉","安哥拉","安道爾","宏都拉斯","尼加拉瓜","巴基斯坦","巴拉圭","巴拿馬","巴林","巴西","希臘","德國","愛沙尼亞","愛爾蘭","拉脫維亞","挪威","捷克","摩洛哥","摩爾多瓦","摩納哥","敘利亞","教廷 (梵蒂岡)","斯洛伐克","斯洛維尼亞","新加坡","日本","智利","歐盟","比利時","汶萊","沙烏地阿拉伯","法國","波士尼亞赫塞哥維納","波多黎各","波蘭","泰國","澳大利亞","澳門特別行政區","烏克蘭","烏拉圭","牙買加","猶太教假日","玻利維亞","瑞典","瑞士","瓜地馬拉","白俄羅斯","盧森堡","科威特","秘魯","突尼西亞","立陶宛","約旦","紐西蘭","羅馬尼亞","美國","義大利","聖馬利諾","肯亞","芬蘭","英國","荷蘭","菲律賓","葉門","葡萄牙","蒙特內哥羅","薩爾瓦多","西班牙","賽普勒斯","赤道幾內亞","越南","阿拉伯聯合大公國","阿曼","阿根廷","阿爾及利亞","阿爾巴尼亞","韓國","香港特別行政區","馬來西亞","馬爾他","黎巴嫩"))
$Cultures.Add("Croatian (Croatia)", @("Albanija","Alžir","Andora","Angola","Argentina","Armenija","Australija","Austrija","Azerbajdžan","Bahrein","Belgija","Bjelarus","Bolivija","Bosna i Hercegovina","Brazil","Brunej","Bugarska","Češka","Čile","Cipar","Crna Gora","Danska","Dominikanska Republika","Egipat","Ekvador","Ekvatorijalna Gvineja","El Salvador","Estonija","Europska unija","Filipini","Finska","Francuska","Grčka","Gruzija","Gvatemala","Honduras","Hrvatska","Indija","Irak","Irska","Islamski (šijitski) vjerski blagdani","Islamski (sunitski) vjerski blagdani","Island","Italija","Izrael","Jamajka","Japan","Jemen","Jordan","Južnoafrička Republika","Kanada","Katar","Kazahstan","Kenija","Kina","Kolumbija","Kongo (Demokratska Republika)","Koreja","Kostarika","Kršćanski vjerski blagdani","Kuvajt","Latvija","Libanon","Lihtenštajn","Litva","Luksemburg","Mađarska","Malezija","Malta","Maroko","Meksiko","Moldova","Monako","Nigerija","Nikaragva","Nizozemska","Njemačka","Norveška","Novi Zeland","Oman","Pakistan","Panama","Paragvaj","Peru","Poljska","Portoriko","Portugal","Posebna upr. jedinica Hong Kong","Posebna upravna jedinica Makao","Rumunjska","Rusija","San Marino","Saudijska Arabija","Singapur","Sirija","Sjedinjene Američke Države","Sjeverna Makedonija","Slovačka","Slovenija","Španjolska","Srbija","Švedska","Sveta Stolica (Vatikan)","Švicarska","Tajland","Trinidad i Tobago","Tunis","Turska","Ujedinjeni Arapski Emirati","Ujedinjeno Kraljevstvo","Ukrajina","Urugvaj","Venezuela","Vijetnam","Židovski vjerski blagdani"))
$Cultures.Add("Czech (Czech Republic)", @("Albánie","Alžírsko","Andorra","Angola","Argentina","Arménie","Austrálie","Ázerbájdžán","Bahrajn","Belgie","Bělorusko","Bolívie","Bosna a Hercegovina","Brazílie","Brunej","Bulharsko","Černá Hora","Česko","Chile","Chorvatsko","Čína","Dánsko","Dominikánská republika","Egypt","Ekvádor","Estonsko","Evropská unie","Filipíny","Finsko","Francie","Gruzie","Guatemala","Honduras","Hongkong – zvláštní správní oblast","Indie","Irák","Irsko","Islámské náboženské svátky (šíité)","Islámské náboženské svátky (sunnité)","Island","Itálie","Izrael","Jamajka","Japonsko","Jemen","Jižní Afrika","Jordánsko","Kanada","Katar","Kazachstán","Keňa","Kolumbie","Kongo (Konžská demokratická republika)","Korea","Kostarika","Křesťanské náboženské svátky","Kuvajt","Kypr","Libanon","Lichtenštejnsko","Litva","Lotyšsko","Lucembursko","Macao – zvláštní správní oblast","Maďarsko","Malajsie","Malta","Maroko","Mexiko","Moldavsko","Monako","Německo","Nigérie","Nikaragua","Nizozemsko","Norsko","Nový Zéland","Omán","Pákistán","Panama","Paraguay","Peru","Polsko","Portoriko","Portugalsko","Rakousko","Řecko","Rovníková Guinea","Rumunsko","Rusko","Salvador","San Marino","Saúdská Arábie","Severní Makedonie","Singapur","Slovensko","Slovinsko","Španělsko","Spojené arabské emiráty","Spojené království","Spojené státy","Srbsko","Svatý stolec (Vatikán)","Švédsko","Švýcarsko","Sýrie","Thajsko","Trinidad a Tobago","Tunisko","Turecko","Ukrajina","Uruguay","Venezuela","Vietnam","Židovské náboženské svátky"))
$Cultures.Add("Danish (Denmark)", @("Ækvatorialguinea","Albanien","Algeriet","Andorra","Angola","Argentina","Armenien","Aserbajdsjan","Australien","Bahrain","Belgien","Bolivia","Bosnien-Hercegovina","Brasilien","Brunei","Bulgarien","Canada","Chile","Colombia","Congo (Den Demokratiske Republik)","Costa Rica","Cypern","Danmark","De Forenede Arabiske Emirater","Den Dominikanske Republik","Ecuador","Egypten","El Salvador","Estland","European Union (EU)","Filippinerne","Finland","Frankrig","Georgien","Grækenland","Guatemala","Honduras","Hviderusland","Indien","Irak","Irland","Islamiske (shia) religiøse helligdage","Islamiske (sunni) religiøse helligdage","Island","Israel","Italien","Jamaica","Japan","Jødiske helligdage","Jordan","Kasakhstan","Kenya","Kina","Korea","Kristne helligdage","Kroatien","Kuwait","Letland","Libanon","Liechtenstein","Litauen","Luxembourg","Macao","Malaysia","Malta","Marokko","Mexico","Moldova","Monaco","Montenegro","Nederlandene","New Zealand","Nicaragua","Nigeria","Nordmakedonien","Norge","Oman","Østrig","Pakistan","Panama","Paraguay","Pavestolen (Vatikanstaten)","Peru","Polen","Portugal","Puerto Rico","Qatar","Rumænien","Rusland","San Marino","SAR Hongkong","Saudi-Arabien","Schweiz","Serbien","Singapore","Slovakiet","Slovenien","Spanien","Storbritannien","Sverige","Sydafrika","Syrien","Thailand","Tjekkiet","Trinidad og Tobago","Tunesien","Tyrkiet","Tyskland","Ukraine","Ungarn","Uruguay","USA","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Dutch (Netherlands)", @("Albanië","Algerije","Andorra","Angola","Argentinië","Armenië","Australië","Azerbeidzjan","Bahrein","Belarus","België","Bolivia","Bosnië en Herzegovina","Brazilië","Brunei","Bulgarije","Canada","Chili","China","Christelijke feestdagen","Colombia","Congo (Democratische Republiek)","Costa Rica","Cyprus","Denemarken","Dominicaanse Republiek","Duitsland","Ecuador","Egypte","El Salvador","Equatoriaal-Guinea","Estland","Europese Unie","Filipijnen","Finland","Frankrijk","Georgië","Griekenland","Guatemala","Heilige Stoel (Vaticaanstad)","Honduras","Hongarije","Hongkong SAR","Ierland","IJsland","India","Irak","Islamitische (sjiitische) religieuze feestdagen","Islamitische (soennitische) religieuze feestdagen","Israël","Italië","Jamaica","Japan","Jemen","Joodse feestdagen","Jordanië","Kazachstan","Kenia","Koeweit","Kroatië","Letland","Libanon","Liechtenstein","Litouwen","Luxemburg","Macau SAR","Maleisië","Malta","Marokko","Mexico","Moldavië","Monaco","Montenegro","Nederland","Nicaragua","Nieuw-Zeeland","Nigeria","Noord-Macedonië","Noorwegen","Oekraïne","Oman","Oostenrijk","Pakistan","Panama","Paraguay","Peru","Polen","Porto Rico","Portugal","Qatar","Roemenië","Rusland","San Marino","Saudi-Arabië","Servië","Singapore","Slovenië","Slowakije","Spanje","Syrië","Thailand","Trinidad en Tobago","Tsjechië","Tunesië","Turkije","Uruguay","Venezuela","Verenigd Koninkrijk","Verenigde Arabische Emiraten","Verenigde Staten","Vietnam","Zuid-Afrika","Zuid-Korea","Zweden","Zwitserland"))
$Cultures.Add("English (United States)", @("Albania","Algeria","Andorra","Angola","Argentina","Armenia","Australia","Austria","Azerbaijan","Bahrain","Belarus","Belgium","Bolivia","Bosnia and Herzegovina","Brazil","Brunei","Bulgaria","Canada","Chile","China","Christian Religious Holidays","Colombia","Congo (Democratic Republic of)","Costa Rica","Croatia","Cyprus","Czechia","Denmark","Dominican Republic","Ecuador","Egypt","El Salvador","Equatorial Guinea","Estonia","European Union","Finland","France","Georgia","Germany","Greece","Guatemala","Holy See (Vatican City)","Honduras","Hong Kong S.A.R.","Hungary","Iceland","India","Iraq","Ireland","Islamic (Shia) Religious Holidays","Islamic (Sunni) Religious Holidays","Israel","Italy","Jamaica","Japan","Jewish Religious Holidays","Jordan","Kazakhstan","Kenya","Korea","Kuwait","Latvia","Lebanon","Liechtenstein","Lithuania","Luxembourg","Macao S.A.R.","Malaysia","Malta","Mexico","Moldova","Monaco","Montenegro","Morocco","Netherlands","New Zealand","Nicaragua","Nigeria","North Macedonia","Norway","Oman","Pakistan","Panama","Paraguay","Peru","Philippines","Poland","Portugal","Puerto Rico","Qatar","Romania","Russia","San Marino","Saudi Arabia","Serbia","Singapore","Slovakia","Slovenia","South Africa","Spain","Sweden","Switzerland","Syria","Thailand","Trinidad and Tobago","Tunisia","Turkey","Ukraine","United Arab Emirates","United Kingdom","United States","Uruguay","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Estonian (Estonia)", @("Albaania","Alžeeria","Ameerika Ühendriigid","Andorra","Angola","Aomeni erihalduspiirkond","Araabia Ühendemiraadid","Argentina","Armeenia","Aserbaidžaan","Austraalia","Austria","Bahrein","Belgia","Boliivia","Bosnia ja Hertsegoviina","Brasiilia","Brunei","Bulgaaria","Colombia","Costa Rica","Dominikaani Vabariik","Ecuador","Eesti","Egiptus","Ekvatoriaal-Guinea","El Salvador","Euroopa Liit","Filipiinid","Gruusia","Guatemala","Hiina","Hispaania","Holland","Honduras","Hongkongi erihalduspiirkond","Horvaatia","Iirimaa","Iisrael","India","Iraak","Islami (šiia) religioossed pühad","Islami (sunni) religioossed pühad","Island","Itaalia","Jaapan","Jamaica","Jeemen","Jordaania","Juudi pühad","Kanada","Kasahstan","Katar","Kenya","Kongo Demokraatlik Vabariik","Kreeka","Kristlikud pühad","Küpros","Kuveit","Läti","Leedu","Liechtenstein","Liibanon","Lõuna-Aafrika Vabariik","Lõuna-Korea","Luksemburg","Malaisia","Malta","Maroko","Mehhiko","Moldova","Monaco","Montenegro","Nicaragua","Nigeeria","Norra","Omaan","Pakistan","Panama","Paraguay","Peruu","Põhja-Makedoonia","Poola","Portugal","Prantsusmaa","Puerto Rico","Püha Tool (Vatikani Linnriik)","Rootsi","Rumeenia","Saksamaa","San Marino","Saudi Araabia","Serbia","Singapur","Slovakkia","Sloveenia","Soome","Süüria","Šveits","Taani","Tai","Trinidad ja Tobago","Tšehhia","Tšiili","Tuneesia","Türgi","Ühendkuningriik","Ukraina","Ungari","Uruguay","Uus-Meremaa","Valgevene","Venemaa","Venezuela","Vietnam"))
$Cultures.Add("Finnish (Finland)", @("Alankomaat","Albania","Algeria","Andorra","Angola","Argentiina","Armenia","Australia","Azerbaidžan","Bahrain","Belgia","Bolivia","Bosnia ja Hertsegovina","Brasilia","Brunei","Bulgaria","Chile","Costa Rica","Dominikaaninen tasavalta","Ecuador","Egypti","El Salvador","Espanja","Etelä-Afrikka","Euroopan unioni","Filippiinit","Georgia","Guatemala","Honduras","Intia","Irak","Irlanti","Islamilaiset juhlapyhät (shiia)","Islamilaiset juhlapyhät (sunni)","Islanti","Israel","Italia","Itävalta","Jamaika","Japani","Jemen","Jordania","Juutalaiset pyhäpäivät","Kanada","Kazakstan","Kenia","Kiina","Kiinan kansantasavallan erityishallintoalue Hongkong","Kolumbia","Kongo (demokraattinen tasavalta)","Korea","Kreikka","Kristilliset pyhäpäivät","Kroatia","Kuwait","Kypros","Latvia","Libanon","Liechtenstein","Liettua","Luxemburg","Macaon erityishallintoalue","Malesia","Malta","Marokko","Meksiko","Moldova","Monaco","Montenegro","Nicaragua","Nigeria","Norja","Oman","Päiväntasaajan Guinea","Pakistan","Panama","Paraguay","Peru","Pohjois-Makedonia","Portugali","Puerto Rico","Puola","Pyhä istuin (Vatikaanivaltio)","Qatar","Ranska","Romania","Ruotsi","Saksa","San Marino","Saudi-Arabia","Serbia","Singapore","Slovakia","Slovenia","Suomi","Sveitsi","Syyria","Tanska","Thaimaa","Trinidad ja Tobago","Tšekki","Tunisia","Turkki","Ukraina","Unkari","Uruguay","Uusi-Seelanti","Valko-Venäjä","Venäjä","Venezuela","Vietnam","Viro","Yhdistyneet arabiemiirikunnat","Yhdistynyt kuningaskunta","Yhdysvallat"))
$Cultures.Add("French (France)", @("Afrique du Sud","Albanie","Algérie","Allemagne","Andorre","Angola","Arabie saoudite","Argentine","Arménie","Australie","Autriche","Azerbaïdjan","Bahreïn","Bélarus","Belgique","Bolivie","Bosnie-Herzégovine","Brésil","Brunei","Bulgarie","Canada","Chili","Chine","Chypre","Colombie","Congo (République démocratique du)","Corée du Sud","Costa Rica","Croatie","Danemark","Égypte","Émirats arabes unis","Équateur","Espagne","Estonie","États-Unis","Fêtes religieuses chrétiennes","Fêtes religieuses islamiques (shia)","Fêtes religieuses islamiques (sunni)","Fêtes religieuses juives","Finlande","France","Géorgie","Grèce","Guatemala","Guinée équatoriale","Honduras","Hong Kong RAS","Hongrie","Inde","Irak","Irlande","Islande","Israël","Italie","Jamaïque","Japon","Jordanie","Kazakhstan","Kenya","Koweït","Lettonie","Liban","Liechtenstein","Lituanie","Luxembourg","Macao (R.A.S.)","Macédoine du Nord","Malaisie","Malte","Maroc","Mexique","Moldova","Monaco","Monténégro","Nicaragua","Nigeria","Norvège","Nouvelle-Zélande","Oman","Pakistan","Panama","Paraguay","Pays-Bas","Pérou","Philippines","Pologne","Portugal","Puerto Rico","Qatar","République Dominicaine","Roumanie","Royaume-Uni","Russie","Saint-Marin","Saint-Siège (Cité du Vatican)","Salvador","Serbie","Singapour","Slovaquie","Slovénie","Suède","Suisse","Syrie","Tchéquie","Thaïlande","Trinité-et-Tobago","Tunisie","Turquie","Ukraine","Union européenne","Uruguay","Venezuela","Vietnam","Yémen"))
$Cultures.Add("German (Germany)", @("Ägypten","Albanien","Algerien","Andorra","Angola","Äquatorialguinea","Argentinien","Armenien","Aserbaidschan","Australien","Bahrain","Belarus","Belgien","Bolivien","Bosnien und Herzegowina","Brasilien","Brunei Darussalam","Bulgarien","Chile","China","Christliche religiöse Feiertage","Costa Rica","Dänemark","Deutschland","Dominikanische Republik","Ecuador","El Salvador","Estland","Europäische Union","Finnland","Frankreich","Georgien","Griechenland","Guatemala","Heiliger Stuhl (Vatikanstadt)","Honduras","Indien","Irak","Irland","Islamische religiöse Feiertage (schiitisch)","Islamische religiöse Feiertage (sunnitisch)","Island","Israel","Italien","Jamaika","Japan","Jemen","Jordanien","Jüdische religiöse Feiertage","Kanada","Kasachstan","Katar","Kenia","Kolumbien","Kongo (Demokratische Republik)","Kroatien","Kuwait","Lettland","Libanon","Liechtenstein","Litauen","Luxemburg","Malaysia","Malta","Marokko","Mexiko","Monaco","Montenegro","Neuseeland","Nicaragua","Niederlande","Nigeria","Nordmazedonien","Norwegen","Oman","Österreich","Pakistan","Panama","Paraguay","Peru","Philippinen","Polen","Portugal","Puerto Rico","Republik Korea","Republik Moldau","Rumänien","Russische Föderation","San Marino","Saudi-Arabien","Schweden","Schweiz","Serbien","Singapur","Slowakei","Slowenien","Sonderverwaltungsregion Hongkong","Sonderverwaltungsregion Macau","Spanien","Südafrika","Syrien","Thailand","Trinidad und Tobago","Tschechien","Tunesien","Türkei","Ukraine","Ungarn","Uruguay","Venezuela","Vereinigte Arabische Emirate","Vereinigte Staaten","Vereinigtes Königreich","Vietnam","Zypern"))
$Cultures.Add("Greek (Greece)", @("Αγία Έδρα (Πόλη του Βατικανού)","Άγιος Μαρίνος","Αγκόλα","Αζερμπαϊτζάν","Αίγυπτος","Αλβανία","Αλγερία","Ανδόρα","Αργεντινή","Αρμενία","Αυστραλία","Αυστρία","Βέλγιο","Βενεζουέλα","Βιετνάμ","Βολιβία","Βόρεια Μακεδονία","Βοσνία και Ερζεγοβίνη","Βουλγαρία","Βραζιλία","Γαλλία","Γερμανία","Γεωργία","Γουατεμάλα","Δανία","Δομινικανή Δημοκρατία","Εβραϊκές θρησκευτικές εορτές","Ελ Σαλβαδόρ","Ελβετία","Ελλάδα","Εσθονία","Ευρωπαϊκή Ένωση","Ηνωμένα Αραβικά Εμιράτα","Ηνωμένες Πολιτείες","Ηνωμένο Βασίλειο","Ιαπωνία","Ινδία","Ιορδανία","Ιράκ","Ιρλανδία","Ισημερινή Γουινέα","Ισημερινός","Ισλαμικές θρησκευτικές αργίες (Σιίτες)","Ισλαμικές θρησκευτικές αργίες (Σουνίτες)","Ισλανδία","Ισπανία","Ισραήλ","Ιταλία","Καζαχστάν","Καναδάς","Κατάρ","Κάτω Χώρες","Κένυα","Κίνα","Κολομβία","Κονγκό (Λαϊκή Δημοκρατία)","Κορέα","Κόστα Ρίκα","Κουβέιτ","Κροατία","Κύπρος","Λετονία","Λευκορωσία","Λίβανος","Λιθουανία","Λιχτενστάιν","Λουξεμβούργο","Μακάο ΕΔΠ","Μαλαισία","Μάλτα","Μαρόκο","Μαυροβούνιο","Μεξικό","Μολδαβία","Μονακό","Μπαχρέιν","Μπρουνέι","Νέα Ζηλανδία","Νιγηρία","Νικαράγουα","Νορβηγία","Νότια Αφρική","Ομάν","Ονδούρα","Ουγγαρία","Ουκρανία","Ουρουγουάη","Πακιστάν","Παναμάς","Παραγουάη","Περού","Πολωνία","Πορτογαλία","Πουέρτο Ρίκο","Ρουμανία","Ρωσία","Σαουδική Αραβία","Σερβία","Σιγκαπούρη","Σλοβακία","Σλοβενία","Σουηδία","Συρία","Ταϊλάνδη","Τζαμάικα","Τουρκία","Τρινιντάντ και Τομπάγκο","Τσεχία","Τυνησία","Υεμένη","Φιλιππίνες","Φινλανδία","Χιλή","Χονγκ Κονγκ ΕΔΠ","Χριστιανικές θρησκευτικές εορτές"))
$Cultures.Add("Hebrew (Israel)", @("אוסטריה","אוסטרליה","אוקראינה","אורוגוואי","אזרבייג'ן","איחוד האמירויות הערביות","איטליה","איסלנד","אירלנד","אל סלבדור","אלבניה","אלג'יריה","אנגולה","אנדורה","אסטוניה","אקוודור","ארגנטינה","ארמניה","ארצות הברית","בולגריה","בוליביה","בוסניה והרצגובינה","בחריין","בלגיה","בלרוס","ברוניי","ברזיל","בריטניה","גואטמלה","גיאורגיה","גינאה המשוונית","ג'מייקה","גרמניה","דנמרק","דרום אפריקה","האיחוד האירופי","הודו","הולנד","הונג קונג S.A.R.‎","הונגריה","הונדורס","הכס הקדוש (קרית הוותיקן)","הפיליפינים","הרפובליקה הדומיניקנית","וייטנאם","ונצואלה","חגים דתיים יהודיים","חגים דתיים מוסלמיים (סוני)","חגים דתיים מוסלמיים (שיעה)","חגים דתיים נוצריים","טוניסיה","טורקיה","טרינידד וטובגו","יוון","יפן","ירדן","ישראל","כוויית","לבנון","לוקסמבורג","לטביה","ליטא","ליכטנשטיין","מולדובה","מונטנגרו","מונקו","מלזיה","מלטה","מצרים","מקאו S.A.R.‎","מקדוניה הצפונית","מקסיקו","מרוקו","נורווגיה","ניגריה","ניו זילנד","ניקרגואה","סוריה","סין","סינגפור","סלובניה","סלובקיה","סן מרינו","ספרד","סרביה","עומאן","עיראק","ערב הסעודית","פוארטו ריקו","פולין","פורטוגל","פינלנד","פנמה","פקיסטן","פרגוואי","פרו","צ'ילה","צ'כיה","צרפת","קולומביה","קונגו (הרפובליקה הדמוקרטית)","קוסטה ריקה","קוריאה","קזחסטן","קטאר","קנדה","קניה","קפריסין","קרואטיה","רומניה","רוסיה","שוודיה","שוויץ","תאילנד","תימן"))
$Cultures.Add("Hindi (India)", @("अंगोला","अंडोरा","अज़रबैजान","अर्जेंटीना","अर्मेनिया","अल सल्वाडोर","अल्जेरिया","अल्बानिया","आईसलैंड","आयरलैंड","इक्वेटोरियल गीनिया","इक्वेडोर","इज़रायल","इटली","इराक","इस्लामी (शिया) धार्मिक हॉलिडे","इस्लामी (सुन्नी) धार्मिक हॉलिडे","ईसाई धार्मिक हॉलिडे","उक्रैन","उत्तरी मकदूनिया","उरुग्वे","एस्टोनिया","ऑस्ट्रिया","ऑस्ट्रेलिया","ओमान","कजाकस्तान","कतर","कनाडा","काँगो (का लोकतांत्रिक गणराज्य)","कुवैत","केन्या","कोरिया","कोलंबिया","कोस्टा रिका","क्रोएशिया","ग्रीस","ग्वाटेमाला","चिली","चीन","चेकिया","जमैका","जर्मनी","जापान","जॉर्जिया","जॉर्डन","ट्यूनीशिया","डेन्मार्क","डोमिनिकन गणराज्य","तुर्कस्तान","त्रिनिदाद और टोबेगो","थायलंड","दक्षिण आफ़्रिका","नाइजीरिया","निकारागुआ","नीदरलैंड","नॉर्वे","न्यूज़ीलैंड","पनामा","पाकिस्तान","पुर्तगाल","पेरू","पैराग्वे","पोलैंड","प्युर्तो रिको","फ़िनलैंड","फ़िलीपीन्स","फ़्रांस","बहरीन","बुल्गारिया","बेलारूस","बेल्जियम","बोलीविया","बोस्निया और हर्ज़ेगोविना","ब्राज़ील","ब्रुनेई","भारत","मकाऊ S.A.R.","मलेशिया","माल्टा","मिस्र","मेक्सिको","मॉन्टेंगरो","मोनाको","मोरोक्को","मोल्डोवा","यमन","यहूदी धार्मिक छुट्टियाँ","युनाइटेड किंगडम","योरपीय यूनियन","रुमानिया","रूस","लक्ज़ेम्बर्ग","लाटविया","लिचेंस्टीन","लिथुआनिया","लेबनान","वियतनाम","वेनेज़ुएला","संयुक्त अरब अमारात","संयुक्त राज्य अमरीका","सऊदी अरबस्तान","सर्बिया","साइप्रस","सिंगापुर","सीरिया","सैन मारीनो","स्पेन","स्लोवाक गणतंत्र","स्लोवेनिया","स्वित्ज़र्लैंड","स्वीडन","हंगरी","हाँग काँग S.A.R.","होंडुरस","होली सी (वेटिकन सिटी)"))
$Cultures.Add("Hungarian (Hungary)", @("Albánia","Algéria","Amerikai Egyesült Államok","Andorra","Angola","Argentína","Ausztrália","Ausztria","Azerbajdzsán","Bahrein","Belarusz","Belgium","Bolívia","Bosznia-Hercegovina","Brazília","Brunei","Bulgária","Chile","Ciprus","Costa Rica","Csehország","Dánia","Dél-Afrika","Dominikai Köztársaság","Ecuador","Egyenlítői Guinea","Egyesült Arab Emírségek","Egyesült Királyság","Egyiptom","Észak-Macedónia","Észtország","Európai Unió","Finnország","Franciaország","Fülöp-szigetek","Görögország","Grúzia","Guatemala","Hollandia","Honduras","Hongkong (KKT)","Horvátország","India","Irak","Írország","Iszlám (síita) vallási ünnepek","Iszlám (szunnita) vallási ünnepek","Izland","Izrael","Jamaica","Japán","Jemen","Jordánia","Kanada","Katar","Kazahsztán","Kenya","Keresztény vallási ünnepek","Kína","Kolumbia","Kongói Demokratikus Köztársaság","Korea","Kuvait","Lengyelország","Lettország","Libanon","Liechtenstein","Litvánia","Luxemburg","Magyarország","Makaó (KKT)","Malajzia","Málta","Marokkó","Mexikó","Moldova","Monaco","Montenegró","Németország","Nicaragua","Nigéria","Norvégia","Olaszország","Omán","Örményország","Oroszország","Pakisztán","Panama","Paraguay","Peru","Portugália","Puerto Rico","Románia","Salvador","San Marino","Spanyolország","Svájc","Svédország","Szaúd-Arábia","Szentszék (Vatikánváros)","Szerbia","Szingapúr","Szíria","Szlovákia","Szlovénia","Thaiföld","Törökország","Trinidad és Tobago","Tunézia","Új-Zéland","Ukrajna","Uruguay","Venezuela","Vietnam","Zsidó vallási ünnepek"))
$Cultures.Add("Indonesian (Indonesia)", @("Afrika Selatan","Albania","Aljazair","Amerika Serikat","Andorra","Angola","Arab Saudi","Argentina","Armenia","Australia","Austria","Azerbaijan","Bahrain","Belanda","Belarus","Belgia","Bolivia","Bosnia dan Herzegovina","Brasil","Brunei","Bulgaria","Ceko","Chili","China","Daerah Administratif Khusus Macau","Denmark","Ekuador","El Salvador","Estonia","Filipina","Finlandia","Georgia","Guatemala","Guinea Ekuatorial","Hari Libur Agama Yahudi","Hari Libur Kristen","Holy See (Kota Suci Vatikan)","Honduras","Hongaria","India","Irak","Irlandia","Islandia","Israel","Italia","Jamaika","Jepang","Jerman","Kanada","Kazakhstan","Kenya","Kerajaan Inggris Bersatu","Kolombia","Kongo (Republik Demokratik)","Korea","Kosta Rika","Kroasia","Kuwait","Latvia","Lebanon","Libur Agama Islam (Sunni)","Libur Agama Islam (Syiah)","Liechtenstein","Lithuania","Luksemburg","Makedonia Utara","Malaysia","Malta","Maroko","Meksiko","Mesir","Moldova","Monako","Montenegro","Nigeria","Nikaragua","Norwegia","Oman","Pakistan","Panama","Paraguay","Perserikatan Eropa","Peru","Polandia","Portugal","Prancis","Puerto Riko","Qatar","Republik Dominika","Rumania","Rusia","S.A.R. Hong Kong","San Marino","Selandia Baru","Serbia","Singapura","Siprus","Slovenia","Slowakia","Spanyol","Suriah","Swedia","Swiss","Thailand","Trinidad dan Tobago","Tunisia","Turki","Ukraina","Uni Emirat Arab","Uruguay","Venezuela","Vietnam","Yaman","Yordania","Yunani"))
$Cultures.Add("Italian (Italy)", @("Albania","Algeria","Andorra","Angola","Arabia Saudita","Argentina","Armenia","Australia","Austria","Azerbaigian","Bahrein","Belarus","Belgio","Bolivia","Bosnia ed Erzegovina","Brasile","Brunei","Bulgaria","Canada","Cechia","Cile","Cina","Cipro","Colombia","Corea","Costa Rica","Croazia","Danimarca","Ecuador","Egitto","El Salvador","Emirati Arabi Uniti","Estonia","Filippine","Finlandia","Francia","Georgia","Germania","Giamaica","Giappone","Giordania","Grecia","Guatemala","Guinea Equatoriale","Honduras","Hong Kong R.A.S.","India","Iraq","Irlanda","Islanda","Israele","Italia","Kazakhstan","Kenya","Kuwait","Lettonia","Libano","Liechtenstein","Lituania","Lussemburgo","Macao - R.A.S.","Macedonia del Nord","Malaysia","Malta","Marocco","Messico","Moldova","Monaco","Montenegro","Nicaragua","Nigeria","Norvegia","Nuova Zelanda","Oman","Paesi Bassi","Pakistan","Panama","Paraguay","Perù","Polonia","Portogallo","Portorico","Qatar","Regno Unito","Religione cristiana","Religione ebraica","Religione islamica (sciita)","Religione islamica (sunnita)","Repubblica democratica del Congo","Repubblica dominicana","Romania","Russia","San Marino","Santa Sede (Stato della Città del Vaticano)","Serbia","Singapore","Siria","Slovacchia","Slovenia","Spagna","Stati Uniti","Sudafrica","Svezia","Svizzera","Thailandia","Trinidad e Tobago","Tunisia","Turchia","Ucraina","Ungheria","Unione Europea","Uruguay","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Japanese (Japan)", @("アイスランド","アイルランド","アゼルバイジャン","アラブ首長国連邦","アルジェリア","アルゼンチン","アルバニア","アルメニア","アンゴラ","アンドラ","イエメン","イスラエル","イスラム教 (シーア派) 祝祭日","イスラム教 (スンニ派) 祝祭日","イタリア","イラク","インド","ウクライナ","ウルグアイ","エクアドル","エジプト","エストニア","エルサルバドル","オーストラリア","オーストリア","オマーン","オランダ","カザフスタン","カタール","カナダ","キプロス","ギリシャ","キリスト教祝祭日","グアテマラ","クウェート","クロアチア","ケニア","コスタリカ","コロンビア","コンゴ民主共和国","サウジアラビア","サンマリノ","ジャマイカ","ジョージア","シリア","シンガポール","スイス","スウェーデン","スペイン","スロバキア","スロベニア","セルビア","タイ","チェコ","チュニジア","チリ","デンマーク","ドイツ","ドミニカ共和国","トリニダード・トバゴ","トルコ","ナイジェリア","ニカラグア","ニュージーランド","ノルウェー","バーレーン","パキスタン","バチカン","パナマ","パラグアイ","ハンガリー","フィリピン","フィンランド","プエルトリコ","ブラジル","フランス","ブルガリア","ブルネイ","ベトナム","ベネズエラ","ベラルーシ","ペルー","ベルギー","ポーランド","ボスニア・ヘルツェゴビナ","ボリビア","ポルトガル","ホンジュラス","マカオ特別行政区","マルタ","マレーシア","メキシコ","モナコ","モルドバ","モロッコ","モンテネグロ","ユダヤ教祝祭日","ヨルダン","ラトビア","リトアニア","リヒテンシュタイン","ルーマニア","ルクセンブルク","レバノン","ロシア","中国","北マケドニア","南アフリカ","日本","欧州連合創設記念日","米国","英国","赤道ギニア","韓国","香港特別行政区"))
$Cultures.Add("Kazakh (Kazakhstan)", @("Австралия","Австрия","Албания","Алжир","Америка Құрама Штаттары","Ангола","Андорра","Аргентина","Армения","Бахрейн","Беларусь","Бельгия","Біріккен Араб Әмірліктері","Болгария","Боливия","Босния және Герцеговина","Бразилия","Бруней","Венгрия","Венесуэла","Вьетнам","Гватемала","Германия","Гондурас","Гонконг АӘА","Грекия","Грузия","Дания","Доминикан Республикасы","Еврей ұлттық мейрамдары","Еуропалық Одақ","Әзірбайжан","Әулие тақ (Ватикан)","Жаңа Зеландия","Жапония","Йемен","Израиль","Иордания","Ирак","Ирландия","Исландия","Испания","Италия","Канада","Катар","Кения","Кипр","Колумбия","Конго (Демократиялық республикасы)","Корея","Коста-Рика","Кувейт","Қазақстан","Құрама Корольдік","Қытай","Латвия","Ливан","Литва","Лихтенштейн","Люксембург","Макао АӘА","Малайзия","Мальта","Марокко","Мексика","Молдова","Монако","Мұсылмандық (Сунна) діни мерекелер","Мұсылмандық (Шия) діни мерекелер","Мысыр","Нигерия","Нидерланд","Никарагуа","Норвегия","Оман","Оңтүстік Африка","Панама","Парагвай","Перу","Пәкістан","Польша","Португалия","Пуэрто-Рико","Ресей","Румыния","Сальвадор","Сан-Марино","Сауд Арабиясы","Сербия","Сингапур","Сирия","Словакия","Словения","Солтүстік Македония","Таиланд","Тринидад және Тобаго","Тунис","Түркия","Украина","Уругвай","Үндістан","Филиппин","Финляндия","Франция","Хорватия","Христиан діни мейрамдары","Черногория","Чехия","Чили","Швейцария","Швеция","Эквадор","Экваторлық Гвинея","Эстония","Ямайка"))
$Cultures.Add("Korean (Korea)", @("과테말라","교황청(바티칸 시국)","교회 기념일","그리스","나이지리아","남아프리카 공화국","네덜란드","노르웨이","뉴질랜드","니카라과","덴마크","도미니카 공화국","독일","라트비아","러시아","레바논","루마니아","룩셈부르크","리투아니아","리히텐슈타인","마카오 특별 행정구","말레이시아","멕시코","모나코","모로코","몬테네그로","몰도바","몰타","미국","바레인","베네수엘라","베트남","벨기에","벨로루시","보스니아 헤르체고비나","볼리비아","북마케도니아","불가리아","브라질","브루나이","사우디아라비아","산마리노","세르비아","스웨덴","스위스","스페인","슬로바키아","슬로베니아","시리아","싱가포르","아랍에미리트","아르메니아","아르헨티나","아이슬란드","아일랜드","아제르바이잔","안도라","알바니아","알제리","앙골라","에스토니아","에콰도르","엘살바도르","영국","예멘","오만","오스트레일리아","오스트리아","요르단","우루과이","우크라이나","유럽 연합","유태교 기념일","이라크","이스라엘","이슬람(수니파) 종교 휴일","이슬람(시아파) 종교 휴일","이집트","이탈리아","인도","일본","자메이카","적도 기니","조지아","중국","체코","칠레","카자흐스탄","카타르","캐나다","케냐","코스타리카","콜롬비아","콩고(민주공화국)","쿠웨이트","크로아티아","키프로스","태국","터키","튀니지","트리니다드 토바고","파나마","파라과이","파키스탄","페루","포르투갈","폴란드","푸에르토리코","프랑스","핀란드","필리핀","한국","헝가리","혼두라스","홍콩 특별 행정구"))
$Cultures.Add("Latvian (Latvia)", @("Albānija","Alžīrija","Amerikas Savienotās Valstis","Andora","Angola","Apvienotā Karaliste","Apvienotie Arābu Emirāti","Argentīna","Armēnija","Austrālija","Austrija","Azerbaidžāna","Bahreina","Baltkrievija","Beļģija","Bolīvija","Bosnija un Hercegovina","Brazīlija","Bruneja","Bulgārija","Čehija","Čīle","Dānija","Dienvidāfrika","Dominikānas Republika","Ebreju reliģiskie svētki","Ēģipte","Eiropas Savienība","Ekvadora","Ekvatoriālā Gvineja","Filipīnas","Francija","Grieķija","Gruzija","Gvatemala","Hondurasa","Horvātija","Igaunija","Indija","Īpašais administratīvais reģions Honkonga","Īpašais administratīvais reģions Makao","Irāka","Īrija","Islāma (šiisms) reliģiskie svētki","Islāmistu (sunnisms) reliģiskie svētki","Islande","Itālija","Izraēla","Jamaika","Japāna","Jaunzēlande","Jemena","Jordānija","Kanāda","Katara","Kazahstāna","Kenija","Ķīna","Kipra","Kolumbija","Kongo Demokrātiskā Republika","Koreja","Kostarika","Krievija","Kristiešu reliģiskie svētki","Kuveita","Latvija","Libāna","Lietuva","Lihtenšteina","Luksemburga","Malaizija","Malta","Maroka","Meksika","Melnkalne","Moldova","Monako","Nīderlande","Nigērija","Nikaragva","Norvēģija","Omāna","Pakistāna","Panama","Paragvaja","Peru","Polija","Portugāle","Puertoriko","Rumānija","Salvadora","Sanmarīno","Saūda Arābija","Serbija","Singapūra","Sīrija","Slovākija","Slovēnija","Somija","Spānija","Šveice","Svētais Krēsls (Vatikāns)","Taizeme","Trinidāda un Tobāgo","Tunisija","Turcija","Ukraina","Ungārija","Urugvaja","Vācija","Venecuēla","Vjetnama","Ziemeļmaķedonija","Zviedrija"))
$Cultures.Add("Lithuanian (Lithuania)", @("Airija","Albanija","Alžyras","Andora","Angola","Argentina","Armėnija","Australija","Austrija","Azerbaidžanas","Bahreinas","Baltarusija","Belgija","Bolivija","Bosnija ir Hercegovina","Brazilija","Brunėjus","Bulgarija","Čekija","Čilė","Danija","Dominikos Respublika","Egiptas","Ekvadoras","Estija","Europos Sąjunga","Filipinai","Graikija","Gruzija","Gvatemala","Hondūras","Indija","Irakas","Islamo (Shia) religijos šventės","Islamo (Sunni) religijos šventės","Islandija","Ispanija","Italija","Izraelis","Jamaika","Japonija","Jemenas","Jordanija","Jungtinė Karalystė","Jungtinės Valstijos","Jungtiniai Arabų Emyratai","Juodkalnija","Kanada","Kataras","Kazachija","Kenija","Kinija","Kipras","Kolumbija","Kongas (Demokratinė Respublika)","Korėja","Kosta Rika","Krikščionių religinės šventės","Kroatija","Kuveitas","Latvija","Lenkija","Libanas","Lichtenšteinas","Lietuva","Liuksemburgas","Malaizija","Malta","Marokas","Meksika","Moldova","Monakas","Naujoji Zelandija","Nigerija","Nikaragva","Norvegija","Nyderlandai","Omanas","Pakistanas","Panama","Paragvajus","Peru","Pietų Afrika","Portugalija","Prancūzija","Puerto Rikas","Pusiaujo Gvinėja","Rumunija","Rusija","Salvadoras","San Marinas","Saudo Arabija","Serbija","Šiaurės Makedonija","Singapūras","Sirija","Slovakija","Slovėnija","Suomija","Švedija","Šveicarija","Šventasis Sostas (Vatikano Miesto Valstybė)","Tailandas","Trinidadas ir Tobagas","Tunisas","Turkija","Ukraina","Urugvajus","Venesuela","Vengrija","Vietnamas","Vokietija","YAKR Honkongas","YAKR Makao","Žydų religinės šventės"))
$Cultures.Add("Malay (Malaysia)", @("Afrika Selatan","Albania","Algeria","Amerika Syarikat","Andorra","Angola","Arab Saudi","Argentina","Armenia","Australia","Austria","Azerbaijan","Bahrain","Belanda","Belarus","Belgium","Bolivia","Bosnia dan Herzegovina","Brazil","Brunei","Bulgaria","Chile","China","Colombia","Congo (Republik Demokratik)","Costa Rica","Croatia","Cuti Keagamaan Kristian","Cuti-cuti Keagamaan Yahudi","Cyprus","Czechia","Denmark","Ecuador","El Salvador","Emiriyah Arab Bersatu","Equatorial Guinea","Estonia","Filipina","Finland","Georgia","Greece","Guatemala","Hari Keagamaan Islam (Shia)","Hari Keagamaan Islam (Sunni)","Holy See (Kota Vatican)","Honduras","Hungary","Iceland","India","Iraq","Ireland","Israel","Itali","Jamaica","Jepun","Jerman","Jordan","Kanada","Kazakhstan","Kenya","Kesatuan Eropah","Korea","Kuwait","Latvia","Liechtenstein","Lithuania","Lubnan","Luxembourg","Macau S.A.R.","Macedonia Utara","Maghribi","Malaysia","Malta","Mesir","Mexico","Moldova","Monaco","Montenegro","New Zealand","Nicaragua","Nigeria","Norway","Oman","Pakistan","Panama","Paraguay","Perancis","Peru","Poland","Portugal","Puerto Rico","Qatar","Republik Dominica","Romania","Rusia","S.A.R Hong Kong","San Marino","Sepanyol","Serbia","Singapura","Slovakia","Slovenia","Sweden","Switzerland","Syria","Thailand","Trinidad dan Tobago","Tunisia","Turki","Ukraine","United Kingdom","Uruguay","Venezuela","Vietnam","Yaman"))
$Cultures.Add("Norwegian, Bokmål (Norway)", @("Albania","Algerie","Andorra","Angola","Argentina","Armenia","Aserbajdsjan","Australia","Bahrain","Belgia","Bolivia","Bosnia-Hercegovina","Brasil","Brunei","Bulgaria","Canada","Chile","Colombia","Costa Rica","Danmark","De forente arabiske emirater","Den demokratiske republikken Kongo","Den dominikanske republikken","Den europeiske union","Ecuador","Egypt","Ekvatorial-Guinea","El Salvador","Estland","Filippinene","Finland","Frankrike","Georgia","Guatemala","Hellas","Honduras","Hongkong SAR","Hviterussland","India","Irak","Irland","Island","Israel","Italia","Jamaica","Japan","Jemen","Jødiske helligdager","Jordan","Kasakhstan","Kenya","Kina","Kristne helligdager","Kroatia","Kuwait","Kypros","Latvia","Libanon","Liechtenstein","Litauen","Luxemburg","Macao SAR","Malaysia","Malta","Marokko","Mexico","Moldova","Monaco","Montenegro","Muslimske religiøse helligdagar (shia)","Muslimske religiøse helligdager (sunni)","Nederland","New Zealand","Nicaragua","Nigeria","Nord-Makedonia","Norge","Oman","Østerrike","Pakistan","Panama","Paraguay","Peru","Polen","Portugal","Puerto Rico","Qatar","Romania","Russland","San Marino","Saudi-Arabia","Serbia","Singapore","Slovakia","Slovenia","Sør-Afrika","Sør-Korea","Spania","Storbritannia","Sveits","Sverige","Syria","Thailand","Trinidad og Tobago","Tsjekkia","Tunisia","Tyrkia","Tyskland","Ukraina","Ungarn","Uruguay","USA","Vatikanstaten","Venezuela","Vietnam"))
$Cultures.Add("Polish (Poland)", @("Albania","Algieria","Andora","Angola","Arabia Saudyjska","Argentyna","Armenia","Australia","Austria","Azerbejdżan","Bahrajn","Belgia","Białoruś","Boliwia","Bośnia i Hercegowina","Brazylia","Brunei","Bułgaria","Chile","Chiny","Chorwacja","Chrześcijańskie święta religijne","Cypr","Czarnogóra","Czechy","Dania","Dominikana","Egipt","Ekwador","Estonia","Filipiny","Finlandia","Francja","Grecja","Gruzja","Gwatemala","Gwinea Równikowa","Hiszpania","Holandia","Honduras","Hongkong SAR","Indie","Irak","Irlandia","Islamskie święta religijne (sunnizm)","Islamskie święta religijne (szyizm)","Islandia","Izrael","Jamajka","Japonia","Jemen","Jordania","Kanada","Katar","Kazachstan","Kenia","Kolumbia","Kongo (Demokratyczna Republika)","Korea Południowa","Kostaryka","Kuwejt","Liban","Liechtenstein","Litwa","Łotwa","Luksemburg","Macedonia Północna","Makau SAR","Malezja","Malta","Maroko","Meksyk","Mołdawia","Monako","Niemcy","Nigeria","Nikaragua","Norwegia","Nowa Zelandia","Oman","Pakistan","Panama","Paragwaj","Peru","Polska","Portoryko","Portugalia","Rosja","RPA","Rumunia","Salwador","San Marino","Serbia","Singapur","Słowacja","Słowenia","Stany Zjednoczone","Stolica Apostolska (Państwo Watykańskie)","Syria","Szwajcaria","Szwecja","Tajlandia","Trynidad i Tobago","Tunezja","Turcja","Ukraina","Unia Europejska","Urugwaj","Węgry","Wenezuela","Wietnam","Włochy","Zjednoczone Emiraty Arabskie","Zjednoczone Królestwo","Żydowskie święta religijne"))
$Cultures.Add("Portuguese (Brazil)", @("África do Sul","Albânia","Alemanha","Andorra","Angola","Arábia Saudita","Argélia","Argentina","Armênia","Austrália","Áustria","Azerbaijão","Bahrein","Belarus","Bélgica","Bolívia","Bósnia e Herzegovina","Brasil","Brunei","Bulgária","Canadá","Catar","Cazaquistão","Chile","China","Chipre","Cingapura","Colômbia","Congo (República Democrática do)","Coreia","Costa Rica","Croácia","Czechia","Dinamarca","Egito","El Salvador","Emirados Árabes Unidos","Equador","Eslováquia","Eslovênia","Espanha","Estados Unidos","Estônia","Feriados Cristãos","Feriados Judaicos","Feriados Religiosos Islâmicos (Shia)","Feriados Religiosos Islâmicos (Sunni)","Filipinas","Finlândia","França","Geórgia","Grécia","Guatemala","Guiné Equatorial","Honduras","Hungria","Iêmen","Índia","Iraque","Irlanda","Islândia","Israel","Itália","Jamaica","Japão","Jordânia","Kuwait","Letônia","Líbano","Liechtenstein","Lituânia","Luxemburgo","Macedônia do Norte","Malásia","Malta","Marrocos","México","Moldova","Mônaco","Montenegro","Nicarágua","Nigéria","Noruega","Nova Zelândia","Omã","Países Baixos","Panamá","Paquistão","Paraguai","Peru","Polônia","Porto Rico","Portugal","Quênia","RAE de Hong Kong","RAE de Macau","Reino Unido","República Dominicana","Romênia","Rússia","San Marino","Santa Sé (Cidade do Vaticano)","Sérvia","Síria","Suécia","Suíça","Tailândia","Trinidad e Tobago","Tunísia","Turquia","Ucrânia","União Europeia","Uruguai","Venezuela","Vietnã"))
$Cultures.Add("Portuguese (Portugal)", @("África do Sul","Albânia","Alemanha","Andorra","Angola","Arábia Saudita","Argélia","Argentina","Arménia","Austrália","Áustria","Azerbaijão","Barém","Bélgica","Bielorrússia","Bolívia","Bósnia e Herzegovina","Brasil","Brunei","Bulgária","Canadá","Catar","Cazaquistão","Chéquia","Chile","China","Chipre","Colômbia","Congo (República Democrática do)","Coreia","Costa Rica","Croácia","Dinamarca","Egito","Emirados Árabes Unidos","Equador","Eslováquia","Eslovénia","Espanha","Estados Unidos","Estónia","Feriados religiosos cristãos","Feriados Religiosos Islâmicos (Sunitas)","Feriados Religiosos Islâmicos (Xiitas)","Feriados religiosos judaicos","Filipinas","Finlândia","França","Geórgia","Grécia","Guatemala","Guiné Equatorial","Honduras","Hungria","Iémen","Índia","Iraque","Irlanda","Islândia","Israel","Itália","Jamaica","Japão","Jordânia","Kuwait","Letónia","Líbano","Listenstaine","Lituânia","Luxemburgo","Macedónia do Norte","Malásia","Malta","Marrocos","México","Moldova","Mónaco","Montenegro","Nicarágua","Nigéria","Noruega","Nova Zelândia","Omã","Países Baixos","Panamá","Paquistão","Paraguai","Peru","Polónia","Porto Rico","Portugal","Quénia","RAE de Hong Kong","RAE de Macau","Reino Unido","República Dominicana","Roménia","Rússia","Salvador","Santa Sé (Cidade do Vaticano)","São Marinho","Sérvia","Singapura","Síria","Suécia","Suíça","Tailândia","Trindade e Tobago","Tunísia","Turquia","Ucrânia","União Europeia","Uruguai","Venezuela","Vietname"))
$Cultures.Add("Romanian (Romania)", @("Africa de Sud","Albania","Algeria","Andorra","Angola","Arabia Saudită","Argentina","Armenia","Australia","Austria","Azerbaidjan","Bahrain","Belarus","Belgia","Bolivia","Bosnia și Herțegovina","Brazilia","Brunei","Bulgaria","Canada","Cehia","Chile","China","Cipru","Columbia","Congo (Republica Democrată)","Coreea","Costa Rica","Croația","Danemarca","Ecuador","Egipt","El Salvador","Elveția","Emiratele Arabe Unite","Estonia","Filipine","Finlanda","Franța","Georgia","Germania","Grecia","Guatemala","Guineea Ecuatorială","Honduras","India","Iordania","Irak","Irlanda","Islanda","Israel","Italia","Jamaica","Japonia","Kazahstan","Kenya","Kuweit","Letonia","Liban","Liechtenstein","Lituania","Luxemburg","Macedonia de Nord","Malaysia","Malta","Maroc","Mexic","Moldova","Monaco","Muntenegru","Nicaragua","Nigeria","Norvegia","Noua Zeelandă","Oman","Pakistan","Panama","Paraguay","Peru","Polonia","Portugalia","Puerto Rico","Qatar","RAS Hong Kong","RAS Macao","Regatul Unit","Republica Dominicană","România","Rusia","San Marino","Sărbătorile religioase creștine","Sărbătorile religioase evreiești","Sărbătorile religioase islamice (Șiite)","Sărbătorile religioase islamice (Sunnite)","Serbia","Sfântul Scaun (Statul Cetății Vaticanului)","Singapore","Siria","Slovacia","Slovenia","Spania","Statele Unite ale Americii","Suedia","Țările de Jos","Thailanda","Trinidad Tobago","Tunisia","Turcia","Ucraina","Ungaria","Uniunea Europeană","Uruguay","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Russian (Russia)", @("Австралия","Австрия","Азербайджан","Албания","Алжир","Ангола","Андорра","Аргентина","Армения","Бахрейн","Беларусь","Бельгия","Болгария","Боливия","Босния и Герцеговина","Бразилия","Бруней-Даруссалам","Венгрия","Венесуэла","Вьетнам","Гватемала","Германия","Гондурас","Гонконг (САР)","Греция","Грузия","Дания","Демократическая Республика Конго","Доминиканская Республика","Еврейские религиозные праздники","Египет","ЕС","Йемен","Израиль","Индия","Иордания","Ирак","Ирландия","Исламские (суннитские) религиозные праздники","Исламские (шиитские) религиозные праздники","Исландия","Испания","Италия","Казахстан","Канада","Катар","Кения","Кипр","Китай","Колумбия","Корея","Коста-Рика","Кувейт","Латвия","Ливан","Литва","Лихтенштейн","Люксембург","Малайзия","Мальта","Марокко","Мексика","Молдова","Монако","Нигерия","Нидерланды","Никарагуа","Новая Зеландия","Норвегия","ОАЭ","Оман","Пакистан","Панама","Папский Престол (Ватикан)","Парагвай","Перу","Польша","Португалия","Пуэрто-Рико","Россия","Румыния","Сан-Марино","САР Макао","Саудовская Аравия","Северная Македония","Сербия","Сингапур","Сирийская Арабская Республика","Словакия","Словения","Соединенное Королевство","США","Таиланд","Тринидад и Тобаго","Тунис","Турция","Украина","Уругвай","Филиппины","Финляндия","Франция","Хорватия","Христианские религиозные праздники","Черногория","Чехия","Чили","Швейцария","Швеция","Эквадор","Экваториальная Гвинея","Эль-Сальвадор","Эстония","Южная Африка","Ямайка","Япония"))
$Cultures.Add("Serbian (Latin, Serbia and Montenegro (Former))", @("Albania","Algeria","Andorra","Angola","Argentina","Armenia","Australia","Austria","Azerbaijan","Bahrain","Belarus","Belgium","Bolivia","Bosnia and Herzegovina","Brazil","Brunei","Bulgaria","Canada","Chile","China","Christian Religious Holidays","Colombia","Congo (Democratic Republic of)","Costa Rica","Croatia","Cyprus","Czechia","Denmark","Dominican Republic","Ecuador","Egypt","El Salvador","Equatorial Guinea","Estonia","European Union","Finland","France","Georgia","Germany","Greece","Guatemala","Holy See (Vatican City)","Honduras","Hong Kong S.A.R.","Hungary","Iceland","India","Iraq","Ireland","Islamic (Shia) Religious Holidays","Islamic (Sunni) Religious Holidays","Israel","Italy","Jamaica","Japan","Jewish Religious Holidays","Jordan","Kazakhstan","Kenya","Korea","Kuwait","Latvia","Lebanon","Liechtenstein","Lithuania","Luxembourg","Macao S.A.R.","Malaysia","Malta","Mexico","Moldova","Monaco","Montenegro","Morocco","Netherlands","New Zealand","Nicaragua","Nigeria","North Macedonia","Norway","Oman","Pakistan","Panama","Paraguay","Peru","Philippines","Poland","Portugal","Puerto Rico","Qatar","Romania","Russia","San Marino","Saudi Arabia","Serbia","Singapore","Slovakia","Slovenia","South Africa","Spain","Sweden","Switzerland","Syria","Thailand","Trinidad and Tobago","Tunisia","Turkey","Ukraine","United Arab Emirates","United Kingdom","United States","Uruguay","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Serbian (Latin, Serbia)", @("Albanija","Alžir","Andora","Angola","Argentina","Australija","Austrija","Azerbejdžan","Bahrein","Belgija","Belorusija","Bolivija","Bosna i Hercegovina","Brazil","Brunej","Bugarska","Češka","Čile","Crna Gora","Danska","Dominikanska Republika","Egipat","Ekvador","Ekvatorijalna Gvineja","Estonija","Evropska unija","Filipini","Finska","Francuska","Grad-država Vatikan","Grčka","Gruzija","Gvatemala","Holandija","Honduras","Hrišćanski verski praznici","Hrvatska","Indija","Irak","Irska","Islamski (šiitski) religijski praznici","Islamski (sunitski) religijski praznici","Island","Italija","Izrael","Jamajka","Japan","Jemen","Jermenija","Jevrejski verski praznici","Jordan","Južna Afrika","Kanada","Katar","Kazahstan","Kenija","Kina","Kipar","Kolumbija","Kongo (Demokratska Republika)","Koreja","Kostarika","Kuvajt","Letonija","Liban","Lihtenštajn","Litvanija","Luksemburg","Mađarska","Makao S.A.O.","Malezija","Malta","Maroko","Meksiko","Moldavija","Monako","Nemačka","Nigerija","Nikaragva","Norveška","Novi Zeland","Oman","Pakistan","Panama","Paragvaj","Peru","Poljska","Portoriko","Portugalija","Rumunija","Rusija","SAD","Salvador","San Marino","SAO Hongkong","Saudijska Arabija","Severna Makedonija","Singapur","Sirija","Slovačka","Slovenija","Španija","Srbija","Švajcarska","Švedska","Tajland","Trinidad i Tobago","Tunis","Turska","Ujedinjeni Arapski Emirati","Ujedinjeno Kraljevstvo","Ukrajina","Urugvaj","Venecuela","Vijetnam"))
$Cultures.Add("Slovak (Slovakia)", @("Albánsko","Alžírsko","Andorra","Angola","Argentína","Arménsko","Austrália","Azerbajdžan","Bahrajn","Belgicko","Bielorusko","Bolívia","Bosna a Hercegovina","Brazília","Brunej","Bulharsko","Česko","Chorvátsko","Čierna Hora","Čile","Čína","Cyprus","Dánsko","Dominikánska republika","Egypt","Ekvádor","Estónsko","Európska únia","Filipíny","Fínsko","Francúzsko","Grécko","Gruzínsko","Guatemala","Holandsko","Honduras","Hongkong OAO","India","Irak","Írsko","Islamské cirkevné sviatky (sunitské)","Islamské náboženské sviatky (šiítske)","Island","Izrael","Jamajka","Japonsko","Jemen","Jordánsko","Južná Afrika","Kanada","Katar","Kazachstan","Keňa","Kolumbia","Kongo (Konžská demokratická republika)","Kórejská republika","Kostarika","Kresťanské náboženské sviatky","Kuvajt","Libanon","Lichtenštajnsko","Litva","Lotyšsko","Luxembursko","Maďarsko","Makao OAO","Malajzia","Malta","Maroko","Mexiko","Moldavsko","Monako","Nemecko","Nigéria","Nikaragua","Nórsko","Nový Zéland","Omán","Pakistan","Panama","Paraguaj","Peru","Poľsko","Portoriko","Portugalsko","Rakúsko","Rovníková Guinea","Rumunsko","Rusko","Salvádor","San Maríno","Saudská Arábia","﻿Severné Macedónsko","Singapur","Slovensko","Slovinsko","Španielsko","Spojené arabské emiráty","Spojené kráľovstvo","Spojené štáty","Srbsko","Švajčiarsko","Svätá stolica (Vatikán)","Švédsko","Sýria","Taliansko","Thajsko","Trinidad a Tobago","Tunisko","Turecko","Ukrajina","Uruguaj","Venezuela","Vietnam","Židovské cirkevné sviatky"))
$Cultures.Add("Slovenian (Slovenia)", @("Albanija","Alžirija","Andora","Angola","Argentina","Armenija","Avstralija","Avstrija","Azerbajdžan","Bahrajn","Belgija","Belorusija","Bolgarija","Bolivija","Bosna in Hercegovina","Brazilija","Brunej","Češka","Čile","Ciper","Črna gora","Danska","Dominikanska republika","Egipt","Ekvador","Ekvatorialna Gvineja","Estonija","Evropska unija","Filipini","Finska","Francija","Grčija","Gruzija","Gvatemala","Honduras","Hongkong posebna administrativna regija","Hrvaška","Indija","Irak","Irska","Islamski (šiitski) verski prazniki","Islamski (sunitski) verski prazniki","Islandija","Italija","Izrael","Jamajka","Japonska","Jemen","Jordanija","Judovski verski prazniki","Južna Afrika","Južna Koreja","Kanada","Katar","Katoliški verski prazniki","Kazahstan","Kenija","Kitajska","Kolumbija","Kongo (demokratična republika)","Kostarika","Kuvajt","Latvija","Libanon","Lihtenštajn","Litva","Luksemburg","Madžarska","Malezija","Malta","Maroko","Mehika","Moldavija","Monako","Nemčija","Nigerija","Nikaragva","Nizozemska","Norveška","Nova Zelandija","Oman","Pakistan","Panama","Paragvaj","Peru","Poljska","Portoriko","Portugalska","Posebno administativno območje Macao","Romunija","Rusija","Salvador","San Marino","Saudova Arabija","Severna Makedonija","Singapur","Sirija","Slovaška","Slovenija","Španija","Srbija","Švedska","Sveti sedež (Vatikan)","Švica","Tajska","Trinidad in Tobago","Tunizija","Turčija","Ukrajina","Urugvaj","Venezuela","Vietnam","Združene države","Združeni arabski emirati","Združeno kraljestvo"))
$Cultures.Add("Spanish (Spain)", @("Albania","Alemania","Andorra","Angola","Arabia Saudí","Argelia","Argentina","Armenia","Australia","Austria","Azerbaiyán","Baréin","Belarús","Bélgica","Bolivia","Bosnia y Herzegovina","Brasil","Brunéi","Bulgaria","Canadá","Chequia","Chile","China","Chipre","Colombia","Congo (República Democrática del)","Corea del Sur","Costa Rica","Croacia","Dinamarca","Ecuador","Egipto","El Salvador","Emiratos Árabes Unidos","Eslovaquia","Eslovenia","España","Estados Unidos","Estonia","Festividades religiosas cristianas","Festividades religiosas islámicas (chiíes)","Festividades religiosas islámicas (suní)","Festividades religiosas judías","Filipinas","Finlandia","Francia","Georgia","Grecia","Guatemala","Guinea Ecuatorial","Honduras","Hong Kong (RAE)","Hungría","India","Irak","Irlanda","Islandia","Israel","Italia","Jamaica","Japón","Jordania","Kazajistán","Kenia","Kuwait","Letonia","Líbano","Liechtenstein","Lituania","Luxemburgo","Macao RAE","Macedonia del Norte","Malasia","Malta","Marruecos","México","Moldova","Mónaco","Montenegro","Nicaragua","Nigeria","Noruega","Nueva Zelanda","Omán","Países Bajos","Pakistán","Panamá","Paraguay","Perú","Polonia","Portugal","Puerto Rico","Qatar","Reino Unido","República Dominicana","Rumania","Rusia","San Marino","Santa Sede (Ciudad del Vaticano)","Serbia","Singapur","Siria","Sudáfrica","Suecia","Suiza","Tailandia","Trinidad y Tobago","Túnez","Turquía","Ucrania","Unión Europea","Uruguay","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Swedish (Sweden)", @("Albanien","Algeriet","Andorra","Angola","Argentina","Armenien","Australien","Azerbajdzjan","Bahrain","Belgien","Bolivia","Bosnien och Hercegovina","Brasilien","Brunei","Bulgarien","Chile","Colombia","Costa Rica","Cypern","Czechia","Danmark","Dominikanska republiken","Ecuador","Egypten","Ekvatorialguinea","El Salvador","Estland","Europeiska unionen","Filippinerna","Finland","Förenade Arabemiraten","Frankrike","Georgien","Grekland","Guatemala","Heliga stolen (Vatikanstaten)","Honduras","Hongkong","Indien","Irak","Irland","Islamiska (shia) religiösa helgdagar","Islamiska (sunni) religiösa helgdagar","Island","Israel","Italien","Jamaica","Japan","Jemen","Jordanien","Judiska helgdagar","Kanada","Kazakstan","Kenya","Kina","Kongo (Demokratiska republiken)","Kristna helgdagar","Kroatien","Kuwait","Lettland","Libanon","Liechtenstein","Litauen","Luxemburg","Macao","Malaysia","Malta","Marocko","Mexiko","Moldavien","Monaco","Montenegro","Nederländerna","Nicaragua","Nigeria","Nordmakedonien","Norge","Nya Zeeland","Oman","Österrike","Pakistan","Panama","Paraguay","Peru","Polen","Portugal","Puerto Rico","Qatar","Rumänien","Ryssland","San Marino","Saudiarabien","Schweiz","Serbien","Singapore","Slovakien","Slovenien","Spanien","Storbritannien","Sverige","Sydafrika","Sydkorea","Syrien","Thailand","Trinidad och Tobago","Tunisien","Turkiet","Tyskland","Ukraina","Ungern","Uruguay","USA","Venezuela","Vietnam","Vitryssland"))
$Cultures.Add("Thai (Thailand)", @("เกาหลี","เขตบริหารพิเศษมาเก๊า","เขตบริหารพิเศษฮ่องกง","เคนยา","เช็ก","เซอร์เบีย","เดนมาร์ก","เนเธอร์แลนด์","เบลเยียม","เบลารุส","เปรู","เปอร์โตริโก","เม็กซิโก","เยเมน","เยอรมนี","เลบานอน","เวเนซุเอลา","เวียดนาม","เอกวาดอร์","เอลซัลวาดอร์","เอสโตเนีย","แคนาดา","แองโกลา","แอฟริกาใต้","แอลเบเนีย","แอลจีเรีย","โครเอเชีย","โคลัมเบีย","โบลิเวีย","โปแลนด์","โปรตุเกส","โมนาโก","โมร็อกโก","โรมาเนีย","โอมาน","ไซปรัส","ไทย","ไนจีเรีย","ไอซ์แลนด์","ไอร์แลนด์","กรีซ","กัวเตมาลา","กาตาร์","คองโก (สาธารณรัฐประชาธิปไตย)","คอสตาริกา","คาซัคสถาน","คูเวต","จอร์เจีย","จอร์แดน","จาเมกา","จีน","ชิลี","ซานมารีโน","ซาอุดีอาระเบีย","ซีเรีย","ญี่ปุ่น","ตรินิแดดและโตเบโก","ตุรกี","ตูนิเซีย","นอร์เวย์","นิการากัว","นิวซีแลนด์","บราซิล","บรูไน","บอสเนียและเฮอร์เซโกวีนา","บัลแกเรีย","บาห์เรน","ปากีสถาน","ปานามา","ปารากวัย","ฝรั่งเศส","ฟินแลนด์","ฟิลิปปินส์","มอนเตเนโกร","มอลโดวา","มอลตา","มาเลเซีย","มาซิโดเนียเหนือ","ยูเครน","รัสเซีย","ลักเซมเบิร์ก","ลัตเวีย","ลิกเตนสไตน์","ลิทัวเนีย","วันหยุดของศาสนาคริสต์","วันหยุดของศาสนายิว","วันหยุดของศาสนาอิสลาม (ชีอะฮ์)","วันหยุดของศาสนาอิสลาม (ชุนนี)","สเปน","สโลวะเกีย","สโลวีเนีย","สวิตเซอร์แลนด์","สวีเดน","สหภาพยุโรป","สหรัฐอเมริกา","สหรัฐอาหรับเอมิเรตส์","สหราชอาณาจักร","สันตะสำนัก (นครวาติกัน)","สาธารณรัฐโดมินิกัน","สิงคโปร์","ออสเตรเลีย","ออสเตรีย","อันดอร์รา","อาเซอร์ไบจาน","อาร์เจนตินา","อาร์เมเนีย","อิเควทอเรียลกินี","อิตาลี","อินเดีย","อิรัก","อิสราเอล","อียิปต์","อุรุกวัย","ฮอนดูรัส","ฮังการี"))
$Cultures.Add("Turkish (Turkey)", @("Almanya","Andorra","Angola","Arjantin","Arnavutluk","Avrupa Birliği","Avustralya","Avusturya","Azerbaycan","Bahreyn","Belçika","Beyaz Rusya","Birleşik Arap Emirlikleri","Birleşik Devletler","Birleşik Krallık","Bolivya","Bosna-Hersek","Brezilya","Brunei","Bulgaristan","Çekya","Cezayir","Çin","Danimarka","Dominik Cumhuriyeti","Ekvador","Ekvator Ginesi","El Salvador","Ermenistan","Estonya","Fas","Filipinler","Finlandiya","Fransa","Guatemala","Güney Afrika","Gürcistan","Hindistan","Hırvatistan","Hollanda","Honduras","Hong Kong Çin ÖİB","Hristiyan Dini Günleri","Irak","İrlanda","İspanya","İsrail","İsveç","İsviçre","İtalya","İzlanda","Jamaika","Japonya","Kanada","Karadağ","Katar","Kazakistan","Kenya","Kıbrıs","Kolombiya","Kongo (Demokratik Cumhuriyeti)","Kore","Kosta Rika","Kuveyt","Kuzey Makedonya","Letonya","Liechtenstein","Litvanya","Lübnan","Lüksemburg","Macaristan","Makau Çin ÖİB","Malezya","Malta","Meksika","Mısır","Moldova","Monako","Musevi Dini Günleri","Müslüman (Şii) Dini Tatilleri","Müslüman (Sünni) Dini Tatilleri","Nijerya","Nikaragua","Norveç","Pakistan","Panama","Papalık Makamı (Vatikat Şehri)","Paraguay","Peru","Polonya","Portekiz","Porto Riko","Romanya","Rusya","San Marino","Şili","Singapur","Sırbistan","Slovakya","Slovenya","Suriye","Suudi Arabistan","Tayland","Trinidad ve Tobago","Tunus","Türkiye","Ukrayna","Umman","Ürdün","Uruguay","Venezuela","Vietnam","Yemen","Yeni Zelanda","Yunanistan"))
$Cultures.Add("Ukrainian (Ukraine)", @("Австралія","Австрія","Азербайджан","Албанія","Алжир","Анґола","Андорра","Аргентина","Бахрейн","Бельґія","Білорусь","Болгарія","Болівія","Боснія та Герцеґовина","Бразілія","Бруней","В’єтнам","Венесуела","Вірменія","Гондурас","Греція","Грузія","Ґватемала","Данія","Демократична Республіка Конґо","Домініканська Республіка","Еквадор","Екваторіальна Ґвінея","Естонія","Єврейські релігійні свята","Європейський Союз","Єгипет","Ємен","Йорданія","Ізраїль","Індія","Ірак","Ірландія","Ісламські релігійні свята (суніти)","Ісламські релігійні свята (шиїти)","Ісландія","Іспанія","Італія","Казахстан","Канада","Катар","Кенія","Китай","Кіпр","Колумбія","Корея","Коста-Ріка","Кувейт","Латвія","Литва","Ліван","Ліхтенштейн","Люксембурґ","Малайзія","Мальта","Марокко","Мексика","Молдова","Монако","Ніґерія","Нідерланди","Нікараґуа","Німеччина","Нова Зеландія","Норвеґія","Об’єднані Арабські Емірати","Оман","Пакистан","Панама","Параґвай","Перу","Південна Африка","Північна Македонія","Польща","Портуґалія","Пуерто-Ріко","Росія","Румунія","Сальвадор","Сан-Маріно","САР Гонконґ","САР Макао","Саудівська Аравія","Святійший Престол (Ватикан)","Сербія","Сінґапур","Сірія","Словаччина","Словенія","Сполучене Королівство","Сполучені Штати","Таїланд","Тринідад і Тобаґо","Туніс","Туреччина","Угорщина","Україна","Уруґвай","Філіппіни","Фінляндія","Франція","Хорватія","Християнські релігійні свята","Чехія","Чілі","Чорногорія","Швейцарія","Швеція","Ямайка","Японія"))
$Cultures.Add("Vietnamese (Vietnam)", @("Ả rập Saudi","Ai Cập","Ai-len","Albania","Algeria","Ấn Độ","Andorra","Angola","Áo","Argentina","Armenia","Azerbaijan","Ba Lan","Bắc Macedonia","Ba-ranh","Belarus","Bỉ","Bồ Đào Nha","Bolivia","Bosnia và Herzegovina","Bra-zin","Brunei","Bulgaria","Các Tiểu Vương Quốc Ả Rập","Canada","Chi-lê","Colombia","Cộng hòa Dominica","Cộng hòa Sip","Congo (Cộng hòa Dân chủ)","Costa Rica","Croat-ti-a","Đặc Khu Hành chính Hồng Kông","Đặc khu Hành chính Macao","Đan Mạch","Đức","Ecuador","El Salvador","Estonia","Georgia","Guatemala","Guinea Xích đạo","Hà Lan","Hàn Quốc","Hi Lạp","Hoa Kỳ","Honduras","Hungary","Iceland","Iraq","Israel","Italy","Jamaica","Jordan","Kazakhstan","Kenya","Kuwait","Lát-vi-a","Lễ hội tôn giáo đạo Hồi (Shia)","Lễ hội tôn giáo đạo Hồi (Sunni)","Li-băng","Liechtenstein","Liên Hiệp Vương Quốc Anh","Liên Minh Châu Âu","Lít-va","Luxembourg","Malaysia","Malta","Marốc","Mexico","Moldova","Monaco","Montenegro","Na Uy","Nam Phi","New Zealand","Nga","Ngày Cơ đốc giáo","Ngày lễ đạo Do thái","Nhật Bản","Nicaragua","Nigeria","Ô man","Pakistan","Panama","Pa-ra-guay","Peru","Phần Lan","Pháp","Philippines","Puerto Rico","Qatar","Romania","San Marino","Séc","Serbia","Singapore","Slovakia","Slovenia","Syria","Tây Ban Nha","Thái Lan","Thổ Nhĩ Kỳ","Thụy Điển","Thụy Sỹ","Tòa Thánh (Thành Vatican)","Trinidad và Tobago","Trung Quốc","Tunisia","Úc","Ukraine","U-ru-guay","Venezuela","Việt Nam","Yemen"))

$cultures[$Culture]
}


$cultures = @{}
$Cultures.Add("Arabic (Saudi Arabia)", @("أذربيجان","أرمينيا","أسبانيا","أستراليا","إستونيا","الاتحاد الأوروبي","الأرجنتين","الأردن","الأعياد الدينية الإسلامية (السنّة)","الأعياد الدينية الإسلامية (الشيعة)","الأعياد الدينية المسيحية","الإكوادور","الإمارات العربية المتحدة","ألبانيا","البحرين","البرازيل","البرتغال","البوسنة والهرسك","التشيك","الجزائر","الدنمارك","السعودية","السلفادور","السويد","الصين","العراق","الفلبين","الكرسي الرسولي (دولة الفاتيكان)","الكونغو (جمهورية الكونغو الديمقراطية)","الكويت","ألمانيا","المغرب","المكسيك","المملكة المتحدة","النرويج","النمسا","الهند","الولايات المتحدة","اليابان","اليمن","اليونان","أندورا","أنغولا","أوروجواي","أوكرانيا","إيرلندا","أيسلندا","إيطاليا","باراجواي","باكستان","بروناي","بلجيكا","بلغاريا","بنما","بورتو ريكو","بولندا","بوليفيا","بيرو","بيلاروس","تايلاند","تركيا","ترينداد وتوباغو","تشيلي","تونس","جامايكا","جمهورية الدومينيكان","جنوب أفريقيا","جورجيا","روسيا","رومانيا","سان مارينو","سلوفاكيا","سلوفينيا","سنغافورة","سوريا","سويسرا","شمال مقدونيا","صربيا","عُمان","غواتيمالا","غينيا الاستوائية","فرنسا","فنزويلا","فنلندا","فيتنام","قبرص","قطر","كازاخستان","كرواتيا","كندا","كوريا","كوستاريكا","كولومبيا","كينيا","لاتفيا","لبنان","لتوانيا","لكسمبورغ","ليشتنشتاين","مالطا","ماليزيا","مصر","منطقة ماكاو الإدارية الخاصة","منطقة هونغ كونغ الإدارية الخاصة","مولدافا","موناكو","مونتينيغرو","نيجيريا","نيكاراجوا","نيوزيلندا","هندوراس","هنغاريا","هولندا"))
$Cultures.Add("Bulgarian (Bulgaria)", @("Австралия","Австрия","Азeрбайджан","Албания","Алжир","Ангола","Андора","Аржентина","Армения","Бахрейн","Беларус","Белгия","Боливия","Босна и Херцеговина","Бразилия","Бруней","България","Великобритания","Венецуела","Виетнам","Гватемала","Германия","Грузия","Гърция","Дания","Доминиканска република","Еврейски религиозни празници","Европейски съюз","Египет","Еквадор","Екваториална Гвинея","Естония","Йемен","Израел","Индия","Йордания","Ирак","Ирландия","Исландия","Ислямски (сунитски) религиозни празници","Ислямски (шиитски) религиозни празници","Испания","Италия","Казахстан","Канада","Катар","Кения","Кипър","Китай","Колумбия","Конго (Демократична република)","Корея","Коста Рика","Кувейт","Латвия","Ливан","Литва","Лихтенщайн","Люксембург","Малайзия","Малта","Мароко","Мексико","Молдова","Монако","Нигерия","Нидерландия","Никарагуа","Нова Зеландия","Норвегия","Обединени арабски емирства","Оман","Пакистан","Панама","Парагвай","Перу","Полша","Португалия","Пуерто Рико","Румъния","Русия","Салвадор","Сан Марино","САР Макао","САР Хонконг","Саудитска Арабия","Светия престол (Ватикан)","Северна Македония","Сингапур","Сирия","Словакия","Словения","Съединени щати","Сърбия","Тайланд","Тринидад и Тобаго","Тунис","Турция","Украйна","Унгария","Уругвай","Филипини","Финландия","Франция","Хондурас","Християнски религиозни празници","Хърватска","Черна гора","Чехия","Чили","Швейцария","Швеция","Южна Африка","Ямайка","Япония"))
$Cultures.Add("Chinese (Simplified, PRC)", @("中国","丹麦","乌克兰","乌拉圭","也门","亚美尼亚","以色列","伊拉克","伊斯兰(什叶派)宗教假日","伊斯兰(逊尼派)宗教假日","俄罗斯","保加利亚","克罗地亚","冰岛","列支敦士登","刚果民主共和国","加拿大","匈牙利","北马其顿","南非","卡塔尔","卢森堡","印度","危地马拉","厄瓜多尔","叙利亚","哈萨克斯坦","哥伦比亚","哥斯达黎加","土耳其","圣座(梵蒂冈)","圣马力诺","埃及","基督教节日","塞尔维亚","塞浦路斯","墨西哥","多米尼加共和国","奥地利","委内瑞拉","安哥拉","安道尔","尼加拉瓜","尼日利亚","巴基斯坦","巴拉圭","巴拿马","巴林","巴西","希腊","德国","意大利","拉脱维亚","挪威","捷克","摩尔多瓦","摩洛哥","摩纳哥","文莱","斯洛伐克","斯洛文尼亚","新加坡","新西兰","日本","智利","格鲁吉亚","欧盟","比利时","沙特阿拉伯","法国","波兰","波多黎各","波斯尼亚和黑塞哥维那","泰国","洪都拉斯","澳大利亚","澳门特别行政区","爱尔兰","爱沙尼亚","牙买加","特立尼达和多巴哥","犹太教节日","玻利维亚","瑞典","瑞士","白俄罗斯","科威特","秘鲁","突尼斯","立陶宛","约旦","罗马尼亚","美国","肯尼亚","芬兰","英国","荷兰","菲律宾","萨尔瓦多","葡萄牙","西班牙","赤道几内亚","越南","阿塞拜疆","阿尔及利亚","阿尔巴尼亚","阿拉伯联合酋长国","阿曼","阿根廷","韩国","香港特别行政区","马来西亚","马耳他","黎巴嫩","黑山"))
$Cultures.Add("Chinese (Traditional, Taiwan)", @("丹麥","亞塞拜然","亞美尼亞","以色列","伊拉克","俄羅斯","保加利亞","克羅埃西亞","冰島","列支敦斯登","剛果民主共和國","加拿大","匈牙利","北馬其頓","千里達及托巴哥","南非","卡達","印度","厄瓜多","哈薩克","哥倫比亞","哥斯大黎加","喬治亞","回教 (什葉派) 假日","回教 (遜尼派) 假日","土耳其","埃及","基督教假日","塞爾維亞","墨西哥","多明尼加","奈及利亞","奧地利","委內瑞拉","安哥拉","安道爾","宏都拉斯","尼加拉瓜","巴基斯坦","巴拉圭","巴拿馬","巴林","巴西","希臘","德國","愛沙尼亞","愛爾蘭","拉脫維亞","挪威","捷克","摩洛哥","摩爾多瓦","摩納哥","敘利亞","教廷 (梵蒂岡)","斯洛伐克","斯洛維尼亞","新加坡","日本","智利","歐盟","比利時","汶萊","沙烏地阿拉伯","法國","波士尼亞赫塞哥維納","波多黎各","波蘭","泰國","澳大利亞","澳門特別行政區","烏克蘭","烏拉圭","牙買加","猶太教假日","玻利維亞","瑞典","瑞士","瓜地馬拉","白俄羅斯","盧森堡","科威特","秘魯","突尼西亞","立陶宛","約旦","紐西蘭","羅馬尼亞","美國","義大利","聖馬利諾","肯亞","芬蘭","英國","荷蘭","菲律賓","葉門","葡萄牙","蒙特內哥羅","薩爾瓦多","西班牙","賽普勒斯","赤道幾內亞","越南","阿拉伯聯合大公國","阿曼","阿根廷","阿爾及利亞","阿爾巴尼亞","韓國","香港特別行政區","馬來西亞","馬爾他","黎巴嫩"))
$Cultures.Add("Croatian (Croatia)", @("Albanija","Alžir","Andora","Angola","Argentina","Armenija","Australija","Austrija","Azerbajdžan","Bahrein","Belgija","Bjelarus","Bolivija","Bosna i Hercegovina","Brazil","Brunej","Bugarska","Češka","Čile","Cipar","Crna Gora","Danska","Dominikanska Republika","Egipat","Ekvador","Ekvatorijalna Gvineja","El Salvador","Estonija","Europska unija","Filipini","Finska","Francuska","Grčka","Gruzija","Gvatemala","Honduras","Hrvatska","Indija","Irak","Irska","Islamski (šijitski) vjerski blagdani","Islamski (sunitski) vjerski blagdani","Island","Italija","Izrael","Jamajka","Japan","Jemen","Jordan","Južnoafrička Republika","Kanada","Katar","Kazahstan","Kenija","Kina","Kolumbija","Kongo (Demokratska Republika)","Koreja","Kostarika","Kršćanski vjerski blagdani","Kuvajt","Latvija","Libanon","Lihtenštajn","Litva","Luksemburg","Mađarska","Malezija","Malta","Maroko","Meksiko","Moldova","Monako","Nigerija","Nikaragva","Nizozemska","Njemačka","Norveška","Novi Zeland","Oman","Pakistan","Panama","Paragvaj","Peru","Poljska","Portoriko","Portugal","Posebna upr. jedinica Hong Kong","Posebna upravna jedinica Makao","Rumunjska","Rusija","San Marino","Saudijska Arabija","Singapur","Sirija","Sjedinjene Američke Države","Sjeverna Makedonija","Slovačka","Slovenija","Španjolska","Srbija","Švedska","Sveta Stolica (Vatikan)","Švicarska","Tajland","Trinidad i Tobago","Tunis","Turska","Ujedinjeni Arapski Emirati","Ujedinjeno Kraljevstvo","Ukrajina","Urugvaj","Venezuela","Vijetnam","Židovski vjerski blagdani"))
$Cultures.Add("Czech (Czech Republic)", @("Albánie","Alžírsko","Andorra","Angola","Argentina","Arménie","Austrálie","Ázerbájdžán","Bahrajn","Belgie","Bělorusko","Bolívie","Bosna a Hercegovina","Brazílie","Brunej","Bulharsko","Černá Hora","Česko","Chile","Chorvatsko","Čína","Dánsko","Dominikánská republika","Egypt","Ekvádor","Estonsko","Evropská unie","Filipíny","Finsko","Francie","Gruzie","Guatemala","Honduras","Hongkong – zvláštní správní oblast","Indie","Irák","Irsko","Islámské náboženské svátky (šíité)","Islámské náboženské svátky (sunnité)","Island","Itálie","Izrael","Jamajka","Japonsko","Jemen","Jižní Afrika","Jordánsko","Kanada","Katar","Kazachstán","Keňa","Kolumbie","Kongo (Konžská demokratická republika)","Korea","Kostarika","Křesťanské náboženské svátky","Kuvajt","Kypr","Libanon","Lichtenštejnsko","Litva","Lotyšsko","Lucembursko","Macao – zvláštní správní oblast","Maďarsko","Malajsie","Malta","Maroko","Mexiko","Moldavsko","Monako","Německo","Nigérie","Nikaragua","Nizozemsko","Norsko","Nový Zéland","Omán","Pákistán","Panama","Paraguay","Peru","Polsko","Portoriko","Portugalsko","Rakousko","Řecko","Rovníková Guinea","Rumunsko","Rusko","Salvador","San Marino","Saúdská Arábie","Severní Makedonie","Singapur","Slovensko","Slovinsko","Španělsko","Spojené arabské emiráty","Spojené království","Spojené státy","Srbsko","Svatý stolec (Vatikán)","Švédsko","Švýcarsko","Sýrie","Thajsko","Trinidad a Tobago","Tunisko","Turecko","Ukrajina","Uruguay","Venezuela","Vietnam","Židovské náboženské svátky"))
$Cultures.Add("Danish (Denmark)", @("Ækvatorialguinea","Albanien","Algeriet","Andorra","Angola","Argentina","Armenien","Aserbajdsjan","Australien","Bahrain","Belgien","Bolivia","Bosnien-Hercegovina","Brasilien","Brunei","Bulgarien","Canada","Chile","Colombia","Congo (Den Demokratiske Republik)","Costa Rica","Cypern","Danmark","De Forenede Arabiske Emirater","Den Dominikanske Republik","Ecuador","Egypten","El Salvador","Estland","European Union (EU)","Filippinerne","Finland","Frankrig","Georgien","Grækenland","Guatemala","Honduras","Hviderusland","Indien","Irak","Irland","Islamiske (shia) religiøse helligdage","Islamiske (sunni) religiøse helligdage","Island","Israel","Italien","Jamaica","Japan","Jødiske helligdage","Jordan","Kasakhstan","Kenya","Kina","Korea","Kristne helligdage","Kroatien","Kuwait","Letland","Libanon","Liechtenstein","Litauen","Luxembourg","Macao","Malaysia","Malta","Marokko","Mexico","Moldova","Monaco","Montenegro","Nederlandene","New Zealand","Nicaragua","Nigeria","Nordmakedonien","Norge","Oman","Østrig","Pakistan","Panama","Paraguay","Pavestolen (Vatikanstaten)","Peru","Polen","Portugal","Puerto Rico","Qatar","Rumænien","Rusland","San Marino","SAR Hongkong","Saudi-Arabien","Schweiz","Serbien","Singapore","Slovakiet","Slovenien","Spanien","Storbritannien","Sverige","Sydafrika","Syrien","Thailand","Tjekkiet","Trinidad og Tobago","Tunesien","Tyrkiet","Tyskland","Ukraine","Ungarn","Uruguay","USA","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Dutch (Netherlands)", @("Albanië","Algerije","Andorra","Angola","Argentinië","Armenië","Australië","Azerbeidzjan","Bahrein","Belarus","België","Bolivia","Bosnië en Herzegovina","Brazilië","Brunei","Bulgarije","Canada","Chili","China","Christelijke feestdagen","Colombia","Congo (Democratische Republiek)","Costa Rica","Cyprus","Denemarken","Dominicaanse Republiek","Duitsland","Ecuador","Egypte","El Salvador","Equatoriaal-Guinea","Estland","Europese Unie","Filipijnen","Finland","Frankrijk","Georgië","Griekenland","Guatemala","Heilige Stoel (Vaticaanstad)","Honduras","Hongarije","Hongkong SAR","Ierland","IJsland","India","Irak","Islamitische (sjiitische) religieuze feestdagen","Islamitische (soennitische) religieuze feestdagen","Israël","Italië","Jamaica","Japan","Jemen","Joodse feestdagen","Jordanië","Kazachstan","Kenia","Koeweit","Kroatië","Letland","Libanon","Liechtenstein","Litouwen","Luxemburg","Macau SAR","Maleisië","Malta","Marokko","Mexico","Moldavië","Monaco","Montenegro","Nederland","Nicaragua","Nieuw-Zeeland","Nigeria","Noord-Macedonië","Noorwegen","Oekraïne","Oman","Oostenrijk","Pakistan","Panama","Paraguay","Peru","Polen","Porto Rico","Portugal","Qatar","Roemenië","Rusland","San Marino","Saudi-Arabië","Servië","Singapore","Slovenië","Slowakije","Spanje","Syrië","Thailand","Trinidad en Tobago","Tsjechië","Tunesië","Turkije","Uruguay","Venezuela","Verenigd Koninkrijk","Verenigde Arabische Emiraten","Verenigde Staten","Vietnam","Zuid-Afrika","Zuid-Korea","Zweden","Zwitserland"))
$Cultures.Add("English (United States)", @("Albania","Algeria","Andorra","Angola","Argentina","Armenia","Australia","Austria","Azerbaijan","Bahrain","Belarus","Belgium","Bolivia","Bosnia and Herzegovina","Brazil","Brunei","Bulgaria","Canada","Chile","China","Christian Religious Holidays","Colombia","Congo (Democratic Republic of)","Costa Rica","Croatia","Cyprus","Czechia","Denmark","Dominican Republic","Ecuador","Egypt","El Salvador","Equatorial Guinea","Estonia","European Union","Finland","France","Georgia","Germany","Greece","Guatemala","Holy See (Vatican City)","Honduras","Hong Kong S.A.R.","Hungary","Iceland","India","Iraq","Ireland","Islamic (Shia) Religious Holidays","Islamic (Sunni) Religious Holidays","Israel","Italy","Jamaica","Japan","Jewish Religious Holidays","Jordan","Kazakhstan","Kenya","Korea","Kuwait","Latvia","Lebanon","Liechtenstein","Lithuania","Luxembourg","Macao S.A.R.","Malaysia","Malta","Mexico","Moldova","Monaco","Montenegro","Morocco","Netherlands","New Zealand","Nicaragua","Nigeria","North Macedonia","Norway","Oman","Pakistan","Panama","Paraguay","Peru","Philippines","Poland","Portugal","Puerto Rico","Qatar","Romania","Russia","San Marino","Saudi Arabia","Serbia","Singapore","Slovakia","Slovenia","South Africa","Spain","Sweden","Switzerland","Syria","Thailand","Trinidad and Tobago","Tunisia","Turkey","Ukraine","United Arab Emirates","United Kingdom","United States","Uruguay","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Estonian (Estonia)", @("Albaania","Alžeeria","Ameerika Ühendriigid","Andorra","Angola","Aomeni erihalduspiirkond","Araabia Ühendemiraadid","Argentina","Armeenia","Aserbaidžaan","Austraalia","Austria","Bahrein","Belgia","Boliivia","Bosnia ja Hertsegoviina","Brasiilia","Brunei","Bulgaaria","Colombia","Costa Rica","Dominikaani Vabariik","Ecuador","Eesti","Egiptus","Ekvatoriaal-Guinea","El Salvador","Euroopa Liit","Filipiinid","Gruusia","Guatemala","Hiina","Hispaania","Holland","Honduras","Hongkongi erihalduspiirkond","Horvaatia","Iirimaa","Iisrael","India","Iraak","Islami (šiia) religioossed pühad","Islami (sunni) religioossed pühad","Island","Itaalia","Jaapan","Jamaica","Jeemen","Jordaania","Juudi pühad","Kanada","Kasahstan","Katar","Kenya","Kongo Demokraatlik Vabariik","Kreeka","Kristlikud pühad","Küpros","Kuveit","Läti","Leedu","Liechtenstein","Liibanon","Lõuna-Aafrika Vabariik","Lõuna-Korea","Luksemburg","Malaisia","Malta","Maroko","Mehhiko","Moldova","Monaco","Montenegro","Nicaragua","Nigeeria","Norra","Omaan","Pakistan","Panama","Paraguay","Peruu","Põhja-Makedoonia","Poola","Portugal","Prantsusmaa","Puerto Rico","Püha Tool (Vatikani Linnriik)","Rootsi","Rumeenia","Saksamaa","San Marino","Saudi Araabia","Serbia","Singapur","Slovakkia","Sloveenia","Soome","Süüria","Šveits","Taani","Tai","Trinidad ja Tobago","Tšehhia","Tšiili","Tuneesia","Türgi","Ühendkuningriik","Ukraina","Ungari","Uruguay","Uus-Meremaa","Valgevene","Venemaa","Venezuela","Vietnam"))
$Cultures.Add("Finnish (Finland)", @("Alankomaat","Albania","Algeria","Andorra","Angola","Argentiina","Armenia","Australia","Azerbaidžan","Bahrain","Belgia","Bolivia","Bosnia ja Hertsegovina","Brasilia","Brunei","Bulgaria","Chile","Costa Rica","Dominikaaninen tasavalta","Ecuador","Egypti","El Salvador","Espanja","Etelä-Afrikka","Euroopan unioni","Filippiinit","Georgia","Guatemala","Honduras","Intia","Irak","Irlanti","Islamilaiset juhlapyhät (shiia)","Islamilaiset juhlapyhät (sunni)","Islanti","Israel","Italia","Itävalta","Jamaika","Japani","Jemen","Jordania","Juutalaiset pyhäpäivät","Kanada","Kazakstan","Kenia","Kiina","Kiinan kansantasavallan erityishallintoalue Hongkong","Kolumbia","Kongo (demokraattinen tasavalta)","Korea","Kreikka","Kristilliset pyhäpäivät","Kroatia","Kuwait","Kypros","Latvia","Libanon","Liechtenstein","Liettua","Luxemburg","Macaon erityishallintoalue","Malesia","Malta","Marokko","Meksiko","Moldova","Monaco","Montenegro","Nicaragua","Nigeria","Norja","Oman","Päiväntasaajan Guinea","Pakistan","Panama","Paraguay","Peru","Pohjois-Makedonia","Portugali","Puerto Rico","Puola","Pyhä istuin (Vatikaanivaltio)","Qatar","Ranska","Romania","Ruotsi","Saksa","San Marino","Saudi-Arabia","Serbia","Singapore","Slovakia","Slovenia","Suomi","Sveitsi","Syyria","Tanska","Thaimaa","Trinidad ja Tobago","Tšekki","Tunisia","Turkki","Ukraina","Unkari","Uruguay","Uusi-Seelanti","Valko-Venäjä","Venäjä","Venezuela","Vietnam","Viro","Yhdistyneet arabiemiirikunnat","Yhdistynyt kuningaskunta","Yhdysvallat"))
$Cultures.Add("French (France)", @("Afrique du Sud","Albanie","Algérie","Allemagne","Andorre","Angola","Arabie saoudite","Argentine","Arménie","Australie","Autriche","Azerbaïdjan","Bahreïn","Bélarus","Belgique","Bolivie","Bosnie-Herzégovine","Brésil","Brunei","Bulgarie","Canada","Chili","Chine","Chypre","Colombie","Congo (République démocratique du)","Corée du Sud","Costa Rica","Croatie","Danemark","Égypte","Émirats arabes unis","Équateur","Espagne","Estonie","États-Unis","Fêtes religieuses chrétiennes","Fêtes religieuses islamiques (shia)","Fêtes religieuses islamiques (sunni)","Fêtes religieuses juives","Finlande","France","Géorgie","Grèce","Guatemala","Guinée équatoriale","Honduras","Hong Kong RAS","Hongrie","Inde","Irak","Irlande","Islande","Israël","Italie","Jamaïque","Japon","Jordanie","Kazakhstan","Kenya","Koweït","Lettonie","Liban","Liechtenstein","Lituanie","Luxembourg","Macao (R.A.S.)","Macédoine du Nord","Malaisie","Malte","Maroc","Mexique","Moldova","Monaco","Monténégro","Nicaragua","Nigeria","Norvège","Nouvelle-Zélande","Oman","Pakistan","Panama","Paraguay","Pays-Bas","Pérou","Philippines","Pologne","Portugal","Puerto Rico","Qatar","République Dominicaine","Roumanie","Royaume-Uni","Russie","Saint-Marin","Saint-Siège (Cité du Vatican)","Salvador","Serbie","Singapour","Slovaquie","Slovénie","Suède","Suisse","Syrie","Tchéquie","Thaïlande","Trinité-et-Tobago","Tunisie","Turquie","Ukraine","Union européenne","Uruguay","Venezuela","Vietnam","Yémen"))
$Cultures.Add("German (Germany)", @("Ägypten","Albanien","Algerien","Andorra","Angola","Äquatorialguinea","Argentinien","Armenien","Aserbaidschan","Australien","Bahrain","Belarus","Belgien","Bolivien","Bosnien und Herzegowina","Brasilien","Brunei Darussalam","Bulgarien","Chile","China","Christliche religiöse Feiertage","Costa Rica","Dänemark","Deutschland","Dominikanische Republik","Ecuador","El Salvador","Estland","Europäische Union","Finnland","Frankreich","Georgien","Griechenland","Guatemala","Heiliger Stuhl (Vatikanstadt)","Honduras","Indien","Irak","Irland","Islamische religiöse Feiertage (schiitisch)","Islamische religiöse Feiertage (sunnitisch)","Island","Israel","Italien","Jamaika","Japan","Jemen","Jordanien","Jüdische religiöse Feiertage","Kanada","Kasachstan","Katar","Kenia","Kolumbien","Kongo (Demokratische Republik)","Kroatien","Kuwait","Lettland","Libanon","Liechtenstein","Litauen","Luxemburg","Malaysia","Malta","Marokko","Mexiko","Monaco","Montenegro","Neuseeland","Nicaragua","Niederlande","Nigeria","Nordmazedonien","Norwegen","Oman","Österreich","Pakistan","Panama","Paraguay","Peru","Philippinen","Polen","Portugal","Puerto Rico","Republik Korea","Republik Moldau","Rumänien","Russische Föderation","San Marino","Saudi-Arabien","Schweden","Schweiz","Serbien","Singapur","Slowakei","Slowenien","Sonderverwaltungsregion Hongkong","Sonderverwaltungsregion Macau","Spanien","Südafrika","Syrien","Thailand","Trinidad und Tobago","Tschechien","Tunesien","Türkei","Ukraine","Ungarn","Uruguay","Venezuela","Vereinigte Arabische Emirate","Vereinigte Staaten","Vereinigtes Königreich","Vietnam","Zypern"))
$Cultures.Add("Greek (Greece)", @("Αγία Έδρα (Πόλη του Βατικανού)","Άγιος Μαρίνος","Αγκόλα","Αζερμπαϊτζάν","Αίγυπτος","Αλβανία","Αλγερία","Ανδόρα","Αργεντινή","Αρμενία","Αυστραλία","Αυστρία","Βέλγιο","Βενεζουέλα","Βιετνάμ","Βολιβία","Βόρεια Μακεδονία","Βοσνία και Ερζεγοβίνη","Βουλγαρία","Βραζιλία","Γαλλία","Γερμανία","Γεωργία","Γουατεμάλα","Δανία","Δομινικανή Δημοκρατία","Εβραϊκές θρησκευτικές εορτές","Ελ Σαλβαδόρ","Ελβετία","Ελλάδα","Εσθονία","Ευρωπαϊκή Ένωση","Ηνωμένα Αραβικά Εμιράτα","Ηνωμένες Πολιτείες","Ηνωμένο Βασίλειο","Ιαπωνία","Ινδία","Ιορδανία","Ιράκ","Ιρλανδία","Ισημερινή Γουινέα","Ισημερινός","Ισλαμικές θρησκευτικές αργίες (Σιίτες)","Ισλαμικές θρησκευτικές αργίες (Σουνίτες)","Ισλανδία","Ισπανία","Ισραήλ","Ιταλία","Καζαχστάν","Καναδάς","Κατάρ","Κάτω Χώρες","Κένυα","Κίνα","Κολομβία","Κονγκό (Λαϊκή Δημοκρατία)","Κορέα","Κόστα Ρίκα","Κουβέιτ","Κροατία","Κύπρος","Λετονία","Λευκορωσία","Λίβανος","Λιθουανία","Λιχτενστάιν","Λουξεμβούργο","Μακάο ΕΔΠ","Μαλαισία","Μάλτα","Μαρόκο","Μαυροβούνιο","Μεξικό","Μολδαβία","Μονακό","Μπαχρέιν","Μπρουνέι","Νέα Ζηλανδία","Νιγηρία","Νικαράγουα","Νορβηγία","Νότια Αφρική","Ομάν","Ονδούρα","Ουγγαρία","Ουκρανία","Ουρουγουάη","Πακιστάν","Παναμάς","Παραγουάη","Περού","Πολωνία","Πορτογαλία","Πουέρτο Ρίκο","Ρουμανία","Ρωσία","Σαουδική Αραβία","Σερβία","Σιγκαπούρη","Σλοβακία","Σλοβενία","Σουηδία","Συρία","Ταϊλάνδη","Τζαμάικα","Τουρκία","Τρινιντάντ και Τομπάγκο","Τσεχία","Τυνησία","Υεμένη","Φιλιππίνες","Φινλανδία","Χιλή","Χονγκ Κονγκ ΕΔΠ","Χριστιανικές θρησκευτικές εορτές"))
$Cultures.Add("Hebrew (Israel)", @("אוסטריה","אוסטרליה","אוקראינה","אורוגוואי","אזרבייג'ן","איחוד האמירויות הערביות","איטליה","איסלנד","אירלנד","אל סלבדור","אלבניה","אלג'יריה","אנגולה","אנדורה","אסטוניה","אקוודור","ארגנטינה","ארמניה","ארצות הברית","בולגריה","בוליביה","בוסניה והרצגובינה","בחריין","בלגיה","בלרוס","ברוניי","ברזיל","בריטניה","גואטמלה","גיאורגיה","גינאה המשוונית","ג'מייקה","גרמניה","דנמרק","דרום אפריקה","האיחוד האירופי","הודו","הולנד","הונג קונג S.A.R.‎","הונגריה","הונדורס","הכס הקדוש (קרית הוותיקן)","הפיליפינים","הרפובליקה הדומיניקנית","וייטנאם","ונצואלה","חגים דתיים יהודיים","חגים דתיים מוסלמיים (סוני)","חגים דתיים מוסלמיים (שיעה)","חגים דתיים נוצריים","טוניסיה","טורקיה","טרינידד וטובגו","יוון","יפן","ירדן","ישראל","כוויית","לבנון","לוקסמבורג","לטביה","ליטא","ליכטנשטיין","מולדובה","מונטנגרו","מונקו","מלזיה","מלטה","מצרים","מקאו S.A.R.‎","מקדוניה הצפונית","מקסיקו","מרוקו","נורווגיה","ניגריה","ניו זילנד","ניקרגואה","סוריה","סין","סינגפור","סלובניה","סלובקיה","סן מרינו","ספרד","סרביה","עומאן","עיראק","ערב הסעודית","פוארטו ריקו","פולין","פורטוגל","פינלנד","פנמה","פקיסטן","פרגוואי","פרו","צ'ילה","צ'כיה","צרפת","קולומביה","קונגו (הרפובליקה הדמוקרטית)","קוסטה ריקה","קוריאה","קזחסטן","קטאר","קנדה","קניה","קפריסין","קרואטיה","רומניה","רוסיה","שוודיה","שוויץ","תאילנד","תימן"))
$Cultures.Add("Hindi (India)", @("अंगोला","अंडोरा","अज़रबैजान","अर्जेंटीना","अर्मेनिया","अल सल्वाडोर","अल्जेरिया","अल्बानिया","आईसलैंड","आयरलैंड","इक्वेटोरियल गीनिया","इक्वेडोर","इज़रायल","इटली","इराक","इस्लामी (शिया) धार्मिक हॉलिडे","इस्लामी (सुन्नी) धार्मिक हॉलिडे","ईसाई धार्मिक हॉलिडे","उक्रैन","उत्तरी मकदूनिया","उरुग्वे","एस्टोनिया","ऑस्ट्रिया","ऑस्ट्रेलिया","ओमान","कजाकस्तान","कतर","कनाडा","काँगो (का लोकतांत्रिक गणराज्य)","कुवैत","केन्या","कोरिया","कोलंबिया","कोस्टा रिका","क्रोएशिया","ग्रीस","ग्वाटेमाला","चिली","चीन","चेकिया","जमैका","जर्मनी","जापान","जॉर्जिया","जॉर्डन","ट्यूनीशिया","डेन्मार्क","डोमिनिकन गणराज्य","तुर्कस्तान","त्रिनिदाद और टोबेगो","थायलंड","दक्षिण आफ़्रिका","नाइजीरिया","निकारागुआ","नीदरलैंड","नॉर्वे","न्यूज़ीलैंड","पनामा","पाकिस्तान","पुर्तगाल","पेरू","पैराग्वे","पोलैंड","प्युर्तो रिको","फ़िनलैंड","फ़िलीपीन्स","फ़्रांस","बहरीन","बुल्गारिया","बेलारूस","बेल्जियम","बोलीविया","बोस्निया और हर्ज़ेगोविना","ब्राज़ील","ब्रुनेई","भारत","मकाऊ S.A.R.","मलेशिया","माल्टा","मिस्र","मेक्सिको","मॉन्टेंगरो","मोनाको","मोरोक्को","मोल्डोवा","यमन","यहूदी धार्मिक छुट्टियाँ","युनाइटेड किंगडम","योरपीय यूनियन","रुमानिया","रूस","लक्ज़ेम्बर्ग","लाटविया","लिचेंस्टीन","लिथुआनिया","लेबनान","वियतनाम","वेनेज़ुएला","संयुक्त अरब अमारात","संयुक्त राज्य अमरीका","सऊदी अरबस्तान","सर्बिया","साइप्रस","सिंगापुर","सीरिया","सैन मारीनो","स्पेन","स्लोवाक गणतंत्र","स्लोवेनिया","स्वित्ज़र्लैंड","स्वीडन","हंगरी","हाँग काँग S.A.R.","होंडुरस","होली सी (वेटिकन सिटी)"))
$Cultures.Add("Hungarian (Hungary)", @("Albánia","Algéria","Amerikai Egyesült Államok","Andorra","Angola","Argentína","Ausztrália","Ausztria","Azerbajdzsán","Bahrein","Belarusz","Belgium","Bolívia","Bosznia-Hercegovina","Brazília","Brunei","Bulgária","Chile","Ciprus","Costa Rica","Csehország","Dánia","Dél-Afrika","Dominikai Köztársaság","Ecuador","Egyenlítői Guinea","Egyesült Arab Emírségek","Egyesült Királyság","Egyiptom","Észak-Macedónia","Észtország","Európai Unió","Finnország","Franciaország","Fülöp-szigetek","Görögország","Grúzia","Guatemala","Hollandia","Honduras","Hongkong (KKT)","Horvátország","India","Irak","Írország","Iszlám (síita) vallási ünnepek","Iszlám (szunnita) vallási ünnepek","Izland","Izrael","Jamaica","Japán","Jemen","Jordánia","Kanada","Katar","Kazahsztán","Kenya","Keresztény vallási ünnepek","Kína","Kolumbia","Kongói Demokratikus Köztársaság","Korea","Kuvait","Lengyelország","Lettország","Libanon","Liechtenstein","Litvánia","Luxemburg","Magyarország","Makaó (KKT)","Malajzia","Málta","Marokkó","Mexikó","Moldova","Monaco","Montenegró","Németország","Nicaragua","Nigéria","Norvégia","Olaszország","Omán","Örményország","Oroszország","Pakisztán","Panama","Paraguay","Peru","Portugália","Puerto Rico","Románia","Salvador","San Marino","Spanyolország","Svájc","Svédország","Szaúd-Arábia","Szentszék (Vatikánváros)","Szerbia","Szingapúr","Szíria","Szlovákia","Szlovénia","Thaiföld","Törökország","Trinidad és Tobago","Tunézia","Új-Zéland","Ukrajna","Uruguay","Venezuela","Vietnam","Zsidó vallási ünnepek"))
$Cultures.Add("Indonesian (Indonesia)", @("Afrika Selatan","Albania","Aljazair","Amerika Serikat","Andorra","Angola","Arab Saudi","Argentina","Armenia","Australia","Austria","Azerbaijan","Bahrain","Belanda","Belarus","Belgia","Bolivia","Bosnia dan Herzegovina","Brasil","Brunei","Bulgaria","Ceko","Chili","China","Daerah Administratif Khusus Macau","Denmark","Ekuador","El Salvador","Estonia","Filipina","Finlandia","Georgia","Guatemala","Guinea Ekuatorial","Hari Libur Agama Yahudi","Hari Libur Kristen","Holy See (Kota Suci Vatikan)","Honduras","Hongaria","India","Irak","Irlandia","Islandia","Israel","Italia","Jamaika","Jepang","Jerman","Kanada","Kazakhstan","Kenya","Kerajaan Inggris Bersatu","Kolombia","Kongo (Republik Demokratik)","Korea","Kosta Rika","Kroasia","Kuwait","Latvia","Lebanon","Libur Agama Islam (Sunni)","Libur Agama Islam (Syiah)","Liechtenstein","Lithuania","Luksemburg","Makedonia Utara","Malaysia","Malta","Maroko","Meksiko","Mesir","Moldova","Monako","Montenegro","Nigeria","Nikaragua","Norwegia","Oman","Pakistan","Panama","Paraguay","Perserikatan Eropa","Peru","Polandia","Portugal","Prancis","Puerto Riko","Qatar","Republik Dominika","Rumania","Rusia","S.A.R. Hong Kong","San Marino","Selandia Baru","Serbia","Singapura","Siprus","Slovenia","Slowakia","Spanyol","Suriah","Swedia","Swiss","Thailand","Trinidad dan Tobago","Tunisia","Turki","Ukraina","Uni Emirat Arab","Uruguay","Venezuela","Vietnam","Yaman","Yordania","Yunani"))
$Cultures.Add("Italian (Italy)", @("Albania","Algeria","Andorra","Angola","Arabia Saudita","Argentina","Armenia","Australia","Austria","Azerbaigian","Bahrein","Belarus","Belgio","Bolivia","Bosnia ed Erzegovina","Brasile","Brunei","Bulgaria","Canada","Cechia","Cile","Cina","Cipro","Colombia","Corea","Costa Rica","Croazia","Danimarca","Ecuador","Egitto","El Salvador","Emirati Arabi Uniti","Estonia","Filippine","Finlandia","Francia","Georgia","Germania","Giamaica","Giappone","Giordania","Grecia","Guatemala","Guinea Equatoriale","Honduras","Hong Kong R.A.S.","India","Iraq","Irlanda","Islanda","Israele","Italia","Kazakhstan","Kenya","Kuwait","Lettonia","Libano","Liechtenstein","Lituania","Lussemburgo","Macao - R.A.S.","Macedonia del Nord","Malaysia","Malta","Marocco","Messico","Moldova","Monaco","Montenegro","Nicaragua","Nigeria","Norvegia","Nuova Zelanda","Oman","Paesi Bassi","Pakistan","Panama","Paraguay","Perù","Polonia","Portogallo","Portorico","Qatar","Regno Unito","Religione cristiana","Religione ebraica","Religione islamica (sciita)","Religione islamica (sunnita)","Repubblica democratica del Congo","Repubblica dominicana","Romania","Russia","San Marino","Santa Sede (Stato della Città del Vaticano)","Serbia","Singapore","Siria","Slovacchia","Slovenia","Spagna","Stati Uniti","Sudafrica","Svezia","Svizzera","Thailandia","Trinidad e Tobago","Tunisia","Turchia","Ucraina","Ungheria","Unione Europea","Uruguay","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Japanese (Japan)", @("アイスランド","アイルランド","アゼルバイジャン","アラブ首長国連邦","アルジェリア","アルゼンチン","アルバニア","アルメニア","アンゴラ","アンドラ","イエメン","イスラエル","イスラム教 (シーア派) 祝祭日","イスラム教 (スンニ派) 祝祭日","イタリア","イラク","インド","ウクライナ","ウルグアイ","エクアドル","エジプト","エストニア","エルサルバドル","オーストラリア","オーストリア","オマーン","オランダ","カザフスタン","カタール","カナダ","キプロス","ギリシャ","キリスト教祝祭日","グアテマラ","クウェート","クロアチア","ケニア","コスタリカ","コロンビア","コンゴ民主共和国","サウジアラビア","サンマリノ","ジャマイカ","ジョージア","シリア","シンガポール","スイス","スウェーデン","スペイン","スロバキア","スロベニア","セルビア","タイ","チェコ","チュニジア","チリ","デンマーク","ドイツ","ドミニカ共和国","トリニダード・トバゴ","トルコ","ナイジェリア","ニカラグア","ニュージーランド","ノルウェー","バーレーン","パキスタン","バチカン","パナマ","パラグアイ","ハンガリー","フィリピン","フィンランド","プエルトリコ","ブラジル","フランス","ブルガリア","ブルネイ","ベトナム","ベネズエラ","ベラルーシ","ペルー","ベルギー","ポーランド","ボスニア・ヘルツェゴビナ","ボリビア","ポルトガル","ホンジュラス","マカオ特別行政区","マルタ","マレーシア","メキシコ","モナコ","モルドバ","モロッコ","モンテネグロ","ユダヤ教祝祭日","ヨルダン","ラトビア","リトアニア","リヒテンシュタイン","ルーマニア","ルクセンブルク","レバノン","ロシア","中国","北マケドニア","南アフリカ","日本","欧州連合創設記念日","米国","英国","赤道ギニア","韓国","香港特別行政区"))
$Cultures.Add("Kazakh (Kazakhstan)", @("Австралия","Австрия","Албания","Алжир","Америка Құрама Штаттары","Ангола","Андорра","Аргентина","Армения","Бахрейн","Беларусь","Бельгия","Біріккен Араб Әмірліктері","Болгария","Боливия","Босния және Герцеговина","Бразилия","Бруней","Венгрия","Венесуэла","Вьетнам","Гватемала","Германия","Гондурас","Гонконг АӘА","Грекия","Грузия","Дания","Доминикан Республикасы","Еврей ұлттық мейрамдары","Еуропалық Одақ","Әзірбайжан","Әулие тақ (Ватикан)","Жаңа Зеландия","Жапония","Йемен","Израиль","Иордания","Ирак","Ирландия","Исландия","Испания","Италия","Канада","Катар","Кения","Кипр","Колумбия","Конго (Демократиялық республикасы)","Корея","Коста-Рика","Кувейт","Қазақстан","Құрама Корольдік","Қытай","Латвия","Ливан","Литва","Лихтенштейн","Люксембург","Макао АӘА","Малайзия","Мальта","Марокко","Мексика","Молдова","Монако","Мұсылмандық (Сунна) діни мерекелер","Мұсылмандық (Шия) діни мерекелер","Мысыр","Нигерия","Нидерланд","Никарагуа","Норвегия","Оман","Оңтүстік Африка","Панама","Парагвай","Перу","Пәкістан","Польша","Португалия","Пуэрто-Рико","Ресей","Румыния","Сальвадор","Сан-Марино","Сауд Арабиясы","Сербия","Сингапур","Сирия","Словакия","Словения","Солтүстік Македония","Таиланд","Тринидад және Тобаго","Тунис","Түркия","Украина","Уругвай","Үндістан","Филиппин","Финляндия","Франция","Хорватия","Христиан діни мейрамдары","Черногория","Чехия","Чили","Швейцария","Швеция","Эквадор","Экваторлық Гвинея","Эстония","Ямайка"))
$Cultures.Add("Korean (Korea)", @("과테말라","교황청(바티칸 시국)","교회 기념일","그리스","나이지리아","남아프리카 공화국","네덜란드","노르웨이","뉴질랜드","니카라과","덴마크","도미니카 공화국","독일","라트비아","러시아","레바논","루마니아","룩셈부르크","리투아니아","리히텐슈타인","마카오 특별 행정구","말레이시아","멕시코","모나코","모로코","몬테네그로","몰도바","몰타","미국","바레인","베네수엘라","베트남","벨기에","벨로루시","보스니아 헤르체고비나","볼리비아","북마케도니아","불가리아","브라질","브루나이","사우디아라비아","산마리노","세르비아","스웨덴","스위스","스페인","슬로바키아","슬로베니아","시리아","싱가포르","아랍에미리트","아르메니아","아르헨티나","아이슬란드","아일랜드","아제르바이잔","안도라","알바니아","알제리","앙골라","에스토니아","에콰도르","엘살바도르","영국","예멘","오만","오스트레일리아","오스트리아","요르단","우루과이","우크라이나","유럽 연합","유태교 기념일","이라크","이스라엘","이슬람(수니파) 종교 휴일","이슬람(시아파) 종교 휴일","이집트","이탈리아","인도","일본","자메이카","적도 기니","조지아","중국","체코","칠레","카자흐스탄","카타르","캐나다","케냐","코스타리카","콜롬비아","콩고(민주공화국)","쿠웨이트","크로아티아","키프로스","태국","터키","튀니지","트리니다드 토바고","파나마","파라과이","파키스탄","페루","포르투갈","폴란드","푸에르토리코","프랑스","핀란드","필리핀","한국","헝가리","혼두라스","홍콩 특별 행정구"))
$Cultures.Add("Latvian (Latvia)", @("Albānija","Alžīrija","Amerikas Savienotās Valstis","Andora","Angola","Apvienotā Karaliste","Apvienotie Arābu Emirāti","Argentīna","Armēnija","Austrālija","Austrija","Azerbaidžāna","Bahreina","Baltkrievija","Beļģija","Bolīvija","Bosnija un Hercegovina","Brazīlija","Bruneja","Bulgārija","Čehija","Čīle","Dānija","Dienvidāfrika","Dominikānas Republika","Ebreju reliģiskie svētki","Ēģipte","Eiropas Savienība","Ekvadora","Ekvatoriālā Gvineja","Filipīnas","Francija","Grieķija","Gruzija","Gvatemala","Hondurasa","Horvātija","Igaunija","Indija","Īpašais administratīvais reģions Honkonga","Īpašais administratīvais reģions Makao","Irāka","Īrija","Islāma (šiisms) reliģiskie svētki","Islāmistu (sunnisms) reliģiskie svētki","Islande","Itālija","Izraēla","Jamaika","Japāna","Jaunzēlande","Jemena","Jordānija","Kanāda","Katara","Kazahstāna","Kenija","Ķīna","Kipra","Kolumbija","Kongo Demokrātiskā Republika","Koreja","Kostarika","Krievija","Kristiešu reliģiskie svētki","Kuveita","Latvija","Libāna","Lietuva","Lihtenšteina","Luksemburga","Malaizija","Malta","Maroka","Meksika","Melnkalne","Moldova","Monako","Nīderlande","Nigērija","Nikaragva","Norvēģija","Omāna","Pakistāna","Panama","Paragvaja","Peru","Polija","Portugāle","Puertoriko","Rumānija","Salvadora","Sanmarīno","Saūda Arābija","Serbija","Singapūra","Sīrija","Slovākija","Slovēnija","Somija","Spānija","Šveice","Svētais Krēsls (Vatikāns)","Taizeme","Trinidāda un Tobāgo","Tunisija","Turcija","Ukraina","Ungārija","Urugvaja","Vācija","Venecuēla","Vjetnama","Ziemeļmaķedonija","Zviedrija"))
$Cultures.Add("Lithuanian (Lithuania)", @("Airija","Albanija","Alžyras","Andora","Angola","Argentina","Armėnija","Australija","Austrija","Azerbaidžanas","Bahreinas","Baltarusija","Belgija","Bolivija","Bosnija ir Hercegovina","Brazilija","Brunėjus","Bulgarija","Čekija","Čilė","Danija","Dominikos Respublika","Egiptas","Ekvadoras","Estija","Europos Sąjunga","Filipinai","Graikija","Gruzija","Gvatemala","Hondūras","Indija","Irakas","Islamo (Shia) religijos šventės","Islamo (Sunni) religijos šventės","Islandija","Ispanija","Italija","Izraelis","Jamaika","Japonija","Jemenas","Jordanija","Jungtinė Karalystė","Jungtinės Valstijos","Jungtiniai Arabų Emyratai","Juodkalnija","Kanada","Kataras","Kazachija","Kenija","Kinija","Kipras","Kolumbija","Kongas (Demokratinė Respublika)","Korėja","Kosta Rika","Krikščionių religinės šventės","Kroatija","Kuveitas","Latvija","Lenkija","Libanas","Lichtenšteinas","Lietuva","Liuksemburgas","Malaizija","Malta","Marokas","Meksika","Moldova","Monakas","Naujoji Zelandija","Nigerija","Nikaragva","Norvegija","Nyderlandai","Omanas","Pakistanas","Panama","Paragvajus","Peru","Pietų Afrika","Portugalija","Prancūzija","Puerto Rikas","Pusiaujo Gvinėja","Rumunija","Rusija","Salvadoras","San Marinas","Saudo Arabija","Serbija","Šiaurės Makedonija","Singapūras","Sirija","Slovakija","Slovėnija","Suomija","Švedija","Šveicarija","Šventasis Sostas (Vatikano Miesto Valstybė)","Tailandas","Trinidadas ir Tobagas","Tunisas","Turkija","Ukraina","Urugvajus","Venesuela","Vengrija","Vietnamas","Vokietija","YAKR Honkongas","YAKR Makao","Žydų religinės šventės"))
$Cultures.Add("Malay (Malaysia)", @("Afrika Selatan","Albania","Algeria","Amerika Syarikat","Andorra","Angola","Arab Saudi","Argentina","Armenia","Australia","Austria","Azerbaijan","Bahrain","Belanda","Belarus","Belgium","Bolivia","Bosnia dan Herzegovina","Brazil","Brunei","Bulgaria","Chile","China","Colombia","Congo (Republik Demokratik)","Costa Rica","Croatia","Cuti Keagamaan Kristian","Cuti-cuti Keagamaan Yahudi","Cyprus","Czechia","Denmark","Ecuador","El Salvador","Emiriyah Arab Bersatu","Equatorial Guinea","Estonia","Filipina","Finland","Georgia","Greece","Guatemala","Hari Keagamaan Islam (Shia)","Hari Keagamaan Islam (Sunni)","Holy See (Kota Vatican)","Honduras","Hungary","Iceland","India","Iraq","Ireland","Israel","Itali","Jamaica","Jepun","Jerman","Jordan","Kanada","Kazakhstan","Kenya","Kesatuan Eropah","Korea","Kuwait","Latvia","Liechtenstein","Lithuania","Lubnan","Luxembourg","Macau S.A.R.","Macedonia Utara","Maghribi","Malaysia","Malta","Mesir","Mexico","Moldova","Monaco","Montenegro","New Zealand","Nicaragua","Nigeria","Norway","Oman","Pakistan","Panama","Paraguay","Perancis","Peru","Poland","Portugal","Puerto Rico","Qatar","Republik Dominica","Romania","Rusia","S.A.R Hong Kong","San Marino","Sepanyol","Serbia","Singapura","Slovakia","Slovenia","Sweden","Switzerland","Syria","Thailand","Trinidad dan Tobago","Tunisia","Turki","Ukraine","United Kingdom","Uruguay","Venezuela","Vietnam","Yaman"))
$Cultures.Add("Norwegian, Bokmål (Norway)", @("Albania","Algerie","Andorra","Angola","Argentina","Armenia","Aserbajdsjan","Australia","Bahrain","Belgia","Bolivia","Bosnia-Hercegovina","Brasil","Brunei","Bulgaria","Canada","Chile","Colombia","Costa Rica","Danmark","De forente arabiske emirater","Den demokratiske republikken Kongo","Den dominikanske republikken","Den europeiske union","Ecuador","Egypt","Ekvatorial-Guinea","El Salvador","Estland","Filippinene","Finland","Frankrike","Georgia","Guatemala","Hellas","Honduras","Hongkong SAR","Hviterussland","India","Irak","Irland","Island","Israel","Italia","Jamaica","Japan","Jemen","Jødiske helligdager","Jordan","Kasakhstan","Kenya","Kina","Kristne helligdager","Kroatia","Kuwait","Kypros","Latvia","Libanon","Liechtenstein","Litauen","Luxemburg","Macao SAR","Malaysia","Malta","Marokko","Mexico","Moldova","Monaco","Montenegro","Muslimske religiøse helligdagar (shia)","Muslimske religiøse helligdager (sunni)","Nederland","New Zealand","Nicaragua","Nigeria","Nord-Makedonia","Norge","Oman","Østerrike","Pakistan","Panama","Paraguay","Peru","Polen","Portugal","Puerto Rico","Qatar","Romania","Russland","San Marino","Saudi-Arabia","Serbia","Singapore","Slovakia","Slovenia","Sør-Afrika","Sør-Korea","Spania","Storbritannia","Sveits","Sverige","Syria","Thailand","Trinidad og Tobago","Tsjekkia","Tunisia","Tyrkia","Tyskland","Ukraina","Ungarn","Uruguay","USA","Vatikanstaten","Venezuela","Vietnam"))
$Cultures.Add("Polish (Poland)", @("Albania","Algieria","Andora","Angola","Arabia Saudyjska","Argentyna","Armenia","Australia","Austria","Azerbejdżan","Bahrajn","Belgia","Białoruś","Boliwia","Bośnia i Hercegowina","Brazylia","Brunei","Bułgaria","Chile","Chiny","Chorwacja","Chrześcijańskie święta religijne","Cypr","Czarnogóra","Czechy","Dania","Dominikana","Egipt","Ekwador","Estonia","Filipiny","Finlandia","Francja","Grecja","Gruzja","Gwatemala","Gwinea Równikowa","Hiszpania","Holandia","Honduras","Hongkong SAR","Indie","Irak","Irlandia","Islamskie święta religijne (sunnizm)","Islamskie święta religijne (szyizm)","Islandia","Izrael","Jamajka","Japonia","Jemen","Jordania","Kanada","Katar","Kazachstan","Kenia","Kolumbia","Kongo (Demokratyczna Republika)","Korea Południowa","Kostaryka","Kuwejt","Liban","Liechtenstein","Litwa","Łotwa","Luksemburg","Macedonia Północna","Makau SAR","Malezja","Malta","Maroko","Meksyk","Mołdawia","Monako","Niemcy","Nigeria","Nikaragua","Norwegia","Nowa Zelandia","Oman","Pakistan","Panama","Paragwaj","Peru","Polska","Portoryko","Portugalia","Rosja","RPA","Rumunia","Salwador","San Marino","Serbia","Singapur","Słowacja","Słowenia","Stany Zjednoczone","Stolica Apostolska (Państwo Watykańskie)","Syria","Szwajcaria","Szwecja","Tajlandia","Trynidad i Tobago","Tunezja","Turcja","Ukraina","Unia Europejska","Urugwaj","Węgry","Wenezuela","Wietnam","Włochy","Zjednoczone Emiraty Arabskie","Zjednoczone Królestwo","Żydowskie święta religijne"))
$Cultures.Add("Portuguese (Brazil)", @("África do Sul","Albânia","Alemanha","Andorra","Angola","Arábia Saudita","Argélia","Argentina","Armênia","Austrália","Áustria","Azerbaijão","Bahrein","Belarus","Bélgica","Bolívia","Bósnia e Herzegovina","Brasil","Brunei","Bulgária","Canadá","Catar","Cazaquistão","Chile","China","Chipre","Cingapura","Colômbia","Congo (República Democrática do)","Coreia","Costa Rica","Croácia","Czechia","Dinamarca","Egito","El Salvador","Emirados Árabes Unidos","Equador","Eslováquia","Eslovênia","Espanha","Estados Unidos","Estônia","Feriados Cristãos","Feriados Judaicos","Feriados Religiosos Islâmicos (Shia)","Feriados Religiosos Islâmicos (Sunni)","Filipinas","Finlândia","França","Geórgia","Grécia","Guatemala","Guiné Equatorial","Honduras","Hungria","Iêmen","Índia","Iraque","Irlanda","Islândia","Israel","Itália","Jamaica","Japão","Jordânia","Kuwait","Letônia","Líbano","Liechtenstein","Lituânia","Luxemburgo","Macedônia do Norte","Malásia","Malta","Marrocos","México","Moldova","Mônaco","Montenegro","Nicarágua","Nigéria","Noruega","Nova Zelândia","Omã","Países Baixos","Panamá","Paquistão","Paraguai","Peru","Polônia","Porto Rico","Portugal","Quênia","RAE de Hong Kong","RAE de Macau","Reino Unido","República Dominicana","Romênia","Rússia","San Marino","Santa Sé (Cidade do Vaticano)","Sérvia","Síria","Suécia","Suíça","Tailândia","Trinidad e Tobago","Tunísia","Turquia","Ucrânia","União Europeia","Uruguai","Venezuela","Vietnã"))
$Cultures.Add("Portuguese (Portugal)", @("África do Sul","Albânia","Alemanha","Andorra","Angola","Arábia Saudita","Argélia","Argentina","Arménia","Austrália","Áustria","Azerbaijão","Barém","Bélgica","Bielorrússia","Bolívia","Bósnia e Herzegovina","Brasil","Brunei","Bulgária","Canadá","Catar","Cazaquistão","Chéquia","Chile","China","Chipre","Colômbia","Congo (República Democrática do)","Coreia","Costa Rica","Croácia","Dinamarca","Egito","Emirados Árabes Unidos","Equador","Eslováquia","Eslovénia","Espanha","Estados Unidos","Estónia","Feriados religiosos cristãos","Feriados Religiosos Islâmicos (Sunitas)","Feriados Religiosos Islâmicos (Xiitas)","Feriados religiosos judaicos","Filipinas","Finlândia","França","Geórgia","Grécia","Guatemala","Guiné Equatorial","Honduras","Hungria","Iémen","Índia","Iraque","Irlanda","Islândia","Israel","Itália","Jamaica","Japão","Jordânia","Kuwait","Letónia","Líbano","Listenstaine","Lituânia","Luxemburgo","Macedónia do Norte","Malásia","Malta","Marrocos","México","Moldova","Mónaco","Montenegro","Nicarágua","Nigéria","Noruega","Nova Zelândia","Omã","Países Baixos","Panamá","Paquistão","Paraguai","Peru","Polónia","Porto Rico","Portugal","Quénia","RAE de Hong Kong","RAE de Macau","Reino Unido","República Dominicana","Roménia","Rússia","Salvador","Santa Sé (Cidade do Vaticano)","São Marinho","Sérvia","Singapura","Síria","Suécia","Suíça","Tailândia","Trindade e Tobago","Tunísia","Turquia","Ucrânia","União Europeia","Uruguai","Venezuela","Vietname"))
$Cultures.Add("Romanian (Romania)", @("Africa de Sud","Albania","Algeria","Andorra","Angola","Arabia Saudită","Argentina","Armenia","Australia","Austria","Azerbaidjan","Bahrain","Belarus","Belgia","Bolivia","Bosnia și Herțegovina","Brazilia","Brunei","Bulgaria","Canada","Cehia","Chile","China","Cipru","Columbia","Congo (Republica Democrată)","Coreea","Costa Rica","Croația","Danemarca","Ecuador","Egipt","El Salvador","Elveția","Emiratele Arabe Unite","Estonia","Filipine","Finlanda","Franța","Georgia","Germania","Grecia","Guatemala","Guineea Ecuatorială","Honduras","India","Iordania","Irak","Irlanda","Islanda","Israel","Italia","Jamaica","Japonia","Kazahstan","Kenya","Kuweit","Letonia","Liban","Liechtenstein","Lituania","Luxemburg","Macedonia de Nord","Malaysia","Malta","Maroc","Mexic","Moldova","Monaco","Muntenegru","Nicaragua","Nigeria","Norvegia","Noua Zeelandă","Oman","Pakistan","Panama","Paraguay","Peru","Polonia","Portugalia","Puerto Rico","Qatar","RAS Hong Kong","RAS Macao","Regatul Unit","Republica Dominicană","România","Rusia","San Marino","Sărbătorile religioase creștine","Sărbătorile religioase evreiești","Sărbătorile religioase islamice (Șiite)","Sărbătorile religioase islamice (Sunnite)","Serbia","Sfântul Scaun (Statul Cetății Vaticanului)","Singapore","Siria","Slovacia","Slovenia","Spania","Statele Unite ale Americii","Suedia","Țările de Jos","Thailanda","Trinidad Tobago","Tunisia","Turcia","Ucraina","Ungaria","Uniunea Europeană","Uruguay","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Russian (Russia)", @("Австралия","Австрия","Азербайджан","Албания","Алжир","Ангола","Андорра","Аргентина","Армения","Бахрейн","Беларусь","Бельгия","Болгария","Боливия","Босния и Герцеговина","Бразилия","Бруней-Даруссалам","Венгрия","Венесуэла","Вьетнам","Гватемала","Германия","Гондурас","Гонконг (САР)","Греция","Грузия","Дания","Демократическая Республика Конго","Доминиканская Республика","Еврейские религиозные праздники","Египет","ЕС","Йемен","Израиль","Индия","Иордания","Ирак","Ирландия","Исламские (суннитские) религиозные праздники","Исламские (шиитские) религиозные праздники","Исландия","Испания","Италия","Казахстан","Канада","Катар","Кения","Кипр","Китай","Колумбия","Корея","Коста-Рика","Кувейт","Латвия","Ливан","Литва","Лихтенштейн","Люксембург","Малайзия","Мальта","Марокко","Мексика","Молдова","Монако","Нигерия","Нидерланды","Никарагуа","Новая Зеландия","Норвегия","ОАЭ","Оман","Пакистан","Панама","Папский Престол (Ватикан)","Парагвай","Перу","Польша","Португалия","Пуэрто-Рико","Россия","Румыния","Сан-Марино","САР Макао","Саудовская Аравия","Северная Македония","Сербия","Сингапур","Сирийская Арабская Республика","Словакия","Словения","Соединенное Королевство","США","Таиланд","Тринидад и Тобаго","Тунис","Турция","Украина","Уругвай","Филиппины","Финляндия","Франция","Хорватия","Христианские религиозные праздники","Черногория","Чехия","Чили","Швейцария","Швеция","Эквадор","Экваториальная Гвинея","Эль-Сальвадор","Эстония","Южная Африка","Ямайка","Япония"))
$Cultures.Add("Serbian (Latin, Serbia and Montenegro (Former))", @("Albania","Algeria","Andorra","Angola","Argentina","Armenia","Australia","Austria","Azerbaijan","Bahrain","Belarus","Belgium","Bolivia","Bosnia and Herzegovina","Brazil","Brunei","Bulgaria","Canada","Chile","China","Christian Religious Holidays","Colombia","Congo (Democratic Republic of)","Costa Rica","Croatia","Cyprus","Czechia","Denmark","Dominican Republic","Ecuador","Egypt","El Salvador","Equatorial Guinea","Estonia","European Union","Finland","France","Georgia","Germany","Greece","Guatemala","Holy See (Vatican City)","Honduras","Hong Kong S.A.R.","Hungary","Iceland","India","Iraq","Ireland","Islamic (Shia) Religious Holidays","Islamic (Sunni) Religious Holidays","Israel","Italy","Jamaica","Japan","Jewish Religious Holidays","Jordan","Kazakhstan","Kenya","Korea","Kuwait","Latvia","Lebanon","Liechtenstein","Lithuania","Luxembourg","Macao S.A.R.","Malaysia","Malta","Mexico","Moldova","Monaco","Montenegro","Morocco","Netherlands","New Zealand","Nicaragua","Nigeria","North Macedonia","Norway","Oman","Pakistan","Panama","Paraguay","Peru","Philippines","Poland","Portugal","Puerto Rico","Qatar","Romania","Russia","San Marino","Saudi Arabia","Serbia","Singapore","Slovakia","Slovenia","South Africa","Spain","Sweden","Switzerland","Syria","Thailand","Trinidad and Tobago","Tunisia","Turkey","Ukraine","United Arab Emirates","United Kingdom","United States","Uruguay","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Serbian (Latin, Serbia)", @("Albanija","Alžir","Andora","Angola","Argentina","Australija","Austrija","Azerbejdžan","Bahrein","Belgija","Belorusija","Bolivija","Bosna i Hercegovina","Brazil","Brunej","Bugarska","Češka","Čile","Crna Gora","Danska","Dominikanska Republika","Egipat","Ekvador","Ekvatorijalna Gvineja","Estonija","Evropska unija","Filipini","Finska","Francuska","Grad-država Vatikan","Grčka","Gruzija","Gvatemala","Holandija","Honduras","Hrišćanski verski praznici","Hrvatska","Indija","Irak","Irska","Islamski (šiitski) religijski praznici","Islamski (sunitski) religijski praznici","Island","Italija","Izrael","Jamajka","Japan","Jemen","Jermenija","Jevrejski verski praznici","Jordan","Južna Afrika","Kanada","Katar","Kazahstan","Kenija","Kina","Kipar","Kolumbija","Kongo (Demokratska Republika)","Koreja","Kostarika","Kuvajt","Letonija","Liban","Lihtenštajn","Litvanija","Luksemburg","Mađarska","Makao S.A.O.","Malezija","Malta","Maroko","Meksiko","Moldavija","Monako","Nemačka","Nigerija","Nikaragva","Norveška","Novi Zeland","Oman","Pakistan","Panama","Paragvaj","Peru","Poljska","Portoriko","Portugalija","Rumunija","Rusija","SAD","Salvador","San Marino","SAO Hongkong","Saudijska Arabija","Severna Makedonija","Singapur","Sirija","Slovačka","Slovenija","Španija","Srbija","Švajcarska","Švedska","Tajland","Trinidad i Tobago","Tunis","Turska","Ujedinjeni Arapski Emirati","Ujedinjeno Kraljevstvo","Ukrajina","Urugvaj","Venecuela","Vijetnam"))
$Cultures.Add("Slovak (Slovakia)", @("Albánsko","Alžírsko","Andorra","Angola","Argentína","Arménsko","Austrália","Azerbajdžan","Bahrajn","Belgicko","Bielorusko","Bolívia","Bosna a Hercegovina","Brazília","Brunej","Bulharsko","Česko","Chorvátsko","Čierna Hora","Čile","Čína","Cyprus","Dánsko","Dominikánska republika","Egypt","Ekvádor","Estónsko","Európska únia","Filipíny","Fínsko","Francúzsko","Grécko","Gruzínsko","Guatemala","Holandsko","Honduras","Hongkong OAO","India","Irak","Írsko","Islamské cirkevné sviatky (sunitské)","Islamské náboženské sviatky (šiítske)","Island","Izrael","Jamajka","Japonsko","Jemen","Jordánsko","Južná Afrika","Kanada","Katar","Kazachstan","Keňa","Kolumbia","Kongo (Konžská demokratická republika)","Kórejská republika","Kostarika","Kresťanské náboženské sviatky","Kuvajt","Libanon","Lichtenštajnsko","Litva","Lotyšsko","Luxembursko","Maďarsko","Makao OAO","Malajzia","Malta","Maroko","Mexiko","Moldavsko","Monako","Nemecko","Nigéria","Nikaragua","Nórsko","Nový Zéland","Omán","Pakistan","Panama","Paraguaj","Peru","Poľsko","Portoriko","Portugalsko","Rakúsko","Rovníková Guinea","Rumunsko","Rusko","Salvádor","San Maríno","Saudská Arábia","﻿Severné Macedónsko","Singapur","Slovensko","Slovinsko","Španielsko","Spojené arabské emiráty","Spojené kráľovstvo","Spojené štáty","Srbsko","Švajčiarsko","Svätá stolica (Vatikán)","Švédsko","Sýria","Taliansko","Thajsko","Trinidad a Tobago","Tunisko","Turecko","Ukrajina","Uruguaj","Venezuela","Vietnam","Židovské cirkevné sviatky"))
$Cultures.Add("Slovenian (Slovenia)", @("Albanija","Alžirija","Andora","Angola","Argentina","Armenija","Avstralija","Avstrija","Azerbajdžan","Bahrajn","Belgija","Belorusija","Bolgarija","Bolivija","Bosna in Hercegovina","Brazilija","Brunej","Češka","Čile","Ciper","Črna gora","Danska","Dominikanska republika","Egipt","Ekvador","Ekvatorialna Gvineja","Estonija","Evropska unija","Filipini","Finska","Francija","Grčija","Gruzija","Gvatemala","Honduras","Hongkong posebna administrativna regija","Hrvaška","Indija","Irak","Irska","Islamski (šiitski) verski prazniki","Islamski (sunitski) verski prazniki","Islandija","Italija","Izrael","Jamajka","Japonska","Jemen","Jordanija","Judovski verski prazniki","Južna Afrika","Južna Koreja","Kanada","Katar","Katoliški verski prazniki","Kazahstan","Kenija","Kitajska","Kolumbija","Kongo (demokratična republika)","Kostarika","Kuvajt","Latvija","Libanon","Lihtenštajn","Litva","Luksemburg","Madžarska","Malezija","Malta","Maroko","Mehika","Moldavija","Monako","Nemčija","Nigerija","Nikaragva","Nizozemska","Norveška","Nova Zelandija","Oman","Pakistan","Panama","Paragvaj","Peru","Poljska","Portoriko","Portugalska","Posebno administativno območje Macao","Romunija","Rusija","Salvador","San Marino","Saudova Arabija","Severna Makedonija","Singapur","Sirija","Slovaška","Slovenija","Španija","Srbija","Švedska","Sveti sedež (Vatikan)","Švica","Tajska","Trinidad in Tobago","Tunizija","Turčija","Ukrajina","Urugvaj","Venezuela","Vietnam","Združene države","Združeni arabski emirati","Združeno kraljestvo"))
$Cultures.Add("Spanish (Spain)", @("Albania","Alemania","Andorra","Angola","Arabia Saudí","Argelia","Argentina","Armenia","Australia","Austria","Azerbaiyán","Baréin","Belarús","Bélgica","Bolivia","Bosnia y Herzegovina","Brasil","Brunéi","Bulgaria","Canadá","Chequia","Chile","China","Chipre","Colombia","Congo (República Democrática del)","Corea del Sur","Costa Rica","Croacia","Dinamarca","Ecuador","Egipto","El Salvador","Emiratos Árabes Unidos","Eslovaquia","Eslovenia","España","Estados Unidos","Estonia","Festividades religiosas cristianas","Festividades religiosas islámicas (chiíes)","Festividades religiosas islámicas (suní)","Festividades religiosas judías","Filipinas","Finlandia","Francia","Georgia","Grecia","Guatemala","Guinea Ecuatorial","Honduras","Hong Kong (RAE)","Hungría","India","Irak","Irlanda","Islandia","Israel","Italia","Jamaica","Japón","Jordania","Kazajistán","Kenia","Kuwait","Letonia","Líbano","Liechtenstein","Lituania","Luxemburgo","Macao RAE","Macedonia del Norte","Malasia","Malta","Marruecos","México","Moldova","Mónaco","Montenegro","Nicaragua","Nigeria","Noruega","Nueva Zelanda","Omán","Países Bajos","Pakistán","Panamá","Paraguay","Perú","Polonia","Portugal","Puerto Rico","Qatar","Reino Unido","República Dominicana","Rumania","Rusia","San Marino","Santa Sede (Ciudad del Vaticano)","Serbia","Singapur","Siria","Sudáfrica","Suecia","Suiza","Tailandia","Trinidad y Tobago","Túnez","Turquía","Ucrania","Unión Europea","Uruguay","Venezuela","Vietnam","Yemen"))
$Cultures.Add("Swedish (Sweden)", @("Albanien","Algeriet","Andorra","Angola","Argentina","Armenien","Australien","Azerbajdzjan","Bahrain","Belgien","Bolivia","Bosnien och Hercegovina","Brasilien","Brunei","Bulgarien","Chile","Colombia","Costa Rica","Cypern","Czechia","Danmark","Dominikanska republiken","Ecuador","Egypten","Ekvatorialguinea","El Salvador","Estland","Europeiska unionen","Filippinerna","Finland","Förenade Arabemiraten","Frankrike","Georgien","Grekland","Guatemala","Heliga stolen (Vatikanstaten)","Honduras","Hongkong","Indien","Irak","Irland","Islamiska (shia) religiösa helgdagar","Islamiska (sunni) religiösa helgdagar","Island","Israel","Italien","Jamaica","Japan","Jemen","Jordanien","Judiska helgdagar","Kanada","Kazakstan","Kenya","Kina","Kongo (Demokratiska republiken)","Kristna helgdagar","Kroatien","Kuwait","Lettland","Libanon","Liechtenstein","Litauen","Luxemburg","Macao","Malaysia","Malta","Marocko","Mexiko","Moldavien","Monaco","Montenegro","Nederländerna","Nicaragua","Nigeria","Nordmakedonien","Norge","Nya Zeeland","Oman","Österrike","Pakistan","Panama","Paraguay","Peru","Polen","Portugal","Puerto Rico","Qatar","Rumänien","Ryssland","San Marino","Saudiarabien","Schweiz","Serbien","Singapore","Slovakien","Slovenien","Spanien","Storbritannien","Sverige","Sydafrika","Sydkorea","Syrien","Thailand","Trinidad och Tobago","Tunisien","Turkiet","Tyskland","Ukraina","Ungern","Uruguay","USA","Venezuela","Vietnam","Vitryssland"))
$Cultures.Add("Thai (Thailand)", @("เกาหลี","เขตบริหารพิเศษมาเก๊า","เขตบริหารพิเศษฮ่องกง","เคนยา","เช็ก","เซอร์เบีย","เดนมาร์ก","เนเธอร์แลนด์","เบลเยียม","เบลารุส","เปรู","เปอร์โตริโก","เม็กซิโก","เยเมน","เยอรมนี","เลบานอน","เวเนซุเอลา","เวียดนาม","เอกวาดอร์","เอลซัลวาดอร์","เอสโตเนีย","แคนาดา","แองโกลา","แอฟริกาใต้","แอลเบเนีย","แอลจีเรีย","โครเอเชีย","โคลัมเบีย","โบลิเวีย","โปแลนด์","โปรตุเกส","โมนาโก","โมร็อกโก","โรมาเนีย","โอมาน","ไซปรัส","ไทย","ไนจีเรีย","ไอซ์แลนด์","ไอร์แลนด์","กรีซ","กัวเตมาลา","กาตาร์","คองโก (สาธารณรัฐประชาธิปไตย)","คอสตาริกา","คาซัคสถาน","คูเวต","จอร์เจีย","จอร์แดน","จาเมกา","จีน","ชิลี","ซานมารีโน","ซาอุดีอาระเบีย","ซีเรีย","ญี่ปุ่น","ตรินิแดดและโตเบโก","ตุรกี","ตูนิเซีย","นอร์เวย์","นิการากัว","นิวซีแลนด์","บราซิล","บรูไน","บอสเนียและเฮอร์เซโกวีนา","บัลแกเรีย","บาห์เรน","ปากีสถาน","ปานามา","ปารากวัย","ฝรั่งเศส","ฟินแลนด์","ฟิลิปปินส์","มอนเตเนโกร","มอลโดวา","มอลตา","มาเลเซีย","มาซิโดเนียเหนือ","ยูเครน","รัสเซีย","ลักเซมเบิร์ก","ลัตเวีย","ลิกเตนสไตน์","ลิทัวเนีย","วันหยุดของศาสนาคริสต์","วันหยุดของศาสนายิว","วันหยุดของศาสนาอิสลาม (ชีอะฮ์)","วันหยุดของศาสนาอิสลาม (ชุนนี)","สเปน","สโลวะเกีย","สโลวีเนีย","สวิตเซอร์แลนด์","สวีเดน","สหภาพยุโรป","สหรัฐอเมริกา","สหรัฐอาหรับเอมิเรตส์","สหราชอาณาจักร","สันตะสำนัก (นครวาติกัน)","สาธารณรัฐโดมินิกัน","สิงคโปร์","ออสเตรเลีย","ออสเตรีย","อันดอร์รา","อาเซอร์ไบจาน","อาร์เจนตินา","อาร์เมเนีย","อิเควทอเรียลกินี","อิตาลี","อินเดีย","อิรัก","อิสราเอล","อียิปต์","อุรุกวัย","ฮอนดูรัส","ฮังการี"))
$Cultures.Add("Turkish (Turkey)", @("Almanya","Andorra","Angola","Arjantin","Arnavutluk","Avrupa Birliği","Avustralya","Avusturya","Azerbaycan","Bahreyn","Belçika","Beyaz Rusya","Birleşik Arap Emirlikleri","Birleşik Devletler","Birleşik Krallık","Bolivya","Bosna-Hersek","Brezilya","Brunei","Bulgaristan","Çekya","Cezayir","Çin","Danimarka","Dominik Cumhuriyeti","Ekvador","Ekvator Ginesi","El Salvador","Ermenistan","Estonya","Fas","Filipinler","Finlandiya","Fransa","Guatemala","Güney Afrika","Gürcistan","Hindistan","Hırvatistan","Hollanda","Honduras","Hong Kong Çin ÖİB","Hristiyan Dini Günleri","Irak","İrlanda","İspanya","İsrail","İsveç","İsviçre","İtalya","İzlanda","Jamaika","Japonya","Kanada","Karadağ","Katar","Kazakistan","Kenya","Kıbrıs","Kolombiya","Kongo (Demokratik Cumhuriyeti)","Kore","Kosta Rika","Kuveyt","Kuzey Makedonya","Letonya","Liechtenstein","Litvanya","Lübnan","Lüksemburg","Macaristan","Makau Çin ÖİB","Malezya","Malta","Meksika","Mısır","Moldova","Monako","Musevi Dini Günleri","Müslüman (Şii) Dini Tatilleri","Müslüman (Sünni) Dini Tatilleri","Nijerya","Nikaragua","Norveç","Pakistan","Panama","Papalık Makamı (Vatikat Şehri)","Paraguay","Peru","Polonya","Portekiz","Porto Riko","Romanya","Rusya","San Marino","Şili","Singapur","Sırbistan","Slovakya","Slovenya","Suriye","Suudi Arabistan","Tayland","Trinidad ve Tobago","Tunus","Türkiye","Ukrayna","Umman","Ürdün","Uruguay","Venezuela","Vietnam","Yemen","Yeni Zelanda","Yunanistan"))
$Cultures.Add("Ukrainian (Ukraine)", @("Австралія","Австрія","Азербайджан","Албанія","Алжир","Анґола","Андорра","Аргентина","Бахрейн","Бельґія","Білорусь","Болгарія","Болівія","Боснія та Герцеґовина","Бразілія","Бруней","В’єтнам","Венесуела","Вірменія","Гондурас","Греція","Грузія","Ґватемала","Данія","Демократична Республіка Конґо","Домініканська Республіка","Еквадор","Екваторіальна Ґвінея","Естонія","Єврейські релігійні свята","Європейський Союз","Єгипет","Ємен","Йорданія","Ізраїль","Індія","Ірак","Ірландія","Ісламські релігійні свята (суніти)","Ісламські релігійні свята (шиїти)","Ісландія","Іспанія","Італія","Казахстан","Канада","Катар","Кенія","Китай","Кіпр","Колумбія","Корея","Коста-Ріка","Кувейт","Латвія","Литва","Ліван","Ліхтенштейн","Люксембурґ","Малайзія","Мальта","Марокко","Мексика","Молдова","Монако","Ніґерія","Нідерланди","Нікараґуа","Німеччина","Нова Зеландія","Норвеґія","Об’єднані Арабські Емірати","Оман","Пакистан","Панама","Параґвай","Перу","Південна Африка","Північна Македонія","Польща","Портуґалія","Пуерто-Ріко","Росія","Румунія","Сальвадор","Сан-Маріно","САР Гонконґ","САР Макао","Саудівська Аравія","Святійший Престол (Ватикан)","Сербія","Сінґапур","Сірія","Словаччина","Словенія","Сполучене Королівство","Сполучені Штати","Таїланд","Тринідад і Тобаґо","Туніс","Туреччина","Угорщина","Україна","Уруґвай","Філіппіни","Фінляндія","Франція","Хорватія","Християнські релігійні свята","Чехія","Чілі","Чорногорія","Швейцарія","Швеція","Ямайка","Японія"))
$Cultures.Add("Vietnamese (Vietnam)", @("Ả rập Saudi","Ai Cập","Ai-len","Albania","Algeria","Ấn Độ","Andorra","Angola","Áo","Argentina","Armenia","Azerbaijan","Ba Lan","Bắc Macedonia","Ba-ranh","Belarus","Bỉ","Bồ Đào Nha","Bolivia","Bosnia và Herzegovina","Bra-zin","Brunei","Bulgaria","Các Tiểu Vương Quốc Ả Rập","Canada","Chi-lê","Colombia","Cộng hòa Dominica","Cộng hòa Sip","Congo (Cộng hòa Dân chủ)","Costa Rica","Croat-ti-a","Đặc Khu Hành chính Hồng Kông","Đặc khu Hành chính Macao","Đan Mạch","Đức","Ecuador","El Salvador","Estonia","Georgia","Guatemala","Guinea Xích đạo","Hà Lan","Hàn Quốc","Hi Lạp","Hoa Kỳ","Honduras","Hungary","Iceland","Iraq","Israel","Italy","Jamaica","Jordan","Kazakhstan","Kenya","Kuwait","Lát-vi-a","Lễ hội tôn giáo đạo Hồi (Shia)","Lễ hội tôn giáo đạo Hồi (Sunni)","Li-băng","Liechtenstein","Liên Hiệp Vương Quốc Anh","Liên Minh Châu Âu","Lít-va","Luxembourg","Malaysia","Malta","Marốc","Mexico","Moldova","Monaco","Montenegro","Na Uy","Nam Phi","New Zealand","Nga","Ngày Cơ đốc giáo","Ngày lễ đạo Do thái","Nhật Bản","Nicaragua","Nigeria","Ô man","Pakistan","Panama","Pa-ra-guay","Peru","Phần Lan","Pháp","Philippines","Puerto Rico","Qatar","Romania","San Marino","Séc","Serbia","Singapore","Slovakia","Slovenia","Syria","Tây Ban Nha","Thái Lan","Thổ Nhĩ Kỳ","Thụy Điển","Thụy Sỹ","Tòa Thánh (Thành Vatican)","Trinidad và Tobago","Trung Quốc","Tunisia","Úc","Ukraine","U-ru-guay","Venezuela","Việt Nam","Yemen"))

#set variables
param (
    [Parameter(Mandatory=$false)][string]$appInsightsKey="yourkey",
    [switch]$runLocal=$true,
    [switch]$save=$false
 )

Clear-Host
$global:currentdir = (Split-Path -Parent $MyInvocation.MyCommand.Path)
Write-host -ForegroundColor darkblue "- Running from location: " $currentdir

function Write-AuditLog([string] $message) {
	$timeStamp = Get-Date
	"[$timeStamp] `t $message" | out-file $logfilename -Append
}

#creating seperate audit log to follow progression and check results
if (! (Test-Path -Path "$global:currentdir\AuditLog")) {
	New-Item -ItemType directory -Path "$global:currentdir\AuditLog"
}

$today = Get-Date;
$daydifferentiator = $today.ToString("yyyyMMdd")

$logfilename = "$global:currentdir\AuditLog\log$daydifferentiator.txt"
if (! (Test-Path -Path $logfilename))
{
    write-host "File doesn't exist. Creating logfile"
    $newfile = New-Item -Path "$global:currentdir\AuditLog" -Name "log$daydifferentiator.txt" -ItemType "File" -Verbose
    Write-AuditLog " -- Logfile Created --"
}

Write-AuditLog "Started logging"

function PerformRegEx() {
    param(
        $regEx,
        $content
    )

    $regExOptions = [Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant'
    $matches = [regex]::match($content,$regEx, $regExOptions)
    if($matches -and $matches.Groups.Length  -gt 0){
        return $matches.Groups[1].Value
    }
    return ""
}

function PerformPageRequest(){
    param(
        [string]$URL,
        [boolean]$isModern
    )

    Write-AuditLog "Start of PerformPageRequest for $URL" 

    $webClient = New-Object System.Net.WebClient 
    $webClient.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName, $password)
    $webClient.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
    $response = $webClient.DownloadString($URL);

    $responseHeadersObject = $webClient.ResponseHeaders;
    $responseHeaders = @{};
    $i = 0;
    

    while ($i -lt $responseHeadersObject.Count)
    {
        Write-Host "$($responseHeadersObject.GetKey($i)) : $($responseHeadersObject.Get($i))";
        $responseHeaders.Add($responseHeadersObject.GetKey($i), $responseHeadersObject.Get($i));
        $i++;
    }
    
    # Read header info
    $XSharePointHealthScore = $responseHeaders["X-SharePointHealthScore"]

    # get correlationId of current request from header
    $aCorrelationId = $responseHeaders["sprequestguid"];

    $timeStamp = Get-Date -format "yyyy-MM-dd HH:mm:ss"
    if($isModern){
        # Modern sites don't have all the performance headers
        # Use json objects in the body to read the performance information
        $perfValues = PerformRegEx -regEx '\"perf\"\s:\s(\{.+?})' -content $response

        if($perfValues -and $perfValues -ne ""){           
            $modernPerfomance = ConvertFrom-Json  -InputObject $perfValues 
            Add-Member -InputObject $modernPerfomance  -MemberType NoteProperty -Name "XSharePointHealthScore" -Value $XSharePointHealthScore
            Add-Member -InputObject $modernPerfomance  -MemberType NoteProperty -Name "URL" -Value $URL
            Add-Member -InputObject $modernPerfomance  -MemberType NoteProperty -Name "SiteType" -Value "Modern"

            #save to application insights
            $SPRequestDuration = $modernPerfomance.spRequestDuration;
            
            $aData = @{ 
                correlationId = $aCorrelationId;
                spRequestDuration = $SPRequestDuration;
                healthScore = $XSharePointHealthScore;
                siteUrl = $URL;
                SiteType = "Modern"
            };

            if ($save)
            {
                SaveReportEntryToApplicationInsights -appInsightsKey $appInsightsKey -customData $aData;
            }
        }        
    }else{
        $SPRequestDuration = $responseHeaders["SPRequestDuration"]
        $SPIisLatency = $responseHeaders["SPIisLatency"]

        if($null -eq $SPRequestDuration){

            $SPRequestDuration = PerformRegEx -regEx '\s+?g_duration\s+?=\s+?(\d+)' -content $response;
        }

        if($null -eq $SPIisLatency){
            $SPIisLatency = PerformRegEx -regEx '\s+?g_iislatency\s+?=\s+?(\d+)' -content $response;            
        }

        $cpuDuration = PerformRegEx -regEx '\s+?g_cpuDuration\s+?=\s+?(\d+)' -content $response
        $queryCount = PerformRegEx -regEx '\s+?g_queryCount\s+?=\s+?(\d+)' -content $response
        $queryDuration = PerformRegEx -regEx '\s+?g_queryDuration\s+?=\s+?(\d+)' -content $response


        $siteRecord = New-Object -TypeName PSObject

        $siteRecord | Add-Member -MemberType NoteProperty -Name "IisLatency" -Value $SPIisLatency
        $siteRecord | Add-Member -MemberType NoteProperty -Name "SPRequestDuration" -Value $SPRequestDuration                
        if ($queryCount -and $queryCount -ne "") {
             $siteRecord | Add-Member -MemberType NoteProperty -Name "QueryCount" -Value $queryCount
        }
        if ($queryDuration -and $queryDuration -ne "") {
             $siteRecord | Add-Member -MemberType NoteProperty -Name "QueryDuration" -Value $queryDuration
        }
        if ($cpuDuration -and $cpuDuration -ne "") {
             $siteRecord | Add-Member -MemberType NoteProperty -Name "CPUDuration" -Value $cpuDuration
        }
        $siteRecord | Add-Member -MemberType NoteProperty -Name "XSharePointHealthScore" -Value $XSharePointHealthScore
        $siteRecord | Add-Member -MemberType NoteProperty -Name "URL" -Value $URL
        $siteRecord | Add-Member -MemberType NoteProperty -Name "SiteType" -Value "Classic"

        #save to application insights
        $aData = @{ 
            correlationId = $aCorrelationId;
            spRequestDuration = $SPRequestDuration;
            healthScore = $XSharePointHealthScore;
            siteUrl = $URL;
            SiteType = "Classic"
        };

        if ($save)
        {
            SaveReportEntryToApplicationInsights -appInsightsKey $appInsightsKey -customData $aData;
        }
    }

    $webClient.Dispose();

    Write-AuditLog "End of PerformPageRequest for $URL" 
}

function SaveReportEntryToApplicationInsights($appInsightsKey, $customData)
{
    $baseData = @{ ver =  2; name = "Event sharepoint monitoring script"; properties = $customData };
    $data = @{ baseType = "EventData"; baseData = $baseData };
    $tags = @{ 'ai.cloud.roleInstance' = $env:computername; 'ai.internal.sdkVersion' = "monitoring-ps:1.0.0" };
    $body = @{ name = "Microsoft.ApplicationInsights.Event"; time = $([System.dateTime]::UtcNow.ToString('o')); iKey = $appInsightsKey; tags = $tags; data = $data };

    $json = $body | ConvertTo-Json -Depth 5;

    Invoke-WebRequest -Uri 'https://dc.services.visualstudio.com/v2/track' -Method 'POST' -UseBasicParsing -body $json
}

#########################################
## START LOGIC TO GET PERFORMANCE INFO ##
#########################################

if ($runLocal)
{
    [System.Reflection.Assembly]::LoadFile("C:\PathTo\Microsoft.SharePoint.Client.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFile("C:\PathTo\Microsoft.SharePoint.Client.Runtime.dll") | Out-Null 

    $creds = Get-Credential;
}
else
{
    $creds = Get-AutomationPSCredential -Name "SPOScriptAutomation";
}

$userName = $creds.Username;
$password = $creds.Password;

$modernSites = @("https://yourtenant.sharepoint.com/sites/modernsite1","https://yourtenant.sharepoint.com/sites/modernsite2");
$classicSites = @("https://yourtenant.sharepoint.com/sites/classicsite1/Pages/AllResults.aspx?k=test","https://yourtenant.sharepoint.com/sites/classicsite2/Pages/AllResults.aspx?k=test");

foreach($aModernSite in $modernSites) {
    PerformPageRequest -URL $aModernSite -isModern $true  
}

foreach($aClassicSite in $classicSites) {
    PerformPageRequest -URL $aClassicSite -isModern $false  
}

Write-AuditLog "Ended logging"
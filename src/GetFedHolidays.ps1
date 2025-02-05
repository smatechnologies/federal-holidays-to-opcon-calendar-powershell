<#
This script goes out to the Federal Reserve website, grabs the information about the holidays
and adds the dates to an OpCon calendar.  The parameters can be set here or passed in on
the OpCon job.

You can use traditional MSGIN functionality or the OpCon API.  Also added a "debug" option
so that you can view the dates that will be added.

Author: Bruce Jernell
Version: 1.6
#>
param(
    $msginPath,                                               # Path to MS LSAM MSGIN directory
    $url,                                                     # OpCon API URL
    $apiUser,                                                 # OpCon API User
    $apiPassword,                                             # OpCon API Password
    $extUser,                                                 # OpCon External Event user
    $extPassword,                                             # OpCon External Event password
    $extToken,                                                # OpCon External Token (OpCon Release 20+, can be used instead of API User/PW for API option)
    $calendar,                                                # OpCon Calendar (ex "Master Holiday")
    $option = "debug"                                         # Script option: "api", "msgin", "debug" (displays dates but makes no changes)
)

#Stop on any error
$ErrorActionPreference = 'Stop'

if($option -eq "api")
{
    #Skip Self Signed Certificates in PS 5 or lower
    if($PSVersionTable.PSVersion.Major -le 5)
    {
        try
        {
            Add-Type -TypeDefinition  @"
            using System.Net;
            using System.Security.Cryptography.X509Certificates;
            public class TrustAllCertsPolicy : ICertificatePolicy
            {
                public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem)
                {
                    return true;
                }
            }
"@
            [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
        }
        catch
        { Write-Output "Error Ignoring Self Signed Certs" }
    }

    #Set TLS version
    if($PSVersionTable.PSVersion.Major -eq 7 -and $PSVersionTable.PSVersion.Minor -gt 1) 
    { 
        #Force TLS 1.3
        Write-Output "Using TLS 1.3"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls13 
    }
    else
    {
        #Force TLS 1.2
        Write-Output "Using TLS 1.2"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }

    if($extToken)
    { 
        if($extToken -like "Token*")
        { $token = $extToken }
        else
        { $token = "Token " + $extToken }
    }
    else
    { 
        $body = @{
            "user"= @{"loginName"=$apiuser;"password"=$apipassword};
            "tokenType"=@{"type"="User"}
        }        
                
        try
        {
            if($PSVersionTable.PSVersion.Major -lt 6)
            { $token = "Token " + (Invoke-Restmethod -Method POST -Uri ($url + "/api/tokens") -Body ($body | ConvertTo-Json) -ContentType "application/json").id }
            else
            { $token = "Token " + (Invoke-Restmethod -Method POST -Uri ($url + "/api/tokens") -Body ($body | ConvertTo-Json) -ContentType "application/json" -SkipCertificateCheck).id }
        }
        catch [Exception]
        {
            write-output $_
            write-output $_.Exception.Message
            Exit 1
        }    
    }
}
elseif($option -eq "msgin")
{
    if($msginPath)
    {
        if(test-path $msginPath)
        {   Write-Output "$msginPath path exists" }
        else
        {
            Write-Output "$msginPath path does not exist"
            Exit 101
        }
    }
    else
    {
        Write-Output "MSGIN Path parameter must be specified!"
        Exit 102
    }    
}
elseif($option -ne "debug")
{
    Write-Output 'Invalid option, must be "api", "msgin", or "debug".'
    Exit 100
}

$months = @("January","February","March","April","May","June","July","August","September","October","November","December")
$holidays = @("New Year's Day","Birthday of Martin Luther King","Washington's Birthday","Memorial Day","Juneteenth National","Independence Day","Labor Day","Columbus Day","Veterans Day","Thanksgiving Day","Christmas Day")
$result = Invoke-RestMethod -Uri "https://www.federalreserve.gov/aboutthefed/k8.htm"
$source = $result.Split("`n",[StringSplitOptions]::RemoveEmptyEntries)

# Parses through the HTML source code and extracts the dates
for($x=0;$x -lt ($source.Count-1);$x++)
{
    if($source[$x] -like "*K.8 - Holidays Observed by the Federal Reserve System*")
    {
        $yearstart = $source[$x].Substring(58,4)
        $yearend = $source[$x].Substring(63,4)
    }
    elseif($holidays | ForEach-Object{ if($source[$x] -like ("*" + $_ + "*")){ return $true }})
    {
        $x++
        for($y=0;$y -le ($yearend-$yearstart);$y++)
        {
            If($source[$x] -notlike "*div>*" -and $source[$x] -ne "")
            {
                $source[$x] = $source[$x].Trim() # Remove starting spaces (normalize in case this is changed)
                $date = $source[$x].Substring(4,$source[$x].IndexOf("</td>")-($source[$x].IndexOf("<td>")+4)) #Formerly 7

                $year = [int]$yearstart + $y
                if($date.indexof('>*') -ge 0)   # Formerly ($date -like "*<*")
                {
                    # Match asterisks at end of string
                    # Version 1.32 Bug fix for January **
                    $getStars = ($date | Select-String -Pattern '[*]' -AllMatches).Matches.Value -join ''
                    $date = $date.Substring(0,$date.IndexOf("<a")) # Removes the html
                    $date = [string]($months.IndexOf($date.SubString(0,$date.IndexOf(" ")))+1) + "/" + [string]$date.SubString(($date.IndexOf(" ")+1),($date.Length - ($date.IndexOf(" ")+1))) + "/" + [string]$year + $getStars
                }
                elseif($date.IndexOf('*') -ge 0) 
                { 
                    if($date.IndexOf('***') -ge 0)
                    { $date = [string]($months.IndexOf($date.SubString(0,$date.IndexOf(" ")))+1) + "/" + [string]$date.SubString(($date.IndexOf(" ")+1),($date.Length - ($date.IndexOf(" ")+1))) + "/" + [string]$year }
                    elseif($date.IndexOf('**') -ge 0)
                    { $date = [string]($months.IndexOf($date.SubString(0,$date.IndexOf(" ")))+1) + "/" + [string]$date.SubString(($date.IndexOf(" ")+1),($date.Length - ($date.IndexOf(" ")+1))) + "/" + [string]$year + "**" }
                    else 
                    { $date = [string]($months.IndexOf($date.SubString(0,$date.IndexOf(" ")))+1) + "/" + [string]$date.SubString(($date.IndexOf(" ")+1),($date.Length - ($date.IndexOf(" ")+1))) + "/" + [string]$year + "*" }
                }
                Else
                { $date = [string]($months.IndexOf($date.SubString(0,$date.IndexOf(" ")))+1) + "/" + [string]$date.SubString(($date.IndexOf(" ")+1),($date.Length - ($date.IndexOf(" ")+1))) + "/" + [string]$year }
                
                # Add/Subtract dates based on how the holiday falls
                if($date.Substring($date.Length - 2) -eq "**")
                {
                    $date = $date.Replace('**',"")
                    $date = Get-Date -Date (Get-Date -Date $date).AddDays(+1) -Format "MM/dd/yyyy"
                }
                elseif($date.IndexOf("***") -ge 0) 
                { $date = $date.Replace('***',"") } 
                elseif($date.Substring($date.Length - 2) -eq "*<" -or $date.Substring($date.Length - 1) -eq "*")
                {
                    if($date.Substring($date.Length - 2) -eq "*<")
                    { $date = $date.Replace('*<',"") }
                    else 
                    { $date = $date.Replace('*',"") }
                }

                $date = (Get-Date -date $date -Format "M/d/yyyy")
                Write-Output ("Date to add: " + $date)

                if($option -ne "debug")
                {
                    if(!$calendar)
                    {
                        Write-Output "No calendar name specified!"
                        Exit 103
                    }

                    if($option -eq "api")
                    {
                        $counter = 0

                        try
                        {
                            $counter = 0

                            if($PSVersionTable.PSVersion.Major -lt 6)
                            { $getcalendar = Invoke-RestMethod -Method GET -Uri ($url + "/api/calendars?name=" + $calendar) -Headers @{"authorization" = $token} -ContentType "application/json" }
                            else 
                            { $getcalendar = Invoke-RestMethod -Method GET -Uri ($url + "/api/calendars?name=" + $calendar) -Headers @{"authorization" = $token} -ContentType "application/json" -SkipCertificateCheck }

                            $getcalendar | ForEach-Object{ $counter++ } 
                            
                            if($counter -eq 0)
                            {
                                Write-Output "No calendars found by supplied name!"
                                Exit 3
                            }
                        }
                        catch [Exception]
                        {
                            Write-Output $_
                            Write-Output $_.Exception.Message
                            Exit 3
                        }
                    
                        if($date)
                        {  
                            for($z=0;$z -lt $getcalendar[0].dates.Count;$z++)
                            {
                                $getcalendar[0].dates[$z] = Get-Date -date $getcalendar[0].dates[$z] -Format "M/d/yyyy"
                            }

                            if($date -notin $getcalendar[0].dates)
                            { 
                                $getcalendar[0].dates += $date
                                $body = @{
                                    "id" = $getcalendar[0].Id;
                                    "name" = $calendar;
                                    "dates" = $getcalendar[0].dates
                                }
        
                                try
                                { 
                                    if($PSVersionTable.PSVersion.Major -lt 6)
                                    { $update = Invoke-RestMethod -Method PUT -Uri ($url + "/api/calendars/" + $getcalendar[0].id) -Body ($body | ConvertTo-JSON -Depth 5) -Headers @{"authorization" = $token} -ContentType "application/json" }
                                    else 
                                    { $update = Invoke-RestMethod -Method PUT -Uri ($url + "/api/calendars/" + $getcalendar[0].id) -Body ($body | ConvertTo-JSON -Depth 5) -Headers @{"authorization" = $token} -ContentType "application/json" -SkipCertificateCheck}
                                }
                                catch [Exception]
                                {
                                    Write-Output $_
                                    Write-Output $_.Exception.Message
                                    Exit 2
                                }
                            }
                            else 
                            { Write-Output "Date $date already in calendar $calendar !" }
                        }
                        else
                        { 
                            Write-Output "Issue getting date from Federal Reserve site" 
                            Exit 5
                        }          
                    }
                    elseif($option -eq "msgin")
                    {
                        if($extToken)
                        { $extPassword = $extToken }

                        Write-Output ("Sending date $date to calendar $calendar via MSGIN")
                        "`$CALENDAR:ADD,$calendar,$date,$extUser,$extPassword" | Out-File -FilePath ($msginPath + "\events$x.txt") -Encoding ascii
                    }
                }

                $x++
            }
        }  
    }
}
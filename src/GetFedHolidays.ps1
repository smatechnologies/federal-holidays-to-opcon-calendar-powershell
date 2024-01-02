<#
This script goes out to the Federal Reserve website, grabs the information about the holidays
and adds the dates to an OpCon calendar.  The parameters can be set here or passed in on
the OpCon job.

You can use traditional MSGIN functionality or the OpCon API.  Also added a "debug" option
so that you can view the dates that will be added.

Author: Bruce Jernell
Version: 1.4
#>
param(
    $opconmodule,                                             # Path to OpCon API function module
    $msginPath,                                               # Path to MS LSAM MSGIN directory
    $url,                                                     # OpCon API URL
    $apiUser,                                                 # OpCon API User
    $apiPassword,                                             # OpCon API Password
    $extUser,                                                 # OpCon External Event user
    $extPassword,                                             # OpCon External Event password
    $extToken,                                                # OpCon External Token (OpCon Release 20+)
    $calendar,                                                # OpCon Calendar (ex "Master Holiday")
    $option = "debug"                                         # Script option: "api", "msgin", "debug"
)

$ErrorActionPreference = 'Stop'

#Force TLS 1.2/1.3
if($PSVersionTable.PSVersion.Major -le 5)
{
    Write-Output "The Federal Holiday script for OpCon now only supports Powershell version 6+"
    Write-Output 'Check your PowerShell version or update your OpCon script runner to something like: pwsh.exe -ExecutionPolicy Bypass -File $FILE $ARGUMENTS'
    Exit 99
}
elseif($PSVersionTable.PSVersion.Major -eq 6 -or ($PSVersionTable.PSVersion.Major -eq 7 -and $PSVersionTable.PSVersion.Minor -lt 1))
{ 
    Write-Output "Using TLS 1.2"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}
else 
{ 
    Write-Output "Using TLS 1.3"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls13 
}

if($option -eq "api")
{
    #Verifies opcon module exists and is imported
    if(Test-Path $opconmodule)
    {
        #Import needed module
        Import-Module -Name $opconmodule -Force #-Verbose  #If you uncomment this option you will see a list of all functions

        #Skip self signed certificates (OpCon API default)
        OpCon_SkipCerts

        if($extToken)
        { 
            if($extToken -like "Token*")
            { $token = $extToken }
            else
            { $token = "Token " + $extToken }
        }
        else
        { $token = "Token " + (OpCon_Login -url $url -user $apiUser -password $apiPassword).id }
    }
    else
    {
        Write-Host "Unable to import OpCon API module!"
        Exit 100
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
$holidays = @("New Year*Day","Martin Luther King","Washington*Birthday","Memorial Day","Juneteenth National Independence Day","Independence Day","Labor Day","Columbus Day","Veterans Day","Thanksgiving Day","Christmas Day")
$result = Invoke-RestMethod -Uri "https://www.federalreserve.gov/aboutthefed/k8.htm"
$source = $result.Split("`n",[StringSplitOptions]::RemoveEmptyEntries)

# Parses through the HTML source code and extracts the dates
for($x=0;$x -lt ($source.Count-1);$x++)
{
    if($source[$x] -like "*Holidays Observed by the Federal Reserve System*")
    {
        $yearstart = $source[$x].Substring(58,4)
        $yearend = $source[$x].Substring(63,4)
    }
    elseif($holidays | ForEach-Object{ if($source[$x] -like "*" + $_ + "*"){ return $true }})
    {
        $x++
        for($y=0;$y -le ($yearend-$yearstart);$y++)
        {
            If($source[$x] -notlike "*div>*" -and $source[$x] -ne "")
            {
                $date = $source[$x].Substring(7,$source[$x].IndexOf("</td>")-($source[$x].IndexOf("<td>")+4))

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
                    if($option -eq "api")
                    {   
                        OpCon_UpdateCalendar -url $url -token $token -name $calendar -date $date
                        if($error)
                        { 
                            Write-Output $error
                            Write-Output "There was a problem updating calendar "$calendar" with date "$date 
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
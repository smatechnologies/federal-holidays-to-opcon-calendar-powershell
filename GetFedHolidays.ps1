<#
This script goes out to the Federal Reserve website, grabs the information about the holidays
and adds the dates to an OpCon calendar.  The parameters can be set here or passed in on
the OpCon job.

You can use traditional MSGIN functionality or the OpCon API (if you have the license).

Author: Bruce Jernell
Version: 1.0

#>
param(
    $opconmodule = "C:\ProgramData\OpConxps\Demo\OpCon.psm1", # Path to OpCon API function module
    $msginPath = "C:\ProgramData\OpConxps\MSLSAM\MSGIN",      # Path to MS LSAM MSGIN directory
    $url,                                                     # OpCon API URL
    $apiUser,                                                 # OpCon API User
    $apiPassword,                                             # OpCon API Password
    $extUser,                                                 # OpCon External Event user
    $extPassword,                                             # OpCon External Event password
    $calendar,                                                # OpCon Calendar (ex "Master Holiday")
    $option = "msgin"                                         # Script option, "api" or "msgin"
)

if($option -eq "api")
{
    #Verifies opcon module exists and is imported
    if(Test-Path $opconmodule)
    {
        #Verify PS version is at least 3.0
        if($PSVersionTable.PSVersion.Major -ge 3)
        {
            # Import needed module
            Import-Module -Name $opconmodule #-Verbose  #If you uncomment this option you will see a list of all functions      
        }
        else
        {
            Write-Host "Powershell version needs to be 3.0 or higher!"
            Exit 100
        }
    }
    else
    {
        Write-Host "Unable to import SMA API module!"
        Exit 100
    }

    $token = "Token " + (OpCon_Login -url $url -user $apiUser -password $apiPassword).id
}
elseif($option -eq "msgin")
{
    if($msginPath)
    {
        if(test-path $msginPath)
        {   Write-Host "$msginPath path exists" }
        else
        {
            Write-Host "$msginPath path does not exist"
            Exit 101
        }
    }
    else
    {
        Write-Host "MSGIN Path parameter must be specified!"
        Exit 102
    }    
}
else
{
    Write-Host 'Invalid option, must be "api" or "msgin"!'
    Exit 100
}

$months = @("January","February","March","April","May","June","July","August","September","October","November","December")
$holidays = @("New Year*Day","Martin Luther King","Washington*Birthday","Memorial Day","Independence Day","Labor Day","Columbus Day","Veterans Day","Thanksgiving Day","Christmas Day")
$req = [System.Net.WebRequest]::Create("https://www.federalreserve.gov/aboutthefed/k8.htm")
$resp = $req.GetResponse()
$reqstream = $resp.GetResponseStream()
$stream = New-Object System.IO.StreamReader $reqstream
$result = @($stream.ReadToEnd())
$source = $result.Split([Environment]::NewLine,[StringSplitOptions]::RemoveEmptyEntries)

# Parses through the HTML source code and extracts the dates
for($x=0;$x -lt ($source.Count-1);$x++)
{
    if($source[$x] -like "*Holidays Observed by the Federal Reserve System*")
    {
        $yearstart = $source[$x].Substring(66,4)
        $yearend = $source[$x].Substring(71,4)
    }
    elseif($holidays | ForEach-Object{ if($source[$x] -like "*" + $_ + "*"){ return $true }})
    {
        $x++
        for($y=0;$y -le ($yearend-$yearstart);$y++)
        {
            $date = ($source[$x].Substring(8,$source[$x].IndexOf("</td>")-($source[$x].IndexOf("<td>")+4))).Replace('*',"")
            
            if($date -like "*<*")
            {
                $date = $date.Substring(0,$date.IndexOf("<"))
            }

            $year = [int]$yearstart + $y
            $date = [string]($months.IndexOf($date.SubString(0,$date.IndexOf(" ")))+1) + "/" + [string]$date.SubString(($date.IndexOf(" ")+1),($date.Length - ($date.IndexOf(" ")+1))) + "/" + [string]$year
            
            if($option -eq "api")
            {
                OpCon_UpdateCalendar -url $url -token $token -name $calendar -date $date
            }
            elseif($option -eq "msgin")
            {
                "`$CALENDAR:ADD,$calendar,$date,$extUser,$extPassword" | Out-File -FilePath ($msginPath + "\events$x.txt") -Encoding ascii
            }

            $x++
        }  
    }
}

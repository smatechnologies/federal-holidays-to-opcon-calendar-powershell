# Federal Holidays to OpCon Calendar
This script goes out to the Federal Reserve website: <https://www.federalreserve.gov/aboutthefed/k8.htm>, grabs the information about the holidays and adds the dates to an OpCon calendar.  The parameters can be set here or passed in on the OpCon job.

You can use traditional MSGIN functionality or the OpCon API.  The debug option can also be used if you are testing or wish to see the dates that are being added, without sending anything to OpCon.

# Prerequisites
* Powershell v5.1
* OpCon Release 17+

# Instructions
  * <b>OpConModule</b> - Path to the OpCon API module, if you are using the API to add the holidays
  * <b>MSGINPath</b> - Path to the MSGIN directory on the server you are running the script, if you are using OpCon external events to add the holidays
  * <b>URL</b> - OpCon API url
  * <b>APIUser</b> - OpCon API user 
  * <b>APIPassword</b> - OpCon API password
  * <b>ExtUser</b> - External OpCon user
  * <b>ExtPassword</b> - External OpCon user event password or new token (OpCon Release 20+)
  * <b>ExtToken</b> - External Token (OpCon Release 20+, only valid for API option)
  * <b>Calendar</b> - OpCon calendar to add the holidays too
  * <b>Option</b> - "api", "msgin", "debug"
  
Example MSGIN Pre OpCon20:
```
powershell.exe -ExecutionPolicy Bypass -File "C:\SMA_GetFedHolidays.ps1" -option "msgin" -msginPath "C:\ProgramData\OpConxps\MSLSAM\MSGIN" -extuser "myuser" -extpassword "mypassword" -calendar "Master Holiday"
```  

Example MSGIN OpCon20 and higher:
```
powershell.exe -ExecutionPolicy Bypass -File "C:\SMA_GetFedHolidays.ps1" -option "msgin" -msginPath "C:\ProgramData\OpConxps\MSLSAM\MSGIN" -extuser "myuser" -extToken "mytoken" -calendar "Master Holiday"
```  

Example API Pre OpCon20:
```
powershell.exe -ExecutionPolicy Bypass -File "C:\SMA_GetFedHolidays.ps1" -option "api" -opconmodule "[[apiPath]]" -url "hostname:Port" -apiUser "myuser" -apiPassword "mypassword" -calendar "Master Holiday"
```  

Example API OpCon20 and higher:
```
powershell.exe -ExecutionPolicy Bypass -File "C:\SMA_GetFedHolidays.ps1" -option "api" -opconmodule "[[apiPath]]" -url "hostname:Port" -extToken "mytoken" -calendar "Master Holiday"
```  

# Disclaimer
No Support and No Warranty are provided by SMA Technologies for this project and related material. The use of this project's files is on your own risk.

SMA Technologies assumes no liability for damage caused by the usage of any of the files offered here via this Github repository.

# License
Copyright 2020 SMA Technologies

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at [apache.org/licenses/LICENSE-2.0](http://www.apache.org/licenses/LICENSE-2.0)

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

# Contributing
We love contributions, please read our [Contribution Guide](CONTRIBUTING.md) to get started!

# Code of Conduct
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-v2.0%20adopted-ff69b4.svg)](code-of-conduct.md)
SMA Technologies has adopted the [Contributor Covenant](CODE_OF_CONDUCT.md) as its Code of Conduct, and we expect project participants to adhere to it. Please read the [full text](CODE_OF_CONDUCT.md) so that you can understand what actions will and will not be tolerated.

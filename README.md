# get-fortigate-routes
Export FortiGate Routes to Excel for your Viewing Pleasure

## Description
This script will populate an Excel file with routing information from FortiGates defined in user-provided CSV file.

## On FortiGate, Create API User, Access profile & API Token
### Create API User Access Profile
```
config system accprofile
 edit <accprofile-name>
  set secfabgrp read
  set ftviewgrp read
  set authgrp read
  set sysgrp read
  set netgrp read
  set loggrp read
  set fwgrp read
  set vpngrp read
  set utmgrp read
  set wifi read
 next
end
```

### Create API User
```
config system api-user
 edit <api-username>
  set accprofile <accprofile-name>
  set vdom "root"
 next
end
```

### Generate & Document API Token
```
exec api-user generate-key <api-username>
```

## Create tab-deliminated CSV File. Example below:
```
name&nbspip&nbsptoken&nbspapiuser
FW1	10.0.0.10:1443	qwertyuiopasdfghjklzxcvbnm1234	apiuser
FW2	192.168.1.1	qwertyuiopasdfghjklzxcvbnm1234	apiuser
FW3	100.64.255.1	qwertyuiopasdfghjklzxcvbnm1234	apiuser
FW4	172.16.1.1:8443	qwertyuiopasdfghjklzxcvbnm1234	apiuser
FW5	10.255.0.1:10443	qwertyuiopasdfghjklzxcvbnm1234	apiuser
```

## Installation
1. Install openpyxl and pandas libraries: `pip install openpyxl pandas`.
2. Clone repository: `git clone https://github.com/nicksmom/get-fortigate-routes.git`

## Execution
1. Execute script: `python3 get-fortigate-routes.py`
2. Specify desired VDOMs to collect routing information from.
3. Specify CSV file
4. Script dumps routes to file: `fortigate-routes.xlsx`

# ClevoPyControlCenter
## A python API act as Control Center to control your Clevo Laptop via wmi on windows.
**Only works on Control Center 2.0 models. Control Center 3.0 model use [acpi dsm](https://docs.microsoft.com/en-us/windows-hardware/drivers/bringup/acpi-device-specific-methods) to control device,thus wmi will not work.**
### Usage
```python
# pip3 install pymi
from clevo_wmi import CLEVO_GET
CG = CLEVO_GET()
# set fan1(CPU Fan) speed to max
CG.SetFanSpeedPercent(1,100)
# set fan speed to auto control by ec
CG.SetFansAuto()
# or just get fan info (temperature rpm_percent ..)
# fan2 means GPU fan in my device
print(CG.GetFanInfo(2))
# you can call method in root\WMI namespace
```
### dependents
clevomof.dll that install automatically by control center V2.0
run CheckWMI.py as administrator to check 
### Implement your method
* open cmd
* excute wbemtest
* connect to "root\WMI" name space (if erro happen , run CheckWMI.py)
* open class "CLEVO_GET" or CLEVO_*
* if no error,there are many items in methods ,choose one and call like clevo_wmi.py do

### Acknowledgement


* [PyMI](https://github.com/cloudbase/PyMI)
* [tuxedo-control-center](https://github.com/tuxedocomputers/tuxedo-control-center)
* [clevo-xsm-wmi](https://github.com/rafaelgieschke/clevo-xsm-wmi)
* [LED-manager-for-Clevo](https://github.com/rzrwolf/LED-manager-for-Clevo.git)
Uses the OSD module.

This script will do 3 things. The primary being download and install the latest driver and firmware pack for any MS Surface device running Windows 10/11 (x64) that Microsoft has in its catalog.
If it does not detect a Surface device, it checks to see if it is a Dell device. If a Dell device and serial number is detected, it then pulls up the driver download page for that specific Dell serial number.
If it can not do either of those, it then falls back to pulling drivers directly from Windows Update and matches the drivers via hardware ID, saves them to C:\Techsupp\Drivers, and then installs them using the Windows pnputil.exe

#=======================================================================
#  Driver Management for Microsoft Surface and other devices
#=======================================================================
#Check if the "OSD" module is installed
Write-Host "Checking for presense of OSD PowerShell module..."
if (Get-Module -ListAvailable -Name OSD) {
    Write-Host "OSD module found. Updating to the latest version."
	Write-Host "Please wait..."
    Update-Module -Name OSD -Force 3> $null 6> $null
} else {
    Write-Host "OSD module not found. Installing..."
    Install-Module -Name OSD -Force -Scope CurrentUser
}

# Import the "OSD" module
Import-Module OSD

function Get-RegistryValue {
    param (
        [string]$regQueryOutput,
        [string]$valueName
    )
    $lines = $regQueryOutput -split "`n"
    foreach ($line in $lines) {
        if ($line -match "$valueName\s+REG_SZ\s+(.+)") {
            return $matches[1].Trim()
        }
    }
    return $null
}

$Manufacturer = reg query "HKLM\SYSTEM\HardwareConfig\Current" /v BaseBoardManufacturer
$BaseBoardManufacturer = Get-RegistryValue -regQueryOutput $Manufacturer -valueName "BaseBoardManufacturer"

$model = reg query "HKLM\SYSTEM\HardwareConfig\Current" /v BaseBoardProduct
$BaseBoardProduct = Get-RegistryValue -regQueryOutput $model -valueName "BaseBoardProduct"

$CSName = reg query "HKLM\Software\Microsoft\Windows NT\CurrentVersion" /v ProductName
$ProductName = Get-RegistryValue -regQueryOutput $CSName -valueName "ProductName"

$wmi_os = reg query "HKLM\Software\Microsoft\Windows NT\CurrentVersion" /v CurrentBuildNumber
$CurrentBuildNumber = Get-RegistryValue -regQueryOutput $wmi_os -valueName "CurrentBuildNumber"

switch ($CurrentBuildNumber) {
    '10240' { $wmi_build = "1507" }
    '10586' { $wmi_build = "1511" }
    '14393' { $wmi_build = "1607" }
    '15063' { $wmi_build = "1703" }
    '16299' { $wmi_build = "1709" }
    '17134' { $wmi_build = "1803" }
    '17686' { $wmi_build = "1809" }
    '18362' { $wmi_build = "1903" }
    '18363' { $wmi_build = "1909" }
    '19041' { $wmi_build = "2004" }
    '19042' { $wmi_build = "20H2" }
    '19043' { $wmi_build = "21H1" }
    '19045' { $wmi_build = "22H2" } 
    '22000' { $wmi_build = "21H2"; $ProductName = $ProductName -replace "Windows 10", "Windows 11" }
    '22621' { $wmi_build = "22H2"; $ProductName = $ProductName -replace "Windows 10", "Windows 11" }
    '22631' { $wmi_build = "23H2"; $ProductName = $ProductName -replace "Windows 10", "Windows 11" }
    '26100' { $wmi_build = "24H2"; $ProductName = $ProductName -replace "Windows 10", "Windows 11" }
}

Write-Host ""
Write-Host ""
Write-Host -ForegroundColor Cyan "Device manufacturer is $BaseBoardManufacturer."
Write-Host -ForegroundColor Cyan "Model: $BaseBoardProduct"
Write-Host -ForegroundColor Cyan "OS: $ProductName"
Write-Host -ForegroundColor Cyan "Version: $wmi_build"

$Get_Manufacturer_Info = (Get-WmiObject Win32_ComputerSystem).Manufacturer
if ($Get_Manufacturer_Info -like "*Microsoft*")	
{									
    $Get_Product_Info = (Get-MyComputerProduct)
    Write-Host ""
    Write-Host ""
    Write-Host "Surface device detected."
	Write-Host "Updating Surface driver catalog. Please wait..."
	Invoke-RestMethod "https://raw.githubusercontent.com/everydayintech/OSDCloud-Public/main/Catalogs/Update-OSDCloudSurfaceDriverCatalogJustInTime.ps1" | Invoke-Expression
	Update-OSDCloudSurfaceDriverCatalogJustInTime -UpdateDriverPacksJson 3> $null 6> $null
    
    Write-Host ""
    Write-Host ""
    Write-Host -ForegroundColor Gray "Getting Driver Package for this $BaseBoardProduct"
    $DriverPack = Get-OSDCloudDriverPacks | Where-Object {($_.Product -contains $Get_Product_Info) -and ($_.OS -match $Params.OSVersion)}	

    if ($DriverPack) {
        [System.String]$DownloadPath = 'C:\Techsupp'
        if (-NOT (Test-Path "$DownloadPath")) {
            New-Item $DownloadPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        # Get the current OS version
        $currentOS = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption

        # Determine the OS version string to match
        if ($currentOS -match "Windows 10") {
            $osVersionString = "Win10"
        } elseif ($currentOS -match "Windows 11") {
            $osVersionString = "Win11"
        } else {
            Write-Host "Unsupported OS version"
            exit
        }

        # Function to download driver pack
        function DownloadDriverPack {
            param (
                [Parameter(Mandatory=$true)]
                [PSObject]$pack
            )

            $url = [string]$pack.Url
            
            # Check for null or empty values
            if ([string]::IsNullOrEmpty($url)) {
                Write-Host "DriverPack.Url is null or empty"
            } else {
                # Download only if the URL matches the OS version string
                if ($url -match $osVersionString) {
                    $OutFile = Join-Path -Path $DownloadPath -ChildPath SurfaceDrivers.msi
                    Write-Host "Downloading. Please wait..."
                    Save-WebFile -SourceUrl $url -DestinationDirectory $DownloadPath -DestinationName SurfaceDrivers.msi
                    
                    # Check if the file was downloaded
                    if (Test-Path $OutFile) {
                        $TestFile = Get-Item $OutFile
                        $SizeInMB = $TestFile.Length / 1MB
                        if ($SizeInMB -ge 100) {
                            # Save the DriverPack details to a JSON file
                            $pack | ConvertTo-Json | Out-File "$OutFile.json" -Encoding ascii -Width 2000 -Force
                            Get-Content "$OutFile.json"
                            del "$OutFile.json"
                            Write-Host ""
                            Write-Host ""
                            Write-Host -ForegroundColor Cyan "Installing Microsoft Surface drivers and firmware. Please wait..."
                            Start-Process msiexec -ArgumentList '/i "C:\Techsupp\SurfaceDrivers.msi" /qb /norestart' -Wait
                            Write-Host -ForegroundColor Green "Driver install complete."
                            Write-Host -ForegroundColor Yellow "Press any key to reboot the computer or Ctrl+C to cancel"
                            pause
                            Restart-Computer -Force
                    } else {
                            Write-Host ""
                            Write-Host ""
                            Write-Warning "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Driver Pack failed to download. Pulling drivers from Windows Update..."                            
                            Write-Host ""
                            Write-Host ""
                            Write-Host "Matching and downloading drivers. Please wait..."
                            Write-Host "Saving drivers to C:\Techsupp\Drivers..."
                            Save-MsUpCatDriver -DestinationDirectory C:\TechSupp\Drivers
                            Write-Host -ForegroundColor Green "Windows Update driver download complete."
                            Write-Host "Installing Drivers. Please wait..."
                            pnputil /add-driver C:\Techsupp\Drivers\*.inf /subdirs /install
                            Write-Host -ForegroundColor Green "Driver install complete."
                            Write-Host -ForegroundColor Yellow "Press any key to reboot the computer or Ctrl+C to cancel"
                            pause
                            Restart-Computer -Force
                    }

                        
                        # Path to the downloaded .msi file
                        $msiPath = Join-Path -Path $DownloadPath -ChildPath SurfaceDrivers.msi
                    }
                }
            }
        }

        # Check if $DriverPack is an array
        if ($DriverPack -is [System.Collections.IEnumerable]) {
            foreach ($pack in $DriverPack) {
                DownloadDriverPack -pack $pack
            }
        } else {
            Write-Host "DriverPack is not an array"
            DownloadDriverPack -pack $DriverPack
        }
    }
} elseif ($Get_Manufacturer_Info -like "*Dell*") {
	# Get the BIOS serial number
	$sn = (Get-WmiObject -Query "Select SerialNumber from Win32_BIOS").SerialNumber
	Write-Output "Pulling up Dell webpage driver results based on Serial Number: $sn"
	# Open the Dell support page for the serial number
	Start-Process "https://www.dell.com/support/home/en-us/product-support/servicetag/$sn/drivers"
	exit	
} else {
    # If not Microsoft device, pull generic drivers from Windows Update
    Write-Host "Non-Surface device detected."
    Write-Host "Matching and downloading drivers from Windows Update. Please wait..."
    Write-Host "Saving drivers to C:\Techsupp\Drivers..."
    Save-MsUpCatDriver -DestinationDirectory C:\TechSupp\Drivers
    Write-Host -ForegroundColor Green "Windows Update driver download complete."
    Write-Host "Installing Drivers. Please wait..."
    pnputil /add-driver C:\Techsupp\Drivers\*.inf /subdirs /install
    Write-Host -ForegroundColor Green "Driver install complete."
    Write-Host -ForegroundColor Yellow "Press any key to reboot the computer or Ctrl+C to cancel"
    pause
    Restart-Computer -Force
}
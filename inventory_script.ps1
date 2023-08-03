$NETWORK_SHARE_PATH = "\\network_share\inventory\csv\"

function Decode {
    If ($args[0] -is [System.Array]) {
        [System.Text.Encoding]::ASCII.GetString($args[0])
    }
    Else {
        "Not Found"
    }
}


# MS OFFICE
$office_version = 0
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer)
$reg.OpenSubKey('software\Microsoft\Office').GetSubKeyNames() | ForEach-Object {
    if ($_ -match '(\d+)\.') {
        if ([int]$Matches[1] -gt $office_version) {
            $office_version = [int]$Matches[1]
        }
    }    
}

if ($office_version) {
    $office_keys = Get-CimInstance SoftwareLicensingProduct | Where-Object {
        ($_.PartialProductKey) -and ($_.Name -like "*office*")
    }

    switch ($office_version) {
        11 {$office = "2003 ($office_version.0)" ; break}
        12 {$office = "2007 ($office_version.0)" ; break}
        14 {$office = "2010 ($office_version.0)" ; break}    
        15 {$office = "2013 ($office_version.0)" ; break}
        16 {$office = "2016 ($office_version.0)" ; break}
        Default {$office = "MS Office ($office_version.0)" ; break}
    }
}

# GPU
$gpus = foreach($gpu in Get-WmiObject Win32_VideoController)
{
  @{ Name = $gpu.Description } 
}


# Monitor
class MyMonitor {
    [string]$manufacturer
    [string]$name
    [string]$serial
    [int]$Year
    [int]$Week;
}


$monitors = ForEach ($Monitor in Get-WmiObject WmiMonitorID -Namespace root\wmi){
    [MyMonitor]@{
        manufacturer = Decode $Monitor.ManufacturerName -notmatch 0
        name = Decode $Monitor.UserFriendlyName -notmatch 0
        serial = Decode $Monitor.SerialNumberID -notmatch 0
        Year = $Monitor.YearOfManufacture   
        Week = $Monitor.WeekofManufacture             
        }
}


$monitors = ($monitors | Out-String).Trim() -Replace '\0', '' -Replace '\s{2,}', ' ' -Replace '\s:', ':' -Replace ': ', ':'
$computerSystem = Get-CimInstance CIM_ComputerSystem
$computerIP = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'"
#$computerBIOS = Get-CimInstance CIM_BIOSElement
$computerBIOS = Get-CimInstance Win32_BIOS
$computerOS = Get-CimInstance CIM_OperatingSystem
$computerCPU = Get-CimInstance CIM_Processor
$computerGPU = Get-CimInstance Win32_VideoController
$computerHDD = Get-CimInstance Win32_DiskDrive
$computerSystemDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"



$cimInfo = [PSCustomObject]@{
    PC_Name = $computerSystem.Name
    IP = $computerIP.IPAddress[0]
    Manufacturer = $computerSystem.Manufacturer
    Model = $computerSystem.Model
    BIOS = $computerBIOS.Manufacturer + " " + $computerBIOS.SMBIOSBIOSVersion 
    OS = $computerOS.Caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion
    RAM = "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB)
    CPU = $computerCPU.Name
    GPUs = $gpus.Values -join ', '
    Disks = $computerHDD.Caption -join ', '
    System_Drive_SD = $computerSystemDrive.Name
    SD_Capacity = "{0:N2}" -f ($computerSystemDrive.Size/1GB)
    SD_Free_Space_GB = "{0:N2}" -f ($computerSystemDrive.FreeSpace/1GB)    
    SD_Free_Space_Percentage = "{0:P2}" -f ($computerSystemDrive.FreeSpace/$computerSystemDrive.Size)
    User =  $computerSystem.UserName
    Last_Reboot = $computerOS.LastBootUpTime    
    Office = $office
    Office_Keys = if ($office_keys) {$office_keys.PartialProductKey} else {$null}
    Office_Key_IDs = if ($office_keys) {$office_keys.ProductKeyID} else {$null}
    Monitors = $monitors       
}

# OFFICE checking if more than 1 element
if ($cimInfo.Office_Keys -is [array] -and $cimInfo.Office_Key_IDs -is [array]) {
    $maxIndex = [Math]::Max($cimInfo.Office_Keys.Count, $cimInfo.Office_Key_IDs.Count)
    $result = for ($i = 0; $i -lt $maxIndex; $i++) {
        $key = if ($i -lt $cimInfo.Office_Keys.Count) { $cimInfo.Office_Keys[$i] } else { '' }
        $id = if ($i -lt $cimInfo.Office_Key_IDs.Count) { $cimInfo.Office_Key_IDs[$i] } else { '' }
        [PSCustomObject]@{
            Office_Key = $key
            Office_Key_ID = $id
        }
    }
    $cimInfo.Office_Keys = $result.Office_Key -join ', '
    $cimInfo.Office_Key_IDs = $result.Office_Key_ID -join ', '
}

$csvPath = $NETWORK_SHARE_PATH + $cimInfo.PC_Name + ".csv"
$cimInfo | Export-Csv -NoTypeInformation -Path $csvPath

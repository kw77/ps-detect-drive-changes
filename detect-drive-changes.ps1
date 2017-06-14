Register-WmiEvent -Class win32_VolumeChangeEvent -SourceIdentifier volumeChange
$newEvent = Wait-Event -SourceIdentifier volumeChange

do{
    $newEvent = Wait-Event -SourceIdentifier volumeChange
    $eventType = $newEvent.SourceEventArgs.NewEvent.EventType
    $eventTypeName = switch($eventType)
    {
        1 {"Configuration changed"}
        2 {"Device Added"}
        3 {"Device Removed"}
        4 {"Docking"}
    }
    write-host (get-date -format s) " Event detected = " $eventTypeName

    if ($eventType -eq 2) # New Device Added...
    {
        # Get the drive name (letter) and query the logical disks for more information
        $driveName     = $newEvent.SourceEventArgs.NewEvent.DriveName
        $volume        = [wmi]"Win32_LogicalDisk='$driveName'"
        $driveType     = $volume.DriveType
        $driveTypeName = switch($driveType)
        {
            0 {"Unknown"}
            1 {"No Root Directory"}
            2 {"Removable Disk"}
            3 {"Local Disk"}
            4 {"Network Drive"}
            5 {"Compact Disc"}
            6 {"RAM Disk"}
        }

        ############
        # USB INFO #
        ############
        if($driveType -eq 2)
        {
            # Query extended volume info
            Get-WmiObject win32_Volume | where { $_.DriveLetter -eq $driveName } | foreach {
                $serialNumber = $_.SerialNumber
                $label        = $_.Label
                $fileSystem   = $_.FileSystem
            }

            # Query disk to partition mapping to determine the disk/partition numbers for this drive letter
            Get-WMIObject -class Win32_LogicalDiskToPartition | where { $_.Dependent -match ".*DeviceID=""$driveName""" } | foreach {
                $_.Antecedent -match '.*Disk\s#(\d+)\,\sPartition\s#(\d+).*' | Out-Null
                $deviceNumber = $matches[1]
                $partitionNumber = $matches[2]
            }

            # Query the disk drive to get the physical device details
            Get-WMIObject -class Win32_DiskDrive | where {$_.DeviceID -eq “\\.\PHYSICALDRIVE$deviceNumber”} | foreach {
                $deviceID      = $_.DeviceID
                $interfaceType = $_.interfaceType
                $model         = $_.Model
                $caption       = $_.Caption
                $partitions    = $_.Partitions
                $size          = ([math]::Round($disk.Size/1gb,2))
            }
        }

        ############
        # CD INFO #
        ############
        if($driveType -eq 5)
        {
            # Query extended volume info
            Get-WmiObject win32_Volume | where { $_.DriveLetter -eq $driveName } | foreach {
                $serialNumber = $_.SerialNumber
                $label        = $_.Label
                $fileSystem   = $_.FileSystem
                
                # REVIEW WMI Win32_CDROMDrive Class to see if more info can be taken from there
                $deviceNumber    = 'N/A'
                $partitionNumber = 'N/A'
                $deviceID        = 'N/A'
                $interfaceType   = 'N/A'
                $model           = 'N/A'
                $caption         = 'N/A'
                $partitions      = 'N/A'
                $size            = 'N/A'
            }
        }

        # Create a custom object to hold all of the drive informaiton (including parent device information)
        # Note that one device can have multiple partitions, some of which may be assigned a drive letter
        # Note that some may be hidden and these are *not* detected by this script (Win32_LogicalDiskToPartition shows these)
        $volumeInfo = New-Object -Type PSObject
        $volumeInfo | Add-Member -Name 'event'         -Type NoteProperty -Value 'Device Added'
        $volumeInfo | Add-Member -Name 'timestamp'     -Type NoteProperty -Value (get-date -format s)
        $volumeInfo | Add-Member -Name 'driveName'     -Type NoteProperty -Value $driveName
        $volumeInfo | Add-Member -Name 'driveTypeName' -Type NoteProperty -Value $driveTypeName
        $volumeInfo | Add-Member -Name 'volumeName'    -Type NoteProperty -Value $volume.VolumeName
        $volumeInfo | Add-Member -Name 'size'          -Type NoteProperty -Value ([math]::Round($volume.Size/1gb,2))
        $volumeInfo | Add-Member -Name 'freeSpace'     -Type NoteProperty -Value ([math]::Round($volume.FreeSpace/1gb,2))

        $volumeInfo | Add-Member -Name 'partitionNumber'      -Type NoteProperty -Value $partitionNumber

        $volumeInfo | Add-Member -Name 'deviceID'             -Type NoteProperty -Value $deviceID
        $volumeInfo | Add-Member -Name 'deviceInterfaceType'  -Type NoteProperty -Value $interfaceType
        $volumeInfo | Add-Member -Name 'deviceModel'          -Type NoteProperty -Value $model
        $volumeInfo | Add-Member -Name 'deviceCaption'        -Type NoteProperty -Value $caption
        $volumeInfo | Add-Member -Name 'devicePartitionCount' -Type NoteProperty -Value $partitions
        $volumeInfo | Add-Member -Name 'deviceSize'           -Type NoteProperty -Value ([math]::Round($disk.Size/1gb,2))

        # Write to output...
        $volumeInfo | Format-List
        $volumeInfo | ConvertTo-Json
    }

    if ($eventType -eq 3) # Device Removed...
    {
        $volumeInfo = New-Object -Type PSObject
        $volumeInfo | Add-Member -Name 'event'     -Type NoteProperty -Value 'Device Removed'
        $volumeInfo | Add-Member -Name 'timestamp' -Type NoteProperty -Value (get-date -format s)
        $volumeInfo | Add-Member -Name 'driveName' -Type NoteProperty -Value $driveName
        # The drive is no longer there, only the event information is available

        # Write to output...
        $volumeInfo | Format-List
        $volumeInfo | ConvertTo-Json
    }

    Remove-Event -SourceIdentifier volumeChange

} while (1-eq1) #Loop until next event




#Unregister-Event -SourceIdentifier volumeChange

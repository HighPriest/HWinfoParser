

$output = Get-ChildItem -Path .\ -Filter *.xml -Recurse -File -Name| ForEach-Object { #Begin searching for XML files in directory of the file
    $file = Resolve-Path $_ #Create a path of found file for the XDoc
    [System.Xml.XmlDocument] $xdoc = new-object System.Xml.XmlDocument #Create a XML doc object
    $xdoc.Load($file) #Parse the file into XML object structure

    ##MOST USEFUL NODES FOR INVENTORY CREATION (There are more available, that are not human readable)

    $os_node =       $xdoc.SelectNodes("/HWINFO/COMPUTER") # XPath is case sensitive
    $cpu_node =      $xdoc.SelectNodes("/HWINFO/COMPUTER/SubNodes/CPU")
    $mobo_node =     $xdoc.SelectNodes("/HWINFO/COMPUTER/SubNodes/MOBO")
    $memory_node =   $xdoc.SelectNodes("/HWINFO/COMPUTER/SubNodes/MEMORY")
    $chipset_node =  $xdoc.SelectNodes("/HWINFO/COMPUTER/SubNodes/BUS")
    $video_node =    $xdoc.SelectNodes("/HWINFO/COMPUTER/SubNodes/VIDEO")
    $monitor_node =  $xdoc.SelectNodes("/HWINFO/COMPUTER/SubNodes/MONITOR")
    $storage_node =  $xdoc.SelectNodes("/HWINFO/COMPUTER/SubNodes/DRIVES")
    $sound_node =    $xdoc.SelectNodes("/HWINFO/COMPUTER/SubNodes/SOUND")
    $network_node =  $xdoc.SelectNodes("/HWINFO/COMPUTER/SubNodes/NETWORK")


    New-Object -TypeName PSObject -Property ([ordered]@{ ##Function to create a single XML object (single XML parse) that can be printed out
        'Owner' = Split-Path (Split-Path $file -Parent) -Leaf
        'Division' = Split-Path (Split-Path (Split-Path $file -Parent) -Parent) -Leaf
        'Computer Name' = $os_node.SelectNodes("Property[Entry = 'Computer Name']").Description
        'Computer Brand Name' = $os_node.SelectNodes("Property[Entry = 'Computer Brand Name']").Description
        'Operating System' = $os_node.SelectNodes("Property[Entry = 'Operating System']").Description
        'UEFI Boot' = $os_node.SelectNodes("Property[Entry = 'UEFI Boot']").Description
        'Current User Name' = $os_node.SelectNodes("Property[Entry = 'Current User Name']").Description

        'Number Of Processor Cores' = $cpu_node.SelectNodes("Property[Entry = 'Number Of Processor Cores']").Description
        'Number Of Logical Processors' = $cpu_node.SelectNodes("Property[Entry = 'Number Of Logical Processors']").Description
        'Processor Name' = $cpu_node.SelectNodes("SubNode/Property[Entry = 'Processor Name']").Description -join ','
        'Internal Graphics' = $cpu_node.SelectNodes("SubNode/Property[Entry = 'Internal Graphics']").Description

        'Total Memory Size [MB]' = $memory_node.SelectNodes("Property[Entry = 'Total Memory Size [MB]']").Description
        'Memory Channels Supported' = $memory_node.SelectNodes("Property[Entry = 'Memory Channels Supported']").Description
        'Memory Channels Active' = $memory_node.SelectNodes("Property[Entry = 'Memory Channels Active']").Description
        'Total Width' = $mobo_node.SelectNodes("SubNode[NodeName = 'SMBIOS DMI']/SubNode[NodeName = 'Memory Devices']/SubNode[NodeName = 'Memory Device']/Property[Entry = 'Total Width']").Description -join ','
        'Device Type' = $mobo_node.SelectNodes("SubNode[NodeName = 'SMBIOS DMI']/SubNode[NodeName = 'Memory Devices']/SubNode[NodeName = 'Memory Device']/Property[Entry = 'Device Type']").Description -join ','

        'Motherboard Model' = $mobo_node.SelectNodes("Property[Entry = 'Motherboard Model']").Description
        'Motherboard Chipset' = $mobo_node.SelectNodes("Property[Entry = 'Motherboard Chipset']").Description
        'USB Version Supported' = $mobo_node.SelectNodes("Property[Entry = 'USB Version Supported']").Description
        'RAID Capability' = $mobo_node.SelectNodes("Property[Entry = 'RAID Capability']").Description
        'UEFI Bios' = $mobo_node.SelectNodes("Property[Entry = 'UEFI BIOS']").Description #????? Returns odd information
        'Processor Socket' = $mobo_node.SelectNodes("SubNode[NodeName = 'SMBIOS DMI']/SubNode[NodeName = 'Processor']/Property[Entry = 'Processor Upgrade']").Description
        'ECC' = If ($mobo_node.SelectNodes("Property[Entry = 'ECC']").Description) {"Supported"} Else {"Not Supported"} #If not found then not supported
        'DDR3' = If ($mobo_node.SelectNodes("Property[Entry = 'DDR3']").Description) {"Not Supported"} Else {"Supported"} #If not found then supported
        
        'Video Chipset' = $video_node.SelectNodes("SubNode/Property[Entry = 'Video Chipset']").Description -join ','
        'Video Chipset Codename' = $video_node.SelectNodes("SubNode/Property[Entry = 'Video Chipset Codename']").Description -join ','

        'Drive Model' = $storage_node.SelectNodes("SubNode[NodeName = 'IDE Drives']/SubNode/Property[Entry = 'Drive Model']").Description -join ','
        'Drive Capacity [MB]' = $storage_node.SelectNodes("SubNode[NodeName = 'IDE Drives']/SubNode/Property[Entry = 'Drive Capacity [MB]']").Description -join ','
        'Media Rotation Rate' = $storage_node.SelectNodes("SubNode[NodeName = 'IDE Drives']/SubNode/Property[Entry = 'Media Rotation Rate']").Description -join ','
        'ATA Transport Version Supported' = $storage_node.SelectNodes("SubNode[NodeName = 'IDE Drives']/SubNode/Property[Entry = 'ATA Transport Version Supported']").Description -join ','
        
        'Network Card' = ($network_node.SelectNodes("SubNode/Property[Entry = 'Network Card']")).Description -join ','
        'MAC Address' = ($network_node.SelectNodes("SubNode/Property[Entry = 'MAC Address']")).Description -join ','
    })

    
}

$output | Export-Csv -Path .\HWiNFO-Export.csv -Encoding UTF8 -NoTypeInformation -Delimiter ';' ##Export to CSV all found objects

﻿function Folder_Size {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RootPath
    )



            #$targetfolde= '$RootPath'
            $dataColl = @()
            gci -force $RootPath -ErrorAction SilentlyContinue | ? { $_ -is [io.directoryinfo] } | % {
            $len = 0
            gci -recurse -force $_.fullname -ErrorAction SilentlyContinue | % { $len += $_.length }
            $foldername = $_.fullname
            $foldersize= '{0:N2}' -f ($len / 1Gb)
            $dataObject = New-Object PSObject
            Add-Member -inputObject $dataObject -memberType NoteProperty -name “Folder_Name” -value $foldername
            Add-Member -inputObject $dataObject -memberType NoteProperty -name “Folder_Size_GB” -value $foldersize
            $dataColl += $dataObject
            }
            $dataColl | Out-GridView -Title “Size of subdirectories”
            $dataColl | Export-Csv C:\Temp\foldersize.csv -NoTypeInformation




}

Folder_Size
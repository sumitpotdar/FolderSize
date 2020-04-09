Install-Module -Name PSFolderSize -RequiredVersion 1.0 

function Folder_Size {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RootPath
    )
    Get-FolderSize -BasePath $RootPath | Export-Csv C:\Temp\foldersize.csv -NoTypeInformation
    Write-Host "I am done Yippee!"
}

Folder_Size
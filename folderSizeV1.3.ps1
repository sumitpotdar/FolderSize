Param (
	[string]$Path = "c:\Windows",
	[string]$ReportPath = "c:\utils"
)

Function AddObject {
	Param ( 
		$FileObject
	)
	$Size = [double]($FSO.GetFolder($FileObject.FullName).Size)
	$Script:TotSize += $Size
	If ($Size)
	{	$NiceSize = CalculateSize $Size
	}
	Else
	{	$NiceSize = "0.00 MB"
        $Size = 0
	}
	$Script:Report += New-Object PSObject -Property @{
		'Folder Name' = $FileObject.FullName
		'Created on' = $FileObject.CreationTime
		'Last Updated' = $FileObject.LastWriteTime
		Size = $NiceSize
        RawSize = $Size
		Owner = (Get-Acl $FileObject.FullName).Owner
	}
}

Function CalculateSize {
	Param (
		[double]$Size
	)
	If ($Size -gt 1000000000)
	{	$ReturnSize = "{0:N2} GB" -f ($Size / 1GB)
	}
	Else
	{	$ReturnSize = "{0:N2} MB" -f ($Size / 1MB)
	}
	Return $ReturnSize
}

cls
$report | Export-Csv C:\Temp\FolderSizeV1.csv
$Report = @()
$TotSize = 0
$FSO = New-Object -ComObject Scripting.FileSystemObject


$Root = Get-Item -Path $Path 
AddObject $Root


ForEach ($Folder in (Get-ChildItem -Path $Path | Where { $_.PSisContainer }))
{	AddObject $Folder
}


$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
</style>
<Title>
Folder Sizes for "$Path"
</Title>
"@

$TotSize = CalculateSize $TotSize

$Pre = "<h1>Folder Sizes for ""$Path""</h1><h2>Run on $(Get-Date -f 'MM/dd/yyyy hh:mm:ss tt')</h2>"
$Post = "<h2>Total Space Used In ""$($Path)"":  $TotSize</h2>"

#Create the report and save it to a file
$Report | Sort RawSize -Descending | Select 'Folder Name',Owner,'Created On','Last Updated',Size | ConvertTo-Html -PreContent $Pre -PostContent $Post -Head $Header | Out-File $ReportPath\FolderSizes.html

#Display the report in your default browser
& $ReportPath\FolderSizes.html
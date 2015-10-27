#
# Script.ps1
#
function Get-Sum ($a) {
    return ($a | Measure-Object -Sum).Sum
}

Param(
  [string]$version="15",
  [string]$outpath="folderstatistics.csv",
  [string]$folder="sentitems"
)

$mailboxServers = Get-ExchangeServer |where {$_.ServerRole -like "*mailbox*" -and $_.AdminDisplayVersion -like "*$version*" }
$mailboxDatabases = $mailboxServers = Get-MailboxDatabase 
$foldername = "/"+$folder -replace "items"," items"

$mailboxes = @($mailboxDatabases | Get-Mailbox -ResultSize Unlimited)
$report = @()

foreach ($mailbox in $mailboxes)
{
    $folderstats = Get-MailboxFolderStatistics $mailbox -FolderScope $folder |where {$_.FolderPath -eq $foldername}
	
    $folderObj = New-Object PSObject
    $folderObj | Add-Member -MemberType NoteProperty -Name "Name" -Value $mailbox.DisplayName
    $folderObj | Add-Member -MemberType NoteProperty -Name "Size (Mb)" -Value $folderstats.FolderandSubFolderSize.ToMB()
    $folderObj | Add-Member -MemberType NoteProperty -Name "Items" -Value $folderstats.ItemsinFolderandSubfolders
    $report += $folderObj
}

$report | Export-Csv $outpath
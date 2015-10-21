#
# Script.ps1
#

Param(
  [string]$version="15",
  [string]$outpath="folderstatistics.csv",
  [string]$folder="sentitems"
)

$mailboxServers = Get-ExchangeServer |where {$_.ServerRole -like "*mailbox*" -and $_.AdminDisplayVersion -like "*$version*" }
$mailboxDatabases = $mailboxServers = Get-MailboxDatabase 

$mailboxes = @($mailboxDatabases | Get-Mailbox -ResultSize Unlimited)
$report = @()

foreach ($mailbox in $mailboxes)
{
    $folderstats = Get-MailboxFolderStatistics $mailbox -FolderScope $folder

    $folderObj = New-Object PSObject
    $folderObj | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $mailbox.DisplayName
    $folderObj | Add-Member -MemberType NoteProperty -Name "Inbox Size (Mb)" -Value $folderstats.FolderandSubFolderSize.ToMB()
    $folderObj | Add-Member -MemberType NoteProperty -Name "Inbox Items" -Value $folderstats.ItemsinFolderandSubfolders
    $report += $folderObj
}

$report | Export-Csv $outpath
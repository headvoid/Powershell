<#

Locates all transport servers in the environment (typically the MBX role holders)
Reads all the message tracking logs between the start and end date
stores teh important SMTP and STOREDRIVER information in to an array
Repeats for the number of days required

V0.1 - Inital creation - Justin Whelan
V0.2 - Added counters to total size of send and receive
     - Removed noisy HealthMailbox entries
V0.3 - Cleaned up date and time to allow Number of Days to work. Now requires a negative value
	 - Added "testing" option, which changes the days to hours
	 - Added quick stat displays at the end of the script run
#>

$NumberOfDays = -10
$RefDate = Get-Date

# set to 1 if testing the script, turns days in to hours
$testing = 0

$SMTP = @()
$STORE = @()

$Servers = Get-TransportServer

function CountMessageSize($s)
{
    $totalbytes=0

    foreach($i in $s)
    {
        $totalbytes = $totalbytes + $i.TotalBytes
    }
    return $totalbytes
}


foreach($server in $servers)
{
    Write-Host "Server: $server"
    Write-Host "Start: "$RefDate.AddDays($NumberOfDays)" to "$RefDate
  
	    If($testing -ne 1)
    {
        $a = Get-MessageTrackingLog -Start $RefDate.AddDays($NumberOfDays) -End $RefDate -Server $Server.Name -ResultSize unlimited
    }
    else
    {
        $a = Get-MessageTrackingLog -Start $RefDate.AddHours($NumberOfDays) -End $RefDate -Server $Server.Name -ResultSize unlimited
    }
            

    <#

    EventId (SMTP)
    -------
    SEND
    HAREDIRECT
    RECEIVE
    HADISCARD
    HARECEIVE
    FAIL

    EventId (STORE)
    -------
    NOTIFYMAPI
    RECEIVE
    SUBMIT
    SUBMITFAIL
    DELIVER
    DUPLICATEDELIVER
    
    #>

    # This section sorts out all of the SMTP related traffic

    $SEND = $a |where {$_.EventId -eq "SEND" -and $_.Source -eq "SMTP" -and $_.Sender -notlike "HealthMailbox*" -and $_.MessageSubjet -notlike "* probe"}
    $HAREDIRECT = $a |where {$_.EventId -eq "HAREDIRECT" -and $_.Source -eq "SMTP" -and $_.Sender -notlike "HealthMailbox*" -and $_.MessageSubjet -notlike "* probe"}
    $RECEIVE = $a |where {$_.EventId -eq "RECEIVE" -and $_.Source -eq "SMTP" -and $_.Sender -notlike "HealthMailbox*" -and $_.MessageSubjet -notlike "* probe"}
    $HADISCARD = $a |where {$_.EventId -eq "HADISCARD" -and $_.Source -eq "SMTP" -and $_.Sender -notlike "HealthMailbox*" -and $_.MessageSubjet -notlike "* probe"}
    $HARECEIVE = $a |where {$_.EventId -eq "HARECEIVE" -and $_.Source -eq "SMTP" -and $_.Sender -notlike "HealthMailbox*" -and $_.MessageSubjet -notlike "* probe"}
    $FAIL = $a |where {$_.EventId -eq "FAIL" -and $_.Source -eq "SMTP" -and $_.Sender -notlike "HealthMailbox*" -and $_.MessageSubjet -notlike "* probe"}


    $SMTPO = @{
        STARTDATE=$StartDate.AddDays($CurrentDay)
        ENDDATE=$EndDate.AddDays($CurrentDay)
        SEND=$SEND.Count
        HAREDIRECT=$HAREDIRECT.Count
        RECEIVE=$RECEIVE.Count
        HADISCARD=$HADISCARD.Count
        HARECEIVE=$HARECEIVE.Count
        FAIL=$FAIL.Count
        SERVER=$Server.Name
        SENDBYTES=CountMessageSize($SEND)
        RECEIVEBYTES=CountMessageSize($RECEIVE)
    }
        
    # This section sorts out all of the STOREDRIVER traffic
    
    $NOTIFYMAPI = $a |where {$_.EventId -eq "NOTIFYMAPI" -and $_.Source -eq "STOREDRIVER"}
    $RECEIVE = $a |where {$_.EventId -eq "RECEIVE" -and $_.Source -eq "STOREDRIVER" -and $_.Sender -notlike "HealthMailbox*"}
    $SUBMIT = $a |where {$_.EventId -eq "SUBMIT" -and $_.Source -eq "STOREDRIVER"}
    $SUBMITFAIL = $a |where {$_.EventId -eq "SUBMITFAIL" -and $_.Source -eq "STOREDRIVER"}
    $DELIVER = $a |where {$_.EventId -eq "DELIVER" -and $_.Source -eq "STOREDRIVER" -and $_.Sender -notlike "HealthMailbox*"}
    $DUPLICATEDELIVER = $a |where {$_.EventId -eq "DUPLICATEDELIVER" -and $_.Source -eq "STOREDRIVER"}

    $STOREO = @{
        STARTDATE=$StartDate.AddDays($CurrentDay)
        ENDDATE=$EndDate.AddDays($CurrentDay)
        NOTIFYMAPI=$NOTIFYMAPI.Count
        RECEIVE=$RECEIVE.Count
        SUBMIT=$SUBMIT.Count
        SUBMITFAIL=$SUBMITFAIL.Count
        DELIVER=$DELIVER.Count
        DUPLICATEDELIVER=$DUPLICATEDELIVER.Count
        SERVER=$Server.Name
        RECEIVEBYTES=CountMessageSize($RECEIVE)
    }

    $SMTP += New-Object -TypeName PSObject -Property $SMTPO
    $STORE += New-Object -TypeName PSObject -Property $STOREO
    $CurrentDay--
}

# output some quick stats
Write-Host "Preparing quick stats..."

$mailboxes = (Get-Mailbox).Count

# prepare variables
$sendreceive = 0
$sendreceivebytes = 0
$NumberOfDays = $NumberOfDays * -1

Foreach($mailsend in $SMTP)
{
    # add up all of the send and receive columns for SMTP

    $sendreceive = $sendreceive + $mailsend.SEND
    $sendreceive = $sendreceive + $mailsend.RECEIVE

    # add up all of the sent and received bytes

    $sendreceivebytes = $sendreceivebytes + $mailsend.SENDBYTES
    $sendreceivebytes = $sendreceivebytes + $mailsend.RECEIVEBYTES
}

Foreach($mailsend in $STORE)
{
    # add up all of the send and receive columns for SMTP
    $sendreceive = $sendreceive + $mailsend.DELIVER
    $sendreceive = $sendreceive + $mailsend.RECEIVE
    $sendreceive = $sendreceive + $mailsend.SUBMIT

    # add up all of the sent and received bytes

    $sendreceivebytes = $sendreceivebytes + $mailsend.RECEIVEBYTES
}

Write-Host "Total Mails Sent and Received = "$sendreceive

Write-Host "Total Mail Bytes Sent and Received = "$sendreceivebytes
Write-Host "Total Mail GB Sent and Received = "($sendreceivebytes/1024/1024/1024)
Write-Host "Number of Mailboxes = "$mailboxes
Write-Host "Number of Days Average taken over = "$NumberOfDays
$AvgPerUser = ($sendreceive/$NumberOfDays/$mailboxes)
Write-Host "Number of Emails per user per day = "$AvgPerUser
Write-Host "Size of Average Email in kb = "($sendreceivebytes/$NumberOfDays/$mailboxes/$AvgPerUser/1024)

$SMTP |Export-Csv smtp.csv
$STORE | Export-Csv store.csv
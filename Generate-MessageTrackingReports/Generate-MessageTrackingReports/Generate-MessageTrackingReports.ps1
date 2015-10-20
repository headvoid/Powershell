<#

Locates all transport servers in the environment (typically the MBX role holders)
Reads all the message tracking logs between the start and end date
stores teh important SMTP and STOREDRIVER information in to an array
Repeats for the number of days required

V0.1 - Inital creation - Justin Whelan
V0.2 - Added counters to total size of send and receive
     - Removed noisy HealthMailbox entries
#>

$NumberOfDays = 1
$CurrentDay = 0
$StartDate = Get-Date "8/11/2015 00:01:00"
$EndDate = Get-Date "8/11/2015 23:59:00"

$SMTP = @()
$STORE = @()

# TODO - This needs to be replaced as microsoft are removing the command
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
    while($CurrentDay -ne $NumberOfDays)
    {
        $a = Get-MessageTrackingLog -Start $StartDate.AddDays($CurrentDay) -End $EndDate.AddDays($CurrentDay) -Server $Server.Name -ResultSize unlimited

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
        $CurrentDay++
    }
    $CurrentDay = 0
}
$SMTP |Export-Csv smtp.csv
$STORE | Export-Csv store.csv
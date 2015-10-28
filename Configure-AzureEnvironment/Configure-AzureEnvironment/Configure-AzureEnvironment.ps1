<#

AzureRM : https://azure.microsoft.com/en-us/blog/azps-1-0-pre/

#>

PARAM
(
	[string]$LocalNetworksFile="D:\Google Drive\Acland\Azure Migration\Design\Script\LocalNetworks.csv",
	[string]$AzureSubnetsFile="D:\Google Drive\Acland\Azure Migration\Design\Script\AzureSubnets.csv",
	[string]$AzureNetworkFile="D:\Google Drive\Acland\Azure Migration\Design\Script\AzureNetworks.csv",
	[string]$AzureNetworkModulePath="D:\Google Drive\Scripts\AzureNetworking-1.0.2\AzureNetworking.psd1"
)

function Load-AzureModule()
{
    $modulePath = "C:\Program Files (x86)\Microsoft SDKs\Azure\PowerShell\ServiceManagement\Azure"

    $curdir = Get-Location
    Set-Location $modulePath

    Import-Module ".\Services\Azure.psd1"
    Set-Location $curdir
}

function Select-TextItem 
{ 
PARAM  
( 
    [Parameter(Mandatory=$true)] 
    $options, 
    $displayProperty 
) 
 
    [int]$optionPrefix = 1 
    # Create menu list 
    foreach ($option in $options) 
    { 
        if ($displayProperty -eq $null) 
        { 
            Write-Host ("{0,3}: {1}" -f $optionPrefix,$option) 
        } 
        else 
        { 
            Write-Host ("{0,3}: {1}" -f $optionPrefix,$option.$displayProperty) 
        } 
        $optionPrefix++ 
    } 
    Write-Host ("{0,3}: {1}" -f 0,"To cancel")  
    [int]$response = Read-Host "Enter Selection" 
    $val = $null 
    if ($response -gt 0 -and $response -le $options.Count) 
    { 
        $val = $options[$response-1] 
    } 
    return $val 
}    

#Load-AzureModule
#Import-Module $AzureNetworkModulePath
#Import-Module Azure
#Import-Module AzureRM

#Login-AzureRmAccount

#Add-AzureAccount

#$subscriptions = Get-AzureSubscription
#$SelectedSub = Select-TextItem $subscriptions "SubscriptionName"
#Select-AzureSubscription -SubscriptionId $SelectedSub.SubscriptionId

$LocalNetworks = Import-CSV $LocalNetworksFile 
$AzureSubnets = Import-CSV $AzureSubnetsFile
$AzureNetworks = Import-CSV $AzureNetworkFile

#Switch-AzureMode -Name AzureResourceManager

#TODO - improve this section - hard coded AUS EAST
$ResourceGroups = $AzureNetworks.ResourceGroup |select -Unique

# Create the Resource Groups
foreach($RG in $ResourceGroups)
{
	# check to see if it already exists

	New-AzureRMResourceGroup -Name $RG -location "Australia East" -Force
}
# Create the VNET
foreach($vnets in $AzureNetworks)
{
	$GatewaySubnet = $AzureSubnets |where {$_.Subnets -eq "GatewaySubnet" -and $_.Name -eq $vnets.Name}
	$GatewaySubnetConfig = New-AzureRmVirtualNetworkSubnetConfig -Name "GatewaySubnet" -AddressPrefix $GatewaySubnet.AddressPrefix
	New-AzureRmVirtualNetwork -Name $vnets.Name -ResourceGroupName $vnets.ResourceGroup -Location $vnets.Location -AddressPrefix $vnets.AddressPrefix -Subnet $GatewaySubnetConfig

	$RemainingSubnets = $AzureSubnets |where {$_.Subnets -ne "GatewaySubnet" -and $_.Name -eq $vnets.Name}

	$vnetlink = Get-AzureRmVirtualNetwork -ResourceGroupName $vnets.ResourceGroup -Name $vnets.Name
	foreach($subnet in $RemainingSubnets)
	{
		$subnet
		Add-AzureRmVirtualNetworkSubnetConfig -Name $subnet.Subnets -AddressPrefix $subnet.AddressPrefix -VirtualNetwork $vnetlink
	}
	Set-AzureRMVirtualNetwork -VirtualNetwork $vnetlink

	# Request Public IP Address
	$locationofRG = (Get-AzureRmResourceGroup -Name $vnets.ResourceGroup).Location
	$gwpip = New-AzureRmPublicIpAddress -Name $vnets.Gateway -ResourceGroupName $vnets.ResourceGroup -Location $vnets.Location -AllocationMethod Dynamic -DomainNameLabel $vnets.ResourceGroup.toLower()
	$InternalNetwork = Get-AzureRMVirtualNetwork -Name $vnets.Name -ResourceGroupName $vnets.ResourceGroup
	$subnet = Get-AzureRMVirtualNetworkSubnetConfig -Name 'GatewaySubnet' -VirtualNetwork $InternalNetwork
	$gwipconfig = New-AzureRMVirtualNetworkGatewayIpConfig -Name gwipconfig1 -SubnetId $subnet.Id -PublicIpAddressId $gwpip.Id 
	New-AzureRMVirtualNetworkGateway -Name vnetgw1 -ResourceGroupName $vnets.ResourceGroup -Location $locationofRG -IpConfigurations $gwipconfig -GatewayType Vpn -VpnType RouteBased
}


foreach($localsite in $LocalNetworks)
{
	$locationofRG = (Get-AzureRmResourceGroup -Name $localsite.Resource).Location
	New-AzureRmLocalNetworkGateway -Name $localsite.Name -ResourceGroupName $localsite.Resource -GatewayIpAddress $localsite.Gateway -AddressPrefix $localsite.Internal -Location $locationofRG
}

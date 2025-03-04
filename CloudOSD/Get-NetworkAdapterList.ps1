function Get-NetworkAdapterList {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$ComputerName = $env:COMPUTERNAME
    )
    process {
        $CimSession = New-CimSession -ComputerName $ComputerName
        $NetworkAdapterList = Get-CimInstance -CimSession $CimSession -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'" | 
        Where-Object {
            ($null -ne $_.IPAddress) -and ($null -ne $_.DefaultIPGateway) -and ($null -ne $_.DNSServerSearchOrder) -and ($_.Description -notlike '*Wireless*')
        }
        foreach ($NetworkAdapter in $NetworkAdapterList) {
            [PSCustomObject]@{
                ComputerName = $NetworkAdapter.DNSHostName
                MACAddress = $NetworkAdapter.MACAddress
                AdapterName = $NetworkAdapter.Description
                IPv4Address = $NetworkAdapter.IPAddress | Where-Object {$_.Contains('.')}
                IPv6Address = $NetworkAdapter.IPAddress | Where-Object {$_.Contains(':')}
                SubnetMask = $NetworkAdapter.IPSubnet
                DefaultGateway = $NetworkAdapter.DefaultIPGateway
                DNSServer = $NetworkAdapter.DNSServerSearchOrder
                ServiceName = $NetworkAdapter.ServiceName
            }   
        }
    }
}
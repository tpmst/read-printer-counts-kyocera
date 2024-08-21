Import-Module SNMP
 
# Define your devices with name and IP address
$devices = @(
    @{ Name = "Printer Name"; IPAddress = "IP Address"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "Printer Name"; IPAddress = "IP Address"; ColorPrinter = $false; OIDBlack = "1.3.6.1.4.1.18334.1.1.1.5.7.2.1.8.0" },
    @{ Name = "Printer Name"; IPAddress = "IP Address"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.7.2.2")}
)

#; OIDBlack = ""
# Function to get SNMP data and output counts, including checking online status
function Get-SNMPDataForDevice {
    param(
        [string]$Name,
        [string]$IPAddress,
        [bool]$ColorPrinter,
        [string[]]$OIDBlack,
        [string[]]$OIDColor
    )
    $pingResult = Test-Connection -ComputerName $IPAddress -Count 1 -Quiet
    $totalBlackWhite = 0
    $totalColor = 0

    if ($pingResult) {
        # For black and white counts
        foreach ($oid in $OIDBlack) {
            $totalBlackWhite += (Get-SNMPData -IP $IPAddress -Community public -OID $oid).Data
        }

        if ($ColorPrinter) {
            # For color counts
            foreach ($oid in $OIDColor) {
                $totalColor += (Get-SNMPData -IP $IPAddress -Community public -OID $oid).Data
            }

            # Output results for color printer
            [PSCustomObject]@{
                Name = $Name
                IP = $IPAddress
                PrintsBlackWhite = $totalBlackWhite
                PrintsColor = $totalColor
                Status = "Online"
            }
        } else {
            # Output results for black and white printer
            [PSCustomObject]@{
                Name = $Name
                IP = $IPAddress
                PrintsBlackWhite = $totalBlackWhite
                PrintsColor = "N/A"
                Status = "Online"
            }
        }
    } else {
        # Output results if offline
        [PSCustomObject]@{
            Name = $Name
            IP = $IPAddress
            PrintsBlackWhite = "-"
            PrintsColor = "-"
            Status = "Offline"
        }
    }
} 

# Array to collect output
$output = @()

# Loop through each device and retrieve SNMP data
foreach ($device in $devices) {
    $newDeviceName = "`"$($device.Name)`""

    # Append the data to the output using the new device name
    $output += Get-SNMPDataForDevice -Name $newDeviceName -IPAddress $device.IPAddress -ColorPrinter $device.ColorPrinter -OIDBlack $device.OIDBlack -OIDColor $device.OIDColor
}

# Display the collected output
$output | Format-Table -AutoSize


$Date = Get-Date -Format "yyyy-MM"
$Path = 'C:\temp\Printercounts_' + $Date + '.csv'
# Export the output to a CSV file with specified headers
$output | Export-Csv -Path $Path -NoTypeInformation

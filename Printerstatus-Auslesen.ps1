Import-Module SNMP


# Define your devices with name and IP address
$devices = @(
    @{ Name = "BIZHUB C3350i (281)"; IPAddress = "192.173.253.92"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB C258 COLOR"; IPAddress = "192.173.253.98"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB C3350i EG"; IPAddress = "192.173.253.97"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB C3351 COLOR"; IPAddress = "192.173.253.88"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB C364e COLOR m. fiery"; IPAddress = "192.173.253.83"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB C360i COLOR"; IPAddress = "192.173.253.85"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB C3350i (254)"; IPAddress = "192.173.253.93"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB C3350i (753)"; IPAddress = "192.173.253.94"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB C3320i (663)"; IPAddress = "192.173.253.95"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB C308e"; IPAddress = "192.173.253.99"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB C250i"; IPAddress = "192.173.253.80"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.12.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.4.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.3.1.7.1")},
    @{ Name = "BIZHUB 223"; IPAddress = "192.173.253.87"; ColorPrinter = $false; OIDBlack = "1.3.6.1.4.1.18334.1.1.1.5.7.2.1.8.0" },
    @{ Name = "BIZHUB 40P - 712"; IPAddress = "192.173.253.86"; ColorPrinter = $false ; OIDBlack = ""},
    @{ Name = "BIZHUB 40P - 492"; IPAddress = "192.173.253.89"; ColorPrinter = $false ; OIDBlack = "1.3.6.1.4.1.18334.1.1.1.5.7.2.1.1.0"},
    @{ Name = "BIZHUB 40P- 422"; IPAddress = "192.173.253.91"; ColorPrinter = $false ; OIDBlack = "1.3.6.1.4.1.18334.1.1.1.5.7.2.1.1.0"},
    @{ Name = "BIZHUB 4050"; IPAddress = "192.173.253.90"; ColorPrinter = $false; OIDBlack = "1.3.6.1.4.1.18334.1.1.1.5.7.2.1.1.0"},
    @{ Name = "BIZHUB C3110 COLOR"; IPAddress = "192.173.253.66"; ColorPrinter = $true; OIDBlack = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2"); OIDColor = @("1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", "1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.7.2.2")}
)



#Invoke-SnmpWalk -IP 192.173.253.66 -Community public -OIDStart 1.3.6.1.4.1.18334.1.1.1.5.7.2

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

# Function to create a fully quoted CSV
function ConvertTo-QuotedCSV {
    param(
        [Parameter(ValueFromPipeline=$true)]
        $InputObject
    )

    begin {
        $output = New-Object System.Text.StringBuilder
        $first = $true
    }
    process {
        if ($first) {
            # Get property names and add quotes around each
            $header = ($InputObject | Get-Member -MemberType NoteProperty).Name |
                      ForEach-Object { "`"$_`"" }
            [void]$output.AppendLine(($header -join ","))
            $first = $false
        }
        # Process each property value, adding quotes and escaping existing quotes
        $data = $InputObject.PSObject.Properties.Value |
                ForEach-Object { 
                    $value = $_ -replace '"', '""'  # Escape quotes in data
                    "`"$value`""  # Enclose data in quotes
                }
        [void]$output.AppendLine(($data -join ","))
    }
    end {
        return $output.ToString()
    }
}



$Date = Get-Date -Format "yyyy-MM"
$Path = 'C:\temp\Drucker-Zählerstände_' + $Date + '.csv'
# Export the output to a CSV file with specified headers
$output | ConvertTo-QuotedCSV | Set-Content -Path $Path



#Get-SNMPData -IP 192.173.253.86 -Community public -OID "1.3.6.1.4.1.18334.1.1.1.5.7.2.1.1.0"
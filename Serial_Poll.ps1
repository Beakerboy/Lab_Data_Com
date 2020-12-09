function Serial-Poll {
    param (
        $Value,
        $Port
    )
    # Both devices are configured with CF/LF line endings
    $Port.Write("${Value}`r`n")
    Start-Sleep 1
    $Response = ""
    While ($Response -eq "") {
        $Response = $Port.ReadExisting()
    }
    Write-Output $Response
}

# Configure Serial Ports
$DMA_Port = New-Object System.IO.Ports.SerialPort COM5,9600,None,8,1
$Spec_Port = New-Object System.IO.Ports.SerialPort COM6,9600,None,8,1
$Spec_Port.DtrEnable = "false"

# Open Ports
$DMA_Port.Open()
$Spec_Port.Open()

$data_path = "D:\"

# Create empty CSV File using the current unix timestamp as the name
$timestamp = Get-Date -UFormat "%s"
$filename = "${timestamp}.csv"
New-Item -Path $data_path -Name $filename -ItemType "file"

# Initialize Variables
$vial_position = "0"
$not_Finished = 1

# Loop while the DMA is running
While ($not_Finished) {
    # Request data from DMA
    $DMA_Response = Serial-Poll -Port $DMA_Port -Value "getdata"

    # If the response is anything other than "no new data available" poll for the color and save the data
    if ($DMA_Response -ne "no new data available`r") {
        # Request abs data from Spectrophotometer
        $Spec_Response = Serial-Poll -Port $Spec_Port -Value "SND"
        $split = $Spec_Response -Split ("`r`n")
        $abs = $split[1]
        # Get current Datetime
        $date = Get-Date
        $vial_position = $DMA_Response.Split(";")[1]
        if ($DMA_Response.Split(";")[6] -eq "valid`r") {
            $DMA_Response = $DMA_Response.Split(";")[2]
        }
        # Concatinate values
        $line =  "${date},${DMA_Response},${abs}"
        # Write line to csv file
        Add-Content -Path "${data_path}${filename}" -Value $line
    }
    $DMA_Response = Serial-Poll -Port $DMA_Port -Value "finished"
    $not_Finished = ($vial_position -ne "48") -and ($DMA_Response -ne "measurement finished`r")
}

# Close Ports
$DMA_Port.Close()
$Spec_Port.Close()

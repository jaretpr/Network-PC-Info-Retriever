Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Network PC Info Retriever"
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)

# Create a label for the title
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Network PC Info Retriever"
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.AutoSize = $true
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($titleLabel)

# Create a text box to display the process log
$textBoxLog = New-Object System.Windows.Forms.TextBox
$textBoxLog.Location = New-Object System.Drawing.Point(50, 200)
$textBoxLog.Size = New-Object System.Drawing.Size(500, 150)
$textBoxLog.Multiline = $true
$textBoxLog.ScrollBars = "Vertical"
$textBoxLog.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$textBoxLog.ForeColor = [System.Drawing.Color]::White
$textBoxLog.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($textBoxLog)

# Create the Retrieve PC Names button
$btnRetrievePCNames = New-Object System.Windows.Forms.Button
$btnRetrievePCNames.Location = New-Object System.Drawing.Point(50, 80)
$btnRetrievePCNames.Size = New-Object System.Drawing.Size(240, 40)
$btnRetrievePCNames.Text = "Retrieve PC Names"
$btnRetrievePCNames.BackColor = [System.Drawing.Color]::FromArgb(28, 151, 234)
$btnRetrievePCNames.ForeColor = [System.Drawing.Color]::White
$btnRetrievePCNames.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRetrievePCNames)

# Create the Retrieve PC Hardware button
$btnRetrievePCHardware = New-Object System.Windows.Forms.Button
$btnRetrievePCHardware.Location = New-Object System.Drawing.Point(310, 80)
$btnRetrievePCHardware.Size = New-Object System.Drawing.Size(240, 40)
$btnRetrievePCHardware.Text = "Retrieve PC Hardware"
$btnRetrievePCHardware.BackColor = [System.Drawing.Color]::FromArgb(28, 151, 234)
$btnRetrievePCHardware.ForeColor = [System.Drawing.Color]::White
$btnRetrievePCHardware.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRetrievePCHardware)

# Create the Open PC Names File button
$btnOpenPCNamesFile = New-Object System.Windows.Forms.Button
$btnOpenPCNamesFile.Location = New-Object System.Drawing.Point(50, 130)
$btnOpenPCNamesFile.Size = New-Object System.Drawing.Size(150, 30)
$btnOpenPCNamesFile.Text = "Open Names File"
$btnOpenPCNamesFile.BackColor = [System.Drawing.Color]::FromArgb(63, 63, 70)
$btnOpenPCNamesFile.ForeColor = [System.Drawing.Color]::White
$btnOpenPCNamesFile.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($btnOpenPCNamesFile)

# Create the Open PC Hardware File button
$btnOpenPCHardwareFile = New-Object System.Windows.Forms.Button
$btnOpenPCHardwareFile.Location = New-Object System.Drawing.Point(400, 130)
$btnOpenPCHardwareFile.Size = New-Object System.Drawing.Size(150, 30)
$btnOpenPCHardwareFile.Text = "Open Hardware File"
$btnOpenPCHardwareFile.BackColor = [System.Drawing.Color]::FromArgb(63, 63, 70)
$btnOpenPCHardwareFile.ForeColor = [System.Drawing.Color]::White
$btnOpenPCHardwareFile.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($btnOpenPCHardwareFile)

# Define the Retrieve PC Names action
$btnRetrievePCNames.Add_Click({
    $textBoxLog.AppendText("Starting to retrieve PC names...`r`n")
	# Change output directory path, if necessary
    $outputDir = "C:\Active Directory Reports\Input"
    $outputFile = Join-Path -Path $outputDir -ChildPath "COMPUTERS.TXT"

    if (-not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
    }

    $PCList = Get-ADComputer -Filter 'OperatingSystem -notlike "*SERVER*" -and Enabled -eq $true' |
              Select-Object -ExpandProperty Name |
              Sort-Object

    $PCList | Out-File -FilePath $outputFile

    $textBoxLog.AppendText("List of PCs saved to $outputFile`r`n")
})

# Define the Retrieve PC Hardware action
$btnRetrievePCHardware.Add_Click({
    $textBoxLog.AppendText("Starting to retrieve PC hardware information...`r`n")
	# Change input path, if necessary
    $inputPath = "C:\Active Directory Reports\Input\COMPUTERS.TXT"
	# Change output path, if necessary
    $outputPath = "C:\Active Directory Reports\Output\ComputerInfoReport.txt"

    function Convert-WmiDateToDateTime {
        param ([string]$wmiDate)
        if (-not $wmiDate) { return $null }
        $year = $wmiDate.Substring(0, 4)
        $month = $wmiDate.Substring(4, 2)
        $day = $wmiDate.Substring(6, 2)
        $hour = $wmiDate.Substring(8, 2)
        $minute = $wmiDate.Substring(10, 2)
        $second = $wmiDate.Substring(12, 2)
        $datetime = "${year}-${month}-${day} ${hour}:${minute}:${second}"
        return [datetime]$datetime
    }

    $outputDir = Split-Path -Path $outputPath
    if (-not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
    }

    if (Test-Path $outputPath) {
        Remove-Item $outputPath
    }

    $serverList = Get-Content $inputPath

    foreach ($server in $serverList) {
        $textBoxLog.AppendText("Collecting information from $server...`r`n")
        $output = "System Name: $server`r`n"
        $scriptBlock = {
            param ($server)
            $result = @{}
            try {
                $result.sysInfo = Get-WmiObject Win32_ComputerSystem -ComputerName $server -ErrorAction Stop
                $result.bios = Get-WmiObject Win32_BIOS -ComputerName $server -ErrorAction Stop
                $result.os = Get-WmiObject Win32_OperatingSystem -ComputerName $server -ErrorAction Stop
                $result.processor = Get-WmiObject Win32_Processor -ComputerName $server -ErrorAction Stop
                $result.baseBoard = Get-WmiObject Win32_BaseBoard -ComputerName $server -ErrorAction Stop
                $result.logicalDisk = Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter "DeviceID='C:'" -ErrorAction Stop
            } catch {
                $result.error = $_.Exception.Message
            }
            $result
        }

        $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $server
        $completed = Wait-Job $job -Timeout 10
        if ($completed) {
            $data = Receive-Job $job
            if ($data.error) {
                $output += $data.error + "`r`n"
            } else {
                $biosDate = Convert-WmiDateToDateTime -wmiDate $data.bios.ReleaseDate
                $output += "User Name: $($data.sysInfo.UserName)`r`n"
                $output += "System Model: $($data.sysInfo.Model)`r`n"
                $output += "Processor: $($data.processor.Name), $($data.processor.MaxClockSpeed) Mhz, $($data.processor.NumberOfCores) Core(s), $($data.processor.NumberOfLogicalProcessors) Logical Processor(s)`r`n"
                $output += "Installed Physical Memory (RAM): $([math]::Round($data.sysInfo.TotalPhysicalMemory / 1GB, 2)) GB`r`n"
                $output += "Storage: $([math]::Round($data.logicalDisk.Size / 1GB, 2)) GB`r`n"
                $output += "BaseBoard Manufacturer: $($data.baseBoard.Manufacturer)`r`n"
                $output += "Bios Version/Date: $($data.bios.SMBIOSBIOSVersion), $biosDate`r`n"
                $output += "OS Name: $($data.os.Caption)`r`n`r`n"
            }
        } else {
            $output += "Timeout reached when querying $server`r`n"
        }
        Remove-Job $job -Force

        Add-Content -Path $outputPath -Value $output
        $textBoxLog.AppendText("Information for $server collected.`r`n")
    }

    $textBoxLog.AppendText("Inventory completed. Check $outputPath`r`n")
})

# Define the Open PC Names File action
$btnOpenPCNamesFile.Add_Click({
	# Change input directory path, if necessary
    Start-Process explorer.exe "C:\Active Directory Reports\Input"
})

# Define the Open PC Hardware File action
$btnOpenPCHardwareFile.Add_Click({
	# Change output directory path, if necessary
    Start-Process explorer.exe "C:\Active Directory Reports\Output"
})

# Show the form and adjust title label position
$form.Add_Shown({
    $titleLabel.Location = New-Object System.Drawing.Point([math]::Round(($form.ClientSize.Width - $titleLabel.Width) / 2), 20)
})

# Show the form
$form.ShowDialog()

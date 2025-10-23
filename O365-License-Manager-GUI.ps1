<#
.SYNOPSIS
    Office 365 License Report Generator - GUI Version
.DESCRIPTION
    GUI tool to check available O365 licenses and export reports
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set error handling
$ErrorActionPreference = "Continue"

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Office 365 License Manager"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Create status bar
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Text = "Ready"
$form.Controls.Add($statusBar)

# Create title label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$titleLabel.Size = New-Object System.Drawing.Size(860, 40)
$titleLabel.Text = "Office 365 License Manager"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$form.Controls.Add($titleLabel)

# Create connection panel
$connectionPanel = New-Object System.Windows.Forms.GroupBox
$connectionPanel.Location = New-Object System.Drawing.Point(20, 70)
$connectionPanel.Size = New-Object System.Drawing.Size(860, 80)
$connectionPanel.Text = "Connection"
$connectionPanel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($connectionPanel)

# Connection status label
$connectionStatusLabel = New-Object System.Windows.Forms.Label
$connectionStatusLabel.Location = New-Object System.Drawing.Point(20, 30)
$connectionStatusLabel.Size = New-Object System.Drawing.Size(600, 25)
$connectionStatusLabel.Text = "Status: Not Connected"
$connectionStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$connectionPanel.Controls.Add($connectionStatusLabel)

# Connect button
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Location = New-Object System.Drawing.Point(650, 25)
$connectButton.Size = New-Object System.Drawing.Size(180, 35)
$connectButton.Text = "Connect to Azure AD"
$connectButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$connectButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$connectButton.ForeColor = [System.Drawing.Color]::White
$connectButton.FlatStyle = 'Flat'
$connectionPanel.Controls.Add($connectButton)

# Create data grid view
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(20, 160)
$dataGridView.Size = New-Object System.Drawing.Size(860, 400)
$dataGridView.AutoSizeColumnsMode = 'Fill'
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = 'FullRowSelect'
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.BorderStyle = 'Fixed3D'
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$dataGridView.EnableHeadersVisualStyles = $false
$form.Controls.Add($dataGridView)

# Create button panel
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(20, 570)
$buttonPanel.Size = New-Object System.Drawing.Size(860, 60)
$form.Controls.Add($buttonPanel)

# Refresh button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(0, 10)
$refreshButton.Size = New-Object System.Drawing.Size(150, 40)
$refreshButton.Text = "Refresh Licenses"
$refreshButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$refreshButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$refreshButton.ForeColor = [System.Drawing.Color]::White
$refreshButton.FlatStyle = 'Flat'
$refreshButton.Enabled = $false
$buttonPanel.Controls.Add($refreshButton)

# Export CSV button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(160, 10)
$exportButton.Size = New-Object System.Drawing.Size(150, 40)
$exportButton.Text = "Export to CSV"
$exportButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$exportButton.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$exportButton.ForeColor = [System.Drawing.Color]::White
$exportButton.FlatStyle = 'Flat'
$exportButton.Enabled = $false
$buttonPanel.Controls.Add($exportButton)

# Open folder button
$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Location = New-Object System.Drawing.Point(320, 10)
$openFolderButton.Size = New-Object System.Drawing.Size(150, 40)
$openFolderButton.Text = "Open Reports Folder"
$openFolderButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$openFolderButton.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
$openFolderButton.ForeColor = [System.Drawing.Color]::White
$openFolderButton.FlatStyle = 'Flat'
$buttonPanel.Controls.Add($openFolderButton)

# Close button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(710, 10)
$closeButton.Size = New-Object System.Drawing.Size(150, 40)
$closeButton.Text = "Close"
$closeButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$closeButton.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$closeButton.FlatStyle = 'Flat'
$buttonPanel.Controls.Add($closeButton)

# Global variables
$script:isConnected = $false
$script:licenseData = $null
$script:reportFolder = "$env:USERPROFILE\Desktop\O365_Reports"

# Create reports folder if it doesn't exist
if (-not (Test-Path $script:reportFolder)) {
    New-Item -ItemType Directory -Path $script:reportFolder -Force | Out-Null
}

# Function to check and install AzureAD module
function Install-AzureADModuleIfNeeded {
    $statusBar.Text = "Checking for AzureAD module..."
    [System.Windows.Forms.Application]::DoEvents()
    
    $module = Get-Module -ListAvailable -Name AzureAD | Select-Object -First 1
    if (-not $module) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "AzureAD module is not installed. Would you like to install it now?`n`nThis may take a few minutes.",
            "Module Required",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -eq 'Yes') {
            try {
                $statusBar.Text = "Installing AzureAD module... Please wait..."
                [System.Windows.Forms.Application]::DoEvents()
                
                Install-Module AzureAD -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                
                [System.Windows.Forms.MessageBox]::Show(
                    "AzureAD module installed successfully!",
                    "Success",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                $statusBar.Text = "Module installed successfully"
                return $true
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to install AzureAD module:`n`n$($_.Exception.Message)",
                    "Installation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                $statusBar.Text = "Module installation failed"
                return $false
            }
        } else {
            return $false
        }
    } else {
        $statusBar.Text = "AzureAD module found (Version: $($module.Version))"
        return $true
    }
}

# Function to connect to Azure AD
function Connect-ToAzureAD {
    if (-not (Install-AzureADModuleIfNeeded)) {
        return
    }
    
    try {
        $statusBar.Text = "Importing AzureAD module..."
        [System.Windows.Forms.Application]::DoEvents()
        Import-Module AzureAD -ErrorAction Stop
        
        $statusBar.Text = "Connecting to Azure AD... Please complete authentication in the browser window."
        [System.Windows.Forms.Application]::DoEvents()
        
        $connection = Connect-AzureAD
        
        $script:isConnected = $true
        $connectionStatusLabel.Text = "Status: Connected to $($connection.TenantDomain)"
        $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Green
        $connectButton.Text = "Disconnect"
        $connectButton.BackColor = [System.Drawing.Color]::FromArgb(200, 50, 50)
        $refreshButton.Enabled = $true
        $exportButton.Enabled = $true
        $statusBar.Text = "Connected successfully"
        
        # Auto-load licenses
        Load-Licenses
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Azure AD:`n`n$($_.Exception.Message)",
            "Connection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $statusBar.Text = "Connection failed"
    }
}

# Function to disconnect
function Disconnect-FromAzureAD {
    try {
        Disconnect-AzureAD
        $script:isConnected = $false
        $connectionStatusLabel.Text = "Status: Not Connected"
        $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Black
        $connectButton.Text = "Connect to Azure AD"
        $connectButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        $refreshButton.Enabled = $false
        $exportButton.Enabled = $false
        $dataGridView.Rows.Clear()
        $statusBar.Text = "Disconnected"
    } catch {
        $statusBar.Text = "Error during disconnect"
    }
}

# Function to load licenses
function Load-Licenses {
    try {
        $statusBar.Text = "Loading license information..."
        [System.Windows.Forms.Application]::DoEvents()
        
        $licenses = Get-AzureADSubscribedSku
        $script:licenseData = $licenses
        
        # Clear existing data
        $dataGridView.Rows.Clear()
        $dataGridView.Columns.Clear()
        
        # Add columns
        $dataGridView.Columns.Add("LicenseType", "License Type") | Out-Null
        $dataGridView.Columns.Add("Total", "Total") | Out-Null
        $dataGridView.Columns.Add("Assigned", "Assigned") | Out-Null
        $dataGridView.Columns.Add("Available", "Available") | Out-Null
        $dataGridView.Columns.Add("PercentUsed", "% Used") | Out-Null
        $dataGridView.Columns.Add("Status", "Status") | Out-Null
        
        # Set column widths
        $dataGridView.Columns["LicenseType"].Width = 300
        $dataGridView.Columns["Total"].Width = 100
        $dataGridView.Columns["Assigned"].Width = 100
        $dataGridView.Columns["Available"].Width = 100
        $dataGridView.Columns["PercentUsed"].Width = 100
        $dataGridView.Columns["Status"].Width = 150
        
        # Add data
        foreach ($license in $licenses) {
            $total = $license.PrepaidUnits.Enabled
            $consumed = $license.ConsumedUnits
            $available = $total - $consumed
            $percentUsed = if ($total -gt 0) { [math]::Round(($consumed / $total) * 100, 2) } else { 0 }
            
            $status = if ($available -eq 0) { "OUT OF LICENSES" }
                     elseif ($available -le 5) { "LOW INVENTORY" }
                     else { "OK" }
            
            $rowIndex = $dataGridView.Rows.Add(
                $license.SkuPartNumber,
                $total,
                $consumed,
                $available,
                "$percentUsed%",
                $status
            )
            
            # Color code the status
            $row = $dataGridView.Rows[$rowIndex]
            if ($status -eq "OUT OF LICENSES") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 220, 220)
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::Red
                $row.Cells["Status"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            } elseif ($status -eq "LOW INVENTORY") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 205)
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::FromArgb(200, 100, 0)
                $row.Cells["Status"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            } else {
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::Green
            }
        }
        
        $statusBar.Text = "Loaded $($licenses.Count) license types"
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load licenses:`n`n$($_.Exception.Message)",
            "Load Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $statusBar.Text = "Failed to load licenses"
    }
}

# Function to export to CSV
function Export-ToCSV {
    if ($script:licenseData -eq $null) {
        [System.Windows.Forms.MessageBox]::Show(
            "No license data to export. Please load licenses first.",
            "No Data",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $outputPath = Join-Path $script:reportFolder "O365_License_Report_$timestamp.csv"
        
        $reportData = $script:licenseData | Select-Object `
            @{Name="License Type";Expression={$_.SkuPartNumber}},
            @{Name="SKU ID";Expression={$_.SkuId}},
            @{Name="Total Licenses";Expression={$_.PrepaidUnits.Enabled}},
            @{Name="Assigned";Expression={$_.ConsumedUnits}},
            @{Name="Available";Expression={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}},
            @{Name="Percentage Used";Expression={
                if ($_.PrepaidUnits.Enabled -gt 0) {
                    [math]::Round(($_.ConsumedUnits / $_.PrepaidUnits.Enabled) * 100, 2)
                } else { 0 }
            }},
            @{Name="Status";Expression={
                $avail = $_.PrepaidUnits.Enabled - $_.ConsumedUnits
                if ($avail -eq 0) { "OUT OF LICENSES" }
                elseif ($avail -le 5) { "LOW INVENTORY" }
                else { "OK" }
            }}
        
        $reportData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Report exported successfully to:`n`n$outputPath`n`nWould you like to open the file?",
            "Export Successful",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        if ($result -eq 'Yes') {
            Start-Process $outputPath
        }
        
        $statusBar.Text = "Report exported successfully"
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to export report:`n`n$($_.Exception.Message)",
            "Export Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $statusBar.Text = "Export failed"
    }
}

# Event handlers
$connectButton.Add_Click({
    if ($script:isConnected) {
        Disconnect-FromAzureAD
    } else {
        Connect-ToAzureAD
    }
})

$refreshButton.Add_Click({
    if ($script:isConnected) {
        Load-Licenses
    }
})

$exportButton.Add_Click({
    Export-ToCSV
})

$openFolderButton.Add_Click({
    if (Test-Path $script:reportFolder) {
        Start-Process $script:reportFolder
    } else {
        New-Item -ItemType Directory -Path $script:reportFolder -Force | Out-Null
        Start-Process $script:reportFolder
    }
})

$closeButton.Add_Click({
    if ($script:isConnected) {
        Disconnect-FromAzureAD
    }
    $form.Close()
})

# Form closing event
$form.Add_FormClosing({
    if ($script:isConnected) {
        try {
            Disconnect-AzureAD
        } catch {}
    }
})

# Show the form
[void]$form.ShowDialog()
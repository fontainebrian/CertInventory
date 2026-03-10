# Laptop Inventory Management Application
# Allows importing CSV, comparing with in-stock inventory, and managing inventory

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
# Handle $PSScriptRoot being empty (e.g., when run in ISE or unsaved script)
$scriptPath = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { 
    Get-Location 
} else { 
    $PSScriptRoot 
}
$script:inventoryFile = Join-Path $scriptPath "inventory.json"
$script:missingFile = Join-Path $scriptPath "missing.json"
$script:awaitingFile = Join-Path $scriptPath "awaiting_verification.json"
$script:csvData = @()
$script:inventory = @()
$script:missingLaptops = @()
$script:awaitingVerification = @()

# Load inventory from JSON file
function Load-Inventory {
    if (Test-Path $script:inventoryFile) {
        try {
            $content = Get-Content $script:inventoryFile -Raw
            if ([string]::IsNullOrWhiteSpace($content)) {
                return @()
            }
            $json = $content | ConvertFrom-Json
            # Ensure we always return an array
            if ($null -eq $json) {
                return @()
            } elseif ($json -is [array]) {
                return @($json)
            } else {
                return @($json)
            }
        }
        catch {
            Write-Host "Error loading inventory: $($_.Exception.Message)"
            return @()
        }
    }
    return @()
}

# Save inventory to JSON file
function Save-Inventory {
    if ($null -eq $script:inventory -or $script:inventory.Count -eq 0) {
        "[]" | Set-Content $script:inventoryFile
    } else {
        $script:inventory | ConvertTo-Json -Depth 10 | Set-Content $script:inventoryFile
    }
}

# Load missing laptops from JSON file
function Load-MissingLaptops {
    if (Test-Path $script:missingFile) {
        try {
            $content = Get-Content $script:missingFile -Raw
            if ([string]::IsNullOrWhiteSpace($content)) { return @() }
            $json = $content | ConvertFrom-Json
            if ($null -eq $json) { return @() }
            return @($json)
        }
        catch {
            Write-Host "Error loading missing laptops: $($_.Exception.Message)"
            return @()
        }
    }
    return @()
}

# Save missing laptops to JSON file
function Save-MissingLaptops {
    if ($null -eq $script:missingLaptops -or $script:missingLaptops.Count -eq 0) {
        "[]" | Set-Content $script:missingFile
    } else {
        $script:missingLaptops | ConvertTo-Json -Depth 10 | Set-Content $script:missingFile
    }
}

# Load awaiting verification entries from JSON file
function Load-AwaitingVerification {
    if (Test-Path $script:awaitingFile) {
        try {
            $content = Get-Content $script:awaitingFile -Raw
            if ([string]::IsNullOrWhiteSpace($content)) { return @() }
            $json = $content | ConvertFrom-Json
            if ($null -eq $json) { return @() }
            return @($json)
        }
        catch {
            Write-Host "Error loading awaiting verification: $($_.Exception.Message)"
            return @()
        }
    }
    return @()
}

# Save awaiting verification entries to JSON file
function Save-AwaitingVerification {
    if ($null -eq $script:awaitingVerification -or $script:awaitingVerification.Count -eq 0) {
        "[]" | Set-Content $script:awaitingFile
    } else {
        $script:awaitingVerification | ConvertTo-Json -Depth 10 | Set-Content $script:awaitingFile
    }
}

# Initialize inventory
$script:inventory = Load-Inventory
$script:missingLaptops = Load-MissingLaptops
$script:awaitingVerification = Load-AwaitingVerification

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Laptop Inventory Manager"
$form.Size = New-Object System.Drawing.Size(1200, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Create tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(1165, 640)

# Tab 1: CSV Import and Comparison
$tabImport = New-Object System.Windows.Forms.TabPage
$tabImport.Text = "CSV Import & Compare"

# CSV Import Section
$lblCsvFile = New-Object System.Windows.Forms.Label
$lblCsvFile.Text = "CSV File:"
$lblCsvFile.Location = New-Object System.Drawing.Point(10, 20)
$lblCsvFile.Size = New-Object System.Drawing.Size(80, 20)
$tabImport.Controls.Add($lblCsvFile)

$txtCsvPath = New-Object System.Windows.Forms.TextBox
$txtCsvPath.Location = New-Object System.Drawing.Point(100, 18)
$txtCsvPath.Size = New-Object System.Drawing.Size(800, 20)
$txtCsvPath.ReadOnly = $true
$tabImport.Controls.Add($txtCsvPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Location = New-Object System.Drawing.Point(910, 16)
$btnBrowse.Size = New-Object System.Drawing.Size(80, 25)
$tabImport.Controls.Add($btnBrowse)

$btnLoadCsv = New-Object System.Windows.Forms.Button
$btnLoadCsv.Text = "Load CSV"
$btnLoadCsv.Location = New-Object System.Drawing.Point(1000, 16)
$btnLoadCsv.Size = New-Object System.Drawing.Size(80, 25)
$btnLoadCsv.Enabled = $false
$tabImport.Controls.Add($btnLoadCsv)

# CSV Data Grid
$lblCsvData = New-Object System.Windows.Forms.Label
$lblCsvData.Text = "CSV Data:"
$lblCsvData.Location = New-Object System.Drawing.Point(10, 55)
$lblCsvData.Size = New-Object System.Drawing.Size(100, 20)
$tabImport.Controls.Add($lblCsvData)

# Add to Inventory Button
$btnAddFromCsv = New-Object System.Windows.Forms.Button
$btnAddFromCsv.Text = "Add Selected to Inventory"
$btnAddFromCsv.Location = New-Object System.Drawing.Point(120, 53)
$btnAddFromCsv.Size = New-Object System.Drawing.Size(160, 25)
$btnAddFromCsv.Enabled = $false
$tabImport.Controls.Add($btnAddFromCsv)

# Show Details Button (CSV)
$btnShowDetailsCsv = New-Object System.Windows.Forms.Button
$btnShowDetailsCsv.Text = "Show Details"
$btnShowDetailsCsv.Location = New-Object System.Drawing.Point(290, 53)
$btnShowDetailsCsv.Size = New-Object System.Drawing.Size(100, 25)
$btnShowDetailsCsv.Enabled = $false
$tabImport.Controls.Add($btnShowDetailsCsv)

# Compare Button
$btnCompare = New-Object System.Windows.Forms.Button
$btnCompare.Text = "Compare with Inventory"
$btnCompare.Location = New-Object System.Drawing.Point(400, 53)
$btnCompare.Size = New-Object System.Drawing.Size(150, 25)
$btnCompare.Enabled = $false
$tabImport.Controls.Add($btnCompare)

# Email User Button (CSV tab)
$btnEmailCsvUser = New-Object System.Windows.Forms.Button
$btnEmailCsvUser.Text = "Email User"
$btnEmailCsvUser.Location = New-Object System.Drawing.Point(560, 53)
$btnEmailCsvUser.Size = New-Object System.Drawing.Size(100, 25)
$btnEmailCsvUser.Enabled = $false
$tabImport.Controls.Add($btnEmailCsvUser)

# Add to Missing Laptops Button (CSV tab)
$btnAddToMissing = New-Object System.Windows.Forms.Button
$btnAddToMissing.Text = "Add to Missing Laptops"
$btnAddToMissing.Location = New-Object System.Drawing.Point(670, 53)
$btnAddToMissing.Size = New-Object System.Drawing.Size(160, 25)
$btnAddToMissing.Enabled = $false
$tabImport.Controls.Add($btnAddToMissing)

# Check Ping Button (CSV tab) - separated from load for speed
$btnCheckPing = New-Object System.Windows.Forms.Button
$btnCheckPing.Text = "Check Ping"
$btnCheckPing.Location = New-Object System.Drawing.Point(840, 53)
$btnCheckPing.Size = New-Object System.Drawing.Size(100, 25)
$btnCheckPing.Enabled = $false
$tabImport.Controls.Add($btnCheckPing)

# Run Verification Button (CSV tab)
$btnRunVerification = New-Object System.Windows.Forms.Button
$btnRunVerification.Text = "Run Verification"
$btnRunVerification.Location = New-Object System.Drawing.Point(950, 53)
$btnRunVerification.Size = New-Object System.Drawing.Size(130, 25)
$btnRunVerification.Enabled = $false
$tabImport.Controls.Add($btnRunVerification)

$dgvCsv = New-Object System.Windows.Forms.DataGridView
$dgvCsv.Location = New-Object System.Drawing.Point(10, 85)
$dgvCsv.Size = New-Object System.Drawing.Size(1135, 545)
$dgvCsv.AllowUserToAddRows = $false
$dgvCsv.AllowUserToDeleteRows = $false
$dgvCsv.ReadOnly = $true
$dgvCsv.AutoSizeColumnsMode = "Fill"
$dgvCsv.SelectionMode = "FullRowSelect"

# Context menu for CSV grid - copy serial number
$contextMenuCsv = New-Object System.Windows.Forms.ContextMenuStrip
$menuItemCopyCsvSerial = New-Object System.Windows.Forms.ToolStripMenuItem
$menuItemCopyCsvSerial.Text = "Copy Serial Number"
$menuItemCopyCsvSerial.Add_Click({
    if ($dgvCsv.SelectedRows.Count -gt 0) {
        $serialValue = $dgvCsv.SelectedRows[0].Cells["serial_number"].Value
        if ($null -ne $serialValue -and $serialValue -ne "") {
            [System.Windows.Forms.Clipboard]::SetText($serialValue)
        }
    }
})
$contextMenuCsv.Items.Add($menuItemCopyCsvSerial) | Out-Null
$dgvCsv.ContextMenuStrip = $contextMenuCsv

$tabImport.Controls.Add($dgvCsv)


# Tab 2: Inventory Management
$tabInventory = New-Object System.Windows.Forms.TabPage
$tabInventory.Text = "Manage Inventory"

# Add Laptop Section
$grpAdd = New-Object System.Windows.Forms.GroupBox
$grpAdd.Text = "Add Laptop to Inventory"
$grpAdd.Location = New-Object System.Drawing.Point(10, 10)
$grpAdd.Size = New-Object System.Drawing.Size(1135, 120)

$lblSerial = New-Object System.Windows.Forms.Label
$lblSerial.Text = "Serial Number:"
$lblSerial.Location = New-Object System.Drawing.Point(20, 30)
$lblSerial.Size = New-Object System.Drawing.Size(100, 20)
$grpAdd.Controls.Add($lblSerial)

$txtSerial = New-Object System.Windows.Forms.TextBox
$txtSerial.Location = New-Object System.Drawing.Point(130, 28)
$txtSerial.Size = New-Object System.Drawing.Size(200, 20)
$grpAdd.Controls.Add($txtSerial)

$lblModel = New-Object System.Windows.Forms.Label
$lblModel.Text = "Model:"
$lblModel.Location = New-Object System.Drawing.Point(350, 30)
$lblModel.Size = New-Object System.Drawing.Size(50, 20)
$grpAdd.Controls.Add($lblModel)

$txtModel = New-Object System.Windows.Forms.TextBox
$txtModel.Location = New-Object System.Drawing.Point(410, 28)
$txtModel.Size = New-Object System.Drawing.Size(300, 20)
$grpAdd.Controls.Add($txtModel)

$lblLocation = New-Object System.Windows.Forms.Label
$lblLocation.Text = "Location:"
$lblLocation.Location = New-Object System.Drawing.Point(20, 65)
$lblLocation.Size = New-Object System.Drawing.Size(100, 20)
$grpAdd.Controls.Add($lblLocation)

$cmbLocation = New-Object System.Windows.Forms.ComboBox
$cmbLocation.Location = New-Object System.Drawing.Point(130, 63)
$cmbLocation.Size = New-Object System.Drawing.Size(200, 20)
$cmbLocation.DropDownStyle = "DropDownList"
$cmbLocation.Items.AddRange(@("Denver", "Greenwood Village", "Salt Lake City"))
$cmbLocation.SelectedIndex = 0
$grpAdd.Controls.Add($cmbLocation)

$chkLoaner = New-Object System.Windows.Forms.CheckBox
$chkLoaner.Text = "Loaner Laptop"
$chkLoaner.Location = New-Object System.Drawing.Point(350, 63)
$chkLoaner.Size = New-Object System.Drawing.Size(120, 25)
$grpAdd.Controls.Add($chkLoaner)

$btnAddLaptop = New-Object System.Windows.Forms.Button
$btnAddLaptop.Text = "Add to Inventory"
$btnAddLaptop.Location = New-Object System.Drawing.Point(500, 60)
$btnAddLaptop.Size = New-Object System.Drawing.Size(120, 30)
$grpAdd.Controls.Add($btnAddLaptop)

$tabInventory.Controls.Add($grpAdd)

# Search Section
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search:"
$lblSearch.Location = New-Object System.Drawing.Point(10, 145)
$lblSearch.Size = New-Object System.Drawing.Size(60, 20)
$tabInventory.Controls.Add($lblSearch)

$txtSearchInventory = New-Object System.Windows.Forms.TextBox
$txtSearchInventory.Location = New-Object System.Drawing.Point(75, 143)
$txtSearchInventory.Size = New-Object System.Drawing.Size(300, 20)
$tabInventory.Controls.Add($txtSearchInventory)

$btnClearSearch = New-Object System.Windows.Forms.Button
$btnClearSearch.Text = "Clear"
$btnClearSearch.Location = New-Object System.Drawing.Point(385, 141)
$btnClearSearch.Size = New-Object System.Drawing.Size(60, 25)
$tabInventory.Controls.Add($btnClearSearch)

# Current Inventory Section
$lblInventory = New-Object System.Windows.Forms.Label
$lblInventory.Text = "Current Inventory:"
$lblInventory.Location = New-Object System.Drawing.Point(460, 145)
$lblInventory.Size = New-Object System.Drawing.Size(150, 20)
$tabInventory.Controls.Add($lblInventory)

$dgvInventory = New-Object System.Windows.Forms.DataGridView
$dgvInventory.Location = New-Object System.Drawing.Point(10, 170)
$dgvInventory.Size = New-Object System.Drawing.Size(1135, 370)
$dgvInventory.AllowUserToAddRows = $false
$dgvInventory.AllowUserToDeleteRows = $false
$dgvInventory.ReadOnly = $true
$dgvInventory.AutoSizeColumnsMode = "Fill"
$dgvInventory.SelectionMode = "FullRowSelect"

# Create context menu for copying serial number
$contextMenuInventory = New-Object System.Windows.Forms.ContextMenuStrip
$menuItemCopySerial = New-Object System.Windows.Forms.ToolStripMenuItem
$menuItemCopySerial.Text = "Copy Serial Number"
$menuItemCopySerial.Add_Click({
    if ($dgvInventory.SelectedRows.Count -gt 0) {
        $selectedRow = $dgvInventory.SelectedRows[0]
        $serialNumber = $selectedRow.Cells["SerialNumber"].Value
        if ($null -ne $serialNumber) {
            [System.Windows.Forms.Clipboard]::SetText($serialNumber)
            [System.Windows.Forms.MessageBox]::Show("Serial number copied to clipboard: $serialNumber", "Copied", "OK", "Information")
        }
    }
})
$contextMenuInventory.Items.Add($menuItemCopySerial) | Out-Null
$dgvInventory.ContextMenuStrip = $contextMenuInventory

$tabInventory.Controls.Add($dgvInventory)

$btnRemoveLaptop = New-Object System.Windows.Forms.Button
$btnRemoveLaptop.Text = "Remove Selected"
$btnRemoveLaptop.Location = New-Object System.Drawing.Point(10, 550)
$btnRemoveLaptop.Size = New-Object System.Drawing.Size(120, 30)
$tabInventory.Controls.Add($btnRemoveLaptop)

$btnRefreshInventory = New-Object System.Windows.Forms.Button
$btnRefreshInventory.Text = "Refresh"
$btnRefreshInventory.Location = New-Object System.Drawing.Point(140, 550)
$btnRefreshInventory.Size = New-Object System.Drawing.Size(80, 30)
$tabInventory.Controls.Add($btnRefreshInventory)

# Tab 3: Device Lookup
$tabLookup = New-Object System.Windows.Forms.TabPage
$tabLookup.Text = "Device Lookup"

# Lookup Section
$grpLookup = New-Object System.Windows.Forms.GroupBox
$grpLookup.Text = "Check Device Status"
$grpLookup.Location = New-Object System.Drawing.Point(10, 10)
$grpLookup.Size = New-Object System.Drawing.Size(1135, 100)

$lblLookupSerial = New-Object System.Windows.Forms.Label
$lblLookupSerial.Text = "Serial Number:"
$lblLookupSerial.Location = New-Object System.Drawing.Point(20, 35)
$lblLookupSerial.Size = New-Object System.Drawing.Size(100, 20)
$grpLookup.Controls.Add($lblLookupSerial)

$txtLookupSerial = New-Object System.Windows.Forms.TextBox
$txtLookupSerial.Location = New-Object System.Drawing.Point(130, 33)
$txtLookupSerial.Size = New-Object System.Drawing.Size(300, 20)
$grpLookup.Controls.Add($txtLookupSerial)

$btnCheckDevice = New-Object System.Windows.Forms.Button
$btnCheckDevice.Text = "Check Device"
$btnCheckDevice.Location = New-Object System.Drawing.Point(440, 30)
$btnCheckDevice.Size = New-Object System.Drawing.Size(120, 30)
$grpLookup.Controls.Add($btnCheckDevice)

$btnClearLookup = New-Object System.Windows.Forms.Button
$btnClearLookup.Text = "Clear"
$btnClearLookup.Location = New-Object System.Drawing.Point(570, 30)
$btnClearLookup.Size = New-Object System.Drawing.Size(80, 30)
$grpLookup.Controls.Add($btnClearLookup)

$tabLookup.Controls.Add($grpLookup)

# Results Section
$lblLookupResults = New-Object System.Windows.Forms.Label
$lblLookupResults.Text = "Device Information:"
$lblLookupResults.Location = New-Object System.Drawing.Point(10, 120)
$lblLookupResults.Size = New-Object System.Drawing.Size(150, 20)
$tabLookup.Controls.Add($lblLookupResults)

$txtLookupResults = New-Object System.Windows.Forms.TextBox
$txtLookupResults.Location = New-Object System.Drawing.Point(10, 145)
$txtLookupResults.Size = New-Object System.Drawing.Size(1135, 450)
$txtLookupResults.Multiline = $true
$txtLookupResults.ScrollBars = "Vertical"
$txtLookupResults.ReadOnly = $true
$txtLookupResults.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLookupResults.Text = "Enter a serial number and click 'Check Device' to view information."
$tabLookup.Controls.Add($txtLookupResults)

# Tab 4: Loaner Laptops
$tabLoaner = New-Object System.Windows.Forms.TabPage
$tabLoaner.Text = "Loaner Laptops"

# Loaner Laptops Section
$lblLoanerList = New-Object System.Windows.Forms.Label
$lblLoanerList.Text = "Loaner Laptops:"
$lblLoanerList.Location = New-Object System.Drawing.Point(10, 20)
$lblLoanerList.Size = New-Object System.Drawing.Size(150, 20)
$tabLoaner.Controls.Add($lblLoanerList)

# DataGridView for loaner laptops
$dgvLoaners = New-Object System.Windows.Forms.DataGridView
$dgvLoaners.Location = New-Object System.Drawing.Point(10, 45)
$dgvLoaners.Size = New-Object System.Drawing.Size(1135, 500)
$dgvLoaners.AllowUserToAddRows = $false
$dgvLoaners.AllowUserToDeleteRows = $false
$dgvLoaners.ReadOnly = $true
$dgvLoaners.AutoSizeColumnsMode = "Fill"
$dgvLoaners.SelectionMode = "FullRowSelect"
$tabLoaner.Controls.Add($dgvLoaners)

# Add Button
$btnAddLoaner = New-Object System.Windows.Forms.Button
$btnAddLoaner.Text = "Add Loaner"
$btnAddLoaner.Location = New-Object System.Drawing.Point(10, 555)
$btnAddLoaner.Size = New-Object System.Drawing.Size(120, 30)
$tabLoaner.Controls.Add($btnAddLoaner)

# Remove Button
$btnRemoveLoaner = New-Object System.Windows.Forms.Button
$btnRemoveLoaner.Text = "Remove Selected"
$btnRemoveLoaner.Location = New-Object System.Drawing.Point(140, 555)
$btnRemoveLoaner.Size = New-Object System.Drawing.Size(120, 30)
$tabLoaner.Controls.Add($btnRemoveLoaner)

# Refresh Button
$btnRefreshLoaners = New-Object System.Windows.Forms.Button
$btnRefreshLoaners.Text = "Refresh"
$btnRefreshLoaners.Location = New-Object System.Drawing.Point(270, 555)
$btnRefreshLoaners.Size = New-Object System.Drawing.Size(100, 30)
$tabLoaner.Controls.Add($btnRefreshLoaners)

# Check Status Button
$btnCheckStatus = New-Object System.Windows.Forms.Button
$btnCheckStatus.Text = "Check Status"
$btnCheckStatus.Location = New-Object System.Drawing.Point(380, 555)
$btnCheckStatus.Size = New-Object System.Drawing.Size(100, 30)
$tabLoaner.Controls.Add($btnCheckStatus)

# Assign Button
$btnAssignLoaner = New-Object System.Windows.Forms.Button
$btnAssignLoaner.Text = "Assign"
$btnAssignLoaner.Location = New-Object System.Drawing.Point(490, 555)
$btnAssignLoaner.Size = New-Object System.Drawing.Size(100, 30)
$tabLoaner.Controls.Add($btnAssignLoaner)

# Unassign Button
$btnUnassignLoaner = New-Object System.Windows.Forms.Button
$btnUnassignLoaner.Text = "Unassign"
$btnUnassignLoaner.Location = New-Object System.Drawing.Point(600, 555)
$btnUnassignLoaner.Size = New-Object System.Drawing.Size(100, 30)
$tabLoaner.Controls.Add($btnUnassignLoaner)

# Show Details Button
$btnShowDetailsLoaner = New-Object System.Windows.Forms.Button
$btnShowDetailsLoaner.Text = "Show Details"
$btnShowDetailsLoaner.Location = New-Object System.Drawing.Point(710, 555)
$btnShowDetailsLoaner.Size = New-Object System.Drawing.Size(100, 30)
$tabLoaner.Controls.Add($btnShowDetailsLoaner)

# Email User Button
$btnEmailLoanerUser = New-Object System.Windows.Forms.Button
$btnEmailLoanerUser.Text = "Email User"
$btnEmailLoanerUser.Location = New-Object System.Drawing.Point(820, 555)
$btnEmailLoanerUser.Size = New-Object System.Drawing.Size(100, 30)
$tabLoaner.Controls.Add($btnEmailLoanerUser)

# Tab 5: Missing Laptops
$tabMissing = New-Object System.Windows.Forms.TabPage
$tabMissing.Text = "Missing Laptops"

$lblMissingList = New-Object System.Windows.Forms.Label
$lblMissingList.Text = "Missing Laptops:"
$lblMissingList.Location = New-Object System.Drawing.Point(10, 20)
$lblMissingList.Size = New-Object System.Drawing.Size(150, 20)
$tabMissing.Controls.Add($lblMissingList)

$dgvMissing = New-Object System.Windows.Forms.DataGridView
$dgvMissing.Location = New-Object System.Drawing.Point(10, 45)
$dgvMissing.Size = New-Object System.Drawing.Size(1135, 500)
$dgvMissing.AllowUserToAddRows = $false
$dgvMissing.AllowUserToDeleteRows = $false
$dgvMissing.ReadOnly = $true
$dgvMissing.AutoSizeColumnsMode = "Fill"
$dgvMissing.SelectionMode = "FullRowSelect"

# Context menu for Missing grid - copy serial number
$contextMenuMissing = New-Object System.Windows.Forms.ContextMenuStrip
$menuItemCopyMissingSerial = New-Object System.Windows.Forms.ToolStripMenuItem
$menuItemCopyMissingSerial.Text = "Copy Serial Number"
$menuItemCopyMissingSerial.Add_Click({
    if ($dgvMissing.SelectedRows.Count -gt 0) {
        $val = $dgvMissing.SelectedRows[0].Cells["SerialNumber"].Value
        if ($null -ne $val -and $val -ne "") {
            [System.Windows.Forms.Clipboard]::SetText($val)
        }
    }
})
$contextMenuMissing.Items.Add($menuItemCopyMissingSerial) | Out-Null
$dgvMissing.ContextMenuStrip = $contextMenuMissing
$tabMissing.Controls.Add($dgvMissing)

$btnRemoveMissing = New-Object System.Windows.Forms.Button
$btnRemoveMissing.Text = "Remove Selected"
$btnRemoveMissing.Location = New-Object System.Drawing.Point(10, 555)
$btnRemoveMissing.Size = New-Object System.Drawing.Size(130, 30)
$tabMissing.Controls.Add($btnRemoveMissing)

$btnRefreshMissing = New-Object System.Windows.Forms.Button
$btnRefreshMissing.Text = "Refresh"
$btnRefreshMissing.Location = New-Object System.Drawing.Point(150, 555)
$btnRefreshMissing.Size = New-Object System.Drawing.Size(80, 30)
$tabMissing.Controls.Add($btnRefreshMissing)

$btnEmailMissingUser = New-Object System.Windows.Forms.Button
$btnEmailMissingUser.Text = "Email User"
$btnEmailMissingUser.Location = New-Object System.Drawing.Point(240, 555)
$btnEmailMissingUser.Size = New-Object System.Drawing.Size(100, 30)
$tabMissing.Controls.Add($btnEmailMissingUser)

# Tab 6: Awaiting Verification
$tabAwaiting = New-Object System.Windows.Forms.TabPage
$tabAwaiting.Text = "Awaiting Verification"

$lblAwaitingList = New-Object System.Windows.Forms.Label
$lblAwaitingList.Text = "Laptops awaiting verification (email sent, response pending):"
$lblAwaitingList.Location = New-Object System.Drawing.Point(10, 20)
$lblAwaitingList.Size = New-Object System.Drawing.Size(600, 20)
$tabAwaiting.Controls.Add($lblAwaitingList)

$dgvAwaiting = New-Object System.Windows.Forms.DataGridView
$dgvAwaiting.Location = New-Object System.Drawing.Point(10, 45)
$dgvAwaiting.Size = New-Object System.Drawing.Size(1135, 500)
$dgvAwaiting.AllowUserToAddRows = $false
$dgvAwaiting.AllowUserToDeleteRows = $false
$dgvAwaiting.ReadOnly = $true
$dgvAwaiting.AutoSizeColumnsMode = "Fill"
$dgvAwaiting.SelectionMode = "FullRowSelect"

# Context menu for Awaiting grid - copy serial number
$contextMenuAwaiting = New-Object System.Windows.Forms.ContextMenuStrip
$menuItemCopyAwaitingSerial = New-Object System.Windows.Forms.ToolStripMenuItem
$menuItemCopyAwaitingSerial.Text = "Copy Serial Number"
$menuItemCopyAwaitingSerial.Add_Click({
    if ($dgvAwaiting.SelectedRows.Count -gt 0) {
        $val = $dgvAwaiting.SelectedRows[0].Cells["SerialNumber"].Value
        if ($null -ne $val -and $val -ne "") {
            [System.Windows.Forms.Clipboard]::SetText($val)
        }
    }
})
$contextMenuAwaiting.Items.Add($menuItemCopyAwaitingSerial) | Out-Null
$dgvAwaiting.ContextMenuStrip = $contextMenuAwaiting
$tabAwaiting.Controls.Add($dgvAwaiting)

$btnRemoveAwaiting = New-Object System.Windows.Forms.Button
$btnRemoveAwaiting.Text = "Mark Verified (Remove)"
$btnRemoveAwaiting.Location = New-Object System.Drawing.Point(10, 555)
$btnRemoveAwaiting.Size = New-Object System.Drawing.Size(160, 30)
$tabAwaiting.Controls.Add($btnRemoveAwaiting)

$btnRefreshAwaiting = New-Object System.Windows.Forms.Button
$btnRefreshAwaiting.Text = "Refresh"
$btnRefreshAwaiting.Location = New-Object System.Drawing.Point(180, 555)
$btnRefreshAwaiting.Size = New-Object System.Drawing.Size(80, 30)
$tabAwaiting.Controls.Add($btnRefreshAwaiting)

$btnEmailAwaitingUser = New-Object System.Windows.Forms.Button
$btnEmailAwaitingUser.Text = "Re-send Email"
$btnEmailAwaitingUser.Location = New-Object System.Drawing.Point(270, 555)
$btnEmailAwaitingUser.Size = New-Object System.Drawing.Size(120, 30)
$tabAwaiting.Controls.Add($btnEmailAwaitingUser)

# Tab 7: Needs Verification
$tabNeedsVerification = New-Object System.Windows.Forms.TabPage
$tabNeedsVerification.Text = "Needs Verification"

$lblNeedsVerification = New-Object System.Windows.Forms.Label
$lblNeedsVerification.Text = "Devices that need verification (run 'Run Verification' on the CSV tab first):"
$lblNeedsVerification.Location = New-Object System.Drawing.Point(10, 20)
$lblNeedsVerification.Size = New-Object System.Drawing.Size(700, 20)
$tabNeedsVerification.Controls.Add($lblNeedsVerification)

$dgvNeedsVerification = New-Object System.Windows.Forms.DataGridView
$dgvNeedsVerification.Location = New-Object System.Drawing.Point(10, 45)
$dgvNeedsVerification.Size = New-Object System.Drawing.Size(1135, 490)
$dgvNeedsVerification.AllowUserToAddRows = $false
$dgvNeedsVerification.AllowUserToDeleteRows = $false
$dgvNeedsVerification.ReadOnly = $true
$dgvNeedsVerification.AutoSizeColumnsMode = "Fill"
$dgvNeedsVerification.SelectionMode = "FullRowSelect"

# Context menu for Needs Verification grid
$contextMenuNeedsVerif = New-Object System.Windows.Forms.ContextMenuStrip
$menuItemCopyNVSerial = New-Object System.Windows.Forms.ToolStripMenuItem
$menuItemCopyNVSerial.Text = "Copy Serial Number"
$menuItemCopyNVSerial.Add_Click({
    if ($dgvNeedsVerification.SelectedRows.Count -gt 0) {
        $val = $dgvNeedsVerification.SelectedRows[0].Cells["SerialNumber"].Value
        if ($null -ne $val -and $val -ne "") { [System.Windows.Forms.Clipboard]::SetText($val) }
    }
})
$contextMenuNeedsVerif.Items.Add($menuItemCopyNVSerial) | Out-Null
$dgvNeedsVerification.ContextMenuStrip = $contextMenuNeedsVerif
$tabNeedsVerification.Controls.Add($dgvNeedsVerification)

$btnRefreshNeedsVerif = New-Object System.Windows.Forms.Button
$btnRefreshNeedsVerif.Text = "Refresh"
$btnRefreshNeedsVerif.Location = New-Object System.Drawing.Point(10, 545)
$btnRefreshNeedsVerif.Size = New-Object System.Drawing.Size(80, 30)
$tabNeedsVerification.Controls.Add($btnRefreshNeedsVerif)

$btnEmailSelectedNeedsVerif = New-Object System.Windows.Forms.Button
$btnEmailSelectedNeedsVerif.Text = "Email Selected"
$btnEmailSelectedNeedsVerif.Location = New-Object System.Drawing.Point(100, 545)
$btnEmailSelectedNeedsVerif.Size = New-Object System.Drawing.Size(120, 30)
$tabNeedsVerification.Controls.Add($btnEmailSelectedNeedsVerif)

$btnEmailAllNeedsVerif = New-Object System.Windows.Forms.Button
$btnEmailAllNeedsVerif.Text = "Send Email to All"
$btnEmailAllNeedsVerif.Location = New-Object System.Drawing.Point(230, 545)
$btnEmailAllNeedsVerif.Size = New-Object System.Drawing.Size(140, 30)
$btnEmailAllNeedsVerif.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
$btnEmailAllNeedsVerif.ForeColor = [System.Drawing.Color]::White
$btnEmailAllNeedsVerif.FlatStyle = "Flat"
$tabNeedsVerification.Controls.Add($btnEmailAllNeedsVerif)

$lblEmailAllStatus = New-Object System.Windows.Forms.Label
$lblEmailAllStatus.Text = ""
$lblEmailAllStatus.Location = New-Object System.Drawing.Point(385, 550)
$lblEmailAllStatus.Size = New-Object System.Drawing.Size(750, 20)
$lblEmailAllStatus.ForeColor = [System.Drawing.Color]::DarkGreen
$tabNeedsVerification.Controls.Add($lblEmailAllStatus)

# Add tabs to control
$tabControl.TabPages.Add($tabImport)
$tabControl.TabPages.Add($tabInventory)
$tabControl.TabPages.Add($tabLookup)
$tabControl.TabPages.Add($tabLoaner)
$tabControl.TabPages.Add($tabMissing)
$tabControl.TabPages.Add($tabAwaiting)
$tabControl.TabPages.Add($tabNeedsVerification)
$form.Controls.Add($tabControl)

# Event Handlers

# Browse for CSV file
$btnBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV File"
    
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $txtCsvPath.Text = $openFileDialog.FileName
        $btnLoadCsv.Enabled = $true
    }
})

# Load CSV file
$btnLoadCsv.Add_Click({
    try {
        $script:csvData = Import-Csv -Path $txtCsvPath.Text
        
        # Create DataTable for display
        $dt = New-Object System.Data.DataTable
        
        if ($script:csvData.Count -gt 0) {
            # Add columns from CSV
            $script:csvData[0].PSObject.Properties.Name | ForEach-Object {
                $dt.Columns.Add($_) | Out-Null
            }
            
            # Add extra columns
            $dt.Columns.Add("Pingable")     | Out-Null
            $dt.Columns.Add("In AD")        | Out-Null
            $dt.Columns.Add("User in AD?")  | Out-Null
            $dt.Columns.Add("Missing?")     | Out-Null
            $dt.Columns.Add("Verified")     | Out-Null

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $dgvCsv.DataSource = $null

            # Pre-load missing list once for O(1) lookups
            $currentMissing = Load-MissingLaptops
            $missingSerials = @{}
            foreach ($m in $currentMissing) { $missingSerials[$m.SerialNumber.Trim()] = $true }

            # Pre-load awaiting verification serials for O(1) lookups
            $script:awaitingVerification = Load-AwaitingVerification
            $awaitingSerials = @{}
            foreach ($a in $script:awaitingVerification) { $awaitingSerials[$a.SerialNumber.Trim()] = $true }

            # Batch-fetch ALL AD computers in one query and build a hashtable for O(1) lookup
            $adComputerSet = @{}
            try {
                Get-ADComputer -Filter * -Properties Name -ErrorAction Stop | ForEach-Object {
                    $adComputerSet[$_.Name.ToUpper()] = $true
                }
            } catch { }

            # AD user lookup cache - avoids re-querying the same name multiple times
            $adUserCache = @{}

            foreach ($row in $script:csvData) {
                $dr = $dt.NewRow()
                foreach ($prop in $row.PSObject.Properties) {
                    $dr[$prop.Name] = $prop.Value
                }

                # Pingable starts as blank - populated on demand via Check Ping button
                $dr["Pingable"] = "-"

                # In AD - O(1) hashtable lookup, no network call per row
                $inADVal = if ($adComputerSet.ContainsKey($row.serial_number.ToUpper())) { "Yes" } else { "No" }
                $dr["In AD"] = $inADVal

                # User in AD? - cached so each unique name is only queried once
                $assignedUser = if ($row.PSObject.Properties['assigned_to']) { $row.assigned_to.Trim() } else { "" }
                if (-not [string]::IsNullOrWhiteSpace($assignedUser)) {
                    if (-not $adUserCache.ContainsKey($assignedUser)) {
                        $foundUser = $false
                        $safeName = $assignedUser -replace "'", "''"

                        if (-not $foundUser) {
                            try {
                                $r = @(Get-ADUser -Filter "DisplayName -eq '$safeName'" -ErrorAction Stop)
                                if ($r.Count -gt 0) { $foundUser = $true }
                            } catch { }
                        }
                        if (-not $foundUser) {
                            try {
                                $r = @(Get-ADUser -Filter "Name -eq '$safeName'" -ErrorAction Stop)
                                if ($r.Count -gt 0) { $foundUser = $true }
                            } catch { }
                        }
                        if (-not $foundUser -and $assignedUser -match ' ') {
                            try {
                                $nameParts = $assignedUser -split '\s+', 2
                                $fn = $nameParts[0].Trim() -replace "'", "''"
                                $ln = $nameParts[1].Trim() -replace "'", "''"
                                $r = @(Get-ADUser -Filter "GivenName -eq '$fn' -and Surname -eq '$ln'" -ErrorAction Stop)
                                if ($r.Count -gt 0) { $foundUser = $true }
                            } catch { }
                        }
                        $adUserCache[$assignedUser] = $foundUser
                    }
                    $dr["User in AD?"] = if ($adUserCache[$assignedUser]) { "Yes" } else { "No" }
                } else {
                    $dr["User in AD?"] = "N/A"
                }

                # Missing list - O(1) hashtable lookup
                $dr["Missing?"] = if ($missingSerials.ContainsKey($row.serial_number.Trim())) { "Yes" } else { "No" }

                # Verified - show Email Sent if already in awaiting list, else Pending
                $verifiedVal = if ($awaitingSerials.ContainsKey($row.serial_number.Trim())) { "Email Sent" } else { "Pending" }
                $dr["Verified"] = $verifiedVal

                $dt.Rows.Add($dr)
            }

            $form.Cursor = [System.Windows.Forms.Cursors]::Default

            # Suspend redraws while binding and colouring
            $dgvCsv.SuspendLayout()
            $dgvCsv.DataSource = $dt

            foreach ($row in $dgvCsv.Rows) {
                # In AD
                if ($row.Cells["In AD"].Value -eq "Yes") {
                    $row.Cells["In AD"].Style.ForeColor = [System.Drawing.Color]::Green
                } else {
                    $row.Cells["In AD"].Style.ForeColor = [System.Drawing.Color]::Red
                }

                # User in AD?
                switch ($row.Cells["User in AD?"].Value) {
                    "Yes" { $row.Cells["User in AD?"].Style.ForeColor = [System.Drawing.Color]::Green }
                    "No"  { $row.Cells["User in AD?"].Style.ForeColor = [System.Drawing.Color]::Red }
                    "N/A" { $row.Cells["User in AD?"].Style.ForeColor = [System.Drawing.Color]::Gray }
                }

                # Missing?
                if ($row.Cells["Missing?"].Value -eq "Yes") {
                    $row.Cells["Missing?"].Style.ForeColor = [System.Drawing.Color]::Red
                    $row.Cells["Missing?"].Style.Font = New-Object System.Drawing.Font($dgvCsv.Font, [System.Drawing.FontStyle]::Bold)
                } else {
                    $row.Cells["Missing?"].Style.ForeColor = [System.Drawing.Color]::Green
                }

                # Verified
                switch ($row.Cells["Verified"].Value) {
                    "Email Sent" {
                        $row.Cells["Verified"].Style.ForeColor = [System.Drawing.Color]::DarkBlue
                        $row.Cells["Verified"].Style.Font = New-Object System.Drawing.Font($dgvCsv.Font, [System.Drawing.FontStyle]::Bold)
                    }
                    default {
                        $row.Cells["Verified"].Style.ForeColor = [System.Drawing.Color]::Gray
                    }
                }
            }

            $dgvCsv.ResumeLayout()

            $btnCompare.Enabled = $true
            $btnAddFromCsv.Enabled = $true
            $btnShowDetailsCsv.Enabled = $true
            $btnEmailCsvUser.Enabled = $true
            $btnAddToMissing.Enabled = $true
            $btnCheckPing.Enabled = $true
            $btnRunVerification.Enabled = $true
            [System.Windows.Forms.MessageBox]::Show("CSV file loaded successfully! Found $($script:csvData.Count) records.`n`nClick 'Check Ping' to test connectivity or 'Run Verification' to verify laptops.", "Success", "OK", "Information")
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading CSV file: $($_.Exception.Message)", "Error", "OK", "Error")
    }
})

# Check Ping button (CSV tab) - on-demand connectivity check
$btnCheckPing.Add_Click({
    $dt = $dgvCsv.DataSource
    if (-not $dt -or $dt.Rows.Count -eq 0) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $dgvCsv.SuspendLayout()

    for ($i = 0; $i -lt $script:csvData.Count; $i++) {
        $serial = $script:csvData[$i].serial_number
        $pingable = "Offline"
        try {
            $pingTest = Test-Connection -ComputerName $serial -Count 1 -Quiet -ErrorAction Stop
            if ($pingTest -eq $true) { $pingable = "Online" }
        }
        catch {
            if ($_.Exception.Message -like "*could not be resolved*" -or $_.Exception.Message -like "*host not found*") {
                $pingable = "DNS Error"
            }
        }
        $dt.Rows[$i]["Pingable"] = $pingable

        $cell = $dgvCsv.Rows[$i].Cells["Pingable"]
        switch ($pingable) {
            "Online"    { $cell.Style.ForeColor = [System.Drawing.Color]::Green }
            "Offline"   { $cell.Style.ForeColor = [System.Drawing.Color]::Red }
            "DNS Error" { $cell.Style.ForeColor = [System.Drawing.Color]::DarkOrange }
        }
    }

    $dgvCsv.ResumeLayout()
    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    $onlineCount = ($script:csvData.Count..1 | ForEach-Object { $dt.Rows[$_ - 1]["Pingable"] } | Where-Object { $_ -eq "Online" }).Count
    [System.Windows.Forms.MessageBox]::Show("Ping check complete!", "Ping Done", "OK", "Information")
})

# Run Verification button - checks all CSV rows and marks Verified / Needs Verification
$btnRunVerification.Add_Click({
    $dt = $dgvCsv.DataSource
    if (-not $dt -or $dt.Rows.Count -eq 0) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    # Batch-fetch AD computers with LastLogonDate for O(1) lookups
    $adComputerInfo = @{}
    try {
        Get-ADComputer -Filter * -Properties Name, LastLogonDate -ErrorAction Stop | ForEach-Object {
            $adComputerInfo[$_.Name.ToUpper()] = $_.LastLogonDate
        }
    } catch { }

    # Build a set of serial numbers in local inventory (non-loaner, unassigned = in stock)
    $inventorySerials = @{}
    $localInventory = @(Load-Inventory)
    foreach ($item in $localInventory) {
        if ($item.SerialNumber) {
            $inventorySerials[$item.SerialNumber.Trim().ToUpper()] = $true
        }
    }

    # Reload missing serials
    $currentMissing = Load-MissingLaptops
    $missingSerials = @{}
    foreach ($m in $currentMissing) { $missingSerials[$m.SerialNumber.Trim().ToUpper()] = $true }

    $cutoff = (Get-Date).AddDays(-7)
    $verifiedCount = 0
    $needsCount = 0

    $dgvCsv.SuspendLayout()

    for ($i = 0; $i -lt $script:csvData.Count; $i++) {
        $row = $script:csvData[$i]
        $serial = $row.serial_number.Trim()
        $assignedUser = if ($row.PSObject.Properties['assigned_to']) { $row.assigned_to.Trim() } else { "" }
        $serialUpper = $serial.ToUpper()

        $status = "Needs Verification"

        # Rule 1: Flagged as missing → always Needs Verification
        if ($missingSerials.ContainsKey($serialUpper)) {
            $status = "Missing"
        }
        # Rule 2: In local inventory with no assigned user → in stock, verified
        elseif ($inventorySerials.ContainsKey($serialUpper) -and [string]::IsNullOrWhiteSpace($assignedUser)) {
            $status = "Verified (In Stock)"
            $verifiedCount++
        }
        # Rule 3: In AD with a LastLogonDate within the past 7 days
        elseif ($adComputerInfo.ContainsKey($serialUpper)) {
            $lastLogon = $adComputerInfo[$serialUpper]
            if ($lastLogon -and $lastLogon -gt $cutoff) {
                $status = "Verified (Active)"
                $verifiedCount++
            } else {
                $status = "Needs Verification"
                $needsCount++
            }
        }
        else {
            $needsCount++
        }

        $dt.Rows[$i]["Verified"] = $status

        $cell = $dgvCsv.Rows[$i].Cells["Verified"]
        switch -Wildcard ($status) {
            "Verified*" {
                $cell.Style.ForeColor = [System.Drawing.Color]::Green
                $cell.Style.Font = New-Object System.Drawing.Font($dgvCsv.Font, [System.Drawing.FontStyle]::Bold)
            }
            "Missing" {
                $cell.Style.ForeColor = [System.Drawing.Color]::Red
                $cell.Style.Font = New-Object System.Drawing.Font($dgvCsv.Font, [System.Drawing.FontStyle]::Bold)
            }
            default {
                $cell.Style.ForeColor = [System.Drawing.Color]::DarkOrange
                $cell.Style.Font = New-Object System.Drawing.Font($dgvCsv.Font, [System.Drawing.FontStyle]::Regular)
            }
        }

        # If now verified, auto-remove from awaiting verification list
        if ($status -like "Verified*") {
            $before = $script:awaitingVerification.Count
            $script:awaitingVerification = @($script:awaitingVerification | Where-Object { $_.SerialNumber.Trim() -ne $serial })
            if ($script:awaitingVerification.Count -lt $before) {
                Save-AwaitingVerification
            }
        }
    }

    # Refresh awaiting grid and needs-verification grid after run
    Update-AwaitingVerificationList
    Update-NeedsVerificationList

    $dgvCsv.ResumeLayout()
    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    [System.Windows.Forms.MessageBox]::Show(
        "Verification complete!`n`nVerified: $verifiedCount`nNeeds Verification: $needsCount`nMissing: $($missingSerials.Count)`n`nVerified = Active in AD (logged in within 7 days) or in local stock.`nNeeds Verification = Not recently active in AD.`n`nSee the 'Needs Verification' tab for the full list.",
        "Verification Results", "OK", "Information")
})

# Add selected CSV item(s) to inventory
$btnAddFromCsv.Add_Click({
    if ($dgvCsv.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one item from the CSV to add.", "No Selection", "OK", "Warning")
        return
    }
    
    # Create dialog for location and loaner settings
    $addDialog = New-Object System.Windows.Forms.Form
    $addDialog.Text = "Add to Inventory"
    $addDialog.Size = New-Object System.Drawing.Size(400, 200)
    $addDialog.StartPosition = "CenterParent"
    $addDialog.FormBorderStyle = "FixedDialog"
    $addDialog.MaximizeBox = $false
    $addDialog.MinimizeBox = $false
    
    $lblDialogLocation = New-Object System.Windows.Forms.Label
    $lblDialogLocation.Text = "Location:"
    $lblDialogLocation.Location = New-Object System.Drawing.Point(20, 20)
    $lblDialogLocation.Size = New-Object System.Drawing.Size(80, 20)
    $addDialog.Controls.Add($lblDialogLocation)
    
    $cmbDialogLocation = New-Object System.Windows.Forms.ComboBox
    $cmbDialogLocation.Location = New-Object System.Drawing.Point(110, 18)
    $cmbDialogLocation.Size = New-Object System.Drawing.Size(250, 20)
    $cmbDialogLocation.DropDownStyle = "DropDownList"
    $cmbDialogLocation.Items.AddRange(@("Denver", "Greenwood Village", "Salt Lake City"))
    $cmbDialogLocation.SelectedIndex = 0
    $addDialog.Controls.Add($cmbDialogLocation)
    
    $chkDialogLoaner = New-Object System.Windows.Forms.CheckBox
    $chkDialogLoaner.Text = "Mark as Loaner Laptop"
    $chkDialogLoaner.Location = New-Object System.Drawing.Point(20, 60)
    $chkDialogLoaner.Size = New-Object System.Drawing.Size(200, 25)
    $addDialog.Controls.Add($chkDialogLoaner)
    
    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "Adding $($dgvCsv.SelectedRows.Count) item(s) to inventory"
    $lblCount.Location = New-Object System.Drawing.Point(20, 100)
    $lblCount.Size = New-Object System.Drawing.Size(350, 20)
    $addDialog.Controls.Add($lblCount)
    
    $btnDialogOk = New-Object System.Windows.Forms.Button
    $btnDialogOk.Text = "Add"
    $btnDialogOk.Location = New-Object System.Drawing.Point(150, 130)
    $btnDialogOk.Size = New-Object System.Drawing.Size(80, 30)
    $btnDialogOk.DialogResult = "OK"
    $addDialog.Controls.Add($btnDialogOk)
    
    $btnDialogCancel = New-Object System.Windows.Forms.Button
    $btnDialogCancel.Text = "Cancel"
    $btnDialogCancel.Location = New-Object System.Drawing.Point(240, 130)
    $btnDialogCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnDialogCancel.DialogResult = "Cancel"
    $addDialog.Controls.Add($btnDialogCancel)
    
    $addDialog.AcceptButton = $btnDialogOk
    $addDialog.CancelButton = $btnDialogCancel
    
    $dialogResult = $addDialog.ShowDialog()
    
    if ($dialogResult -eq "OK") {
        # Reload inventory to ensure we have the latest data
        $script:inventory = @(Load-Inventory)
        
        $location = $cmbDialogLocation.SelectedItem
        $isLoaner = $chkDialogLoaner.Checked
        $addedCount = 0
        $skippedCount = 0
        $skippedItems = @()
        
        foreach ($selectedRow in $dgvCsv.SelectedRows) {
            try {
                # Get the row index to find corresponding CSV data
                $rowIndex = $selectedRow.Index
                $csvItem = $script:csvData[$rowIndex]
                
                $serial = $csvItem.serial_number
                $model = $csvItem.model
                
                # Debug output
                Write-Host "Checking serial: '$serial'"
                Write-Host "Inventory is null: $($null -eq $script:inventory)"
                Write-Host "Inventory type: $($script:inventory.GetType().Name)"
                Write-Host "Current inventory count: $($script:inventory.Count)"
                if ($script:inventory.Count -gt 0) {
                    Write-Host "Inventory serials: $($script:inventory.SerialNumber -join ', ')"
                }
                
                # Check if already in inventory (trim and case-insensitive)
                $existing = $null
                if ($script:inventory.Count -gt 0) {
                    $existing = $script:inventory | Where-Object { $_.SerialNumber.Trim() -eq $serial.Trim() }
                }
                
                if ($existing) {
                    Write-Host "Found existing: $($existing.SerialNumber)"
                    $skippedCount++
                    $skippedItems += $serial
                    continue
                }
                
                Write-Host "Adding to inventory: $serial"
                
                # Add to inventory
                $newItem = [PSCustomObject]@{
                    SerialNumber = $serial
                    Model = $model
                    Location = $location
                    IsLoaner = $isLoaner
                    DateAdded = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                }
                
                $script:inventory = @($script:inventory) + @($newItem)
                Write-Host "Added successfully. New count: $($script:inventory.Count)"
                $addedCount++
            }
            catch {
                Write-Host "Error adding item: $($_.Exception.Message)"
                $skippedCount++
            }
        }
        
        Write-Host "Save-Inventory about to be called. Inventory count: $($script:inventory.Count)"
        Save-Inventory
        Write-Host "Save-Inventory completed"
        Update-InventoryGrid
        
        $message = "Successfully added $addedCount item(s) to inventory."
        if ($skippedCount -gt 0) {
            $message += "`n`nSkipped $skippedCount item(s) (already in inventory):"
            $skippedItems | ForEach-Object { $message += "`n- $_" }
        }
        
        [System.Windows.Forms.MessageBox]::Show($message, "Add Complete", "OK", "Information")
    }
})

# Compare CSV with Inventory
$btnCompare.Add_Click({
    try {
        $script:inventory = @(Load-Inventory)
        $inCount  = 0
        $outCount = 0

        for ($i = 0; $i -lt $script:csvData.Count; $i++) {
            $serial = $script:csvData[$i].serial_number
            $inInventory = $script:inventory | Where-Object { $_.SerialNumber -eq $serial }
            $row = $dgvCsv.Rows[$i]

            if ($inInventory) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
                $inCount++
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
                $outCount++
            }
        }

        [System.Windows.Forms.MessageBox]::Show("Comparison complete!`n`nIn our inventory: $inCount`nNot in our inventory: $outCount", "Comparison Results", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error comparing data: $($_.Exception.Message)", "Error", "OK", "Error")
    }
})

# Add laptop to inventory
$btnAddLaptop.Add_Click({
    $serial = $txtSerial.Text.Trim()
    $model = $txtModel.Text.Trim()
    $location = $cmbLocation.SelectedItem
    $isLoaner = $chkLoaner.Checked
    
    if ([string]::IsNullOrWhiteSpace($serial)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a serial number.", "Validation Error", "OK", "Warning")
        return
    }
    
    # Reload inventory to ensure we have the latest data
    $script:inventory = @(Load-Inventory)
    
    # Check if serial already exists
    $existing = $script:inventory | Where-Object { $_.SerialNumber -eq $serial }
    if ($existing) {
        [System.Windows.Forms.MessageBox]::Show("A laptop with serial number '$serial' already exists in inventory.", "Duplicate Entry", "OK", "Warning")
        return
    }
    
    # Add to inventory
    $newItem = [PSCustomObject]@{
        SerialNumber = $serial
        Model = $model
        Location = $location
        IsLoaner = $isLoaner
        DateAdded = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }
    
    $script:inventory = @($script:inventory) + @($newItem)
    Save-Inventory
    
    # Clear form
    $txtSerial.Text = ""
    $txtModel.Text = ""
    $cmbLocation.SelectedIndex = 0
    $chkLoaner.Checked = $false
    
    # Refresh display
    Update-InventoryGrid
    Update-LoanerList
    
    [System.Windows.Forms.MessageBox]::Show("Laptop added to inventory successfully!", "Success", "OK", "Information")
})

# Remove laptop from inventory
$btnRemoveLaptop.Add_Click({
    if ($dgvInventory.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a laptop to remove.", "No Selection", "OK", "Warning")
        return
    }
    
    $selectedSerial = $dgvInventory.SelectedRows[0].Cells["SerialNumber"].Value
    
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove laptop with serial number '$selectedSerial'?", "Confirm Removal", "YesNo", "Question")
    
    if ($result -eq "Yes") {
        $script:inventory = @($script:inventory | Where-Object { $_.SerialNumber -ne $selectedSerial })
        Save-Inventory
        Update-InventoryGrid
        [System.Windows.Forms.MessageBox]::Show("Laptop removed from inventory.", "Success", "OK", "Information")
    }
})

# Refresh inventory display
$btnRefreshInventory.Add_Click({
    $txtSearchInventory.Text = ""
    Update-InventoryGrid
})

# Search inventory as user types
$txtSearchInventory.Add_TextChanged({
    $searchText = $txtSearchInventory.Text.Trim()
    Update-InventoryGrid -SearchText $searchText
})

# Clear search button
$btnClearSearch.Add_Click({
    $txtSearchInventory.Text = ""
    Update-InventoryGrid
})

# Show Details for selected laptop from CSV
$btnShowDetailsCsv.Add_Click({
    if ($dgvCsv.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a computer to view details.", "No Selection", "OK", "Warning")
        return
    }
    
    $rowIndex = $dgvCsv.SelectedRows[0].Index
    $csvItem = $script:csvData[$rowIndex]
    $selectedSerial = $csvItem.serial_number
    $selectedItem = $csvItem
    
    # Create details form
    $detailsForm = New-Object System.Windows.Forms.Form
    $detailsForm.Text = "Computer Details - $selectedSerial"
    $detailsForm.Size = New-Object System.Drawing.Size(500, 400)
    $detailsForm.StartPosition = "CenterParent"
    $detailsForm.FormBorderStyle = "FixedDialog"
    $detailsForm.MaximizeBox = $false
    $detailsForm.MinimizeBox = $false
    
    # Details text box
    $txtDetails = New-Object System.Windows.Forms.TextBox
    $txtDetails.Location = New-Object System.Drawing.Point(10, 10)
    $txtDetails.Size = New-Object System.Drawing.Size(465, 300)
    $txtDetails.Multiline = $true
    $txtDetails.ScrollBars = "Vertical"
    $txtDetails.ReadOnly = $true
    $txtDetails.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtDetails.Text = "Gathering information, please wait..."
    $detailsForm.Controls.Add($txtDetails)
    
    # Close button
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(200, 320)
    $btnClose.Size = New-Object System.Drawing.Size(80, 30)
    $btnClose.DialogResult = "OK"
    $detailsForm.Controls.Add($btnClose)
    $detailsForm.AcceptButton = $btnClose
    
    # Show form and gather info in background
    $detailsForm.Add_Shown({
        $details = @()
        $details += "=" * 60
        $details += "COMPUTER DETAILS"
        $details += "=" * 60
        $details += ""
        $details += "Serial Number: $selectedSerial"
        $details += "Model: $($selectedItem.model)"
        $details += "CSV Status: $($selectedItem.install_status)"
        $details += "CSV Location: $($selectedItem.stockroom)"
        $details += "Assigned To: $(if ($selectedItem.assigned_to) { $selectedItem.assigned_to } else { 'Unassigned' })"
        $details += "Asset Function: $(if ($selectedItem.asset_function) { $selectedItem.asset_function } else { 'N/A' })"
        $details += "Substatus: $(if ($selectedItem.substatus) { $selectedItem.substatus } else { 'N/A' })"
        $details += ""
        $details += "-" * 60
        $details += "NETWORK STATUS"
        $details += "-" * 60
        
        # Try to ping the computer
        try {
            $details += ""
            $details += "Testing connectivity to $selectedSerial..."
            $pingResult = Test-Connection -ComputerName $selectedSerial -Count 2 -Quiet -ErrorAction SilentlyContinue
            if ($pingResult) {
                $details += "Status: ONLINE (Ping successful)"
                $pingDetails = Test-Connection -ComputerName $selectedSerial -Count 1 -ErrorAction SilentlyContinue
                if ($pingDetails) {
                    $details += "IP Address: $($pingDetails.IPV4Address)"
                    $details += "Response Time: $($pingDetails.ResponseTime)ms"
                }
            } else {
                $details += "Status: OFFLINE (No response to ping)"
            }
        }
        catch {
            $details += "Status: ERROR - Unable to ping ($($_.Exception.Message))"
        }
        
        $details += ""
        $details += "-" * 60
        $details += "ACTIVE DIRECTORY INFORMATION"
        $details += "-" * 60
        
        # Try to get AD computer information
        try {
            $details += ""
            $details += "Querying Active Directory..."
            $adComputer = Get-ADComputer -Identity $selectedSerial -Properties LastLogonDate, Description, OperatingSystem, OperatingSystemVersion -ErrorAction Stop
            
            $details += "Computer Name: $($adComputer.Name)"
            $details += "Distinguished Name: $($adComputer.DistinguishedName)"
            $details += "Operating System: $($adComputer.OperatingSystem)"
            $details += "OS Version: $($adComputer.OperatingSystemVersion)"
            $details += "Last Logon: $(if ($adComputer.LastLogonDate) { $adComputer.LastLogonDate } else { 'Never' })"
            $details += "Description: $(if ($adComputer.Description) { $adComputer.Description } else { 'None' })"
            $details += "Enabled: $($adComputer.Enabled)"
            
            # Try to get logged-on user
            $details += ""
            $details += "Checking for logged-on user..."
            try {
                if ($pingResult) {
                    $loggedOnUser = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $selectedSerial -ErrorAction Stop
                    if ($loggedOnUser.UserName) {
                        $details += "Current User: $($loggedOnUser.UserName)"
                        
                        # Try to get user details from AD
                        try {
                            $userName = $loggedOnUser.UserName.Split('\')[1]
                            $adUser = Get-ADUser -Identity $userName -Properties DisplayName, EmailAddress, Title, Department, LastLogonDate, Manager -ErrorAction Stop
                            $details += "Display Name: $($adUser.DisplayName)"
                            $details += "Email: $(if ($adUser.EmailAddress) { $adUser.EmailAddress } else { 'N/A' })"
                            $details += "Title: $(if ($adUser.Title) { $adUser.Title } else { 'N/A' })"
                            $details += "Department: $(if ($adUser.Department) { $adUser.Department } else { 'N/A' })"
                            $details += "User Last Logon: $(if ($adUser.LastLogonDate) { $adUser.LastLogonDate } else { 'Unknown' })"
                            
                            # Get manager information
                            if ($adUser.Manager) {
                                try {
                                    $manager = Get-ADUser -Identity $adUser.Manager -Properties DisplayName, EmailAddress -ErrorAction Stop
                                    $details += ""
                                    $details += "Manager Information:"
                                    $details += "  Manager Name: $($manager.DisplayName)"
                                    $details += "  Manager Email: $(if ($manager.EmailAddress) { $manager.EmailAddress } else { 'N/A' })"
                                }
                                catch {
                                    $details += ""
                                    $details += "Manager: Unable to retrieve manager details"
                                }
                            } else {
                                $details += ""
                                $details += "Manager: No manager assigned"
                            }
                        }
                        catch {
                            $details += "Unable to retrieve user details: $($_.Exception.Message)"
                        }
                    } else {
                        $details += "Current User: No user logged on"
                    }
                } else {
                    $details += "Cannot check logged-on user (computer offline)"
                }
            }
            catch {
                $details += "Unable to query logged-on user: $($_.Exception.Message)"
            }
        }
        catch {
            $details += "Unable to retrieve AD information: $($_.Exception.Message)"
            $details += ""
            $details += "Note: Active Directory PowerShell module may not be installed,"
            $details += "or you may not have permission to query AD."
        }
        
        $details += ""
        $details += "=" * 60
        
        $txtDetails.Text = $details -join "`r`n"
    })
    
    [void]$detailsForm.ShowDialog()
})

# Email User button (CSV tab)
$btnEmailCsvUser.Add_Click({
    if ($dgvCsv.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select an asset from the CSV list.", "No Selection", "OK", "Warning")
        return
    }

    $rowIndex = $dgvCsv.SelectedRows[0].Index
    $csvItem = $script:csvData[$rowIndex]

    $selectedSerial = $csvItem.serial_number
    $model = if ($csvItem.model) { $csvItem.model } else { "Unknown Model" }
    $assetTag = if ($csvItem.PSObject.Properties['asset_tag'] -and $csvItem.asset_tag) { $csvItem.asset_tag } else { "N/A" }
    $assignedTo = if ($csvItem.PSObject.Properties['assigned_to']) { $csvItem.assigned_to } else { "" }

    if ([string]::IsNullOrWhiteSpace($assignedTo)) {
        [System.Windows.Forms.MessageBox]::Show("This asset does not have an assigned user in the CSV data.", "No User Assigned", "OK", "Warning")
        return
    }

    # Look up user in Active Directory by name to get email and manager
    $userEmail = ""
    $managerEmail = ""
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $adUser = $null
    $safeName = $assignedTo -replace "'", "''"

    # Try DisplayName exact match
    try {
        $r = @(Get-ADUser -Filter "DisplayName -eq '$safeName'" -Properties EmailAddress, Manager -ErrorAction Stop)
        if ($r.Count -gt 0) { $adUser = $r[0] }
    } catch { }

    # Try Name (CN) exact match
    if (-not $adUser) {
        try {
            $r = @(Get-ADUser -Filter "Name -eq '$safeName'" -Properties EmailAddress, Manager -ErrorAction Stop)
            if ($r.Count -gt 0) { $adUser = $r[0] }
        } catch { }
    }

    # Try GivenName + Surname
    if (-not $adUser -and $assignedTo -match ' ') {
        try {
            $nameParts = $assignedTo -split '\s+', 2
            $firstName = $nameParts[0].Trim() -replace "'", "''"
            $lastName  = $nameParts[1].Trim() -replace "'", "''"
            $r = @(Get-ADUser -Filter "GivenName -eq '$firstName' -and Surname -eq '$lastName'" -Properties EmailAddress, Manager -ErrorAction Stop)
            if ($r.Count -gt 0) { $adUser = $r[0] }
        } catch { }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    if (-not $adUser) {
        [System.Windows.Forms.MessageBox]::Show("Could not find user '$assignedTo' in Active Directory.", "AD Lookup Error", "OK", "Error")
        return
    }

    $userEmail = if ($adUser.EmailAddress) { $adUser.EmailAddress } else { "" }
    if ($adUser.Manager) {
        try {
            $mgr = Get-ADUser -Identity $adUser.Manager -Properties EmailAddress -ErrorAction Stop
            $managerEmail = if ($mgr.EmailAddress) { $mgr.EmailAddress } else { "" }
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($userEmail)) {
        [System.Windows.Forms.MessageBox]::Show("User '$assignedTo' was found in Active Directory but has no email address.", "No Email Address", "OK", "Warning")
        return
    }

    # Dialog to capture sender email
    $emailDialog = New-Object System.Windows.Forms.Form
    $emailDialog.Text = "Send Laptop Verification Email"
    $emailDialog.Size = New-Object System.Drawing.Size(500, 220)
    $emailDialog.StartPosition = "CenterParent"
    $emailDialog.FormBorderStyle = "FixedDialog"
    $emailDialog.MaximizeBox = $false
    $emailDialog.MinimizeBox = $false

    $lblFrom = New-Object System.Windows.Forms.Label
    $lblFrom.Text = "From (your email):"
    $lblFrom.Location = New-Object System.Drawing.Point(20, 20)
    $lblFrom.Size = New-Object System.Drawing.Size(120, 20)
    $emailDialog.Controls.Add($lblFrom)

    $txtFrom = New-Object System.Windows.Forms.TextBox
    $txtFrom.Location = New-Object System.Drawing.Point(150, 18)
    $txtFrom.Size = New-Object System.Drawing.Size(320, 20)
    $emailDialog.Controls.Add($txtFrom)

    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = "To (user):"
    $lblTo.Location = New-Object System.Drawing.Point(20, 55)
    $lblTo.Size = New-Object System.Drawing.Size(120, 20)
    $emailDialog.Controls.Add($lblTo)

    $txtTo = New-Object System.Windows.Forms.TextBox
    $txtTo.Location = New-Object System.Drawing.Point(150, 53)
    $txtTo.Size = New-Object System.Drawing.Size(320, 20)
    $txtTo.ReadOnly = $true
    $txtTo.Text = $userEmail
    $emailDialog.Controls.Add($txtTo)

    $lblCc = New-Object System.Windows.Forms.Label
    $lblCc.Text = "CC (manager + you):"
    $lblCc.Location = New-Object System.Drawing.Point(20, 90)
    $lblCc.Size = New-Object System.Drawing.Size(120, 20)
    $emailDialog.Controls.Add($lblCc)

    $txtCcDisplay = New-Object System.Windows.Forms.TextBox
    $txtCcDisplay.Location = New-Object System.Drawing.Point(150, 88)
    $txtCcDisplay.Size = New-Object System.Drawing.Size(320, 20)
    $txtCcDisplay.ReadOnly = $true
    $txtCcDisplay.Text = if ($managerEmail) { "$managerEmail + you" } else { "you" }
    $emailDialog.Controls.Add($txtCcDisplay)

    $btnSend = New-Object System.Windows.Forms.Button
    $btnSend.Text = "Send"
    $btnSend.Location = New-Object System.Drawing.Point(260, 135)
    $btnSend.Size = New-Object System.Drawing.Size(80, 30)
    $btnSend.DialogResult = "OK"
    $emailDialog.Controls.Add($btnSend)

    $btnCancelEmail = New-Object System.Windows.Forms.Button
    $btnCancelEmail.Text = "Cancel"
    $btnCancelEmail.Location = New-Object System.Drawing.Point(350, 135)
    $btnCancelEmail.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancelEmail.DialogResult = "Cancel"
    $emailDialog.Controls.Add($btnCancelEmail)

    $emailDialog.AcceptButton = $btnSend
    $emailDialog.CancelButton = $btnCancelEmail

    $dialogResult = $emailDialog.ShowDialog()

    if ($dialogResult -ne "OK") {
        return
    }

    $fromEmail = $txtFrom.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($fromEmail) -or (-not $fromEmail.Contains("@"))) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid sender email address.", "Invalid Email", "OK", "Warning")
        return
    }

    $ccList = @()
    if (-not [string]::IsNullOrWhiteSpace($managerEmail)) {
        $ccList += $managerEmail
    }
    $ccList += $fromEmail

    try {
        Send-Email -SendTo $userEmail -AssetTag $assetTag -SerialNumber $selectedSerial -Model $model -EmailAdd $fromEmail -Cc $ccList

        # Add to Awaiting Verification list (replace any existing entry for same serial)
        $script:awaitingVerification = @($script:awaitingVerification | Where-Object { $_.SerialNumber.Trim() -ne $selectedSerial.Trim() })
        $script:awaitingVerification += [PSCustomObject]@{
            SerialNumber  = $selectedSerial
            Model         = $model
            AssetTag      = $assetTag
            AssignedTo    = $assignedTo
            EmailSentDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            EmailSentBy   = $fromEmail
            Source        = "CSV"
        }
        Save-AwaitingVerification
        Update-AwaitingVerificationList

        # Update the Verified cell in the CSV grid to reflect email sent
        $dt = $dgvCsv.DataSource
        if ($dt) {
            for ($i = 0; $i -lt $script:csvData.Count; $i++) {
                if ($script:csvData[$i].serial_number.Trim() -eq $selectedSerial.Trim()) {
                    $dt.Rows[$i]["Verified"] = "Email Sent"
                    $dgvCsv.Rows[$i].Cells["Verified"].Style.ForeColor = [System.Drawing.Color]::DarkBlue
                    $dgvCsv.Rows[$i].Cells["Verified"].Style.Font = New-Object System.Drawing.Font($dgvCsv.Font, [System.Drawing.FontStyle]::Bold)
                    break
                }
            }
        }

        [System.Windows.Forms.MessageBox]::Show("Email sent successfully to $userEmail.`n`nThis laptop has been added to the 'Awaiting Verification' list.", "Email Sent", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to send email: $($_.Exception.Message)", "Email Error", "OK", "Error")
    }
})

# Check Device button (Device Lookup tab)
$btnCheckDevice.Add_Click({
    $serial = $txtLookupSerial.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($serial)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a serial number.", "Validation Error", "OK", "Warning")
        return
    }
    
    $txtLookupResults.Text = "Gathering information for $serial, please wait..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $form.Refresh()
    
    $details = @()
    $details += "=" * 80
    $details += "DEVICE LOOKUP - $serial"
    $details += "=" * 80
    $details += ""
    $details += "Checked: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $details += ""
    
    # Check Inventory
    $details += "-" * 80
    $details += "INVENTORY STATUS"
    $details += "-" * 80
    $details += ""
    
    # Reload inventory to ensure we have the latest data
    $currentInventory = @(Load-Inventory)
    $inventoryItem = $currentInventory | Where-Object { $_.SerialNumber.Trim() -eq $serial.Trim() }
    
    if ($inventoryItem) {
        $details += "Status: FOUND IN INVENTORY"
        $details += ""
        $details += "Inventory Information:"
        $details += "  Serial Number: $($inventoryItem.SerialNumber)"
        $details += "  Model: $(if ($inventoryItem.Model) { $inventoryItem.Model } else { 'N/A' })"
        $details += "  Location: $(if ($inventoryItem.Location) { $inventoryItem.Location } else { 'N/A' })"
        $details += "  Date Added: $(if ($inventoryItem.DateAdded) { $inventoryItem.DateAdded } else { 'N/A' })"
        if ($inventoryItem.Notes) {
            $details += "  Notes: $($inventoryItem.Notes)"
        }
        if ($inventoryItem.IsLoaner -eq $true) {
            $details += "  Type: LOANER LAPTOP"
        } else {
            $details += "  Type: Regular Inventory"
        }
    } else {
        $details += "Status: NOT FOUND IN INVENTORY"
        $details += ""
        $details += "This device is not currently tracked in the inventory system."
    }
    
    # Test ping
    $details += "-" * 80
    $details += "NETWORK CONNECTIVITY"
    $details += "-" * 80
    $details += ""
    
    $pingResult = $false
    try {
        $details += "Testing ping to $serial..."
        $pingTest = Test-Connection -ComputerName $serial -Count 2 -ErrorAction Stop
        if ($pingTest) {
            $pingResult = $true
            $details += "Status: CONNECTED"
            $details += "IP Address: $($pingTest[0].IPV4Address)"
            $details += "Response Time: $($pingTest[0].ResponseTime)ms"
        }
    }
    catch {
        $details += "Status: NOT CONNECTED"
        $details += "Error: $($_.Exception.Message)"
    }
    
    # Check Active Directory
    $details += ""
    $details += "-" * 80
    $details += "ACTIVE DIRECTORY STATUS"
    $details += "-" * 80
    $details += ""
    
    try {
        $details += "Querying Active Directory for $serial..."
        $adComputer = Get-ADComputer -Identity $serial -Properties LastLogonDate, Description, OperatingSystem, OperatingSystemVersion, Enabled -ErrorAction Stop
        
        $details += "Status: FOUND IN ACTIVE DIRECTORY"
        $details += ""
        $details += "Computer Information:"
        $details += "  Name: $($adComputer.Name)"
        $details += "  Distinguished Name: $($adComputer.DistinguishedName)"
        $details += "  Operating System: $($adComputer.OperatingSystem)"
        $details += "  OS Version: $($adComputer.OperatingSystemVersion)"
        $details += "  Enabled: $($adComputer.Enabled)"
        $details += "  Description: $(if ($adComputer.Description) { $adComputer.Description } else { 'None' })"
        $details += "  Computer Last Logon: $(if ($adComputer.LastLogonDate) { $adComputer.LastLogonDate } else { 'Never' })"

        # Last logged-on user from registry (requires machine to be online)
        if ($pingResult) {
            try {
                $logonUI = Invoke-Command -ComputerName $serial -ScriptBlock {
                    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" -ErrorAction Stop
                } -ErrorAction Stop
                $details += "  Last Logged On User: $(if ($logonUI.LastLoggedOnUser) { $logonUI.LastLoggedOnUser } else { 'Unknown' })"
                $details += "  Last Logged On Display Name: $(if ($logonUI.LastLoggedOnDisplayName) { $logonUI.LastLoggedOnDisplayName } else { 'Unknown' })"
            } catch { }
        }

        # Get logged-on user
        $details += ""
        $details += "-" * 80
        $details += "LOGGED-ON USER INFORMATION"
        $details += "-" * 80
        $details += ""

        if ($pingResult) {
            try {
                $loggedOnUser = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $serial -ErrorAction Stop

                if ($loggedOnUser.UserName) {
                    $details += "Current User: $($loggedOnUser.UserName)"
                    $details += ""
                    
                    # Get user details from AD
                    try {
                        $userName = $loggedOnUser.UserName.Split('\')[1]
                        $adUser = Get-ADUser -Identity $userName -Properties DisplayName, EmailAddress, Title, Department, LastLogonDate, Manager, SamAccountName -ErrorAction Stop
                        
                        $details += "User Details:"
                        $details += "  User ID: $($adUser.SamAccountName)"
                        $details += "  Display Name: $($adUser.DisplayName)"
                        $details += "  Email: $(if ($adUser.EmailAddress) { $adUser.EmailAddress } else { 'N/A' })"
                        $details += "  Title: $(if ($adUser.Title) { $adUser.Title } else { 'N/A' })"
                        $details += "  Department: $(if ($adUser.Department) { $adUser.Department } else { 'N/A' })"
                        $details += "  User Last Logon: $(if ($adUser.LastLogonDate) { $adUser.LastLogonDate } else { 'Unknown' })"
                        
                        # Get manager information
                        if ($adUser.Manager) {
                            $details += ""
                            $details += "Manager Information:"
                            try {
                                $manager = Get-ADUser -Identity $adUser.Manager -Properties DisplayName, EmailAddress, SamAccountName -ErrorAction Stop
                                $details += "  Manager ID: $($manager.SamAccountName)"
                                $details += "  Manager Name: $($manager.DisplayName)"
                                $details += "  Manager Email: $(if ($manager.EmailAddress) { $manager.EmailAddress } else { 'N/A' })"
                            }
                            catch {
                                $details += "  Unable to retrieve manager details: $($_.Exception.Message)"
                            }
                        } else {
                            $details += ""
                            $details += "Manager Information: No manager assigned in Active Directory"
                        }
                    }
                    catch {
                        $details += "Unable to retrieve user details from Active Directory:"
                        $details += "  Error: $($_.Exception.Message)"
                    }
                } else {
                    $details += "Status: No user currently logged on to this device"
                }
            }
            catch {
                $details += "Unable to query logged-on user:"
                $details += "  Error: $($_.Exception.Message)"
            }
        } else {
            $details += "Cannot query logged-on user - device is not connected to the network"
        }
    }
    catch {
        $details += "Status: NOT FOUND IN ACTIVE DIRECTORY"
        $details += "Error: $($_.Exception.Message)"
        $details += ""
        $details += "Note: The device may not exist in Active Directory, or you may not have"
        $details += "permission to query AD. Ensure the Active Directory PowerShell module is installed."
    }
    
    $details += ""
    $details += "=" * 80
    $details += "END OF REPORT"
    $details += "=" * 80
    
    $txtLookupResults.Text = $details -join "`r`n"
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# Clear Lookup button
$btnClearLookup.Add_Click({
    $txtLookupSerial.Text = ""
    $txtLookupResults.Text = "Enter a serial number and click 'Check Device' to view information."
})

# Add Loaner button
$btnAddLoaner.Add_Click({
    # Create dialog for adding loaner
    $addLoanerDialog = New-Object System.Windows.Forms.Form
    $addLoanerDialog.Text = "Add Loaner Laptop"
    $addLoanerDialog.Size = New-Object System.Drawing.Size(400, 200)
    $addLoanerDialog.StartPosition = "CenterParent"
    $addLoanerDialog.FormBorderStyle = "FixedDialog"
    $addLoanerDialog.MaximizeBox = $false
    $addLoanerDialog.MinimizeBox = $false
    
    # Serial Number
    $lblDialogSerial = New-Object System.Windows.Forms.Label
    $lblDialogSerial.Text = "Serial Number:"
    $lblDialogSerial.Location = New-Object System.Drawing.Point(20, 20)
    $lblDialogSerial.Size = New-Object System.Drawing.Size(100, 20)
    $addLoanerDialog.Controls.Add($lblDialogSerial)
    
    $txtDialogSerial = New-Object System.Windows.Forms.TextBox
    $txtDialogSerial.Location = New-Object System.Drawing.Point(130, 18)
    $txtDialogSerial.Size = New-Object System.Drawing.Size(230, 20)
    $addLoanerDialog.Controls.Add($txtDialogSerial)
    
    # Location
    $lblDialogLoanerLocation = New-Object System.Windows.Forms.Label
    $lblDialogLoanerLocation.Text = "Location:"
    $lblDialogLoanerLocation.Location = New-Object System.Drawing.Point(20, 60)
    $lblDialogLoanerLocation.Size = New-Object System.Drawing.Size(100, 20)
    $addLoanerDialog.Controls.Add($lblDialogLoanerLocation)
    
    $cmbDialogLoanerLocation = New-Object System.Windows.Forms.ComboBox
    $cmbDialogLoanerLocation.Location = New-Object System.Drawing.Point(130, 58)
    $cmbDialogLoanerLocation.Size = New-Object System.Drawing.Size(230, 20)
    $cmbDialogLoanerLocation.DropDownStyle = "DropDownList"
    $cmbDialogLoanerLocation.Items.AddRange(@("Denver", "Greenwood Village", "Salt Lake City"))
    $cmbDialogLoanerLocation.SelectedIndex = 0
    $addLoanerDialog.Controls.Add($cmbDialogLoanerLocation)
    
    # Buttons
    $btnDialogAddLoaner = New-Object System.Windows.Forms.Button
    $btnDialogAddLoaner.Text = "Add"
    $btnDialogAddLoaner.Location = New-Object System.Drawing.Point(150, 110)
    $btnDialogAddLoaner.Size = New-Object System.Drawing.Size(80, 30)
    $btnDialogAddLoaner.DialogResult = "OK"
    $addLoanerDialog.Controls.Add($btnDialogAddLoaner)
    
    $btnDialogCancelLoaner = New-Object System.Windows.Forms.Button
    $btnDialogCancelLoaner.Text = "Cancel"
    $btnDialogCancelLoaner.Location = New-Object System.Drawing.Point(240, 110)
    $btnDialogCancelLoaner.Size = New-Object System.Drawing.Size(80, 30)
    $btnDialogCancelLoaner.DialogResult = "Cancel"
    $addLoanerDialog.Controls.Add($btnDialogCancelLoaner)
    
    $addLoanerDialog.AcceptButton = $btnDialogAddLoaner
    $addLoanerDialog.CancelButton = $btnDialogCancelLoaner
    
    $dialogResult = $addLoanerDialog.ShowDialog()
    
    if ($dialogResult -eq "OK") {
        $serial = $txtDialogSerial.Text.Trim()
        $location = $cmbDialogLoanerLocation.SelectedItem
        
        if ([string]::IsNullOrWhiteSpace($serial)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a serial number.", "Validation Error", "OK", "Warning")
            return
        }
        
        # Reload inventory to ensure we have the latest data
        $script:inventory = @(Load-Inventory)
        
        # Check if serial already exists
        $existing = $script:inventory | Where-Object { $_.SerialNumber.Trim() -eq $serial.Trim() }
        if ($existing) {
            [System.Windows.Forms.MessageBox]::Show("A laptop with serial number '$serial' already exists in inventory.", "Duplicate Entry", "OK", "Warning")
            return
        }
        
        # Add to inventory as loaner
        $newItem = [PSCustomObject]@{
            SerialNumber = $serial
            Model = ""
            Location = $location
            IsLoaner = $true
            DateAdded = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
        
        $script:inventory = @($script:inventory) + @($newItem)
        Save-Inventory
        
        # Refresh displays
        Update-LoanerList
        Update-InventoryGrid
        
        [System.Windows.Forms.MessageBox]::Show("Loaner laptop added successfully!", "Success", "OK", "Information")
    }
})

# Remove Loaner button
$btnRemoveLoaner.Add_Click({
    if ($dgvLoaners.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a loaner laptop to remove.", "No Selection", "OK", "Warning")
        return
    }
    
    $selectedSerial = $dgvLoaners.SelectedRows[0].Cells["SerialNumber"].Value
    
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove loaner laptop with serial number '$selectedSerial'?", "Confirm Removal", "YesNo", "Question")
    
    if ($result -eq "Yes") {
        $script:inventory = @($script:inventory | Where-Object { $_.SerialNumber -ne $selectedSerial })
        Save-Inventory
        Update-LoanerList
        Update-InventoryGrid
        [System.Windows.Forms.MessageBox]::Show("Loaner laptop removed from inventory.", "Success", "OK", "Information")
    }
})

# Refresh Loaners button
$btnRefreshLoaners.Add_Click({
    Update-LoanerList
})

# Assign Loaner button
$btnAssignLoaner.Add_Click({
    if ($dgvLoaners.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a loaner laptop to assign.", "No Selection", "OK", "Warning")
        return
    }
    
    $selectedSerial = $dgvLoaners.SelectedRows[0].Cells["SerialNumber"].Value
    
    # Create dialog for user ID
    $assignDialog = New-Object System.Windows.Forms.Form
    $assignDialog.Text = "Assign Loaner Laptop"
    $assignDialog.Size = New-Object System.Drawing.Size(400, 150)
    $assignDialog.StartPosition = "CenterParent"
    $assignDialog.FormBorderStyle = "FixedDialog"
    $assignDialog.MaximizeBox = $false
    $assignDialog.MinimizeBox = $false
    
    $lblUserId = New-Object System.Windows.Forms.Label
    $lblUserId.Text = "User ID:"
    $lblUserId.Location = New-Object System.Drawing.Point(20, 20)
    $lblUserId.Size = New-Object System.Drawing.Size(80, 20)
    $assignDialog.Controls.Add($lblUserId)
    
    $txtUserId = New-Object System.Windows.Forms.TextBox
    $txtUserId.Location = New-Object System.Drawing.Point(110, 18)
    $txtUserId.Size = New-Object System.Drawing.Size(250, 20)
    $assignDialog.Controls.Add($txtUserId)
    
    $btnAssignOk = New-Object System.Windows.Forms.Button
    $btnAssignOk.Text = "Assign"
    $btnAssignOk.Location = New-Object System.Drawing.Point(150, 70)
    $btnAssignOk.Size = New-Object System.Drawing.Size(80, 30)
    $btnAssignOk.DialogResult = "OK"
    $assignDialog.Controls.Add($btnAssignOk)
    
    $btnAssignCancel = New-Object System.Windows.Forms.Button
    $btnAssignCancel.Text = "Cancel"
    $btnAssignCancel.Location = New-Object System.Drawing.Point(240, 70)
    $btnAssignCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnAssignCancel.DialogResult = "Cancel"
    $assignDialog.Controls.Add($btnAssignCancel)
    
    $assignDialog.AcceptButton = $btnAssignOk
    $assignDialog.CancelButton = $btnAssignCancel
    
    $dialogResult = $assignDialog.ShowDialog()
    
    if ($dialogResult -eq "OK") {
        $userId = $txtUserId.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($userId)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a user ID.", "Validation Error", "OK", "Warning")
            return
        }
        
        # Look up user in Active Directory
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            
            $adUser = Get-ADUser -Identity $userId -Properties DisplayName, EmailAddress, Manager -ErrorAction Stop
            
            $userName = $adUser.DisplayName
            $userEmail = if ($adUser.EmailAddress) { $adUser.EmailAddress } else { "N/A" }
            $managerName = "N/A"
            $managerEmail = "N/A"
            
            # Get manager information if exists
            if ($adUser.Manager) {
                try {
                    $manager = Get-ADUser -Identity $adUser.Manager -Properties DisplayName, EmailAddress -ErrorAction Stop
                    $managerName = $manager.DisplayName
                    $managerEmail = if ($manager.EmailAddress) { $manager.EmailAddress } else { "N/A" }
                }
                catch {
                    Write-Host "Could not retrieve manager details: $($_.Exception.Message)"
                }
            }
            
            # Update inventory
            $script:inventory = @(Load-Inventory)
            $loaner = $script:inventory | Where-Object { $_.SerialNumber -eq $selectedSerial }
            
            if ($loaner) {
                $loaner | Add-Member -MemberType NoteProperty -Name AssignedTo -Value $userId -Force
                $loaner | Add-Member -MemberType NoteProperty -Name AssignedName -Value $userName -Force
                $loaner | Add-Member -MemberType NoteProperty -Name AssignedEmail -Value $userEmail -Force
                $loaner | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerName -Force
                $loaner | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value $managerEmail -Force
                
                Save-Inventory
                Update-LoanerList
                
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
                
                $message = "Loaner laptop assigned successfully!`n`n"
                $message += "Assigned To: $userName ($userId)`n"
                $message += "Email: $userEmail`n"
                $message += "Manager: $managerName`n"
                $message += "Manager Email: $managerEmail"
                
                [System.Windows.Forms.MessageBox]::Show($message, "Assignment Complete", "OK", "Information")
            }
        }
        catch {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("Error looking up user in Active Directory: $($_.Exception.Message)`n`nPlease verify the user ID is correct and you have access to Active Directory.", "AD Lookup Error", "OK", "Error")
        }
    }
})

# Unassign Loaner button
$btnUnassignLoaner.Add_Click({
    if ($dgvLoaners.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a loaner laptop to unassign.", "No Selection", "OK", "Warning")
        return
    }
    
    $selectedSerial = $dgvLoaners.SelectedRows[0].Cells["SerialNumber"].Value
    
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to unassign this loaner laptop?", "Confirm Unassignment", "YesNo", "Question")
    
    if ($result -eq "Yes") {
        $script:inventory = @(Load-Inventory)
        $loaner = $script:inventory | Where-Object { $_.SerialNumber -eq $selectedSerial }
        
        if ($loaner) {
            # Remove assignment properties
            $loaner.PSObject.Properties.Remove('AssignedTo')
            $loaner.PSObject.Properties.Remove('AssignedName')
            $loaner.PSObject.Properties.Remove('AssignedEmail')
            $loaner.PSObject.Properties.Remove('ManagerName')
            $loaner.PSObject.Properties.Remove('ManagerEmail')
            
            Save-Inventory
            Update-LoanerList
            
            [System.Windows.Forms.MessageBox]::Show("Loaner laptop unassigned successfully!", "Unassignment Complete", "OK", "Information")
        }
    }
})

# Check Status button
$btnCheckStatus.Add_Click({
    $script:inventory = @(Load-Inventory)
    $loaners = $script:inventory | Where-Object { $_.IsLoaner -eq $true }
    
    if ($loaners.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No loaner laptops to check.", "No Loaners", "OK", "Information")
        return
    }
    
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $form.Refresh()
    
    # Create a hashtable to store ping results
    $script:loanerStatus = @{}
    
    $counter = 0
    foreach ($loaner in $loaners) {
        $counter++
        Write-Host "Checking $counter of $($loaners.Count): $($loaner.SerialNumber)"
        
        $status = "Offline"
        try {
            $pingTest = Test-Connection -ComputerName $loaner.SerialNumber -Count 1 -Quiet -ErrorAction Stop
            if ($pingTest -eq $true) {
                $status = "Online"
            }
        }
        catch {
            # Check if DNS error or just offline
            if ($_.Exception.Message -like "*could not be resolved*" -or $_.Exception.Message -like "*host not found*") {
                $status = "DNS Error"
            } else {
                $status = "Offline"
            }
        }
        
        $script:loanerStatus[$loaner.SerialNumber] = $status
    }
    
    # Refresh the list with status
    Update-LoanerList -IncludeStatus $true
    
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    
    $onlineCount = ($script:loanerStatus.Values | Where-Object { $_ -eq "Online" }).Count
    $offlineCount = ($script:loanerStatus.Values | Where-Object { $_ -eq "Offline" }).Count
    $dnsErrorCount = ($script:loanerStatus.Values | Where-Object { $_ -eq "DNS Error" }).Count
    
    [System.Windows.Forms.MessageBox]::Show("Status check complete!`n`nOnline: $onlineCount`nOffline: $offlineCount`nDNS Error: $dnsErrorCount", "Status Check Complete", "OK", "Information")
})

# Show Details for Loaner button
$btnShowDetailsLoaner.Add_Click({
    if ($dgvLoaners.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a loaner laptop to view details.", "No Selection", "OK", "Warning")
        return
    }
    
    $selectedSerial = $dgvLoaners.SelectedRows[0].Cells["SerialNumber"].Value
    
    # Create details form
    $detailsForm = New-Object System.Windows.Forms.Form
    $detailsForm.Text = "Loaner Laptop Details - $selectedSerial"
    $detailsForm.Size = New-Object System.Drawing.Size(600, 500)
    $detailsForm.StartPosition = "CenterParent"
    $detailsForm.FormBorderStyle = "FixedDialog"
    $detailsForm.MaximizeBox = $false
    $detailsForm.MinimizeBox = $false
    
    # Details text box
    $txtDetails = New-Object System.Windows.Forms.TextBox
    $txtDetails.Location = New-Object System.Drawing.Point(10, 10)
    $txtDetails.Size = New-Object System.Drawing.Size(565, 410)
    $txtDetails.Multiline = $true
    $txtDetails.ScrollBars = "Vertical"
    $txtDetails.ReadOnly = $true
    $txtDetails.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtDetails.Text = "Gathering information, please wait..."
    $detailsForm.Controls.Add($txtDetails)
    
    # Close button
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(250, 430)
    $btnClose.Size = New-Object System.Drawing.Size(80, 30)
    $btnClose.DialogResult = "OK"
    $detailsForm.Controls.Add($btnClose)
    $detailsForm.AcceptButton = $btnClose
    
    # Show form and gather info in background
    $detailsForm.Add_Shown({
        $details = @()
        $details += "=" * 80
        $details += "LOANER LAPTOP DETAILS"
        $details += "=" * 80
        $details += ""
        $details += "Serial Number: $selectedSerial"
        
        # Get loaner info from inventory
        $script:inventory = @(Load-Inventory)
        $loanerInfo = $script:inventory | Where-Object { $_.SerialNumber -eq $selectedSerial }
        
        if ($loanerInfo) {
            $details += "Model: $(if ($loanerInfo.Model) { $loanerInfo.Model } else { 'N/A' })"
            $details += "Location: $($loanerInfo.Location)"
            $details += "Date Added: $($loanerInfo.DateAdded)"
            
            if ($loanerInfo.PSObject.Properties['AssignedTo'] -and ![string]::IsNullOrWhiteSpace($loanerInfo.AssignedTo)) {
                $details += ""
                $details += "Assignment Information:"
                $details += "  Assigned To: $($loanerInfo.AssignedTo)"
                $details += "  User Name: $(if ($loanerInfo.AssignedName) { $loanerInfo.AssignedName } else { 'N/A' })"
                $details += "  User Email: $(if ($loanerInfo.AssignedEmail) { $loanerInfo.AssignedEmail } else { 'N/A' })"
                $details += "  Manager Name: $(if ($loanerInfo.ManagerName) { $loanerInfo.ManagerName } else { 'N/A' })"
                $details += "  Manager Email: $(if ($loanerInfo.ManagerEmail) { $loanerInfo.ManagerEmail } else { 'N/A' })"
            } else {
                $details += ""
                $details += "Assignment: Not currently assigned"
            }
        }
        
        $details += ""
        $details += "-" * 80
        $details += "NETWORK STATUS"
        $details += "-" * 80
        
        # Try to ping the computer
        $pingResult = $false
        try {
            $details += ""
            $details += "Testing connectivity to $selectedSerial..."
            $pingTest = Test-Connection -ComputerName $selectedSerial -Count 2 -ErrorAction Stop
            if ($pingTest) {
                $pingResult = $true
                $details += "Status: ONLINE (Ping successful)"
                $details += "IP Address: $($pingTest[0].IPV4Address)"
                $details += "Response Time: $($pingTest[0].ResponseTime)ms"
            }
        }
        catch {
            $details += "Status: OFFLINE (No response to ping)"
            $details += "Error: $($_.Exception.Message)"
        }
        
        $details += ""
        $details += "-" * 80
        $details += "ACTIVE DIRECTORY INFORMATION"
        $details += "-" * 80
        
        # Try to get AD computer information
        try {
            $details += ""
            $details += "Querying Active Directory for $selectedSerial..."
            $adComputer = Get-ADComputer -Identity $selectedSerial -Properties LastLogonDate, Description, OperatingSystem, OperatingSystemVersion, Enabled -ErrorAction Stop
            
            $details += "Status: FOUND IN ACTIVE DIRECTORY"
            $details += ""
            $details += "Computer Information:"
            $details += "  Name: $($adComputer.Name)"
            $details += "  Distinguished Name: $($adComputer.DistinguishedName)"
            $details += "  Operating System: $($adComputer.OperatingSystem)"
            $details += "  OS Version: $($adComputer.OperatingSystemVersion)"
            $details += "  Enabled: $($adComputer.Enabled)"
            $details += "  Description: $(if ($adComputer.Description) { $adComputer.Description } else { 'None' })"
            $details += "  Computer Last Logon: $(if ($adComputer.LastLogonDate) { $adComputer.LastLogonDate } else { 'Never' })"

            # Last logged-on user from registry (requires machine to be online)
            if ($pingResult) {
                try {
                    $logonUI = Invoke-Command -ComputerName $selectedSerial -ScriptBlock {
                        Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" -ErrorAction Stop
                    } -ErrorAction Stop
                    $details += "  Last Logged On User: $(if ($logonUI.LastLoggedOnUser) { $logonUI.LastLoggedOnUser } else { 'Unknown' })"
                    $details += "  Last Logged On Display Name: $(if ($logonUI.LastLoggedOnDisplayName) { $logonUI.LastLoggedOnDisplayName } else { 'Unknown' })"
                } catch { }
            }

            # Try to get logged-on user
            $details += ""
            $details += "-" * 80
            $details += "LOGGED-ON USER INFORMATION"
            $details += "-" * 80
            $details += ""

            if ($pingResult) {
                try {
                    $loggedOnUser = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $selectedSerial -ErrorAction Stop

                    if ($loggedOnUser.UserName) {
                        $details += "Current User: $($loggedOnUser.UserName)"
                        $details += ""
                        
                        # Get user details from AD
                        try {
                            $userName = $loggedOnUser.UserName.Split('\\')[1]
                            $adUser = Get-ADUser -Identity $userName -Properties DisplayName, EmailAddress, Title, Department, LastLogonDate, Manager, SamAccountName -ErrorAction Stop
                            
                            $details += "User Details:"
                            $details += "  User ID: $($adUser.SamAccountName)"
                            $details += "  Display Name: $($adUser.DisplayName)"
                            $details += "  Email: $(if ($adUser.EmailAddress) { $adUser.EmailAddress } else { 'N/A' })"
                            $details += "  Title: $(if ($adUser.Title) { $adUser.Title } else { 'N/A' })"
                            $details += "  Department: $(if ($adUser.Department) { $adUser.Department } else { 'N/A' })"
                            $details += "  User Last Logon: $(if ($adUser.LastLogonDate) { $adUser.LastLogonDate } else { 'Unknown' })"
                            
                            # Get manager information
                            if ($adUser.Manager) {
                                $details += ""
                                $details += "Manager Information:"
                                try {
                                    $manager = Get-ADUser -Identity $adUser.Manager -Properties DisplayName, EmailAddress, SamAccountName -ErrorAction Stop
                                    $details += "  Manager ID: $($manager.SamAccountName)"
                                    $details += "  Manager Name: $($manager.DisplayName)"
                                    $details += "  Manager Email: $(if ($manager.EmailAddress) { $manager.EmailAddress } else { 'N/A' })"
                                }
                                catch {
                                    $details += "  Unable to retrieve manager details: $($_.Exception.Message)"
                                }
                            } else {
                                $details += ""
                                $details += "Manager Information: No manager assigned in Active Directory"
                            }
                        }
                        catch {
                            $details += "Unable to retrieve user details from Active Directory:"
                            $details += "  Error: $($_.Exception.Message)"
                        }
                    } else {
                        $details += "Status: No user currently logged on to this device"
                    }
                }
                catch {
                    $details += "Unable to query logged-on user:"
                    $details += "  Error: $($_.Exception.Message)"
                }
            } else {
                $details += "Cannot query logged-on user - device is not connected to the network"
            }
        }
        catch {
            $details += "Status: NOT FOUND IN ACTIVE DIRECTORY"
            $details += "Error: $($_.Exception.Message)"
            $details += ""
            $details += "Note: The device may not exist in Active Directory, or you may not have"
            $details += "permission to query AD. Ensure the Active Directory PowerShell module is installed."
        }
        
        $details += ""
        $details += "=" * 80
        $details += "END OF REPORT"
        $details += "=" * 80
        
        $txtDetails.Text = $details -join "`r`n"
    })
    
    [void]$detailsForm.ShowDialog()
})

# Email User for selected loaner
$btnEmailLoanerUser.Add_Click({
    if ($dgvLoaners.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a loaner laptop to email the user about.", "No Selection", "OK", "Warning")
        return
    }

    $selectedSerial = $dgvLoaners.SelectedRows[0].Cells["SerialNumber"].Value

    # Get loaner info from inventory (for email fields)
    $script:inventory = @(Load-Inventory)
    $loanerInfo = $script:inventory | Where-Object { $_.SerialNumber -eq $selectedSerial }

    if (-not $loanerInfo) {
        [System.Windows.Forms.MessageBox]::Show("Unable to find this loaner in the inventory file.", "Not Found", "OK", "Error")
        return
    }

    $userEmail = if ($loanerInfo.PSObject.Properties['AssignedEmail']) { $loanerInfo.AssignedEmail } else { "" }
    $managerEmail = if ($loanerInfo.PSObject.Properties['ManagerEmail']) { $loanerInfo.ManagerEmail } else { "" }
    $model = if ($loanerInfo.Model) { $loanerInfo.Model } else { "Unknown model" }
    $assetTag = if ($loanerInfo.PSObject.Properties['AssetTag']) { $loanerInfo.AssetTag } else { "N/A" }

    if ([string]::IsNullOrWhiteSpace($userEmail) -or $userEmail -eq "N/A") {
        [System.Windows.Forms.MessageBox]::Show("This loaner does not have a valid user email address. Please ensure it is assigned and has an email.", "Missing User Email", "OK", "Warning")
        return
    }

    # Dialog to capture sender email
    $emailDialog = New-Object System.Windows.Forms.Form
    $emailDialog.Text = "Send Laptop Verification Email"
    $emailDialog.Size = New-Object System.Drawing.Size(500, 220)
    $emailDialog.StartPosition = "CenterParent"
    $emailDialog.FormBorderStyle = "FixedDialog"
    $emailDialog.MaximizeBox = $false
    $emailDialog.MinimizeBox = $false

    $lblFrom = New-Object System.Windows.Forms.Label
    $lblFrom.Text = "From (your email):"
    $lblFrom.Location = New-Object System.Drawing.Point(20, 20)
    $lblFrom.Size = New-Object System.Drawing.Size(120, 20)
    $emailDialog.Controls.Add($lblFrom)

    $txtFrom = New-Object System.Windows.Forms.TextBox
    $txtFrom.Location = New-Object System.Drawing.Point(150, 18)
    $txtFrom.Size = New-Object System.Drawing.Size(320, 20)
    $emailDialog.Controls.Add($txtFrom)

    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = "To (user):"
    $lblTo.Location = New-Object System.Drawing.Point(20, 55)
    $lblTo.Size = New-Object System.Drawing.Size(120, 20)
    $emailDialog.Controls.Add($lblTo)

    $txtTo = New-Object System.Windows.Forms.TextBox
    $txtTo.Location = New-Object System.Drawing.Point(150, 53)
    $txtTo.Size = New-Object System.Drawing.Size(320, 20)
    $txtTo.ReadOnly = $true
    $txtTo.Text = $userEmail
    $emailDialog.Controls.Add($txtTo)

    $lblCc = New-Object System.Windows.Forms.Label
    $lblCc.Text = "CC (manager + you):"
    $lblCc.Location = New-Object System.Drawing.Point(20, 90)
    $lblCc.Size = New-Object System.Drawing.Size(120, 20)
    $emailDialog.Controls.Add($lblCc)

    $txtCc = New-Object System.Windows.Forms.TextBox
    $txtCc.Location = New-Object System.Drawing.Point(150, 88)
    $txtCc.Size = New-Object System.Drawing.Size(320, 20)
    $txtCc.ReadOnly = $true
    $txtCc.Text = $managerEmail
    $emailDialog.Controls.Add($txtCc)

    $btnSend = New-Object System.Windows.Forms.Button
    $btnSend.Text = "Send"
    $btnSend.Location = New-Object System.Drawing.Point(260, 135)
    $btnSend.Size = New-Object System.Drawing.Size(80, 30)
    $btnSend.DialogResult = "OK"
    $emailDialog.Controls.Add($btnSend)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(350, 135)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.DialogResult = "Cancel"
    $emailDialog.Controls.Add($btnCancel)

    $emailDialog.AcceptButton = $btnSend
    $emailDialog.CancelButton = $btnCancel

    $dialogResult = $emailDialog.ShowDialog()

    if ($dialogResult -ne "OK") {
        return
    }

    $fromEmail = $txtFrom.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($fromEmail) -or (-not $fromEmail.Contains("@"))) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid sender email address.", "Invalid Email", "OK", "Warning")
        return
    }

    # Build CC list: manager (if any) + sender
    $ccList = @()
    if (-not [string]::IsNullOrWhiteSpace($managerEmail) -and $managerEmail -ne "N/A") {
        $ccList += $managerEmail
    }
    $ccList += $fromEmail

    try {
        Send-Email -SendTo $userEmail -Date (Get-Date -Format "MM/dd/yyyy") -AssetTag $assetTag -SerialNumber $selectedSerial -Model $model -EmailAdd $fromEmail -Cc $ccList
        [System.Windows.Forms.MessageBox]::Show("Email sent successfully to $userEmail.", "Email Sent", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to send email: $($_.Exception.Message)", "Email Error", "OK", "Error")
    }
})

# Add to Missing Laptops button (CSV tab)
$btnAddToMissing.Add_Click({
    if ($dgvCsv.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a laptop from the CSV list.", "No Selection", "OK", "Warning")
        return
    }

    $script:missingLaptops = Load-MissingLaptops
    $addedCount = 0
    $skippedItems = @()

    foreach ($selectedRow in $dgvCsv.SelectedRows) {
        $rowIndex = $selectedRow.Index
        $csvItem = $script:csvData[$rowIndex]

        $serial = $csvItem.serial_number
        $model = if ($csvItem.model) { $csvItem.model } else { "" }
        $assetTag = if ($csvItem.PSObject.Properties['asset_tag'] -and $csvItem.asset_tag) { $csvItem.asset_tag } else { "N/A" }
        $assignedTo = if ($csvItem.PSObject.Properties['assigned_to'] -and $csvItem.assigned_to) { $csvItem.assigned_to } else { "" }
        $csvStatus = if ($csvItem.PSObject.Properties['install_status'] -and $csvItem.install_status) { $csvItem.install_status } else { "" }
        $csvLocation = if ($csvItem.PSObject.Properties['stockroom'] -and $csvItem.stockroom) { $csvItem.stockroom } else { "" }

        $existing = $script:missingLaptops | Where-Object { $_.SerialNumber.Trim() -eq $serial.Trim() }
        if ($existing) {
            $skippedItems += $serial
            continue
        }

        $newMissing = [PSCustomObject]@{
            SerialNumber = $serial
            Model        = $model
            AssetTag     = $assetTag
            AssignedTo   = $assignedTo
            CSVStatus    = $csvStatus
            CSVLocation  = $csvLocation
            DateMarked   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
        $script:missingLaptops = @($script:missingLaptops) + @($newMissing)
        $addedCount++
    }

    Save-MissingLaptops
    Update-MissingLaptopsList

    $message = "Added $addedCount laptop(s) to Missing Laptops."
    if ($skippedItems.Count -gt 0) {
        $message += "`n`nAlready in Missing Laptops (skipped):`n" + ($skippedItems -join "`n")
    }
    [System.Windows.Forms.MessageBox]::Show($message, "Missing Laptops Updated", "OK", "Information")
})

# Remove from Missing Laptops button
$btnRemoveMissing.Add_Click({
    if ($dgvMissing.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a laptop to remove.", "No Selection", "OK", "Warning")
        return
    }

    $selectedSerial = $dgvMissing.SelectedRows[0].Cells["SerialNumber"].Value
    $result = [System.Windows.Forms.MessageBox]::Show("Remove '$selectedSerial' from Missing Laptops?", "Confirm Removal", "YesNo", "Question")

    if ($result -eq "Yes") {
        $script:missingLaptops = @(Load-MissingLaptops | Where-Object { $_.SerialNumber -ne $selectedSerial })
        Save-MissingLaptops
        Update-MissingLaptopsList
    }
})

# Refresh Missing Laptops button
$btnRefreshMissing.Add_Click({
    Update-MissingLaptopsList
})

# Email User from Missing Laptops tab
$btnEmailMissingUser.Add_Click({
    if ($dgvMissing.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a missing laptop.", "No Selection", "OK", "Warning")
        return
    }

    $selectedSerial = $dgvMissing.SelectedRows[0].Cells["SerialNumber"].Value
    $script:missingLaptops = Load-MissingLaptops
    $missingItem = $script:missingLaptops | Where-Object { $_.SerialNumber -eq $selectedSerial }

    if (-not $missingItem) {
        [System.Windows.Forms.MessageBox]::Show("Unable to find this entry.", "Not Found", "OK", "Error")
        return
    }

    $assignedTo = $missingItem.AssignedTo
    $model = if ($missingItem.Model) { $missingItem.Model } else { "Unknown Model" }
    $assetTag = if ($missingItem.AssetTag) { $missingItem.AssetTag } else { "N/A" }

    if ([string]::IsNullOrWhiteSpace($assignedTo)) {
        [System.Windows.Forms.MessageBox]::Show("No assigned user recorded for this laptop.", "No User", "OK", "Warning")
        return
    }

    $userEmail = ""
    $managerEmail = ""
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $adUser = $null
    $safeName = $assignedTo -replace "'", "''"

    try {
        $r = @(Get-ADUser -Filter "DisplayName -eq '$safeName'" -Properties EmailAddress, Manager -ErrorAction Stop)
        if ($r.Count -gt 0) { $adUser = $r[0] }
    } catch { }

    if (-not $adUser) {
        try {
            $r = @(Get-ADUser -Filter "Name -eq '$safeName'" -Properties EmailAddress, Manager -ErrorAction Stop)
            if ($r.Count -gt 0) { $adUser = $r[0] }
        } catch { }
    }

    if (-not $adUser -and $assignedTo -match ' ') {
        try {
            $nameParts = $assignedTo -split '\s+', 2
            $firstName = $nameParts[0].Trim() -replace "'", "''"
            $lastName  = $nameParts[1].Trim() -replace "'", "''"
            $r = @(Get-ADUser -Filter "GivenName -eq '$firstName' -and Surname -eq '$lastName'" -Properties EmailAddress, Manager -ErrorAction Stop)
            if ($r.Count -gt 0) { $adUser = $r[0] }
        } catch { }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    if (-not $adUser) {
        [System.Windows.Forms.MessageBox]::Show("Could not find '$assignedTo' in Active Directory.", "AD Lookup Error", "OK", "Error")
        return
    }

    $userEmail = if ($adUser.EmailAddress) { $adUser.EmailAddress } else { "" }
    if ($adUser.Manager) {
        try {
            $mgr = Get-ADUser -Identity $adUser.Manager -Properties EmailAddress -ErrorAction Stop
            $managerEmail = if ($mgr.EmailAddress) { $mgr.EmailAddress } else { "" }
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($userEmail)) {
        [System.Windows.Forms.MessageBox]::Show("User '$assignedTo' has no email address in Active Directory.", "No Email", "OK", "Warning")
        return
    }

    $emailDialog = New-Object System.Windows.Forms.Form
    $emailDialog.Text = "Send Laptop Verification Email"
    $emailDialog.Size = New-Object System.Drawing.Size(500, 220)
    $emailDialog.StartPosition = "CenterParent"
    $emailDialog.FormBorderStyle = "FixedDialog"
    $emailDialog.MaximizeBox = $false
    $emailDialog.MinimizeBox = $false

    $lFrom = New-Object System.Windows.Forms.Label; $lFrom.Text = "From (your email):"; $lFrom.Location = New-Object System.Drawing.Point(20,20); $lFrom.Size = New-Object System.Drawing.Size(120,20); $emailDialog.Controls.Add($lFrom)
    $tFrom = New-Object System.Windows.Forms.TextBox; $tFrom.Location = New-Object System.Drawing.Point(150,18); $tFrom.Size = New-Object System.Drawing.Size(320,20); $emailDialog.Controls.Add($tFrom)
    $lTo = New-Object System.Windows.Forms.Label; $lTo.Text = "To (user):"; $lTo.Location = New-Object System.Drawing.Point(20,55); $lTo.Size = New-Object System.Drawing.Size(120,20); $emailDialog.Controls.Add($lTo)
    $tTo = New-Object System.Windows.Forms.TextBox; $tTo.Location = New-Object System.Drawing.Point(150,53); $tTo.Size = New-Object System.Drawing.Size(320,20); $tTo.ReadOnly = $true; $tTo.Text = $userEmail; $emailDialog.Controls.Add($tTo)
    $lCc = New-Object System.Windows.Forms.Label; $lCc.Text = "CC (manager + you):"; $lCc.Location = New-Object System.Drawing.Point(20,90); $lCc.Size = New-Object System.Drawing.Size(120,20); $emailDialog.Controls.Add($lCc)
    $tCc = New-Object System.Windows.Forms.TextBox; $tCc.Location = New-Object System.Drawing.Point(150,88); $tCc.Size = New-Object System.Drawing.Size(320,20); $tCc.ReadOnly = $true; $tCc.Text = if ($managerEmail) { "$managerEmail + you" } else { "you" }; $emailDialog.Controls.Add($tCc)

    $bSend = New-Object System.Windows.Forms.Button; $bSend.Text = "Send"; $bSend.Location = New-Object System.Drawing.Point(260,135); $bSend.Size = New-Object System.Drawing.Size(80,30); $bSend.DialogResult = "OK"; $emailDialog.Controls.Add($bSend)
    $bCancel = New-Object System.Windows.Forms.Button; $bCancel.Text = "Cancel"; $bCancel.Location = New-Object System.Drawing.Point(350,135); $bCancel.Size = New-Object System.Drawing.Size(80,30); $bCancel.DialogResult = "Cancel"; $emailDialog.Controls.Add($bCancel)
    $emailDialog.AcceptButton = $bSend; $emailDialog.CancelButton = $bCancel

    if ($emailDialog.ShowDialog() -ne "OK") { return }

    $fromEmail = $tFrom.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($fromEmail) -or (-not $fromEmail.Contains("@"))) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid sender email address.", "Invalid Email", "OK", "Warning")
        return
    }

    $ccList = @()
    if (-not [string]::IsNullOrWhiteSpace($managerEmail)) { $ccList += $managerEmail }
    $ccList += $fromEmail

    try {
        Send-Email -SendTo $userEmail -AssetTag $assetTag -SerialNumber $selectedSerial -Model $model -EmailAdd $fromEmail -Cc $ccList

        # Add to Awaiting Verification list (replace any existing entry for same serial)
        $script:awaitingVerification = @($script:awaitingVerification | Where-Object { $_.SerialNumber.Trim() -ne $selectedSerial.Trim() })
        $script:awaitingVerification += [PSCustomObject]@{
            SerialNumber  = $selectedSerial
            Model         = $model
            AssetTag      = $assetTag
            AssignedTo    = $assignedTo
            EmailSentDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            EmailSentBy   = $fromEmail
            Source        = "Missing"
        }
        Save-AwaitingVerification
        Update-AwaitingVerificationList

        [System.Windows.Forms.MessageBox]::Show("Email sent successfully to $userEmail.`n`nThis laptop has been added to the 'Awaiting Verification' list.", "Email Sent", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to send email: $($_.Exception.Message)", "Email Error", "OK", "Error")
    }
})

# Function to update loaner list
function Update-LoanerList {
    param(
        [bool]$IncludeStatus = $false
    )
    
    $script:inventory = @(Load-Inventory)
    
    $loaners = $script:inventory | Where-Object { $_.IsLoaner -eq $true }
    
    # Create DataTable for loaner laptops
    $dt = New-Object System.Data.DataTable
    $dt.Columns.Add("SerialNumber") | Out-Null
    $dt.Columns.Add("Model") | Out-Null
    $dt.Columns.Add("Location") | Out-Null
    $dt.Columns.Add("AssignedTo") | Out-Null
    $dt.Columns.Add("UserName") | Out-Null
    $dt.Columns.Add("UserEmail") | Out-Null
    $dt.Columns.Add("ManagerName") | Out-Null
    $dt.Columns.Add("ManagerEmail") | Out-Null
    if ($IncludeStatus) {
        $dt.Columns.Add("Status") | Out-Null
    }
    $dt.Columns.Add("DateAdded") | Out-Null
    
    foreach ($loaner in $loaners) {
        $dr = $dt.NewRow()
        $dr["SerialNumber"] = $loaner.SerialNumber
        $dr["Model"] = if ([string]::IsNullOrWhiteSpace($loaner.Model)) { "" } else { $loaner.Model }
        $dr["Location"] = $loaner.Location
        $dr["AssignedTo"] = if ($loaner.PSObject.Properties['AssignedTo']) { $loaner.AssignedTo } else { "" }
        $dr["UserName"] = if ($loaner.PSObject.Properties['AssignedName']) { $loaner.AssignedName } else { "" }
        $dr["UserEmail"] = if ($loaner.PSObject.Properties['AssignedEmail']) { $loaner.AssignedEmail } else { "" }
        $dr["ManagerName"] = if ($loaner.PSObject.Properties['ManagerName']) { $loaner.ManagerName } else { "" }
        $dr["ManagerEmail"] = if ($loaner.PSObject.Properties['ManagerEmail']) { $loaner.ManagerEmail } else { "" }
        
        # Add status if available
        if ($IncludeStatus) {
            if ($script:loanerStatus -and $script:loanerStatus.ContainsKey($loaner.SerialNumber)) {
                $dr["Status"] = $script:loanerStatus[$loaner.SerialNumber]
            } else {
                $dr["Status"] = "Unknown"
            }
        }
        
        $dr["DateAdded"] = $loaner.DateAdded
        $dt.Rows.Add($dr)
    }
    
    $dgvLoaners.DataSource = $dt
    
    # Color code rows based on status if status is included
    if ($IncludeStatus) {
        foreach ($row in $dgvLoaners.Rows) {
            $status = $row.Cells["Status"].Value
            if ($status -eq "Online") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::Green
            } elseif ($status -eq "Offline") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightCoral
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::Red
            } elseif ($status -eq "DNS Error") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::DarkOrange
            }
        }
    }
}

# Function to update inventory grid
function Update-InventoryGrid {
    param(
        [string]$SearchText = ""
    )
    
    $script:inventory = @(Load-Inventory)
    
    # Filter inventory based on search text
    $filteredInventory = $script:inventory
    if (-not [string]::IsNullOrWhiteSpace($SearchText)) {
        $filteredInventory = $script:inventory | Where-Object {
            ($_.SerialNumber -like "*$SearchText*") -or
            ($_.Model -like "*$SearchText*") -or
            ($_.Location -like "*$SearchText*")
        }
    }
    
    $dt = New-Object System.Data.DataTable
    $dt.Columns.Add("SerialNumber") | Out-Null
    $dt.Columns.Add("Model") | Out-Null
    $dt.Columns.Add("Location") | Out-Null
    $dt.Columns.Add("IsLoaner") | Out-Null
    $dt.Columns.Add("DateAdded") | Out-Null
    
    foreach ($item in $filteredInventory) {
        $dr = $dt.NewRow()
        $dr["SerialNumber"] = $item.SerialNumber
        $dr["Model"] = $item.Model
        $dr["Location"] = $item.Location
        $dr["IsLoaner"] = if ($item.IsLoaner) { "Yes" } else { "No" }
        $dr["DateAdded"] = $item.DateAdded
        $dt.Rows.Add($dr)
    }
    
    $dgvInventory.DataSource = $dt
}

function Send-Email {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SendTo,

        [Parameter(Mandatory=$false)]
        [string]$Date = (Get-Date -Format "MM/dd/yyyy"),

        [Parameter(Mandatory=$true)]
        [string]$AssetTag,

        [Parameter(Mandatory=$true)]
        [string]$SerialNumber,

        [Parameter(Mandatory=$true)]
        [string]$Model,

        [Parameter(Mandatory=$true)]
        [string]$EmailAdd,

        [Parameter(Mandatory=$false)]
        [string[]]$Cc
    )

    $Body = @"
Hi,

I'm checking on the status of a company-issued laptop and need to confirm whether you currently have the following device:

    Model:       $Model
    Asset Tag:   $AssetTag
    Serial #:    $SerialNumber

If you DO have this laptop, please reply and let me know that it's in your possession.

If you DO NOT have this laptop, please reply "No" and include any details you may know (returned to IT, transferred to another user, replaced, etc.).

Thanks for your help.

Regards,
Brian Fontaine
"@

    $SMTP = "apprelay.mcgladrey.rsm.net"

    $mailParams = @{
        From       = $EmailAdd
        To         = $SendTo
        Subject    = "Laptop Verification - $Model ($AssetTag) - $Date"
        Body       = $Body
        SmtpServer = $SMTP
    }

    if ($Cc -and $Cc.Count -gt 0) {
        $mailParams.Cc = $Cc
    }

    Send-MailMessage @mailParams
}

function Update-MissingLaptopsList {
    $script:missingLaptops = Load-MissingLaptops

    $dt = New-Object System.Data.DataTable
    $dt.Columns.Add("SerialNumber") | Out-Null
    $dt.Columns.Add("Model")        | Out-Null
    $dt.Columns.Add("AssetTag")     | Out-Null
    $dt.Columns.Add("AssignedTo")   | Out-Null
    $dt.Columns.Add("CSVStatus")    | Out-Null
    $dt.Columns.Add("CSVLocation")  | Out-Null
    $dt.Columns.Add("DateMarked")   | Out-Null

    foreach ($item in $script:missingLaptops) {
        $dr = $dt.NewRow()
        $dr["SerialNumber"] = $item.SerialNumber
        $dr["Model"]        = if ($item.Model)       { $item.Model }       else { "" }
        $dr["AssetTag"]     = if ($item.AssetTag)    { $item.AssetTag }    else { "N/A" }
        $dr["AssignedTo"]   = if ($item.AssignedTo)  { $item.AssignedTo }  else { "" }
        $dr["CSVStatus"]    = if ($item.CSVStatus)   { $item.CSVStatus }   else { "" }
        $dr["CSVLocation"]  = if ($item.CSVLocation) { $item.CSVLocation } else { "" }
        $dr["DateMarked"]   = $item.DateMarked
        $dt.Rows.Add($dr)
    }

    $dgvMissing.DataSource = $dt

    # Highlight rows red to make it obvious these are missing
    foreach ($row in $dgvMissing.Rows) {
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
        $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
    }
}

function Update-AwaitingVerificationList {
    $script:awaitingVerification = Load-AwaitingVerification

    $dt = New-Object System.Data.DataTable
    $dt.Columns.Add("SerialNumber")  | Out-Null
    $dt.Columns.Add("Model")         | Out-Null
    $dt.Columns.Add("AssetTag")      | Out-Null
    $dt.Columns.Add("AssignedTo")    | Out-Null
    $dt.Columns.Add("EmailSentDate") | Out-Null
    $dt.Columns.Add("EmailSentBy")   | Out-Null
    $dt.Columns.Add("Source")        | Out-Null

    foreach ($item in $script:awaitingVerification) {
        $dr = $dt.NewRow()
        $dr["SerialNumber"]  = $item.SerialNumber
        $dr["Model"]         = if ($item.Model)         { $item.Model }         else { "" }
        $dr["AssetTag"]      = if ($item.AssetTag)      { $item.AssetTag }      else { "N/A" }
        $dr["AssignedTo"]    = if ($item.AssignedTo)    { $item.AssignedTo }    else { "" }
        $dr["EmailSentDate"] = if ($item.EmailSentDate) { $item.EmailSentDate } else { "" }
        $dr["EmailSentBy"]   = if ($item.EmailSentBy)   { $item.EmailSentBy }   else { "" }
        $dr["Source"]        = if ($item.Source)        { $item.Source }        else { "" }
        $dt.Rows.Add($dr)
    }

    $dgvAwaiting.DataSource = $dt

    foreach ($row in $dgvAwaiting.Rows) {
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
        $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkBlue
    }
}

# Awaiting Verification - Remove Selected (Mark Verified)
$btnRemoveAwaiting.Add_Click({
    if ($dgvAwaiting.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a laptop to mark as verified.", "No Selection", "OK", "Warning")
        return
    }

    $selectedSerial = $dgvAwaiting.SelectedRows[0].Cells["SerialNumber"].Value
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Mark '$selectedSerial' as verified and remove from this list?",
        "Confirm", "YesNo", "Question")

    if ($confirm -eq "Yes") {
        $script:awaitingVerification = @($script:awaitingVerification | Where-Object { $_.SerialNumber.Trim() -ne $selectedSerial.Trim() })
        Save-AwaitingVerification
        Update-AwaitingVerificationList
    }
})

# Awaiting Verification - Refresh
$btnRefreshAwaiting.Add_Click({
    Update-AwaitingVerificationList
})

# Awaiting Verification - Re-send Email
$btnEmailAwaitingUser.Add_Click({
    if ($dgvAwaiting.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a laptop to re-send the email.", "No Selection", "OK", "Warning")
        return
    }

    $selectedSerial  = $dgvAwaiting.SelectedRows[0].Cells["SerialNumber"].Value
    $script:awaitingVerification = Load-AwaitingVerification
    $awaitingItem = $script:awaitingVerification | Where-Object { $_.SerialNumber -eq $selectedSerial }

    if (-not $awaitingItem) {
        [System.Windows.Forms.MessageBox]::Show("Unable to find this entry.", "Not Found", "OK", "Error")
        return
    }

    $assignedTo = $awaitingItem.AssignedTo
    $model      = if ($awaitingItem.Model)    { $awaitingItem.Model }    else { "Unknown Model" }
    $assetTag   = if ($awaitingItem.AssetTag) { $awaitingItem.AssetTag } else { "N/A" }

    if ([string]::IsNullOrWhiteSpace($assignedTo)) {
        [System.Windows.Forms.MessageBox]::Show("No assigned user recorded for this laptop.", "No User", "OK", "Warning")
        return
    }

    $userEmail    = ""
    $managerEmail = ""
    $form.Cursor  = [System.Windows.Forms.Cursors]::WaitCursor

    $adUser   = $null
    $safeName = $assignedTo -replace "'", "''"

    try {
        $r = @(Get-ADUser -Filter "DisplayName -eq '$safeName'" -Properties EmailAddress, Manager -ErrorAction Stop)
        if ($r.Count -gt 0) { $adUser = $r[0] }
    } catch { }
    if (-not $adUser) {
        try {
            $r = @(Get-ADUser -Filter "Name -eq '$safeName'" -Properties EmailAddress, Manager -ErrorAction Stop)
            if ($r.Count -gt 0) { $adUser = $r[0] }
        } catch { }
    }
    if (-not $adUser -and $assignedTo -match ' ') {
        try {
            $nameParts = $assignedTo -split '\s+', 2
            $fn = $nameParts[0].Trim() -replace "'", "''"
            $ln = $nameParts[1].Trim() -replace "'", "''"
            $r = @(Get-ADUser -Filter "GivenName -eq '$fn' -and Surname -eq '$ln'" -Properties EmailAddress, Manager -ErrorAction Stop)
            if ($r.Count -gt 0) { $adUser = $r[0] }
        } catch { }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    if (-not $adUser) {
        [System.Windows.Forms.MessageBox]::Show("Could not find '$assignedTo' in Active Directory.", "AD Lookup Error", "OK", "Error")
        return
    }

    $userEmail = if ($adUser.EmailAddress) { $adUser.EmailAddress } else { "" }
    if ($adUser.Manager) {
        try {
            $mgr = Get-ADUser -Identity $adUser.Manager -Properties EmailAddress -ErrorAction Stop
            $managerEmail = if ($mgr.EmailAddress) { $mgr.EmailAddress } else { "" }
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($userEmail)) {
        [System.Windows.Forms.MessageBox]::Show("User '$assignedTo' has no email address in Active Directory.", "No Email", "OK", "Warning")
        return
    }

    $emailDialog = New-Object System.Windows.Forms.Form
    $emailDialog.Text = "Re-send Laptop Verification Email"
    $emailDialog.Size = New-Object System.Drawing.Size(500, 220)
    $emailDialog.StartPosition = "CenterParent"
    $emailDialog.FormBorderStyle = "FixedDialog"
    $emailDialog.MaximizeBox = $false
    $emailDialog.MinimizeBox = $false

    $lFrom = New-Object System.Windows.Forms.Label; $lFrom.Text = "From (your email):"; $lFrom.Location = New-Object System.Drawing.Point(20,20); $lFrom.Size = New-Object System.Drawing.Size(120,20); $emailDialog.Controls.Add($lFrom)
    $tFrom = New-Object System.Windows.Forms.TextBox; $tFrom.Location = New-Object System.Drawing.Point(150,18); $tFrom.Size = New-Object System.Drawing.Size(320,20); $tFrom.Text = $awaitingItem.EmailSentBy; $emailDialog.Controls.Add($tFrom)
    $lTo = New-Object System.Windows.Forms.Label; $lTo.Text = "To (user):"; $lTo.Location = New-Object System.Drawing.Point(20,55); $lTo.Size = New-Object System.Drawing.Size(120,20); $emailDialog.Controls.Add($lTo)
    $tTo = New-Object System.Windows.Forms.TextBox; $tTo.Location = New-Object System.Drawing.Point(150,53); $tTo.Size = New-Object System.Drawing.Size(320,20); $tTo.ReadOnly = $true; $tTo.Text = $userEmail; $emailDialog.Controls.Add($tTo)
    $lCc = New-Object System.Windows.Forms.Label; $lCc.Text = "CC (manager + you):"; $lCc.Location = New-Object System.Drawing.Point(20,90); $lCc.Size = New-Object System.Drawing.Size(120,20); $emailDialog.Controls.Add($lCc)
    $tCc = New-Object System.Windows.Forms.TextBox; $tCc.Location = New-Object System.Drawing.Point(150,88); $tCc.Size = New-Object System.Drawing.Size(320,20); $tCc.ReadOnly = $true; $tCc.Text = if ($managerEmail) { "$managerEmail + you" } else { "you" }; $emailDialog.Controls.Add($tCc)

    $bSend = New-Object System.Windows.Forms.Button; $bSend.Text = "Send"; $bSend.Location = New-Object System.Drawing.Point(260,135); $bSend.Size = New-Object System.Drawing.Size(80,30); $bSend.DialogResult = "OK"; $emailDialog.Controls.Add($bSend)
    $bCancel = New-Object System.Windows.Forms.Button; $bCancel.Text = "Cancel"; $bCancel.Location = New-Object System.Drawing.Point(350,135); $bCancel.Size = New-Object System.Drawing.Size(80,30); $bCancel.DialogResult = "Cancel"; $emailDialog.Controls.Add($bCancel)
    $emailDialog.AcceptButton = $bSend; $emailDialog.CancelButton = $bCancel

    if ($emailDialog.ShowDialog() -ne "OK") { return }

    $fromEmail = $tFrom.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($fromEmail) -or (-not $fromEmail.Contains("@"))) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid sender email address.", "Invalid Email", "OK", "Warning")
        return
    }

    $ccList = @()
    if (-not [string]::IsNullOrWhiteSpace($managerEmail)) { $ccList += $managerEmail }
    $ccList += $fromEmail

    try {
        Send-Email -SendTo $userEmail -AssetTag $assetTag -SerialNumber $selectedSerial -Model $model -EmailAdd $fromEmail -Cc $ccList

        # Update the email sent date and sender in the awaiting list
        $script:awaitingVerification = @($script:awaitingVerification | Where-Object { $_.SerialNumber.Trim() -ne $selectedSerial.Trim() })
        $script:awaitingVerification += [PSCustomObject]@{
            SerialNumber  = $selectedSerial
            Model         = $model
            AssetTag      = $assetTag
            AssignedTo    = $assignedTo
            EmailSentDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            EmailSentBy   = $fromEmail
            Source        = if ($awaitingItem.Source) { $awaitingItem.Source } else { "Awaiting" }
        }
        Save-AwaitingVerification
        Update-AwaitingVerificationList

        [System.Windows.Forms.MessageBox]::Show("Email re-sent successfully to $userEmail.", "Email Sent", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to send email: $($_.Exception.Message)", "Email Error", "OK", "Error")
    }
})

function Update-NeedsVerificationList {
    $dt = New-Object System.Data.DataTable
    $dt.Columns.Add("SerialNumber") | Out-Null
    $dt.Columns.Add("Model")        | Out-Null
    $dt.Columns.Add("AssetTag")     | Out-Null
    $dt.Columns.Add("AssignedTo")   | Out-Null
    $dt.Columns.Add("InAD")         | Out-Null
    $dt.Columns.Add("UserInAD")     | Out-Null
    $dt.Columns.Add("VerifiedStatus")| Out-Null

    if ($script:csvData -and $script:csvData.Count -gt 0) {
        $csvDt = $dgvCsv.DataSource
        for ($i = 0; $i -lt $script:csvData.Count; $i++) {
            $row = $script:csvData[$i]
            $verifiedVal = if ($csvDt) { $csvDt.Rows[$i]["Verified"] } else { "Pending" }

            if ($verifiedVal -eq "Needs Verification" -or $verifiedVal -eq "Missing") {
                $dr = $dt.NewRow()
                $dr["SerialNumber"]    = $row.serial_number
                $dr["Model"]          = if ($row.PSObject.Properties['model_category'] -and $row.model_category) { $row.model_category } elseif ($row.PSObject.Properties['model'] -and $row.model) { $row.model } else { "" }
                $dr["AssetTag"]       = if ($row.PSObject.Properties['asset_tag'] -and $row.asset_tag) { $row.asset_tag } else { "N/A" }
                $dr["AssignedTo"]     = if ($row.PSObject.Properties['assigned_to'] -and $row.assigned_to) { $row.assigned_to } else { "" }
                $dr["InAD"]           = if ($csvDt) { $csvDt.Rows[$i]["In AD"] } else { "" }
                $dr["UserInAD"]       = if ($csvDt) { $csvDt.Rows[$i]["User in AD?"] } else { "" }
                $dr["VerifiedStatus"] = $verifiedVal
                $dt.Rows.Add($dr)
            }
        }
    }

    $dgvNeedsVerification.DataSource = $dt

    foreach ($row in $dgvNeedsVerification.Rows) {
        $statusVal = $row.Cells["VerifiedStatus"].Value
        if ($statusVal -eq "Missing") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
        } else {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkOrange
        }
    }
}

# Shared helper: look up AD user by display name and return adUser object or $null
function Find-ADUserByName {
    param([string]$Name)
    $safeName = $Name -replace "'", "''"
    $adUser = $null

    try {
        $r = @(Get-ADUser -Filter "DisplayName -eq '$safeName'" -Properties EmailAddress, Manager -ErrorAction Stop)
        if ($r.Count -gt 0) { $adUser = $r[0] }
    } catch { }
    if (-not $adUser) {
        try {
            $r = @(Get-ADUser -Filter "Name -eq '$safeName'" -Properties EmailAddress, Manager -ErrorAction Stop)
            if ($r.Count -gt 0) { $adUser = $r[0] }
        } catch { }
    }
    if (-not $adUser -and $Name -match ' ') {
        try {
            $parts = $Name -split '\s+', 2
            $fn = $parts[0].Trim() -replace "'", "''"
            $ln = $parts[1].Trim() -replace "'", "''"
            $r = @(Get-ADUser -Filter "GivenName -eq '$fn' -and Surname -eq '$ln'" -Properties EmailAddress, Manager -ErrorAction Stop)
            if ($r.Count -gt 0) { $adUser = $r[0] }
        } catch { }
    }
    return $adUser
}

# Needs Verification - Refresh
$btnRefreshNeedsVerif.Add_Click({
    Update-NeedsVerificationList
})

# Needs Verification - Email Selected
$btnEmailSelectedNeedsVerif.Add_Click({
    if ($dgvNeedsVerification.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a device to email.", "No Selection", "OK", "Warning")
        return
    }

    $selRow     = $dgvNeedsVerification.SelectedRows[0]
    $serial     = $selRow.Cells["SerialNumber"].Value
    $assignedTo = $selRow.Cells["AssignedTo"].Value
    $model      = $selRow.Cells["Model"].Value
    $assetTag   = $selRow.Cells["AssetTag"].Value

    if ([string]::IsNullOrWhiteSpace($assignedTo)) {
        [System.Windows.Forms.MessageBox]::Show("No assigned user for this device.", "No User", "OK", "Warning")
        return
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $adUser = Find-ADUserByName -Name $assignedTo
    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    if (-not $adUser) {
        [System.Windows.Forms.MessageBox]::Show("Could not find '$assignedTo' in Active Directory.", "AD Lookup Error", "OK", "Error")
        return
    }

    $userEmail = if ($adUser.EmailAddress) { $adUser.EmailAddress } else { "" }
    $managerEmail = ""
    if ($adUser.Manager) {
        try {
            $mgr = Get-ADUser -Identity $adUser.Manager -Properties EmailAddress -ErrorAction Stop
            $managerEmail = if ($mgr.EmailAddress) { $mgr.EmailAddress } else { "" }
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($userEmail)) {
        [System.Windows.Forms.MessageBox]::Show("'$assignedTo' has no email address in AD.", "No Email", "OK", "Warning")
        return
    }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Send Verification Email"
    $dlg.Size = New-Object System.Drawing.Size(500, 220)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

    $lF = New-Object System.Windows.Forms.Label; $lF.Text = "From (your email):"; $lF.Location = New-Object System.Drawing.Point(20,20); $lF.Size = New-Object System.Drawing.Size(120,20); $dlg.Controls.Add($lF)
    $tF = New-Object System.Windows.Forms.TextBox; $tF.Location = New-Object System.Drawing.Point(150,18); $tF.Size = New-Object System.Drawing.Size(320,20); $dlg.Controls.Add($tF)
    $lT = New-Object System.Windows.Forms.Label; $lT.Text = "To (user):"; $lT.Location = New-Object System.Drawing.Point(20,55); $lT.Size = New-Object System.Drawing.Size(120,20); $dlg.Controls.Add($lT)
    $tT = New-Object System.Windows.Forms.TextBox; $tT.Location = New-Object System.Drawing.Point(150,53); $tT.Size = New-Object System.Drawing.Size(320,20); $tT.ReadOnly = $true; $tT.Text = $userEmail; $dlg.Controls.Add($tT)
    $lC = New-Object System.Windows.Forms.Label; $lC.Text = "CC (manager + you):"; $lC.Location = New-Object System.Drawing.Point(20,90); $lC.Size = New-Object System.Drawing.Size(120,20); $dlg.Controls.Add($lC)
    $tC = New-Object System.Windows.Forms.TextBox; $tC.Location = New-Object System.Drawing.Point(150,88); $tC.Size = New-Object System.Drawing.Size(320,20); $tC.ReadOnly = $true; $tC.Text = if ($managerEmail) { "$managerEmail + you" } else { "you" }; $dlg.Controls.Add($tC)
    $bS = New-Object System.Windows.Forms.Button; $bS.Text = "Send"; $bS.Location = New-Object System.Drawing.Point(260,135); $bS.Size = New-Object System.Drawing.Size(80,30); $bS.DialogResult = "OK"; $dlg.Controls.Add($bS)
    $bX = New-Object System.Windows.Forms.Button; $bX.Text = "Cancel"; $bX.Location = New-Object System.Drawing.Point(350,135); $bX.Size = New-Object System.Drawing.Size(80,30); $bX.DialogResult = "Cancel"; $dlg.Controls.Add($bX)
    $dlg.AcceptButton = $bS; $dlg.CancelButton = $bX

    if ($dlg.ShowDialog() -ne "OK") { return }

    $fromEmail = $tF.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($fromEmail) -or (-not $fromEmail.Contains("@"))) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid sender email.", "Invalid Email", "OK", "Warning")
        return
    }

    $ccList = @()
    if (-not [string]::IsNullOrWhiteSpace($managerEmail)) { $ccList += $managerEmail }
    $ccList += $fromEmail

    try {
        Send-Email -SendTo $userEmail -AssetTag $assetTag -SerialNumber $serial -Model $model -EmailAdd $fromEmail -Cc $ccList

        $script:awaitingVerification = @($script:awaitingVerification | Where-Object { $_.SerialNumber.Trim() -ne $serial.Trim() })
        $script:awaitingVerification += [PSCustomObject]@{
            SerialNumber  = $serial
            Model         = $model
            AssetTag      = $assetTag
            AssignedTo    = $assignedTo
            EmailSentDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            EmailSentBy   = $fromEmail
            Source        = "Needs Verification"
        }
        Save-AwaitingVerification
        Update-AwaitingVerificationList

        # Update Verified cell in CSV grid
        $csvDt = $dgvCsv.DataSource
        if ($csvDt) {
            for ($i = 0; $i -lt $script:csvData.Count; $i++) {
                if ($script:csvData[$i].serial_number.Trim() -eq $serial.Trim()) {
                    $csvDt.Rows[$i]["Verified"] = "Email Sent"
                    $dgvCsv.Rows[$i].Cells["Verified"].Style.ForeColor = [System.Drawing.Color]::DarkBlue
                    $dgvCsv.Rows[$i].Cells["Verified"].Style.Font = New-Object System.Drawing.Font($dgvCsv.Font, [System.Drawing.FontStyle]::Bold)
                    break
                }
            }
        }

        Update-NeedsVerificationList
        [System.Windows.Forms.MessageBox]::Show("Email sent to $userEmail and added to Awaiting Verification.", "Email Sent", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to send email: $($_.Exception.Message)", "Email Error", "OK", "Error")
    }
})

# Needs Verification - Send Email to All
$btnEmailAllNeedsVerif.Add_Click({
    $dt = $dgvNeedsVerification.DataSource
    if (-not $dt -or $dt.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("The Needs Verification list is empty. Run 'Run Verification' on the CSV tab first.", "Nothing to Send", "OK", "Warning")
        return
    }

    $totalRows = $dt.Rows.Count
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will send a verification email to the assigned user of all $totalRows device(s) in this list.`n`nDevices with no assigned user or no AD match will be skipped.`n`nDo you want to continue?",
        "Send Email to All - Confirm", "YesNo", "Warning")
    if ($confirm -ne "Yes") { return }

    # Ask for sender email once for the whole batch
    $fromDlg = New-Object System.Windows.Forms.Form
    $fromDlg.Text = "Sender Email"
    $fromDlg.Size = New-Object System.Drawing.Size(420, 130)
    $fromDlg.StartPosition = "CenterParent"
    $fromDlg.FormBorderStyle = "FixedDialog"
    $fromDlg.MaximizeBox = $false; $fromDlg.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Your email address (From):"; $lbl.Location = New-Object System.Drawing.Point(15,18); $lbl.Size = New-Object System.Drawing.Size(170,20); $fromDlg.Controls.Add($lbl)
    $txtSender = New-Object System.Windows.Forms.TextBox; $txtSender.Location = New-Object System.Drawing.Point(190,16); $txtSender.Size = New-Object System.Drawing.Size(200,20); $fromDlg.Controls.Add($txtSender)
    $bOk = New-Object System.Windows.Forms.Button; $bOk.Text = "Continue"; $bOk.Location = New-Object System.Drawing.Point(190,55); $bOk.Size = New-Object System.Drawing.Size(90,28); $bOk.DialogResult = "OK"; $fromDlg.Controls.Add($bOk)
    $bCx = New-Object System.Windows.Forms.Button; $bCx.Text = "Cancel"; $bCx.Location = New-Object System.Drawing.Point(290,55); $bCx.Size = New-Object System.Drawing.Size(80,28); $bCx.DialogResult = "Cancel"; $fromDlg.Controls.Add($bCx)
    $fromDlg.AcceptButton = $bOk; $fromDlg.CancelButton = $bCx

    if ($fromDlg.ShowDialog() -ne "OK") { return }

    $fromEmail = $txtSender.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($fromEmail) -or (-not $fromEmail.Contains("@"))) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid sender email address.", "Invalid Email", "OK", "Warning")
        return
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblEmailAllStatus.Text = "Sending emails..."
    $form.Refresh()

    $sentCount    = 0
    $skippedCount = 0
    $errorCount   = 0
    $adUserCache  = @{}

    $csvDt = $dgvCsv.DataSource

    foreach ($row in $dt.Rows) {
        $serial     = $row["SerialNumber"]
        $assignedTo = $row["AssignedTo"]
        $model      = $row["Model"]
        $assetTag   = $row["AssetTag"]

        if ([string]::IsNullOrWhiteSpace($assignedTo)) {
            $skippedCount++
            continue
        }

        # AD lookup with cache
        if (-not $adUserCache.ContainsKey($assignedTo)) {
            $adUserCache[$assignedTo] = Find-ADUserByName -Name $assignedTo
        }
        $adUser = $adUserCache[$assignedTo]

        if (-not $adUser -or [string]::IsNullOrWhiteSpace($adUser.EmailAddress)) {
            $skippedCount++
            continue
        }

        $userEmail    = $adUser.EmailAddress
        $managerEmail = ""
        if ($adUser.Manager) {
            try {
                $mgr = Get-ADUser -Identity $adUser.Manager -Properties EmailAddress -ErrorAction Stop
                $managerEmail = if ($mgr.EmailAddress) { $mgr.EmailAddress } else { "" }
            } catch { }
        }

        $ccList = @()
        if (-not [string]::IsNullOrWhiteSpace($managerEmail)) { $ccList += $managerEmail }
        $ccList += $fromEmail

        try {
            Send-Email -SendTo $userEmail -AssetTag $assetTag -SerialNumber $serial -Model $model -EmailAdd $fromEmail -Cc $ccList

            # Add / update awaiting verification
            $script:awaitingVerification = @($script:awaitingVerification | Where-Object { $_.SerialNumber.Trim() -ne $serial.Trim() })
            $script:awaitingVerification += [PSCustomObject]@{
                SerialNumber  = $serial
                Model         = $model
                AssetTag      = $assetTag
                AssignedTo    = $assignedTo
                EmailSentDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                EmailSentBy   = $fromEmail
                Source        = "Needs Verification"
            }

            # Update Verified cell in CSV grid
            if ($csvDt) {
                for ($i = 0; $i -lt $script:csvData.Count; $i++) {
                    if ($script:csvData[$i].serial_number.Trim() -eq $serial.Trim()) {
                        $csvDt.Rows[$i]["Verified"] = "Email Sent"
                        $dgvCsv.Rows[$i].Cells["Verified"].Style.ForeColor = [System.Drawing.Color]::DarkBlue
                        $dgvCsv.Rows[$i].Cells["Verified"].Style.Font = New-Object System.Drawing.Font($dgvCsv.Font, [System.Drawing.FontStyle]::Bold)
                        break
                    }
                }
            }

            $sentCount++
            $lblEmailAllStatus.Text = "Sending... $sentCount sent so far"
            $form.Refresh()
        }
        catch {
            $errorCount++
        }
    }

    Save-AwaitingVerification
    Update-AwaitingVerificationList
    Update-NeedsVerificationList

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    $lblEmailAllStatus.Text = "Done: $sentCount sent, $skippedCount skipped (no user/AD match), $errorCount errors."

    [System.Windows.Forms.MessageBox]::Show(
        "Batch email complete!`n`nSent:    $sentCount`nSkipped: $skippedCount (no assigned user or not found in AD)`nErrors:  $errorCount`n`nAll sent devices have been added to 'Awaiting Verification'.",
        "Batch Send Complete", "OK", "Information")
})

# Initialize inventory display
Update-InventoryGrid
Update-LoanerList
Update-MissingLaptopsList
Update-AwaitingVerificationList

# Show form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()


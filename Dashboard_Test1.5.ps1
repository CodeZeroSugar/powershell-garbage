
# Displays vulnerabilities for devices and unit watchlist computers

# Global variable for filtered data
$script:filteredData = @()

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Data Loading ---
# Load vulnerability data from CSV files
$dataPath = "C:\path\to\.csv"
$logPath = "C:\VulnLogs\vulnLog.csv"
$data = Import-Csv -Path $dataPath
$logData = Import-Csv -Path $logPath

# Define list of units
$hurlburtUnits = @(
    "1 SOW", "1 SOCS", "1 SOCES", "1 SOSFS", "492 SOW", "39 IOS", "505 CTS",
    "1 SOFSS", "1 SOCNS", "2 SOS", "11 SOIS", "HQ", "1 SOAMX", "361 ISR",
    "505 CCW", "823 RHS", "556 RHS", "COMP 05"
)

# --- Helper Functions ---
function Update-VulnerabilityCounts {
    param($filteredData)

    # Calculate vulnerability counts
    $vulnCrit = (@($filteredData | Where-Object { $_.'Severity' -like "*Critical*" }).Count) + 0
    $vulnHigh = (@($filteredData | Where-Object { $_.'Severity' -like "*High*" }).Count) + 0
    $vulnMed = (@($filteredData | Where-Object { $_.'Severity' -like "*Medium*" }).Count) + 0
    $vulnLow = (@($filteredData | Where-Object { $_.'Severity' -like "*Low*" }).Count) + 0
    $vulnTotal = ($vulnCrit * 10) + ($vulnHigh * 10) + ($vulnMed * 4) + ($vulnLow * 1)

    # Update vulnTextBox
    $vulnTextBox.Text = "Vulnerabilities".PadRight(15) + "`r`n" +
                        "Critical  : $vulnCrit".PadRight(15) + "`r`n" +
                        "High      : $vulnHigh".PadRight(15) + "`r`n" +
                        "Medium    : $vulnMed".PadRight(15) + "`r`n" +
                        "Low       : $vulnLow".PadRight(15)

    # Update scoreBox
    $scoreBox.Clear()
    $scoreBox.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Center
    $scoreBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 16)
    $scoreBox.AppendText("Score`r`n")
    $scoreBox.SelectionColor = if ($vulnTotal -gt 149) { 'Red' } elseif ($vulnTotal -gt 0) { 'Green' } else { 'Black' }
    $scoreBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 38, [System.Drawing.FontStyle]::Bold)
    $scoreBox.AppendText("$vulnTotal")

    # Update watchlistBox
    $watchlistBox.Clear()
    $watchlistBox.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Center
    $watchlistBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 16)
    $watchlistBox.AppendText("Watchlist Week`r`n")

    # Fetch Watchlist Week
    $targetDevice = if ([string]::IsNullOrWhiteSpace($filteredData[0].'NetBios Name')) { 
        $filteredData[0].'IP address' 
    } else { 
        $filteredData[0].'NetBios Name' 
    }
    $targetDevice = $targetDevice.Trim()
    Write-Host "Target Device: $targetDevice"
    $watchlistDevice = $logData | Where-Object { $_.'NetBIOS Name' -like $targetDevice }
    Write-Host "Watchlist Device: $watchlistDevice"
    $watchlistWeek = if ([string]::IsNullOrWhiteSpace($watchlistDevice)) { 0 } else { $watchlistDevice.'Total Weeks' }

    # Set color based on watchlist week
    $watchlistBox.SelectionColor = switch ($watchlistWeek) {
        1 { 'Green' }
        2 { 'Yellow' }
        3 { 'Orange' }
        4 { 'Red' }
        { $_ -gt 4 } { 'Black' }
        default { 'Black' }
    }
    if ($watchlistWeek -gt 4) { $watchlistWeek = "Q" }
    $watchlistBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 38, [System.Drawing.FontStyle]::Bold)
    $watchlistBox.AppendText("$watchlistWeek")
}

# --- Form Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Comply to Connect Dashboard"
$form.Size = New-Object System.Drawing.Size(1340, 800)
$form.StartPosition = "CenterScreen"

# --- Menu Setup ---
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = "File" }
$exportMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = "Export" }
$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = "Help" }
$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = "About" }
$fileMenu.DropDownItems.Add($exportMenuItem)
$helpMenu.DropDownItems.Add($aboutMenuItem)
$menuStrip.Items.AddRange(@($fileMenu, $helpMenu))
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

# --- UI Controls ---
# Search button
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "Search"
$searchButton.Size = New-Object System.Drawing.Size(120, 23)
$searchButton.Location = New-Object System.Drawing.Point(20, 28)
$form.Controls.Add($searchButton)

# Clear button
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "Clear"
$clearButton.Size = New-Object System.Drawing.Size(120, 23)
$clearButton.Location = New-Object System.Drawing.Point(550, 28)
$form.Controls.Add($clearButton)

# Search text box
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Multiline = $false
$searchBox.Size = New-Object System.Drawing.Size(250, 24)
$searchBox.Font = New-Object System.Drawing.Font("Arial", 10)
$searchBox.Location = New-Object System.Drawing.Point(142, 28)
$form.Controls.Add($searchBox)

# Search type ComboBox
$searchTypeComboBox = New-Object System.Windows.Forms.ComboBox
$searchTypeComboBox.Size = New-Object System.Drawing.Size(150, 23)
$searchTypeComboBox.Location = New-Object System.Drawing.Point(400, 28)
$searchTypeComboBox.Items.AddRange(@("NetBIOS Name", "IP Address"))
$searchTypeComboBox.SelectedIndex = 0
$form.Controls.Add($searchTypeComboBox)

# Unit selection ComboBox
$unitComboBox = New-Object System.Windows.Forms.ComboBox
$unitComboBox.Size = New-Object System.Drawing.Size(100, 23)
$unitComboBox.Location = New-Object System.Drawing.Point(1055, 28)
$unitComboBox.Items.AddRange($hurlburtUnits)
$unitComboBox.Text = "Select a Unit"
$form.Controls.Add($unitComboBox)

# Main DataGridView for vulnerabilities
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(1025, 600)
$dataGridView.Location = New-Object System.Drawing.Point(20, 55)
$dataGridView.ColumnHeadersVisible = $true
$dataGridView.RowHeadersVisible = $false
$dataGridView.AllowUserToAddRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
$dataGridView.ColumnCount = 5
$dataGridView.Columns[0].Name = "NetBIOS Name"
$dataGridView.Columns[1].Name = "IP Address"
$dataGridView.Columns[2].Name = "MAC Address"
$dataGridView.Columns[3].Name = "Severity"
$dataGridView.Columns[4].Name = "Plugin Name"
$form.Controls.Add($dataGridView)

# Unit DataGridView for watchlist computers
$unitGridView = New-Object System.Windows.Forms.DataGridView
$unitGridView.Size = New-Object System.Drawing.Size(250, 600) # Your specified size
$unitGridView.Location = New-Object System.Drawing.Point(1055, 55) # Your specified location
$unitGridView.ColumnHeadersVisible = $true
$unitGridView.RowHeadersVisible = $false
$unitGridView.AllowUserToAddRows = $false
$unitGridView.ReadOnly = $true
$unitGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$unitGridView.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$unitGridView.ColumnCount = 2
$unitGridView.Columns[0].Name = "NetBIOS Name"
$unitGridView.Columns[0].SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$unitGridView.Columns[1].Name = "Score"
$unitGridView.Columns[1].SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$unitGridView.Columns[1].DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$unitGridView.Columns[1].DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(2, 0, 2, 0) # Minimal padding

# Set initial column widths (no scrollbar)
$scoreFont = New-Object System.Drawing.Font("Arial", 10) # Match CellFormatting font
$headerFont = $unitGridView.ColumnHeadersDefaultCellStyle.Font
if ($null -eq $headerFont) { $headerFont = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25) }
$scoreTextSize = [System.Windows.Forms.TextRenderer]::MeasureText("Score", $headerFont)
$scoreValueSize = [System.Windows.Forms.TextRenderer]::MeasureText("150", $scoreFont)
$scoreWidth = [Math]::Max($scoreTextSize.Width, $scoreValueSize.Width) + 15 # Increased padding (10 + 5)
$scoreWidth = [Math]::Max($scoreWidth, 70) # Increased minimum (60 + 5)
$totalUsableWidth = 250 # Full width, no scrollbar
$unitGridView.Columns[1].Width = $scoreWidth
$unitGridView.Columns[0].Width = $totalUsableWidth - $scoreWidth
$form.Controls.Add($unitGridView)

# Vulnerability counts TextBox
$vulnTextBox = New-Object System.Windows.Forms.TextBox
$vulnTextBox.Multiline = $true
$vulnTextBox.ReadOnly = $true
$vulnTextBox.Location = New-Object System.Drawing.Point(20, 662)
$vulnTextBox.Size = New-Object System.Drawing.Size(200, 90)
$vulnTextBox.Font = New-Object System.Drawing.Font("Courier New", 10)
$vulnTextBox.Text = "Vulnerabilities".PadRight(15) + "`r`n" +
                    "Critical  : N/A".PadRight(15) + "`r`n" +
                    "High      : N/A".PadRight(15) + "`r`n" +
                    "Medium    : N/A".PadRight(15) + "`r`n" +
                    "Low       : N/A".PadRight(15)
$form.Controls.Add($vulnTextBox)

# Score RichTextBox
$scoreBox = New-Object System.Windows.Forms.RichTextBox
$scoreBox.ReadOnly = $true
$scoreBox.Multiline = $true
$scoreBox.BorderStyle = 'Fixed3D'
$scoreBox.Location = New-Object System.Drawing.Point(230, 662)
$scoreBox.Size = New-Object System.Drawing.Size(200, 90)
$scoreBox.BackColor = [System.Drawing.SystemColors]::Control
$scoreBox.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Center
$scoreBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 16)
$scoreBox.AppendText("Score`r`n")
$scoreBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 38, [System.Drawing.FontStyle]::Bold)
$scoreBox.AppendText("0")
$form.Controls.Add($scoreBox)

# Watchlist week RichTextBox
$watchlistBox = New-Object System.Windows.Forms.RichTextBox
$watchlistBox.ReadOnly = $true
$watchlistBox.Multiline = $true
$watchlistBox.BorderStyle = 'Fixed3D'
$watchlistBox.Location = New-Object System.Drawing.Point(440, 662)
$watchlistBox.Size = New-Object System.Drawing.Size(200, 90)
$watchlistBox.BackColor = [System.Drawing.SystemColors]::Control
$watchlistBox.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Center
$watchlistBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 16)
$watchlistBox.AppendText("Watchlist Week`r`n")
$watchlistBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 38, [System.Drawing.FontStyle]::Bold)
$watchlistBox.AppendText("0")
$form.Controls.Add($watchlistBox)

# --- Event Handlers ---
# Search box Enter key handler
$searchBox.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter') {
        $searchButton.PerformClick()
    }
})

# Search button click handler
$searchButton.Add_Click({
    $dataGridView.Rows.Clear()
    $filterText = $searchBox.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($filterText)) {
        $dataGridView.Rows.Add("Please enter a search term.", "", "", "", "")
        Update-VulnerabilityCounts -filteredData @()
        $script:filteredData = @()
        return
    }

    $searchType = $searchTypeComboBox.SelectedItem
    $script:filteredData = @()
    if ($searchType -eq "NetBIOS Name") {
        $script:filteredData = $data | Where-Object { $_.'NetBIOS Name' -like "*$filterText*" }
    } elseif ($searchType -eq "IP Address") {
        $script:filteredData = $data | Where-Object { $_.'IP Address' -eq "$filterText" }
    }
    $script:filteredData = @($script:filteredData)

    if ($script:filteredData.Count -gt 0) {
        foreach ($row in $script:filteredData) {
            $dataGridView.Rows.Add($row.'NetBIOS Name', $row.'IP Address', $row.'MAC Address', $row.'Severity', $row.'Plugin Name')
        }
        $dataGridView.Columns[0].FillWeight = 70
        $dataGridView.Columns[1].FillWeight = 50
        $dataGridView.Columns[2].FillWeight = 60
        $dataGridView.Columns[3].FillWeight = 50
        $dataGridView.Columns[4].FillWeight = 270
        Update-VulnerabilityCounts -filteredData $script:filteredData
    } else {
        $dataGridView.Rows.Add("No results found.", "", "", "", "")
        Update-VulnerabilityCounts -filteredData @()
        $script:filteredData = @()
    }

    $dataGridView.Refresh()
})

# Clear button click handler
$clearButton.Add_Click({
    $searchBox.Text = ""
    $searchTypeComboBox.SelectedIndex = 0
    $dataGridView.Rows.Clear()
    $script:filteredData = @()
    $vulnTextBox.Text = "Vulnerabilities".PadRight(15) + "`r`n" +
                        "Critical  : N/A".PadRight(15) + "`r`n" +
                        "High      : N/A".PadRight(15) + "`r`n" +
                        "Medium    : N/A".PadRight(15) + "`r`n" +
                        "Low       : N/A".PadRight(15)
    $scoreBox.Clear()
    $scoreBox.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Center
    $scoreBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 16)
    $scoreBox.AppendText("Score`r`n")
    $scoreBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 38, [System.Drawing.FontStyle]::Bold)
    $scoreBox.AppendText("0")
    $watchlistBox.Clear()
    $watchlistBox.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Center
    $watchlistBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 16)
    $watchlistBox.AppendText("Watchlist Week`r`n")
    $watchlistBox.SelectionFont = New-Object System.Drawing.Font("Courier New", 38, [System.Drawing.FontStyle]::Bold)
    $watchlistBox.AppendText("0")
    $dataGridView.Refresh()
})

# Unit ComboBox selection handler
$unitComboBox.Add_SelectedIndexChanged({
    $unitGridView.Rows.Clear()
    $selectedUnit = $unitComboBox.SelectedItem

    if ([string]::IsNullOrWhiteSpace($selectedUnit)) {
        $unitGridView.Rows.Add("No unit selected.", "")
        Write-Host "No unit selected "
        # Reset widths (no scrollbar)
        $totalUsableWidth = 250
        $unitGridView.Columns[1].Width = $scoreWidth
        $unitGridView.Columns[0].Width = $totalUsableWidth - $scoreWidth
        return
    }

    $selectedUnit = $selectedUnit.Replace(" ", "")
    Write-Host "Selected Unit: $selectedUnit "
    $unitData = $logData | Where-Object { $_.'NetBIOS Name' -like "*$selectedUnit*" }

    $weekData = @{}
    $weeks = @("Week 1", "Week 2", "Week 3", "Week 4", "Quarantine")
    foreach ($week in $weeks) {
        $weekData[$week] = @()
    }

    if ($unitData) {
        Write-Host "Found $($unitData.Count) devices for unit: $selectedUnit "
        foreach ($row in $unitData) {
            $netBiosName = $row.'NetBIOS Name'
            $totalWeeks = $row.'Total Weeks'
            if ([string]::IsNullOrWhiteSpace($totalWeeks)) { $totalWeeks = "0" }
            
            try {
                $weeksNum = [int]$totalWeeks
            } catch {
                $weeksNum = 0
                Write-Host "Invalid Total Weeks for: $netBiosName : $totalWeeks "
            }

            $header = switch ($weeksNum) {
                1 { "Week 1" }
                2 { "Week 2" }
                3 { "Week 3" }
                4 { "Week 4" }
                { $_ -gt 4 } { "Quarantine" }
                default { $null }
            }
            
            if ($header -and $netBiosName) {
                # Calculate score
                $computerVulns = $data | Where-Object { $_.'NetBIOS Name' -eq $netBiosName }
                $vulnCrit = (@($computerVulns | Where-Object { $_.'Severity' -like "*Critical*" }).Count)
                $vulnHigh = (@($computerVulns | Where-Object { $_.'Severity' -like "*High*" }).Count)
                $vulnMed = (@($computerVulns | Where-Object { $_.'Severity' -like "*Medium*" }).Count)
                $vulnLow = (@($computerVulns | Where-Object { $_.'Severity' -like "*Low*" }).Count)
                $score = ($vulnCrit * 10) + ($vulnHigh * 10) + ($vulnMed * 4) + ($vulnLow * 1)
                Write-Host "Computer: $netBiosName , Score: $score (Crit: $vulnCrit , High: $vulnHigh , Med: $vulnMed , Low: $vulnLow )"
                $weekData[$header] += [PSCustomObject]@{ NetBIOSName = $netBiosName; Score = $score }
            } else {
                Write-Host "Skipped: $netBiosName : Invalid header: $header "
            }
        }
    } else {
        Write-Host "No devices found for unit: $selectedUnit "
    }

    $hasData = $false
    foreach ($week in $weeks) {
        if ($weekData[$week].Count -gt 0) {
            $headerRowIndex = $unitGridView.Rows.Add($week, "")
            $unitGridView.Rows[$headerRowIndex].Tag = "Header"
            $hasData = $true
            
            foreach ($computer in $weekData[$week] | Sort-Object { $_.NetBIOSName }) {
                $rowIndex = $unitGridView.Rows.Add($computer.NetBIOSName, $computer.Score)
                $unitGridView.Rows[$rowIndex].Tag = $week
                Write-Host "Added Row: NetBIOS: $($computer.NetBIOSName) , Score: $($computer.Score) "
            }
        }
    }

    if (-not $hasData) {
        $unitGridView.Rows.Add("No results found.", "")
        Write-Host "No watchlist data for unit: $selectedUnit "
    }

    # Adjust column widths based on scrollbar visibility
    $rowHeight = $unitGridView.RowTemplate.Height # Default ~22px
    $visibleRows = [Math]::Floor(600 / $rowHeight) # Approx rows visible in 600px height
    $scrollBarVisible = $unitGridView.RowCount -gt $visibleRows
    $scrollBarWidth = if ($scrollBarVisible) { [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth } else { 0 }
    $totalUsableWidth = 250 - $scrollBarWidth
    $unitGridView.Columns[1].Width = $scoreWidth
    $unitGridView.Columns[0].Width = $totalUsableWidth - $scoreWidth
    Write-Host "Scrollbar: $scrollBarVisible , Scrollbar Width: $scrollBarWidth , Total Usable Width: $totalUsableWidth , Score Width: $scoreWidth , NetBIOS Width: $($unitGridView.Columns[0].Width) "

    # Force refresh
    $unitGridView.Refresh()
    $unitGridView.InvalidateColumn(1) # Invalidate Score column
})
# UnitGridView cell formatting
$unitGridView.Add_CellFormatting({
    param($sender, $e)
    
    if ($e.RowIndex -ge 0) {
        $row = $unitGridView.Rows[$e.RowIndex]
        
        if ($row.Tag -eq "Header") {
            $e.CellStyle.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
            $e.CellStyle.BackColor = [System.Drawing.Color]::LightGray
            $e.CellStyle.ForeColor = [System.Drawing.Color]::Black
            if ($e.ColumnIndex -eq 1) { $e.Value = "" } # Clear Score column for headers
        } elseif ($row.Tag -in @("Week 1", "Week 2", "Week 3", "Week 4", "Quarantine")) {
            switch ($row.Tag) {
                "Week 1" { $e.CellStyle.BackColor = [System.Drawing.Color]::LightGreen }
                "Week 2" { $e.CellStyle.BackColor = [System.Drawing.Color]::LightYellow }
                "Week 3" { $e.CellStyle.BackColor = [System.Drawing.Color]::PeachPuff }
                "Week 4" { $e.CellStyle.BackColor = [System.Drawing.Color]::LightCoral }
                "Quarantine" { $e.CellStyle.BackColor = [System.Drawing.Color]::DarkGray }
            }
            $e.CellStyle.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
            $e.CellStyle.ForeColor = [System.Drawing.Color]::Black
            if ($e.ColumnIndex -eq 1 -and $e.Value -ne $null) {
                $e.CellStyle.Format = "N0"
                Write-Host "Formatting Score: Row=$($e.RowIndex), Value=$($e.Value), Width=$($unitGridView.Columns[1].Width)"
            }
        } else {
            $e.CellStyle.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
            $e.CellStyle.BackColor = [System.Drawing.Color]::White
            $e.CellStyle.ForeColor = [System.Drawing.Color]::Black
            if ($e.ColumnIndex -eq 1 -and $e.Value -ne $null) {
                $e.CellStyle.Format = "N0"
                Write-Host "Formatting Score (No Tag): Row=$($e.RowIndex), Value=$($e.Value), Width=$($unitGridView.Columns[1].Width)"
            }
        }
    }
})

$unitGridView.Add_CellPainting({
    param($sender, $e)
    
    if ($e.RowIndex -ge 0 -and $e.ColumnIndex -in @(0, 1)) {
        $row = $unitGridView.Rows[$e.RowIndex]
        if ($row.Tag -eq "Header") {
            $bounds = $e.CellBounds
            if ($e.ColumnIndex -eq 0) {
                # Merge bounds across both columns
                $nextCellBounds = $unitGridView.GetCellDisplayRectangle(1, $e.RowIndex, $true)
                $bounds.Width += $nextCellBounds.Width
            } else {
                # Skip painting the second column for headers
                $e.Handled = $true
                return
            }

            # Draw the merged cell
            $e.Graphics.FillRectangle([System.Drawing.Brushes]::LightGray, $bounds)
            $text = $e.Value.ToString()
            $textFormat = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter
            [System.Windows.Forms.TextRenderer]::DrawText(
                $e.Graphics,
                $text,
                (New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)),
                $bounds,
                [System.Drawing.Color]::Black,
                $textFormat
            )

            # Draw borders, excluding the inner vertical line
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black)
            $e.Graphics.DrawRectangle($pen, $bounds.X, $bounds.Y, $bounds.Width - 1, $bounds.Height - 1)

            $e.Handled = $true
        } else {
            # Ensure default rendering for non-header rows
            $e.Handled = $false
            if ($e.ColumnIndex -eq 1 -and $e.Value -ne $null) {
                Write-Host "Painting Score: Row=$($e.RowIndex), Value=$($e.Value), Width=$($e.CellBounds.Width)"
            }
        }
    }
})

# UnitGridView cell click handler
$unitGridView.Add_CellClick({
    param($sender, $e)

    if ($e.RowIndex -ge 0 -and $unitGridView.Rows[$e.RowIndex].Tag -ne "Header") {
        $selectedRow = $unitGridView.Rows[$e.RowIndex]
        $filterText = $selectedRow.Cells[0].Value # Always use NetBIOS Name from first column
        Write-Host "Selected Computer: $filterText"

        $dataGridView.Rows.Clear()
        if ([string]::IsNullOrWhiteSpace($filterText)) {
            $dataGridView.Rows.Add("Invalid selection.", "", "", "", "")
            Update-VulnerabilityCounts -filteredData @()
            $script:filteredData = @()
            return
        }

        $script:filteredData = $data | Where-Object { $_.'NetBIOS Name' -eq "$filterText" }
        $script:filteredData = @($script:filteredData)

        if ($script:filteredData.Count -gt 0) {
            foreach ($row in $script:filteredData) {
                $dataGridView.Rows.Add($row.'NetBIOS Name', $row.'IP Address', $row.'MAC Address', $row.'Severity', $row.'Plugin Name')
            }
            $dataGridView.Columns[0].FillWeight = 70
            $dataGridView.Columns[1].FillWeight = 50
            $dataGridView.Columns[2].FillWeight = 60
            $dataGridView.Columns[3].FillWeight = 50
            $dataGridView.Columns[4].FillWeight = 270
            Update-VulnerabilityCounts -filteredData $script:filteredData
        } else {
            $dataGridView.Rows.Add("No results found.", "", "", "", "")
            Update-VulnerabilityCounts -filteredData @()
            $script:filteredData = @()
        }

        $dataGridView.Refresh()
    }
})

# DataGridView cell click handler
$dataGridView.Add_CellClick({
    param($sender, $e)

    if ($e.RowIndex -ge 0) {
        $selectedRow = $dataGridView.Rows[$e.RowIndex]
        $detailsForm = New-Object System.Windows.Forms.Form
        $detailsForm.Text = "Details for " + $selectedRow.Cells["NetBIOS Name"].Value
        $detailsForm.Size = New-Object System.Drawing.Size(500, 525)
        $detailsForm.StartPosition = "CenterParent"

        $labelsText = @("NetBIOS Name", "IP Address", "MAC Address", "Severity", "Plugin Name")
        $yPosition = 20

        foreach ($labelText in $labelsText) {
            $label = New-Object System.Windows.Forms.Label
            $label.Size = New-Object System.Drawing.Size(400, 25)
            $label.Location = New-Object System.Drawing.Point(20, $yPosition)
            $label.Text = "$labelText : " + $selectedRow.Cells[$labelText].Value
            $detailsForm.Controls.Add($label)
            $yPosition += 30
        }

        $originalRow = $data[$e.RowIndex]
        $pluginOutputLabel = New-Object System.Windows.Forms.Label
        $pluginOutputLabel.Size = New-Object System.Drawing.Size(450, 20)
        $pluginOutputLabel.Location = New-Object System.Drawing.Point(20, $yPosition)
        $pluginOutputLabel.Text = "Plugin Output:"
        $pluginOutputLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $detailsForm.Controls.Add($pluginOutputLabel)
        $yPosition += 30

        $pluginOutputBox = New-Object System.Windows.Forms.TextBox
        $pluginOutputBox.Size = New-Object System.Drawing.Size(450, 80)
        $pluginOutputBox.Location = New-Object System.Drawing.Point(20, $yPosition)
        $pluginOutputBox.ReadOnly = $true
        $pluginOutputBox.Multiline = $true
        $pluginOutputBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $pluginOutputBox.Text = $originalRow.'Plugin Output'
        $detailsForm.Controls.Add($pluginOutputBox)
        $yPosition += 90

        $stepsToRemediateLabel = New-Object System.Windows.Forms.Label
        $stepsToRemediateLabel.Size = New-Object System.Drawing.Size(450, 20)
        $stepsToRemediateLabel.Location = New-Object System.Drawing.Point(20, $yPosition)
        $stepsToRemediateLabel.Text = "Steps to Remediate:"
        $stepsToRemediateLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $detailsForm.Controls.Add($stepsToRemediateLabel)
        $yPosition += 30

        $stepsToRemediateBox = New-Object System.Windows.Forms.TextBox
        $stepsToRemediateBox.Size = New-Object System.Drawing.Size(450, 80)
        $stepsToRemediateBox.Location = New-Object System.Drawing.Point(20, $yPosition)
        $stepsToRemediateBox.ReadOnly = $true
        $stepsToRemediateBox.Multiline = $true
        $stepsToRemediateBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $stepsToRemediateBox.Text = $originalRow.'Steps To Remediate'
        $detailsForm.Controls.Add($stepsToRemediateBox)

        $detailsForm.ShowDialog()
    }
})

# Export menu item handler
$exportMenuItem.Add_Click({
    if ($script:filteredData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export. Please perform a search first.", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $exportData = $script:filteredData | Select-Object 'NetBIOS Name', 'IP Address', 'MAC Address', 'Severity', 'Plugin Name', 'Plugin Output', 'Steps To Remediate'
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.FileName = "Vulnerabilities_$(Get-Date -Format 'yyyMMdd_HHmmss').csv"

    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        try {
            $filePath = $saveFileDialog.FileName
            Write-Host "File selected: $filePath"
            $exportData | Export-Csv -Path $filePath -NoTypeInformation -Force
            [System.Windows.Forms.MessageBox]::Show("Data successfully exported to $filePath.", "Export Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to export data: $_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# About menu item handler
$aboutMenuItem.Add_Click({
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "About"
    $aboutForm.Size = New-Object System.Drawing.Size(400, 300)
    $instructionsLabel = New-Object System.Windows.Forms.Label
    $instructionsLabel.Text = "Vulnerability Dashboard Tool`r`n`r`n" +
                             "Created by A1C Ian Velder`r`n`r`n" +
                             "This tool displays vulnerabilities on a device:`r`n" +
                             "1. Enter a search term (NetBIOS or IP Address).`r`n" +
                             "2. Select the desired search type.`r`n" +
                             "3. Click 'Search' to filter the data.`r`n" +
                             "4. View the results in the grid below.`r`n" +
                             "5. Click a row to see detailed information.`r`n`r`n" +
                             "Use the 'Unit' drop-down menu to view a unit's watchlist computers."
    $instructionsLabel.Size = New-Object System.Drawing.Size(350, 200)
    $instructionsLabel.Location = New-Object System.Drawing.Point(20, 20)
    $instructionsLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $aboutForm.Controls.Add($instructionsLabel)
    $aboutForm.ShowDialog()
})

# --- Start Application ---
$form.ShowDialog()
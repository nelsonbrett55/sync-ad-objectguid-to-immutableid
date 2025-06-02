Write-Host -f green @"
  _                           _   _                      ____   _   _   ___                 
 | |       ___     __ _    __| | (_)  _ __     __ _     / ___| | | | | |_ _|                
 | |      / _ \   / _' |  / _' | | | | '_ \   / _' |   | |  _  | | | |  | |                 
 | |___  | (_) | | (_| | | (_| | | | | | | | | (_| |   | |_| | | |_| |  | |   _     _     _ 
 |_____|  \___/   \__,_|  \__,_| |_| |_| |_|  \__, |    \____|  \___/  |___| (_)   (_)   (_)
                                              |___/                                         
"@

# Install the required MSOnline module if not already installed
if (-not (Get-Module -Name MSOnline -ListAvailable)) {
    Install-Module -Name MSOnline -Force
}
# Import the necessary .NET assemblies
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class NativeMethods
{
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
}

"@
function Refresh-ADAzureUserListView {

    # Disable the start button
    $buttonRefresh.Enabled = $false
    $buttonChangeID.Enabled = $false

    # Display Progressbar
    $ListViewChange = $progressBar.Top + $progressBar.Height
    $listView.Top = $ListView.top + $ListViewChange
    $ListView.Height = $ListView.Height - $ListViewChange
    $progressBar.Visible = $true
    $labelStatus.Visible = $true
    $labelStatus.Text = "Progress"
    [System.Windows.Forms.Application]::DoEvents()

    # Clear the ListView
    $listView.Items.Clear()

    # Get all users from on-premises Active Directory
    $labelStatus.text = "Loading AD Users"
    $adUsers = Get-ADUser -Filter * -Properties objectGUID, UserPrincipalName

    # Get all users from Azure Active Directory
    $labelStatus.text = "Progress: Loading Azure Users"
    $azusers = Get-MsolUser

    # Get the total number of users
    $totalUsers = $adUsers.Count

    # Initialize the completed users counter
    $completedUsers = 0

    # Iterate through each AD user
    $listView.BeginUpdate()
    foreach ($user in $adUsers) {
        $backgroundColor = [system.Drawing.Color]::Transparent
        $ForeColor = [system.Drawing.Color]::Black

        # Get the GUID of the user
        $adGUID = $user.objectGUID
        $adGUIDString = [System.Guid]$adGUID
        $adGUIDString = $adGUIDString.ToString()
        
        # Calculate the new ImmutableID
        $newImmutableID = [Convert]::ToBase64String([guid]::New($adGUIDString).ToByteArray())
        
        # Get the user's UserPrincipalName
        $userPrincipalName = $user.UserPrincipalName
        
        if ($userPrincipalName.length -ge 10) {
            # Search for the user in Azure AD using UserPrincipalName
            $azureUser = $azusers | Where-Object { $_.UserPrincipalName -eq $userPrincipalName }

            # If the user is found in Azure AD, retrieve the ImmutableID
            if ($azureUser) {
                $immutableID = $azureUser.ImmutableId
                if ($newImmutableID -eq $ImmutableID) {
                    $ForeColor = [System.Drawing.Color]::DarkGreen
                    $backgroundColor = [System.Drawing.Color]::LightCyan
                }
            } else {
                $immutableID = "User not found in Azure AD"
                $ForeColor = [System.Drawing.Color]::Gray
                $backgroundColor = [System.Drawing.Color]::GhostWhite
            }
        
            if ($immutableID.Length -le 1) {
                $immutableID = " "
            }

            # Add the user information to the ListView
            $item = New-Object System.Windows.Forms.ListViewItem($userPrincipalName)
            $item.SubItems.Add($adGUIDString)
            $item.SubItems.Add($immutableID)
            $item.SubItems.Add($newImmutableID)
            $listView.Items.Add($item) | Out-Null
            $item.BackColor = $backgroundColor
            $item.ForeColor = $ForeColor
        }

        # Increment the completed users counter
        $completedUsers++
        
        # Update the progress bar and label
        $progressBar.Value = ($completedUsers / $totalUsers) * 100
        $labelStatus.Text = "Progress: $($completedUsers) / $($totalUsers)"
        [System.Windows.Forms.Application]::DoEvents()
    }
    $listView.EndUpdate()
    
    # Hide Progressbar
    $progressBar.Value = 0
    $labelStatus.Text = ""
    $progressBar.Visible = $false
    $labelStatus.Visible = $false
    $listView.Top = $ListView.top - $ListViewChange
    $ListView.Height = $ListView.Height + $ListViewChange
    [System.Windows.Forms.Application]::DoEvents()

    # Enable the change ID button
    $buttonRefresh.Enabled = $true
    $buttonRefresh.Text = "Refresh"
    $listView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
    $buttonChangeID.Enabled = $true
    $buttonRefresh.Enabled = $true
}

# Import the MSOnline module
Import-Module MSOnline

Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure AD User Sync"
$form.Size = New-Object System.Drawing.Size(1024, 720)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$form.StartPosition = "CenterScreen"

# Create a label for progress status
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(8, 10)
$labelStatus.Size = New-Object System.Drawing.Size(990, 25)
$labelStatus.Visible = $false
$labelStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor`
[System.Windows.Forms.AnchorStyles]::Right -bor`
[System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($labelStatus)

# Create a progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(8, 40)
$progressBar.Size = New-Object System.Drawing.Size(990, 30)
$progressBar.Minimum = 0
$progressBar.Style = "Continuous"
$progressBar.Visible = $false
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor`
[System.Windows.Forms.AnchorStyles]::Right -bor`
[System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($progressBar)

# Create a ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(8, 8)
$listView.Size = New-Object System.Drawing.Size(990, 628)
$listView.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor`
[System.Windows.Forms.AnchorStyles]::Right -bor`
[System.Windows.Forms.AnchorStyles]::Bottom -bor`
[System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($listView)

# Create columns for the ListView
$listView.Columns.Add("UserPrincipalName", 185) | Out-Null  # Adjusted column size
$listView.Columns.Add("AD GUID", 185) | Out-Null  # Adjusted column size
$listView.Columns.Add("Azure Current ImmutableID", 185) | Out-Null  # Adjusted column size
$listView.Columns.Add("New ImmutableID", 185) | Out-Null  # Adjusted column size

# Event handler for column click
$listView.Add_ColumnClick({
    param($sender, $e)

    $clickedColumn = $listView.Columns[$e.Column]
    if ($clickedColumn.Text.IndexOf("▼") -ne -1 -or $clickedColumn.Text.indexOf("▲") -ne -1){
        $clickedColumn.Text = $clickedColumn.Text -replace " ▲$", "" -replace " ▼$", ""
        # Determine the sort order for the clicked column
        if ($listView.Sorting -eq [System.Windows.Forms.SortOrder]::Ascending) {
            $listView.Sorting = [System.Windows.Forms.SortOrder]::Descending
            $clickedColumn.Text += " ▼"
        } else {
            $listView.Sorting = [System.Windows.Forms.SortOrder]::Ascending
            $clickedColumn.Text += " ▲"
        }
    } else {
        # Clear previous sorting indicators from all columns
        foreach ($column in $listView.Columns) {
            $column.Text = $column.Text -replace " ▲$", "" -replace " ▼$", ""
            $column.Tag = $false
        }
        $listView.Sorting = [System.Windows.Forms.SortOrder]::Ascending
        $clickedColumn.Text += " ▲"

        # Sort the ListView items based on the clicked column
        $listView.ListViewItemSorter = [ListViewItemComparer]::new($e.Column, $listView.Sorting)
    }
    $listView.Sort()
})

# Create a button to connect to Azure
$buttonConnect = New-Object System.Windows.Forms.Button
$buttonConnect.Size = New-Object System.Drawing.Size(150, 30)
$buttonConnect.Top = $form.ClientSize.Height - $buttonConnect.Height - 8
$buttonConnect.Left = (($buttonConnect.Width + 25) * 0) + 8
$buttonConnect.Text = "Connect to Azure"
$buttonConnect.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor`
[System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($buttonConnect)

# Create a button to start the sync
$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Size = New-Object System.Drawing.Size(150, 30)
$buttonRefresh.Top = $form.ClientSize.Height - $buttonRefresh.Height - 8
$buttonRefresh.Left = (($buttonRefresh.Width + 25) * 1) + 8
$buttonRefresh.Text = "Refresh"
$buttonRefresh.Enabled = $false
$buttonRefresh.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor`
[System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($buttonRefresh)

# Create a button to change ImmutableID
$buttonChangeID = New-Object System.Windows.Forms.Button
$buttonChangeID.Size = New-Object System.Drawing.Size(150, 30)
$buttonChangeID.Top = $form.ClientSize.Height - $buttonChangeID.Height - 8
$buttonChangeID.Left = (($buttonChangeID.Width + 25) * 2) + 8
$buttonChangeID.Text = "Set ImmutableID"
$buttonChangeID.Enabled = $false
$buttonChangeID.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor`
[System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($buttonChangeID)

# Create a button to Increase Font Size
$buttonZoomIn = New-Object System.Windows.Forms.Button
$buttonZoomIn.Size = New-Object System.Drawing.Size(100, 30)
$buttonZoomIn.Top = $form.ClientSize.Height - $buttonZoomIn.Height - 8
$buttonZoomIn.Left = $form.Width - ((($buttonZoomIn.Width + 25) * 1) + 8)
$buttonZoomIn.Text = "Zoom In"
$buttonZoomIn.Enabled = $true
$buttonZoomIn.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor`
[System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($buttonZoomIn)

# Create a button to decrese Font Size
$buttonZoomOut = New-Object System.Windows.Forms.Button
$buttonZoomOut.Size = New-Object System.Drawing.Size(100, 30)
$buttonZoomOut.Top = $form.ClientSize.Height - $buttonZoomOut.Height - 8
$buttonZoomOut.Left = $form.Width - ((($buttonZoomOut.Width + 25) * 2) + 8)
$buttonZoomOut.Text = "Zoom Out"
$buttonZoomOut.Enabled = $true
$buttonZoomOut.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor`
[System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($buttonZoomOut)

# Event handler for the connect button
$buttonConnect.Add_Click({
    # Disable the connect button
    $buttonConnect.Enabled = $false
    
    #Display Progressbar
    $ListViewChange = $progressBar.Top + $progressBar.Height
    $listView.Top = $ListView.top + $ListViewChange
    $ListView.Height = $ListView.Height - $ListViewChange
    $progressBar.Visible = $true
    $labelStatus.Visible = $true
    $labelStatus.text = "Progress: Connecting to Microsoft Azure..."
    $progressBar.Value = 50
    [System.Windows.Forms.Application]::DoEvents()

    # Check if already connected to MSOnline
    try{
        Get-MsolDomain -ErrorAction Stop
    }
    catch {
        Connect-MsolService -ErrorAction SilentlyContinue
    }

    
    try{
        Get-MsolDomain -ErrorAction Stop > $null
        # Enable the start button
        $buttonRefresh.Enabled = $true
    }
    catch {
        # Enable the connect button
        $buttonConnect.Enabled = $true
    }

    #Hide Progressbar
    $progressBar.Value = 0
    $labelStatus.Text = ""
    $progressBar.Visible = $false
    $labelStatus.Visible = $false
    $listView.Top = $ListView.top - $ListViewChange
    $ListView.Height = $ListView.Height + $ListViewChange
    Refresh-ADAzureUserListView
    
})

# Event handler for the start button
$buttonRefresh.Add_Click({Refresh-ADAzureUserListView})

# Event handler for the change ID button
$buttonChangeID.Add_Click({
    # Disable the change ID button
    $buttonRefresh.Enabled = $false
    $buttonChangeID.Enabled = $false

    #Display Progressbar
    $ListViewChange = $progressBar.Top + $progressBar.Height
    $listView.Top = $ListView.top + $ListViewChange
    $ListView.Height = $ListView.Height - $ListViewChange
    $progressBar.Visible = $true
    $labelStatus.Visible = $true

    # Get the total number of users
    $totalUsers = $listView.Items.Count

    # Initialize the completed users counter
    $completedUsers = 0

    # Iterate through each item in the ListView
    $listView.BeginUpdate()
    foreach ($item in $listView.Items) {
        # Get the UserPrincipalName and new ImmutableID
        $userPrincipalName = $item.Text
        $ImmutableID = $item.SubItems[2].Text
        $newImmutableID = $item.SubItems[3].Text
        $backgroundColor = [system.Drawing.Color]::Transparent
        $ForeColor = [system.Drawing.Color]::Black

        # Change the ImmutableID to the new value
        if ($ImmutableID -ne "User not found in Azure AD" -and $newImmutableID -ne $ImmutableID) {
            [System.Windows.Forms.Application]::DoEvents()
            $item.SubItems[2].Text = $newImmutableID
            Set-MsolUser -UserPrincipalName $userPrincipalName -ImmutableId $newImmutableID
            [System.Windows.Forms.Application]::DoEvents()
            Write-Host -f green "$("$($userPrincipalName)'s".PadRight(42)) ImmutableID has been updated"
            $ForeColor = [System.Drawing.Color]::White
            $backgroundColor = [System.Drawing.Color]::ForestGreen
        } elseif($newImmutableID -eq $ImmutableID){
            #Write-Host -f cyan "$("$($userPrincipalName)'s".PadRight(42)) ImmutableID was already updated"
            $ForeColor = [System.Drawing.Color]::DarkGreen
            $backgroundColor = [System.Drawing.Color]::LightCyan
        } Elseif ($newImmutableID -ne "User not found in Azure AD"){
            #Write-Host -f yellow "$("$($userPrincipalName)'s".PadRight(42)) isn't in Azure yet"
            $ForeColor = [System.Drawing.Color]::Gray
            $backgroundColor = [System.Drawing.Color]::GhostWhite
        }
        $item.BackColor = $backgroundColor
        $item.ForeColor = $ForeColor

        # Increment the completed users counter
        $completedUsers++
        
        # Update the progress bar and label
        $progressBar.Value = ($completedUsers / $totalUsers) * 100
        $labelStatus.Text = "Progress: $($completedUsers) / $($totalUsers)"
        [System.Windows.Forms.Application]::DoEvents()
    }
    $listView.EndUpdate()
    
    #Hide Progressbar
    $progressBar.Value = 0
    $labelStatus.Text = ""
    $progressBar.Visible = $false
    $labelStatus.Visible = $false
    $listView.Top = $ListView.top - $ListViewChange
    $ListView.Height = $ListView.Height + $ListViewChange
    [System.Windows.Forms.Application]::DoEvents()

    # Enable the change ID button
    $buttonRefresh.Enabled = $true
})

# Event handler for the Zoom In button
$buttonZoomIn.Add_Click({
    If($listView.Font.Size + 2 -le 18){
        $buttonZoomOut.Enabled = $true
        $NewSize = $listView.Font.Size + 2
        $listView.Font = New-Object System.Drawing.Font($ListView.Font.Name, $NewSize)
        If($NewSize -ge 18){
            $buttonZoomIn.Enabled = $false
        }
        $listView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
    }
})

# Event handler for the Zoom Out button
$buttonZoomOut.Add_Click({
    
    If($listView.Font.Size - 2 -ge 6){
        $buttonZoomIn.Enabled = $true
        $NewSize = $listView.Font.Size - 2
        $listView.Font = New-Object System.Drawing.Font($ListView.Font.Name, $NewSize)
        
        If($NewSize -le 6){
            $buttonZoomOut.Enabled = $false
        }
        $listView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
    }
})

# Handle the Resize event of the form
$minimumWidth = 800
$minimumHeight = 400
$form.add_Resize({
    if ($form.Width -lt $minimumWidth) {
        $form.Width = $minimumWidth
    }
    
    if ($form.Height -lt $minimumHeight) {
        $form.Height = $minimumHeight
    }
})

# Hide the console window
$consoleWindowHandle = [NativeMethods]::GetConsoleWindow()
[NativeMethods]::ShowWindow($consoleWindowHandle, 0)

# Show the form
$form.ShowDialog() | Out-Null

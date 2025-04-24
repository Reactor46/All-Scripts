# Import Active Directory module
Import-Module ActiveDirectory

# Get script name (without extension)
$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path)

# Define report directory
$reportDir = "D:\Report Results\$scriptName"

# Create directory if it does not exist
If (!(Test-Path -Path $reportDir)) {
    New-Item -ItemType Directory -Path $reportDir -Force
}

# Define file name with timestamp
$timestamp = Get-Date -Format "ddd_MMM-dd-yyyy_hmm_tt"
$reportFile = "$reportDir\$scriptName`_$timestamp.html"

# Define file paths
$userListPath = "D:\K-Share\SCRIPTS\User-Search.txt"

# Read usernames from the text file
$usernames = Get-Content $userListPath

# Prepare an array to store the results
$report = @()

foreach ($username in $usernames) {
    $username = $username.Trim()

    try {
        # Get AD user object with required properties
        $user = Get-ADUser -Identity $username -Properties LastLogonDate, PasswordLastSet, WhenCreated, Description, Department, Manager, Enabled

        $statusClass = ""
        if (-not $user.Enabled) {
            $statusClass = "disabled-user"  # Assign class for disabled users
        }

        # Fetch Manager's name if available
        $managerName = "N/A"
        if ($user.Manager) {
            $manager = Get-ADUser -Identity $user.Manager -Properties GivenName, Surname
            $managerName = "$($manager.GivenName) $($manager.Surname)"
        }

        # Add user data to the report
        # <td>$($user.LastLogonDate)</td>
        $report += @"
        <tr class='$statusClass'>
            <td>$($user.SamAccountName)</td>
            <td>$($user.LastLogonDate.ToString('MM/dd/yyyy'))</td>
            <td>$($user.PasswordLastSet.ToString('MM/dd/yyyy'))</td>
            <td>$($user.WhenCreated.ToString('MM/dd/yyyy'))</td>
            <td>$($user.Description)</td>
            <td>$($user.Department)</td>
            <td>$managerName</td>
        </tr>
"@
    } catch {
        # User not found
        $report += @"
        <tr class="user-not-found">
            <td>$username</td>
            <td class="empty-cell">N/A</td>
            <td class="empty-cell">N/A</td>
            <td class="empty-cell">N/A</td>
            <td class="empty-cell">N/A</td>
            <td class="empty-cell">N/A</td>
            <td class="empty-cell">N/A</td>
            <td colspan="6">User not found</td>
        </tr>
"@
    }
}

# Generate HTML with CSS and JavaScript for sorting & exporting
$htmlContent = @'
<!DOCTYPE html>
<html>
<head>
    <title>User Account Report</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.3.6/js/dataTables.buttons.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.3.6/js/buttons.html5.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.3.6/js/buttons.print.min.js"></script>

    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.3.6/css/buttons.dataTables.min.css">

    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { text-align: center; }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            margin-top: 20px;
        }
        th, td { 
            padding: 10px; 
            text-align: left; 
            border-bottom: 1px solid #ddd; 
        }
        th { 
            background-color: #f4f4f4; 
            cursor: pointer; 
        }
        td { 
            text-align: left; 
        }
        .dataTables_wrapper .dt-buttons { 
            float: right; 
            margin-bottom: 10px; 
        }

        /* Blinking effect for "User not found" */
        @keyframes blinkRed {
            50% { background-color: red; color: white; }
        }
        .user-not-found td { 
            animation: blinkRed 1s infinite;
        }

        /* Blinking effect for disabled users */
        @keyframes blinkYellow {
            50% { background-color: yellow; color: black; }
        }
        .disabled-user td {
            animation: blinkYellow 1s infinite;
        }

        /* Table styling */
        table.dataTable {
            border-radius: 5px;
            overflow: hidden;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
        }

        /* Additional styling for columns */
        th, td { 
            padding: 12px;
        }

        th {
            text-align: center;
        }

        /* Styling for empty cells */
        .empty-cell {
            color: #888;
        }
    </style>

    <script>
        $(document).ready(function() {
            $('#userTable').DataTable({
                dom: 'Bfrtip',
                paging: true,
                searching: true,
                ordering: true,
                info: true,
                buttons: [
                    { extend: 'csv', text: 'Export CSV' },
                    { extend: 'excel', text: 'Export Excel' },
                    { extend: 'pdf', text: 'Export PDF' },
                    { extend: 'print', text: 'Print' }
                ]
            });
        });
    </script>
</head>
<body>
    <h1>User Account Report</h1>
    <table id="userTable" class="display">
        <thead>
            <tr>
                <th>Username</th>
                <th>Last Logon Date</th>
                <th>Last Password Change</th>
                <th>Created Date</th>
                <th>Description</th>
                <th>Department</th>
                <th>Manager</th>
            </tr>
        </thead>
        <tbody>
'@

# Append report data
$htmlContent += $report

# Close HTML
$htmlContent += @'
        </tbody>
    </table>
</body>
</html>
'@

# Save the file
$htmlContent | Out-File -FilePath $reportFile -Encoding UTF8

Write-Host "Report saved: $reportFile"

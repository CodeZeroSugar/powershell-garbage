# Set console colors: green text on black background
$host.UI.RawUI.ForegroundColor = "Green"
$host.UI.RawUI.BackgroundColor = "Black"
# Clear the screen to apply the background color to the entire window
Clear-Host

# Define the "hacker" text lines, excluding loading bar
$hackerLines = @(
    "AUTHENTICATION MATRIX DECRYPTED",
    "SYSTEM ACCESS GRANTED",
    "Ghosting through the mainframe...",
    "0xFF 0xA3 0x7B - Connection Established",
    "DIR /s",
    "Scanning Network Nodes...",
    "PING 127.0.0.1",
    "Operation Complete. Awaiting Commands."
)

# Define loading bar stages
$loadingBar = @(
    "Loading: [          ]",
    "Loading: [==        ]",
    "Loading: [====      ]",
    "Loading: [======    ]",
    "Loading: [========  ]",
    "Loading: [==========]"
)

# Function to simulate typing
function Write-Typing {
    param ([string]$text)
    foreach ($char in $text.ToCharArray()) {
        Write-Host -NoNewline $char
        Start-Sleep -Milliseconds 50
    }
}

# Function to simulate loading bar updating in place
function Write-LoadingBar {
    foreach ($stage in $loadingBar) {
        Write-Typing $stage
        Start-Sleep -Milliseconds 300
        Write-Host -NoNewline "`r"
    }
    Write-Host ""
}

# Process the hacker text and insert loading bar and commands
foreach ($line in $hackerLines) {
    if ($line -eq "Ghosting through the mainframe...") {
        # Type the line before the loading bar
        Write-Typing $line
        Write-Host ""
        # Show the loading bar
        Write-LoadingBar
    } elseif ($line -eq "DIR /s") {
        # Type and execute DIR /s
        Write-Typing $line
        Write-Host ""
        cmd /c "DIR /s"
        Clear-Host
    } elseif ($line -eq "PING 127.0.0.1") {
        # Type and execute PING
        Write-Typing $line
        Write-Host ""
        cmd /c "PING 127.0.0.1"
    } elseif ($line -eq "Operation Complete. Awaiting Commands.") {
        #Close on any input
        Clear-Host
        Read-Host "Operation Complete. Awaiting Commands"
    } else {
        # Just type and display the line
        Write-Typing $line
        Write-Host ""
    }
}

# Reset console colors (optional)
#$host.UI.RawUI.ForegroundColor = "Gray"
#$host.UI.RawUI.BackgroundColor = "DarkBlue"
#Clear-Host

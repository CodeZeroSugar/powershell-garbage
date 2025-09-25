# Define all the dice frames as an array
$diceFrames = @(
@"
+-------+
| o     |
|       |
|     o |
+-------+
"@,
@"
+-------+
| o   o |
| o   o |
| o   o |
+-------+
"@,
@"
+-------+
|     o |
|   o   |
| o     |
+-------+
"@,
@"
+-------+
|       |
|   o   |
|       |
+-------+
"@,
@"
+-------+
| o   o |
|       |
| o   o |
+-------+
"@,
@"
+-------+
| o   o |
|   o   |
| o   o |
+-------+
"@,
@"
+-------+
| o     |
|       |
|     o |
+-------+
"@,
@"
+-------+
| o   o |
| o   o |
| o   o |
+-------+
"@,
@"
+-------+
|       |
|   o   |
|       |
+-------+
"@,
@"
+-------+
|     o |
|   o   |
| o     |
+-------+
"@,
@"
+-------+
| o   o |
|   o   |
| o   o |
+-------+
"@,
@"
+-------+
| o   o |
|       |
| o   o |
+-------+
"@
)

# Clear the screen before starting animation
Clear-Host

# Loop through each frame with a short delay
foreach ($frame in $diceFrames) {
    Clear-Host
    Write-Host $frame
    Start-Sleep -Milliseconds 300
}

# Add a fun final message
Clear-Host

$diceFrames | Get-Random
Write-Host "`n🎉 The die has landed!"

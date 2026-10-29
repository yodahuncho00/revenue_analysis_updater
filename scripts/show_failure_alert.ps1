param(
    [Parameter(Mandatory = $true)]
    [string]$AlertFile
)

Add-Type -AssemblyName System.Windows.Forms
$Message = Get-Content -LiteralPath $AlertFile -Raw
[System.Windows.Forms.MessageBox]::Show(
    $Message,
    "Revenue Analysis updater needs attention",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error
) | Out-Null

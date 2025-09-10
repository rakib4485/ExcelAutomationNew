# # Create Excel COM object
# $excel = New-Object -ComObject Excel.Application
# $excel.Visible = $false
# $excel.DisplayAlerts = $false

# # Load environment variables from .env
# $envVars = @{}
# Get-Content ".env" | ForEach-Object {
#     if ($_ -match "^\s*#") { return } # skip comments
#     if ($_ -match "^\s*$") { return } # skip empty lines
#     $parts = $_ -split "=", 2
#     $key = $parts[0].Trim()
#     $value = $parts[1].Trim()
#     $envVars[$key] = $value
# }

# # Get the XLSM path and macro name
# $xlsmPath = Resolve-Path $envVars["XLSM_PATH"]
# $macroName = $envVars["MACRO_NAME"]

# # Open workbook and run macro
# $workbook = $excel.Workbooks.Open($xlsmPath)
# $excel.Run($macroName)

# # Save and close
# $workbook.Save()
# $workbook.Close($false)
# $excel.Quit()

# # Clean up COM objects to prevent memory leak
# [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
# [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
# [GC]::Collect()
# [GC]::WaitForPendingFinalizers()

# Write-Host "✅ Macro '$macroName' executed and workbook saved."
# param(
#   [string]$xlsmPath
# )

# $ErrorActionPreference = "Stop"

# try {
#     $excel = New-Object -ComObject Excel.Application
#     $excel.Visible = $false

#     $workbook = $excel.Workbooks.Open($xlsmPath)

#     # Run the macro named refresh_All
#     $excel.Run("refresh_All")

#     $workbook.Save()
#     $workbook.Close($false)
#     $excel.Quit()

#     Write-Output "Macro executed successfully"
# } catch {
#     Write-Error "Error running macro: $_"
#     exit 1
# }

param (
    [string]$xlsmPath,
    [string]$macroName
)

$ErrorActionPreference = "Stop"

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $true

    $workbook = $excel.Workbooks.Open($xlsmPath)

    # Run the macro
    $excel.Run($macroName)

    $workbook.Save()
    $workbook.Close($false)
    $excel.Quit()

    Write-Output "✅ Macro '$macroName' executed successfully."
} catch {
    Write-Error "❌ Error running macro: $_"
    exit 1
} finally {
    if ($excel) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

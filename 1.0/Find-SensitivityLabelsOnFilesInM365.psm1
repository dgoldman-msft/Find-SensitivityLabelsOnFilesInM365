# Dot-source internal helper functions
. (Join-Path $PSScriptRoot "internal\functions\Get-TimeStamp.ps1")
. (Join-Path $PSScriptRoot "internal\functions\Write-ToLogFile.ps1")

# Dot-source public function
. (Join-Path $PSScriptRoot "functions\Find-SensitivityLabelsOnFilesInM365.ps1")

# Export public function and all aliases
Export-ModuleMember -Function Find-SensitivityLabelsOnFilesInM365 `
                    -Alias    Find-SensLabels, Get-FileSensitivityLabels, FSLOF

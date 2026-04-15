@{
    # Module identity
    RootModule        = 'Find-SensitivityLabelsOnFilesInM365.psm1'
    ModuleVersion     = '1.0'
    GUID              = 'eb3879c1-fd95-4e04-b831-c11a00ca89ee'
    Author            = 'Dave Goldman'
    CompanyName       = ' '
    Copyright         = '(c) Dave Goldman. All rights reserved.'

    # Description
    Description       = 'Searches Microsoft Purview Content Explorer to report which sensitivity labels are applied to a given list of files across SharePoint Online, OneDrive for Business, Exchange Online, and Microsoft Teams workloads.'

    # Minimum PowerShell version required
    PowerShellVersion = '7.1'

    # Required modules — ExchangeOnlineManagement is checked/installed at runtime via -ConnectIPPS
    RequiredModules   = @()

    # Format file
    FormatsToProcess  = @('xml\Find-SensitivityLabelsOnFilesInM365.Format.ps1xml')

    # Exports — explicit lists override wildcard behaviour
    FunctionsToExport = @('Find-SensitivityLabelsOnFilesInM365')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @('FSLOF')

    # Private data
    PrivateData       = @{
        PSData = @{
            Tags         = @('Purview', 'SensitivityLabels', 'SharePoint', 'OneDrive', 'Compliance', 'ContentExplorer', 'MicrosoftPurview')
            LicenseUri   = 'https://github.com/dgoldman-msft/Find-SensitivityLabelsOnFilesInM365/blob/main/LICENSE'
            ProjectUri   = 'https://github.com/dgoldman-msft/Find-SensitivityLabelsOnFilesInM365'
            ReleaseNotes = 'Initial release.'
        }
    }
}

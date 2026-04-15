function Find-SensitivityLabelsOnFilesInM365 {
    <#
    .SYNOPSIS
        Searches for sensitivity labels on files in Microsoft 365 Workloads.

    .DESCRIPTION
        This function searches for sensitivity labels applied to files in SharePoint Online (SPO), OneDrive for Business (ODB),
        Exchange Online (EXO), and Microsoft Teams workloads. It can automatically discover available sensitivity labels from
        the current IPPS session or use manually specified labels. Results can be exported to CSV format.

    .PARAMETER FileLocation
        Specifies the path to a text file containing file names or URLs to search for (one per line).
        If not specified, defaults to ".\files.txt".

    .PARAMETER TargetFiles
        Array of file names or URLs to search for directly, without reading from a file.
        Cannot be used together with FileLocation parameter.

    .PARAMETER Labels
        Array of sensitivity label names to search for. Label names must match exactly as shown in
        Purview UI. If not specified, the function will attempt to retrieve labels from the current
        IPPS session.

    .PARAMETER Workloads
        Specifies which Microsoft 365 workloads to search. Valid values are "SPO" (SharePoint Online),
        "ODB" (OneDrive for Business), "EXO" (Exchange Online), and "Teams" (Microsoft Teams).
        Multiple workloads are scanned sequentially.
        Defaults to "EXO".

    .PARAMETER PageSize
        Number of records to retrieve per page from Content Explorer. Default is 1000.
        Larger values may improve performance but could cause timeouts.

    .PARAMETER ExportResults
        When specified, exports results to a CSV file per workload (e.g. EXO_Results.csv, SPO_Results.csv)
        written into the directory specified by -LogDirectory.

    .PARAMETER ExportPath
        Reserved for future use. Currently not used.

    .PARAMETER ConnectIPPS
        When specified, automatically imports ExchangeOnlineManagement module and connects to
        IPPS session if not already connected.

    .PARAMETER UseCachedLabels
        When specified, retrieves available sensitivity labels from the current IPPS session via Get-Label.
        Labels with ContentType 'None' (parent/container labels) are excluded.
        If no labels are found or the cmdlet fails, the function terminates with an error.

    .PARAMETER LogDirectory
        Specifies the full path to the directory where log files will be written.
        A file named Logging.txt will be created or appended to inside this directory.
        Defaults to a subdirectory named 'Find-SensitivityLabelsOnFilesInM365' inside the system temp folder ($env:TEMP).

    .EXAMPLE
        Find-SensitivityLabelsOnFilesInM365 -FileLocation ".\myfiles.txt" -Workloads EXO -UseCachedLabels -ExportResults

        Reads file names from myfiles.txt, retrieves labels from the active IPPS session, searches Exchange Online, and exports results to CSV.

    .EXAMPLE
        Find-SensitivityLabelsOnFilesInM365 -TargetFiles "document1.docx", "report.pdf" -Labels "Confidential", "Public" -Workloads EXO

        Searches for specific files with specific sensitivity labels in Exchange Online.

    .EXAMPLE
        Find-SensitivityLabelsOnFilesInM365 -FileLocation ".\myfiles.txt" -UseCachedLabels -Workloads SPO, ODB, EXO, Teams -ExportResults

        Scans all four workloads sequentially using labels from the current session. Exports one CSV per workload into the log directory.

    .NOTES
        Requires connection to Security & Compliance PowerShell (Connect-IPPSSession).
        The Export-ContentExplorerData cmdlet must be available.

    .LINK
        https://docs.microsoft.com/en-us/powershell/module/exchange/export-contentexplorerdata
    #>

    [CmdletBinding(DefaultParameterSetName = 'FileInput')]
    [Alias('FSLOF')]
    param(
        [Parameter(ParameterSetName = 'FileInput', Position = 0)]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                Write-Host "ERROR: File path '$_' does not exist." -ForegroundColor Red
                return $false
            }
            return $true
        })]
        [string]$FileLocation = ".\files.txt",

        [Parameter(ParameterSetName = 'DirectInput', Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$TargetFiles,

        [Parameter()]
        [string[]]$Labels,

        [Parameter()]
        [ValidateSet("SPO", "ODB", "EXO", 'Teams')]
        [string[]]$Workloads = @("EXO"),

        [Parameter()]
        [ValidateRange(1, 5000)]
        [int]$PageSize = 1000,

        [Parameter()]
        [switch]$ExportResults,

        [Parameter()]
        [string]$ExportPath = ".\FileLabelResults.csv",

        [Parameter()]
        [switch]$ConnectIPPS,

        [Parameter()]
        [switch]$UseCachedLabels,

        [Parameter()]
        [string]$LogDirectory = (Join-Path $env:TEMP 'Find-SensitivityLabelsOnFilesInM365')
    )

    begin {
        # Ensure log directory exists before any logging
        if (-not (Test-Path -Path $LogDirectory)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
        }

        $separator = "$(Get-TimeStamp) " + ("-" * 80)
        Write-ToLogFile -StringObject $separator -LogDirectory $LogDirectory
        Write-Verbose "Starting Find-SensitivityLabelsOnFiles function"
        Write-ToLogFile -StringObject "$(Get-TimeStamp) Starting Find-SensitivityLabelsOnFiles" -LogDirectory $LogDirectory

        # Auto-connect to IPPS if requested
        if ($ConnectIPPS) {
            Write-Verbose "Auto-connecting to IPPS session"
            Write-ToLogFile -StringObject "$(Get-TimeStamp) Checking for ExchangeOnlineManagement module" -LogDirectory $LogDirectory
            try {
                if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
                    Write-ToLogFile -StringObject "$(Get-TimeStamp) ExchangeOnlineManagement not found. Installing from PSGallery..." -LogDirectory $LogDirectory
                    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                    Write-ToLogFile -StringObject "$(Get-TimeStamp) ExchangeOnlineManagement installed successfully" -LogDirectory $LogDirectory
                }
                else {
                    Write-ToLogFile -StringObject "$(Get-TimeStamp) ExchangeOnlineManagement module found" -LogDirectory $LogDirectory
                }
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                Write-ToLogFile -StringObject "$(Get-TimeStamp) ExchangeOnlineManagement imported" -LogDirectory $LogDirectory
                Connect-IPPSSession -ErrorAction Stop
                Write-Verbose "Successfully connected to IPPS session"
                Write-ToLogFile -StringObject "$(Get-TimeStamp) Successfully connected to IPPS session" -LogDirectory $LogDirectory
            }
            catch {
                Write-ToLogFile -StringObject "$(Get-TimeStamp) ERROR: Failed to connect to IPPS session: $($_.Exception.Message)" -LogDirectory $LogDirectory
                Write-Host "ERROR: Failed to connect to IPPS session: $($_.Exception.Message)" -ForegroundColor Red
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.InvalidOperationException]::new("Failed to connect to IPPS session."),
                        'IPPSConnectionFailed',
                        [System.Management.Automation.ErrorCategory]::ConnectionError,
                        $null
                    )
                )
            }
        }

        # Validate that the required IPPS cmdlets are available regardless of how the session was established
        Write-Verbose "Validating IPPS session..."
        if (-not (Get-Command -Name 'Export-ContentExplorerData' -ErrorAction SilentlyContinue)) {
            Write-ToLogFile -StringObject "$(Get-TimeStamp) ERROR: Export-ContentExplorerData cmdlet not found. No active IPPS session detected." -LogDirectory $LogDirectory
            Write-Host "ERROR: No active Security & Compliance PowerShell session detected. Please connect first using Connect-IPPSSession, or run with -ConnectIPPS to connect automatically." -ForegroundColor Red
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.InvalidOperationException]::new("No active IPPS session detected."),
                    'NoIPPSSession',
                    [System.Management.Automation.ErrorCategory]::ConnectionError,
                    $null
                )
            )
        }
        Write-Verbose "IPPS session validated successfully"
        Write-ToLogFile -StringObject "$(Get-TimeStamp) IPPS session validated" -LogDirectory $LogDirectory

        # Get target files based on parameter set
        if ($PSCmdlet.ParameterSetName -eq 'FileInput') {
            Write-Verbose "Reading target files from: $FileLocation"
            Write-ToLogFile -StringObject "$(Get-TimeStamp) Reading target files from: $FileLocation" -LogDirectory $LogDirectory
            $Targets = Get-Content $FileLocation | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        }
        else {
            Write-Verbose "Using directly specified target files"
            Write-ToLogFile -StringObject "$(Get-TimeStamp) Using directly specified target files" -LogDirectory $LogDirectory
            $Targets = $TargetFiles
        }

        if (-not $Targets -or $Targets.Count -eq 0) {
            Write-ToLogFile -StringObject "$(Get-TimeStamp) ERROR: No target files specified or found in input file" -LogDirectory $LogDirectory
            Write-Host "ERROR: No target files specified or found in input file." -ForegroundColor Red
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.InvalidOperationException]::new('No target files specified or found in input file.'),
                    'NoTargetFiles',
                    [System.Management.Automation.ErrorCategory]::InvalidArgument,
                    $null
                )
            )
        }

        Write-Verbose "Found $($Targets.Count) target files to search for"
        Write-ToLogFile -StringObject "$(Get-TimeStamp) Found $($Targets.Count) target file(s) to search for" -LogDirectory $LogDirectory

        # Get sensitivity labels
        if ($UseCachedLabels -or (-not $Labels)) {
            Write-Verbose "Retrieving sensitivity labels from current IPPS session"
            $AvailableLabels = $null
            try {
                $allLabels = Get-Label
                # Build a GUID-to-DisplayName lookup so sublabels can resolve their parent name
                $labelById = @{}
                foreach ($l in $allLabels) {
                    $labelById[$l.Guid.ToString()] = $l.DisplayName
                }
                # Build full label paths for sublabels (e.g. "Highly Confidential/All Employees")
                $AvailableLabels = $allLabels | Where-Object { $_.ContentType -ne 'None' } | ForEach-Object {
                    $parentIdStr = if ($_.ParentId) { $_.ParentId.ToString() } else { $null }
                    if ($parentIdStr -and
                        $parentIdStr -ne '00000000-0000-0000-0000-000000000000' -and
                        $labelById.ContainsKey($parentIdStr)) {
                        "$($labelById[$parentIdStr])/$($_.DisplayName)"
                    }
                    else {
                        $_.DisplayName
                    }
                }
            }
            catch {
                Write-ToLogFile -StringObject "$(Get-TimeStamp) ERROR: Could not retrieve sensitivity labels: $($_.Exception.Message)" -LogDirectory $LogDirectory
                Write-Host "ERROR: Could not retrieve sensitivity labels from the IPPS session: $($_.Exception.Message)" -ForegroundColor Red
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.InvalidOperationException]::new("Could not retrieve sensitivity labels from the IPPS session."),
                        'LabelRetrievalFailed',
                        [System.Management.Automation.ErrorCategory]::NotSpecified,
                        $null
                    )
                )
            }

            if ($AvailableLabels) {
                $Labels = $AvailableLabels
                Write-Verbose "Found $($Labels.Count) sensitivity labels in session"
                Write-ToLogFile -StringObject "$(Get-TimeStamp) Retrieved $($Labels.Count) sensitivity label(s) from IPPS session" -LogDirectory $LogDirectory
            }
            else {
                Write-ToLogFile -StringObject "$(Get-TimeStamp) ERROR: No sensitivity labels found in the current IPPS session." -LogDirectory $LogDirectory
                Write-Host "ERROR: No sensitivity labels were found in the current IPPS session. Ensure your account has the Content Explorer List Viewer or Content Explorer Content Viewer role in Microsoft Purview." -ForegroundColor Red
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.InvalidOperationException]::new("No sensitivity labels found in the current IPPS session."),
                        'NoSensitivityLabels',
                        [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                        $null
                    )
                )
            }
        }

        Write-Verbose "Using sensitivity labels: $($Labels -join ', ')"
        Write-Verbose "Searching workloads: $($Workloads -join ', ')"
        Write-ToLogFile -StringObject "$(Get-TimeStamp) Labels to search: $($Labels -join ', ')" -LogDirectory $LogDirectory
        Write-ToLogFile -StringObject "$(Get-TimeStamp) Workloads to search: $($Workloads -join ', ')" -LogDirectory $LogDirectory
    }

    process {
        $hits = New-Object System.Collections.Generic.List[object]

        foreach ($workload in $Workloads) {
                Write-Verbose "Processing workload: $workload"
                Write-ToLogFile -StringObject "$(Get-TimeStamp) *** Starting scan of workload: $workload ***" -LogDirectory $LogDirectory -ForegroundColor Cyan

                foreach ($label in $Labels) {
                    Write-Verbose "Searching for label '$label' in workload '$workload'"
                    Write-ToLogFile -StringObject "$(Get-TimeStamp) Scanning label '$label' in workload '$workload'" -LogDirectory $LogDirectory

                    $pageCookie = $null
                    $morePages = $true
                    $pageCount = 0

                    while ($morePages) {
                        $pageCount++
                        Write-Verbose "Processing page $pageCount for label '$label'"
                        Write-Verbose "Retrieving page $pageCount for label '$label' in workload '$workload'"

                        try {
                            $resp = Export-ContentExplorerData `
                                -TagType "Sensitivity" `
                                -TagName $label `
                                -Workload $workload `
                                -PageSize $PageSize `
                                -PageCookie $pageCookie `
                                -ErrorAction Stop 3>$null
                        }
                        catch {
                            Write-Warning "Error retrieving data for label '$label' in workload '$workload': $($_.Exception.Message)"
                            Write-ToLogFile -StringObject "$(Get-TimeStamp) ERROR: Failed retrieving data for label '$label' in workload '$workload': $($_.Exception.Message)" -LogDirectory $LogDirectory
                            break
                        }

                        # Per documentation: metadata is item 0; records are from item 1 onward
                        $meta = $resp[0]
                        $records = @()
                        if ($resp.Count -gt 1) {
                            $records = $resp[1..($resp.Count - 1)]
                        }

                        Write-Verbose "Found $($records.Count) records in this page"
                        Write-Verbose "Page $pageCount returned $($records.Count) record(s) for label '$label' in workload '$workload'"

                        foreach ($r in $records) {
                            # Robust match: search across all fields by JSON-serializing the record
                            $json = $r | ConvertTo-Json -Depth 30 -Compress

                            foreach ($t in $Targets) {
                                if ($json -like "*$t*") {
                                    $hits.Add([pscustomobject]@{
                                        Target   = $t
                                        Label    = $label
                                        Workload = $workload
                                        Record   = $r
                                    })
                                    Write-Verbose "Found match: '$t' with label '$label' in workload '$workload'"
                                    Write-Verbose "$(Get-TimeStamp) Match found: '$t' | Label: '$label' | Workload: '$workload'"
                                    Add-Content -Path (Join-Path $LogDirectory 'Logging.txt') -Value "$(Get-TimeStamp) Match found: '$t' | Label: '$label' | Workload: '$workload'" -Encoding UTF8
                                }
                            }
                        }

                        # Early stop for this label/workload if we've already matched all targets
                        $matched = $hits |
                            Where-Object { $_.Label -eq $label -and $_.Workload -eq $workload } |
                            Select-Object -ExpandProperty Target -Unique

                        if ($matched.Count -ge $Targets.Count) {
                            Write-Verbose "All targets found for label '$label' in workload '$workload', stopping early"
                            Write-ToLogFile -StringObject "$(Get-TimeStamp) All targets matched for label '$label' in workload '$workload' - stopping early" -LogDirectory $LogDirectory
                            break
                        }

                        $pageCookie = $meta.PageCookie
                        $morePages = [bool]$meta.MorePagesAvailable

                        if (-not $morePages) {
                            Write-Verbose "No more pages available for label '$label' in workload '$workload'"
                        }
                    }
                }

                $wlMatchedTargets = ($hits | Where-Object { $_.Workload -eq $workload } | Select-Object -ExpandProperty Target -Unique)
                $wlMatchCount = $wlMatchedTargets.Count
                Write-ToLogFile -StringObject "$(Get-TimeStamp) *** Completed scan of workload: $workload | Matches so far: $($hits.Count) ***" -LogDirectory $LogDirectory -ForegroundColor Green
                if ($wlMatchCount -gt 0 -and $ExportResults) {
                    $wlCsvPath = Join-Path $LogDirectory "${workload}_Results.csv"
                    Write-ToLogFile -StringObject "$(Get-TimeStamp) [$workload] $wlMatchCount target(s) with matches found - results will be written to: $wlCsvPath" -LogDirectory $LogDirectory -ForegroundColor Green
                }
        }

        Write-Verbose "Search complete. Found $($hits.Count) total matches"
        Write-ToLogFile -StringObject "$(Get-TimeStamp) Search complete. Found $($hits.Count) total match(es) across all labels and workloads" -LogDirectory $LogDirectory
    }

    end {
        # Generate report (one row per input target)
        Write-Verbose "Generating report"
        $rowNum = 0
        $report = $Targets | ForEach-Object {
            $rowNum++
            $t = $_
            $rows = $hits | Where-Object { $_.Target -eq $t }

            [pscustomobject]@{
                PSTypeName        = 'SensitivityLabelReport'
                FileNumber        = $rowNum
                Target            = $t
                Labels            = (($rows | Select-Object -ExpandProperty Label -Unique) -join "; ")
                Workload          = (($rows | Select-Object -ExpandProperty Workload -Unique) -join "; ")
                MatchCount        = $rows.Count
                HasSensivityLabel = $rows.Count -gt 0
            }
        }

        # Export if requested — one CSV per workload, only when matches exist
        if ($ExportResults) {
            foreach ($wl in $Workloads) {
                $wlRowNum = 0
                $wlReport = $Targets | ForEach-Object {
                    $wlRowNum++
                    $t = $_
                    $rows = $hits | Where-Object { $_.Target -eq $t -and $_.Workload -eq $wl }
                    [pscustomobject]@{
                        PSTypeName        = 'SensitivityLabelReport'
                        FileNumber        = $wlRowNum
                        Target            = $t
                        Labels            = (($rows | Select-Object -ExpandProperty Label -Unique) -join "; ")
                        Workload          = $wl
                        MatchCount        = $rows.Count
                        HasSensivityLabel = $rows.Count -gt 0
                    }
                }
                $wlMatches = $wlReport | Where-Object { $_.HasSensivityLabel }
                if ($wlMatches) {
                    $wlPath = Join-Path $LogDirectory "${wl}_Results.csv"
                    Write-ToLogFile -StringObject "$(Get-TimeStamp) [$wl] $($wlMatches.Count) match(es) found - writing results to: $wlPath" -LogDirectory $LogDirectory
                    try {
                        $wlReport | Export-Csv $wlPath -NoTypeInformation -ErrorAction Stop
                        Write-Host "[$wl] Results exported to: $wlPath" -ForegroundColor Green
                    }
                    catch {
                        Write-Error "Failed to export [$wl] results: $($_.Exception.Message)"
                        Write-ToLogFile -StringObject "$(Get-TimeStamp) ERROR: Failed to export [$wl] results to '$wlPath': $($_.Exception.Message)" -LogDirectory $LogDirectory
                    }
                }
                else {
                    Write-ToLogFile -StringObject "$(Get-TimeStamp) [$wl] No matches found - skipping CSV export" -LogDirectory $LogDirectory
                }
            }
        }

        Write-ToLogFile -StringObject "$(Get-TimeStamp) Find-SensitivityLabelsOnFiles completed" -LogDirectory $LogDirectory
        Write-ToLogFile -StringObject $separator -LogDirectory $LogDirectory
        Write-Host "Log file written to: $(Join-Path $LogDirectory 'Logging.txt')" -ForegroundColor Cyan

        # Only output the results table if at least one match was found
        $matched = $report | Where-Object { $_.HasSensivityLabel }
        if ($matched) {
            $report
        }
        else {
            Write-Host "No sensitivity label matches found for the specified files." -ForegroundColor Yellow
        }
    }
}


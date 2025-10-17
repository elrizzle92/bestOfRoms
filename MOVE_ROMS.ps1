# ====================================================================
# SCRIPT METADATA AND CONFIGURATION FILE DEFINITION
# ====================================================================

# This file stores your persistent configuration using a secure XML format (Clixml)
$ConfigFile = Join-Path -Path $PSScriptRoot -ChildPath "MOVE_ROMS_config.xml"

# ====================================================================
# CONFIGURATION AND MENU FUNCTIONS
# ====================================================================

function Get-UserPath {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [Parameter(Mandatory=$true)][string]$Type, # 'Directory', 'File', or 'Destination'
        [Parameter(Mandatory=$false)][switch]$MustExist,
        [Parameter(Mandatory=$false)][switch]$IsOptional # NEW: Allows blank input (just Enter)
    )
    
    $Path = ""
    do {
        $Path = Read-Host "$Prompt"
        
        # 1. Check for empty input
        if ([string]::IsNullOrWhiteSpace($Path)) {
            if ($IsOptional) {
                return "" # New Feature: Return empty string if optional and blank
            } else {
                Write-Warning "Input is mandatory for first-time setup. Please enter a valid path."
                continue
            }
        }
        
        # 2. Path normalization
        # Removes trailing backslash if present (except for root drives like C:\)
        if ($Path -match '\\$' -and $Path.Length -gt 3) { $Path = $Path.Substring(0, $Path.Length - 1) }

        # 3. Check existence if required (only runs if input was provided)
        if ($MustExist) {
            $Exists = $false
            if ($Type -eq 'File' -and (Test-Path -Path $Path -PathType Leaf)) {
                $Exists = $true
            }
            elseif ($Type -eq 'Directory' -and (Test-Path -Path $Path -PathType Container)) {
                $Exists = $true
            }
            
            if (-not $Exists) {
                Write-Warning "Path not found or is not the correct type ($Type). Please check the path and try again."
                continue
            }
        }
        
        # 4. Valid input provided and passed checks (or checks were skipped for Destination)
        return $Path
        
    } while ($true)
}

function Save-Configuration {
    param(
        [Parameter(Mandatory=$true)][PSObject]$Config
    )
    $Config | Export-Clixml -Path $ConfigFile
    Write-Host "`nConfiguration saved to $ConfigFile" -ForegroundColor Green
}

function Get-ScriptConfiguration {
    if (Test-Path -Path $ConfigFile) {
        try {
            $Config = Import-Clixml -Path $ConfigFile -ErrorAction Stop
            Write-Host "Configuration loaded successfully from $ConfigFile" -ForegroundColor Green
            return $Config
        }
        catch {
            Write-Warning "Error loading configuration file. Running initial setup."
            Remove-Item $ConfigFile -Force -ErrorAction SilentlyContinue
            return Initial-Setup
        }
    } else {
        Write-Host "Configuration file not found. Starting first-run setup." -ForegroundColor Yellow
        return Initial-Setup
    }
}

function Initial-Setup {
    Write-Host "`n--------------------------------------------------------"
    Write-Host "INITIAL SCRIPT SETUP (First Run)" -ForegroundColor Green
    Write-Host "Please specify all paths and thresholds." -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------"
    
    $Config = [PSCustomObject]@{
        SourceDirectory = ""
        DestinationDirectory = ""
        GameListFile = ""
        SimilarityThreshold = 0.70 # Default
        SecondaryThreshold = 0.50  # Default
    }
    
    # Paths are MANDATORY during initial setup (IsOptional is not passed)
    Write-Host "`n-- PATH SETUP --"
    $Config.SourceDirectory = Get-UserPath -Prompt "1. Enter the FULL PATH for your ROM Source Directory (must exist):" -Type 'Directory' -MustExist
    $Config.DestinationDirectory = Get-UserPath -Prompt "2. Enter the FULL PATH for your ROM Destination Directory (will be created):" -Type 'Destination'
    $Config.GameListFile = Get-UserPath -Prompt "3. Enter the FULL PATH for your Game List File (must exist):" -Type 'File' -MustExist

    $Config = Change-Thresholds -Config $Config -InitialRun $true

    Save-Configuration -Config $Config
    return $Config
}

function Change-Paths {
    param(
        [Parameter(Mandatory=$true)][PSObject]$Config,
        [Parameter(Mandatory=$false)][switch]$InitialRun
    )
    
    $CurrentSource = $Config.SourceDirectory
    $CurrentDest = $Config.DestinationDirectory
    $CurrentList = $Config.GameListFile
    
    Write-Host "`n--------------------------------------------------------"
    Write-Host "CHANGE PATHS" -ForegroundColor Green
    
    # --- Source Directory ---
    Write-Host "Current Source: $CurrentSource" -ForegroundColor Yellow
    $NewSource = Get-UserPath -Prompt "1. Enter the FULL PATH for your ROM Source Directory (must exist, **leave blank to keep current**):" -Type 'Directory' -MustExist -IsOptional
    if (-not [string]::IsNullOrWhiteSpace($NewSource)) {
        $Config.SourceDirectory = $NewSource
        Write-Host "Source Directory updated to: $NewSource" -ForegroundColor Green
    } else {
        Write-Host "Source Directory path unchanged." -ForegroundColor Gray
    }
    
    # --- Destination Directory ---
    Write-Host "Current Dest.: $CurrentDest" -ForegroundColor Yellow
    $NewDestination = Get-UserPath -Prompt "2. Enter the FULL PATH for your ROM Destination Directory (will be created, **leave blank to keep current**):" -Type 'Destination' -IsOptional
    if (-not [string]::IsNullOrWhiteSpace($NewDestination)) {
        $Config.DestinationDirectory = $NewDestination
        Write-Host "Destination Directory updated to: $NewDestination" -ForegroundColor Green
    } else {
        Write-Host "Destination Directory path unchanged." -ForegroundColor Gray
    }
    
    # --- Game List File ---
    Write-Host "Current List: $CurrentList" -ForegroundColor Yellow
    $NewList = Get-UserPath -Prompt "3. Enter the FULL PATH for your Game List File (must exist, **leave blank to keep current**):" -Type 'File' -MustExist -IsOptional
    if (-not [string]::IsNullOrWhiteSpace($NewList)) {
        $Config.GameListFile = $NewList
        Write-Host "Game List File updated to: $NewList" -ForegroundColor Green
    } else {
        Write-Host "Game List File path unchanged." -ForegroundColor Gray
    }
    
    Save-Configuration -Config $Config
    return $Config
}

function Change-Thresholds {
    param(
        [Parameter(Mandatory=$true)][PSObject]$Config,
        [Parameter(Mandatory=$false)][switch]$InitialRun
    )
    
    Write-Host "`n--------------------------------------------------------"
    Write-Host "CHANGE THRESHOLDS" -ForegroundColor Green
    Write-Host "We recommend starting the main scan at .70 (70%) and the backup scan at .50 (50%)." -ForegroundColor Yellow
    
    do {
        $CurrentPrimary = if($InitialRun) {"(Default: 0.70)"} else {$Config.SimilarityThreshold}
        $Input = Read-Host "1. Enter the PRIMARY similarity threshold (Current: $CurrentPrimary):"
        if ([string]::IsNullOrWhiteSpace($Input)) {
            Write-Host "Primary threshold unchanged." -ForegroundColor Gray
            break
        }
        
        if ($Input -as [double] -and [double]$Input -gt 0 -and [double]$Input -le 1) {
            $Config.SimilarityThreshold = [double]$Input
            break
        } else {
            Write-Warning "Invalid input. Please enter a decimal number between 0 and 1 (e.g., 0.70)."
        }
    } while ($true)
    
    do {
        $CurrentSecondary = if($InitialRun) {"(Default: 0.50)"} else {$Config.SecondaryThreshold}
        $Input = Read-Host "2. Enter the SECONDARY similarity threshold (Current: $CurrentSecondary):"
        if ([string]::IsNullOrWhiteSpace($Input)) {
            Write-Host "Secondary threshold unchanged." -ForegroundColor Gray
            break
        }
        
        $InputDouble = $Input -as [double]
        
        if ($InputDouble -and $InputDouble -gt 0 -and $InputDouble -le 1) {
            if ($InputDouble -lt $Config.SimilarityThreshold) {
                $Config.SecondaryThreshold = $InputDouble
                break
            } else {
                Write-Warning "Secondary threshold must be lower than the primary threshold ($($Config.SimilarityThreshold))."
            }
        } else {
            Write-Warning "Invalid input. Please enter a decimal number between 0 and 1 (e.g., 0.50)."
        }
    } while ($true)

    if (-not $InitialRun) { Save-Configuration -Config $Config }
    return $Config
}

function Show-Menu {
    $Config = Get-ScriptConfiguration
    
    do {
        Write-Host "`n--------------------------------------------------------"
        Write-Host "ROM SORTER MAIN MENU" -ForegroundColor Yellow
        Write-Host "--------------------------------------------------------"
        Write-Host "Current Source: $($Config.SourceDirectory)"
        Write-Host "Current Dest.:  $($Config.DestinationDirectory)"
        Write-Host "Current List:   $($Config.GameListFile)"
        Write-Host "Initial Threshold: $($Config.SimilarityThreshold * 100)%"
        Write-Host "Secondary Threshold: $($Config.SecondaryThreshold * 100)%"
        Write-Host "`n[ Log Folder: $PSScriptRoot\results ]" -ForegroundColor Gray
        Write-Host "--------------------------------------------------------"
        Write-Host "1) Run script (Start Copy Process)"
        Write-Host "2) Change configuration paths (Leave blank to keep current)"
        Write-Host "3) Change similarity thresholds (Leave blank to keep current)"
        Write-Host "4) Quit script"
        Write-Host "--------------------------------------------------------"
        
        $Choice = Read-Host "Enter your choice (1, 2, 3, or 4)"
        
        switch ($Choice) {
            "1" { Start-RomSorting -Config $Config; break }
            "2" { $Config = Change-Paths -Config $Config }
            "3" { $Config = Change-Thresholds -Config $Config }
            "4" { Write-Host "Exiting script."; return }
            default { Write-Warning "Invalid choice. Please enter 1, 2, 3, or 4." }
        }
        
    } while ($Choice -ne "1")
}

# ====================================================================
# MAIN SCRIPT EXECUTION BLOCK
# ====================================================================

function Start-RomSorting {
    param(
        [Parameter(Mandatory=$true)][PSObject]$Config
    )
    
    $SourceDirectory = $Config.SourceDirectory
    $DestinationDirectory = $Config.DestinationDirectory
    $GameListFile = $Config.GameListFile
    $SimilarityThreshold = $Config.SimilarityThreshold
    $SecondaryThreshold = $Config.SecondaryThreshold

    # --- Output Logging Setup ---
    $Identifier = @(
        [char]([System.Random]::new().Next(0, 26) + [byte][char]'a'),
        [char]([System.Random]::new().Next(0, 26) + [byte][char]'a')
    ) -join ""
    $LogFileName = "results_$Identifier.txt"
    $LogDirectory = Join-Path -Path $PSScriptRoot -ChildPath "results"

    if (-not (Test-Path -Path $LogDirectory)) {
        Write-Host "Creating results folder: $LogDirectory"
        New-Item -Path $LogDirectory -ItemType Directory | Out-Null
    }

    $LogFilePath = Join-Path -Path $LogDirectory -ChildPath $LogFileName
    Start-Transcript -Path $LogFilePath -Append
    Write-Host "Transcript started. Output being saved to: $LogFilePath"
    # ----------------------------------------

    # --- Timer Setup ---
    $StartTime = Get-Date

    # --- Destination Directory Setup & Permissions ---
    try {
        if (-not (Test-Path -Path $DestinationDirectory)) { New-Item -Path $DestinationDirectory -ItemType Directory | Out-Null }
        $acl = Get-Acl $DestinationDirectory -ErrorAction Stop
        $permission = "Everyone","FullControl","Allow"
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
        $acl.SetAccessRule($rule)
        $acl | Set-Acl $DestinationDirectory -ErrorAction Stop
        Write-Host "Confirmed full write permissions on destination directory."
    } catch {
        Write-Warning "Could not modify permissions on $DestinationDirectory. Run script as Administrator if copies fail."
    }
    
    # --- CUSTOM FUNCTION: String Similarity (Fuzzy Matching) ---
    function Get-StringSimilarity {
        param( [string]$String1, [string]$String2 )
        $s = $String1.ToLower(); $t = $String2.ToLower(); $n = $s.Length; $m = $t.Length
        if ($n -eq 0 -or $m -eq 0) { return 0 }
        
        $d = @()
        for ($i = 0; $i -le $n; $i++) { 
            $row = @(); for ($j = 0; $j -le $m; $j++) { $row += 0 }
            $d += ,$row
        }

        for ($i = 0; $i -le $n; $i++) { $d[$i][0] = $i }
        for ($j = 0; $j -le $m; $j++) { $d[0][$j] = $j }

        for ($i = 1; $i -le $n; $i++) {
            for ($j = 1; $j -le $m; $j++) {
                $cost = if ($s[$i-1] -eq $t[$j-1]) { 0 } else { 1 }
                $deletion = $d[$i-1][$j] + 1
                $insertion = $d[$i][$j-1] + 1
                $substitution = $d[$i-1][$j-1] + $cost
                $d[$i][$j] = [math]::Min([math]::Min($deletion, $insertion), $substitution)
            }
        }
        $distance = $d[$n][$m]
        $maxLength = [math]::Max($n, $m)
        return 1 - ($distance / $maxLength)
    }

    # --- ROM SCAN LOGIC ---
    function Run-RomScan {
        param(
            [Parameter(Mandatory=$true)][string[]]$GameList,
            [Parameter(Mandatory=$true)][double]$Threshold,
            [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$AllRoms,
            [Parameter(Mandatory=$true)][string]$SourceDir,
            [Parameter(Mandatory=$true)][string]$DestDir
        )

        $MissedGamesList = @()
        
        $StopWords = " the ", " of ", " in ", " and ", " a ", " an ", " vs ", " versus ", " starring ", " video game ", " turbo ", " tournament ", " secret ", " adventure ", " fantasy ", " world ", " story "
        $StopWordsRegex = ($StopWords | ForEach-Object { [regex]::Escape($_) }) -join '|'
        $LeadingArticleRegex = '^(the|a|an)\s+'
        $NumberRegex = '\d+'

        foreach ($GameName in $GameList) {
            if ([string]::IsNullOrEmpty($GameName)) { continue }
            
            $WorkingGameName = $GameName
            
            # 1. Normalize title by removing leading articles, Roman numerals to Arabic, etc.
            $WorkingGameName = [regex]::Replace($WorkingGameName, $LeadingArticleRegex, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $WorkingGameName = $WorkingGameName -replace ' IV ', ' 4 ' -replace ' III ', ' 3 ' -replace ' II ', ' 2 ' -replace ' I ', ' 1 '

            # --- ENHANCED CLEANING FOR GAME NAME (LIST) ---
            $CleanGameName = ($WorkingGameName | 
                ForEach-Object { $_ -replace '\s*\(.*\)\s*', '' } |
                ForEach-Object { $_ -replace '\s*\[.*\]\s*', '' } |
                ForEach-Object { $_ -replace '\s*\d{2,4}\s*', '' } |
                ForEach-Object { $_ -replace '[^a-zA-Z0-9\s]', ' ' } |
                ForEach-Object { $_ -replace '^\s+|\s+$', '' }
            )
                
            $CleanGameName = $CleanGameName -replace $StopWordsRegex, ' '
            $CleanGameName = $CleanGameName -replace '\s+', ' ' -replace '^\s+|\s+$', ''
            
            $GameNameNumbers = [regex]::Matches($CleanGameName, $NumberRegex) | ForEach-Object {$_.Value}
            $GameNameTokens = $CleanGameName.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)

            if ([string]::IsNullOrEmpty($CleanGameName)) { continue }

            Write-Host "`nProcessing entry: '$GameName' (Searching for '$CleanGameName' at $($Threshold * 100)%)"
            
            $PotentialMatches = @()

            # 3. Search for files and assign priority score
            foreach ($File in $AllRoms) {
                $FullName = $File.Name
                
                # --- ROM FILE CLEANING ---
                $WorkingFileName = $File.BaseName
                $CommaFlipRegex = ', (The|A|An)\s*$'
                $WorkingFileName = [regex]::Replace($WorkingFileName, $CommaFlipRegex, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                $WorkingFileName = [regex]::Replace($WorkingFileName, $LeadingArticleRegex, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                $WorkingFileName = $WorkingFileName -replace ' IV ', ' 4 ' -replace ' III ', ' 3 ' -replace ' II ', ' 2 ' -replace ' I ', ' 1 '
                
                $CleanFileName = ($WorkingFileName |
                    ForEach-Object { $_ -replace '\s*\[([^]]+)\]\s*', '' } |
                    ForEach-Object { $_ -replace '\s*\(([^)]+)\)\s*', '' } |
                    ForEach-Object { $_ -replace '[^a-zA-Z0-9\s]', ' ' } |
                    ForEach-Object { $_ -replace $StopWordsRegex, ' ' }
                )
                $CleanFileName = $CleanFileName -replace '\s+', ' ' -replace '^\s+|\s+$', ''
                
                $FileNameNumbers = [regex]::Matches($CleanFileName, $NumberRegex) | ForEach-Object {$_.Value}
                $FileNameTokens = $CleanFileName.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)

                # --- PRIMARY MATCHING: LEVENSHTEIN (Standard Similarity) ---
                $Similarity = Get-StringSimilarity -String1 $CleanGameName -String2 $CleanFileName

                # --- SEQUENCE/NUMBER CONFLICT CHECK (FP Fix) ---
                $NumberConflict = $false
                if ($GameNameNumbers.Count -gt 0) {
                    $HighestGameNumber = $GameNameNumbers | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
                    if (-not ($FileNameNumbers -contains $HighestGameNumber)) {
                        $NumberConflict = $true
                    }
                }

                # --- STRATEGY 1: TOKEN CONTAINMENT BOOST ---
                if ($FileNameTokens.Length -gt 0) {
                    $AllTokensContained = $FileNameTokens | ForEach-Object { $GameNameTokens -contains $_ }
                    $MatchCount = ($AllTokensContained | Where-Object { $_ }).Count

                    if ($MatchCount -eq $FileNameTokens.Length) {
                        if ($NumberConflict) {
                            $Similarity = [math]::Min($Similarity, 0.50)
                            Write-Host "  -> TOKEN PENALTY APPLIED: Token subset is valid, but missing required sequence number. Similarity dropped to 0.50." -ForegroundColor Red
                        }
                        elseif ($Similarity -lt 0.95) {
                            $Similarity = 0.95 
                            Write-Host "  -> TOKEN MATCH BOOST APPLIED: Score set to 0.95 due to perfect token subset." -ForegroundColor Cyan
                        }
                    }
                }

                # Apply the number conflict penalty to the final score
                if ($NumberConflict -and $Similarity -gt 0.65) {
                     $Similarity = [math]::Min($Similarity, 0.65)
                     Write-Host "  -> NUMBER PENALTY APPLIED: List name requires a number the ROM lacks. Similarity capped at 0.65." -ForegroundColor Red
                }


                $NameMatch = ($Similarity -ge $Threshold)

                if ($NameMatch) {
                    $IsPreferredRegion = ($FullName -like "*(U)*") -or ($FullName -like "*(JUE)*")
                    $IsGoodDump = ($FullName -like "*[!]**")
                    
                    $Priority = 1 
                    if ($IsPreferredRegion) { $Priority = 2 }
                    if ($IsPreferredRegion -and $IsGoodDump) { $Priority = 3 }
                    
                    $PotentialMatches += [PSCustomObject]@{
                        File = $File
                        Similarity = $Similarity
                        Priority = $Priority
                    }
                }
            }

            # 4. Choose the single best file and COPY
            if ($PotentialMatches.Count -gt 0) {
                $PrioritizedMatches = $PotentialMatches | Sort-Object -Property @{Expression={$_.Priority}; Descending=$true}, @{Expression={$_.Similarity}; Descending=$true}
                $BestMatchObject = $PrioritizedMatches[0]
                $BestMatchFile = $BestMatchObject.File
                $ChosenFileName = $BestMatchFile.Name
                
                $PriorityText = switch ($BestMatchObject.Priority) {
                    3 {"Highest Priority ([!] and (U)/(JUE))"}
                    2 {"High Priority ((U) or (JUE) only)"}
                    1 {"Accepted Region (Fallback)"}
                }
                Write-Host "  -> Chosen: $($BestMatchFile.Name) ($PriorityText, Similarity: $($BestMatchObject.Similarity))" -ForegroundColor Green

                # FINAL COPY ATTEMPT WITH .NET METHOD
                try {
                    $SourceFileObject = Get-ChildItem -Path $SourceDir -Filter $ChosenFileName -ErrorAction Stop
                    $SourcePath = $SourceFileObject.FullName
                    $DestinationPath = Join-Path -Path $DestDir -ChildPath $ChosenFileName
                    [System.IO.File]::Copy($SourcePath, $DestinationPath, $true)
                    
                    if (Test-Path -Path $DestinationPath) {
                        Write-Host "  Copy SUCCESSFUL. File confirmed at: $DestinationPath" -ForegroundColor Green
                    } else {
                        Write-Host "  WARNING: Copy succeeded but file was not immediately detectable. Confirm file is in destination." -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Error "  ERROR: Copy failed for $ChosenFileName. Error: $($_.Exception.Message)"
                }
            }
            else {
                $MissedGamesList += $GameName
                Write-Host "  No file found matching '$GameName' above $($Threshold * 100)%." -ForegroundColor Gray
            }
        }

        return $MissedGamesList
    }
    
    # 1. Setup
    Write-Host "Starting file copy process using System.IO.File::Copy method..."
    Write-Host "Destination Directory: $DestinationDirectory"
    Write-Host "Initial Fuzzy Match Threshold: $($SimilarityThreshold * 100)%"
    Write-Host "--------------------------------------------------------"

    if (-not (Test-Path -Path $GameListFile)) { Write-Error "Game list file not found: $GameListFile"; Stop-Transcript; exit 1 }
    
    # Filter out empty strings after trimming
    $InitialGameList = Get-Content -Path $GameListFile | Select-Object -First 100 | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    
    if ($InitialGameList.Count -eq 0) { Write-Host "The game list is empty (or contains only empty lines). Exiting."; Stop-Transcript; exit 1 }
    
    $AllRoms = Get-ChildItem -Path $SourceDirectory -File -ErrorAction SilentlyContinue

    # Execute Initial Scan
    $MissedGames = Run-RomScan -GameList $InitialGameList -Threshold $SimilarityThreshold -AllRoms $AllRoms -SourceDir $SourceDirectory -DestDir $DestinationDirectory

    # ====================================================================
    # POST-EXECUTION REPORTING AND RERUN PROMPT
    # ====================================================================

    # --- Time Elapsed ---
    $EndTime = Get-Date
    $ElapsedTime = New-TimeSpan -Start $StartTime -End $EndTime
    $TimeFormatted = "{0:00}:{1:00}:{2:00}" -f $ElapsedTime.Hours, $ElapsedTime.Minutes, $ElapsedTime.Seconds

    Write-Host "`n--------------------------------------------------------"
    Write-Host "Script finished. Copy operations complete."
    Write-Host "Total time elapsed: $TimeFormatted"
    Write-Host "--------------------------------------------------------"

    # Report Missed Games
    if ($MissedGames.Count -gt 0) {
        Write-Host "`nMissed Games Report: $($MissedGames.Count) games not found (or similarity too low):" -ForegroundColor Red
        $Counter = 1
        foreach ($MissedGame in $MissedGames) {
            Write-Host "  $Counter. $MissedGame" -ForegroundColor Yellow
            $Counter++
        }
        
        Write-Host "`n--------------------------------------------------------"
        # --- Rerun Prompt ---
        $RerunPrompt = "Do you want to re-scan these $($MissedGames.Count) missed games using the secondary similarity threshold of $($SecondaryThreshold * 100)% (Y/N)?"
        $Response = Read-Host $RerunPrompt

        if ($Response -eq 'Y' -or $Response -eq 'y') {
            Write-Host "`nStarting secondary scan for missed games at $($SecondaryThreshold * 100)% similarity..." -ForegroundColor Cyan
            
            # Reset timer for the secondary scan
            $StartTimeRerun = Get-Date
            
            $MissedGamesAfterRerun = Run-RomScan -GameList $MissedGames -Threshold $SecondaryThreshold -AllRoms $AllRoms -SourceDir $SourceDirectory -DestDir $DestinationDirectory
            
            # Final Report after Rerun
            $EndTimeRerun = Get-Date
            $ElapsedTimeRerun = New-TimeSpan -Start $StartTimeRerun -End $EndTimeRerun
            $TimeFormattedRerun = "{0:00}:{1:00}:{2:00}" -f $ElapsedTimeRerun.Hours, $ElapsedTimeRerun.Minutes, $ElapsedTimeRerun.Seconds

            Write-Host "`n--------------------------------------------------------"
            Write-Host "Secondary scan complete."
            Write-Host "Secondary scan time elapsed: $TimeFormattedRerun"
            
            if ($MissedGamesAfterRerun.Count -gt 0) {
                Write-Host "`nMissed Games (Final): $($MissedGamesAfterRerun.Count) games still not found:" -ForegroundColor Red
                $Counter = 1
                foreach ($MissedGame in $MissedGamesAfterRerun) {
                    Write-Host "  $Counter. $MissedGame" -ForegroundColor Yellow
                    $Counter++
                }
            } else {
                Write-Host "`nSUCCESS! All remaining games were matched and processed in the secondary scan!" -ForegroundColor Green
            }
            
        } else {
            Write-Host "Secondary scan skipped." -ForegroundColor Yellow
        }
    } else {
        Write-Host "`nSUCCESS! All top games were matched and processed!" -ForegroundColor Green
    }

    Write-Host "--------------------------------------------------------"
    Write-Host "Check your destination folder for all copied files."
    # --- NEW FEATURE: Stop Logging ---
    Stop-Transcript
}

# ====================================================================
# SCRIPT ENTRY POINT
# ====================================================================

Show-Menu
# ====================================================================
# SCRIPT CONFIGURATION
# ====================================================================
$SourceDirectory = "D:\ROMs\Genesis"
$DestinationDirectory = "D:\ROMs\Genesis\_BEST"
$GameListFile = "D:\ROMs\Genesis\top_genesis_games.txt"
$SimilarityThreshold = 0.70 # Initial Scan Threshold

# --- NEW FEATURE: Output Logging Setup (FIXED LOCATION & RANDOM ID) ---
$Identifier = @(
    [char]([System.Random]::new().Next(0, 26) + [byte][char]'a'),
    [char]([System.Random]::new().Next(0, 26) + [byte][char]'a')
) -join ""
$LogFileName = "results_$Identifier.txt"

# Use $PSScriptRoot to ensure the log folder is created next to the script
$LogDirectory = Join-Path -Path $PSScriptRoot -ChildPath "results"

# Check and create the results folder
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

# --- CRITICAL FIX ATTEMPT: Grant Write Permissions (Optional but Recommended) ---
try {
    $acl = Get-Acl $DestinationDirectory -ErrorAction Stop
    $permission = "Everyone","FullControl","Allow"
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
    $acl.SetAccessRule($rule)
    $acl | Set-Acl $DestinationDirectory -ErrorAction Stop
    Write-Host "Confirmed full write permissions on destination directory."
} catch {
    Write-Warning "Could not modify permissions on $DestinationDirectory. Run script as Administrator if copies fail."
}

# ====================================================================
# CUSTOM FUNCTION: String Similarity (Fuzzy Matching) - UNCHANGED
# ====================================================================
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

# ====================================================================
# MAIN SCRIPT EXECUTION BLOCK
# ====================================================================

function Run-RomScan {
    param(
        [Parameter(Mandatory=$true)][string[]]$GameList,
        [Parameter(Mandatory=$true)][double]$Threshold,
        [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$AllRoms,
        [Parameter(Mandatory=$true)][string]$SourceDir,
        [Parameter(Mandatory=$true)][string]$DestDir
    )

    $MissedGamesList = @()
    
    # Define common English stop words - Agnostic only
    $StopWords = " the ", " of ", " in ", " and ", " a ", " an ", " vs ", " versus ", " starring ", " video game "
    $StopWordsRegex = ($StopWords | ForEach-Object { [regex]::Escape($_) }) -join '|'
    
    # Regex to remove leading articles (A, An, The) for normalization
    $LeadingArticleRegex = '^(the|a|an)\s+'

    foreach ($GameName in $GameList) {
        if ([string]::IsNullOrEmpty($GameName)) { continue }
        
        $WorkingGameName = $GameName
        
        # 1. Normalize title by removing leading articles
        $WorkingGameName = [regex]::Replace($WorkingGameName, $LeadingArticleRegex, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # 2. Numeric Normalization (Roman to Arabic)
        $WorkingGameName = $WorkingGameName -replace ' IV ', ' 4 ' -replace ' III ', ' 3 ' -replace ' II ', ' 2 ' -replace ' I ', ' 1 '

        # --- ENHANCED CLEANING FOR GAME NAME (LIST) ---
        $CleanGameName = ($WorkingGameName | 
            # Remove parentheses content
            ForEach-Object { $_ -replace '\s*\(.*\)\s*', '' } |
            # Remove brackets content
            ForEach-Object { $_ -replace '\s*\[.*\]\s*', '' } |
            # Remove years/numbers 
            ForEach-Object { $_ -replace '\s*\d{2,4}\s*', '' } |
            
            # AGGRESSIVE PUNCTUATION STANDARDIZATION
            ForEach-Object { $_ -replace '[^a-zA-Z0-9\s]', ' ' } |
            
            # Trim leading/trailing whitespace
            ForEach-Object { $_ -replace '^\s+|\s+$', '' }
        )
            
        # Remove stop words and keywords
        $CleanGameName = $CleanGameName -replace $StopWordsRegex, ' '
        $CleanGameName = $CleanGameName -replace '\s+', ' ' -replace '^\s+|\s+$', ''
        
        # Prepare tokens for Strategy 1 (Token Containment)
        $GameNameTokens = $CleanGameName.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)

        if ([string]::IsNullOrEmpty($CleanGameName)) { continue }

        Write-Host "`nProcessing entry: '$GameName' (Searching for '$CleanGameName' at $($Threshold * 100)%)"
        
        $PotentialMatches = @()

        # 3. Search for files and assign priority score
        foreach ($File in $AllRoms) {
            $FullName = $File.Name
            
            # --- ROM FILE CLEANING ---
            $WorkingFileName = $File.BaseName
            
            # COMMA-FLIPPING CORRECTION (Ooze, The)
            $CommaFlipRegex = ', (The|A|An)\s*$'
            $WorkingFileName = [regex]::Replace($WorkingFileName, $CommaFlipRegex, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

            # Normalize ROM file name by removing leading articles
            $WorkingFileName = [regex]::Replace($WorkingFileName, $LeadingArticleRegex, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            
            # Numeric Normalization (Roman to Arabic)
            $WorkingFileName = $WorkingFileName -replace ' IV ', ' 4 ' -replace ' III ', ' 3 ' -replace ' II ', ' 2 ' -replace ' I ', ' 1 '
            
            # --- ENHANCED CLEANING FOR FILE NAME (ROM) ---
            $CleanFileName = ($WorkingFileName |
                # Remove brackets content
                ForEach-Object { $_ -replace '\s*\[([^]]+)\]\s*', '' } |
                # Remove parentheses content (often has region/version info)
                ForEach-Object { $_ -replace '\s*\(([^)]+)\)\s*', '' } |
                
                # AGGRESSIVE PUNCTUATION STANDARDIZATION
                ForEach-Object { $_ -replace '[^a-zA-Z0-9\s]', ' ' } |
                
                # Remove stop words and keywords
                ForEach-Object { $_ -replace $StopWordsRegex, ' ' }
            )
            $CleanFileName = $CleanFileName -replace '\s+', ' ' -replace '^\s+|\s+$', ''
            
            # Get ROM Name Tokens
            $FileNameTokens = $CleanFileName.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)


            # --- PRIMARY MATCHING: LEVENSHTEIN (Standard Similarity) ---
            $Similarity = Get-StringSimilarity -String1 $CleanGameName -String2 $CleanFileName


            # --- STRATEGY 1: TOKEN CONTAINMENT BOOST (The Aladdin/Ecco Fix) ---
            $TokenMatch = $false
            if ($FileNameTokens.Length -gt 0) {
                # Check if EVERY word in the ROM name is present in the List Name's words
                $AllTokensContained = $FileNameTokens | ForEach-Object { 
                    $token = $_ 
                    $GameNameTokens -contains $token 
                }
                
                # COUNT FIX FOR PS 5.1 COMPATIBILITY: Use Where-Object piped to .Count
                $MatchCount = ($AllTokensContained | Where-Object { $_ }).Count

                # If the number of matching tokens equals the total number of ROM tokens, it's a perfect subset match
                if ($MatchCount -eq $FileNameTokens.Length) {
                    $TokenMatch = $true
                    # Boost the similarity score significantly if it is a subset match
                    if ($Similarity -lt 0.95) {
                        $Similarity = 0.95 
                        Write-Host "  -> TOKEN MATCH BOOST APPLIED: Score set to 0.95 due to perfect token subset." -ForegroundColor Cyan
                    }
                }
            }


            $NameMatch = ($Similarity -ge $Threshold)

            if ($NameMatch) {
                $IsPreferredRegion = ($FullName -like "*(U)*") -or ($FullName -like "*(JUE)*")
                $IsGoodDump = ($FullName -like "*[!]**")
                
                # Priority logic remains the same
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

        # 4. Choose the single best file and COPY (using .NET)
        if ($PotentialMatches.Count -gt 0) {
            
            # Sorting logic remains the same
            $PrioritizedMatches = $PotentialMatches | Sort-Object -Property @{Expression={$_.Priority}; Descending=$true}, @{Expression={$_.Similarity}; Descending=$true}
            $BestMatchObject = $PrioritizedMatches[0]
            $BestMatchFile = $BestMatchObject.File
            $ChosenFileName = $BestMatchFile.Name
            
            # Verbose Output
            $PriorityText = switch ($BestMatchObject.Priority) {
                3 {"Highest Priority ([!] and (U)/(JUE))"}
                2 {"High Priority ((U) or (JUE) only)"}
                1 {"Accepted Region (Fallback)"}
            }
            Write-Host "  -> Chosen: $($BestMatchFile.Name) ($PriorityText, Similarity: $($BestMatchObject.Similarity))" -ForegroundColor Green

            # FINAL COPY ATTEMPT WITH .NET METHOD
            try {
                $SourceFileObject = Get-ChildItem -Path $SourceDir -Filter $ChosenFileName -ErrorAction Stop
                
                if (-not $SourceFileObject) {
                     Write-Error "  CRITICAL: Final file object lookup failed for $ChosenFileName. Skipping copy." -ErrorAction Stop
                     continue
                }
                
                $SourcePath = $SourceFileObject.FullName
                $DestinationPath = Join-Path -Path $DestDir -ChildPath $ChosenFileName

                # Use .NET Copy Method
                [System.IO.File]::Copy($SourcePath, $DestinationPath, $true)
                
                # Final verification after copy
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
            # Record missed games
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

if (-not (Test-Path -Path $GameListFile)) { Write-Error "Game list file not found: $GameListFile"; exit 1 }
$InitialGameList = Get-Content -Path $GameListFile | Select-Object -First 100 | ForEach-Object { $_.Trim() }
if ($InitialGameList.Count -eq 0) { Write-Host "The game list is empty. Exiting."; exit 1 }
if (-not (Test-Path -Path $DestinationDirectory)) { New-Item -Path $DestinationDirectory -ItemType Directory | Out-Null }

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
    $SecondaryThreshold = 0.50 # Secondary threshold set to 50%
    $RerunPrompt = "Do you want to re-scan these $($MissedGames.Count) missed games using a lower similarity threshold of $($SecondaryThreshold * 100)% (Y/N)?"
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
# End of Script
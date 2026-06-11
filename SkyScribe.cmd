<# : batch portion
@echo off
cd /d "%~dp0"
if "%SS_LAUNCHED%"=="" (
    set SS_LAUNCHED=1
    mode con: cols=100 lines=30
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression (Get-Content '%~f0' -Raw)"
if %errorlevel% neq 0 pause
exit /b
#>

# --- POWERSHELL STARTS HERE ---
try {
    # --- 1. LOGGING & HELPERS ---
    $SwGlobal = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Log-Info($Msg) { Write-Host "[INFO]  $Msg" -ForegroundColor Cyan }
    function Log-Warn($Msg) { Write-Host "[WARN]  $Msg" -ForegroundColor Yellow }
    function Log-Time($Step, $Sw) {
        $ms = $Sw.ElapsedMilliseconds
        $Sw.Restart()
        Write-Host "[PERF]  $($Step): " -NoNewline -ForegroundColor DarkGray
        Write-Host "${ms}ms" -ForegroundColor Green
    }

    function Clean-FileName($Name) {
        $Name = $Name -replace '[\\/:*?"<>|]', '-'
        $Ext = [System.IO.Path]::GetExtension($Name)
        $Base = [System.IO.Path]::GetFileNameWithoutExtension($Name).Trim(' .')
        return "$Base$Ext"
    }

    Write-Host "`n=== SKYSCRIBE v3 (HUB EDITION) STARTED ===" -ForegroundColor Yellow
    Write-Host "Waiting for user action in the main window...`n" -ForegroundColor DarkGray

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $ScriptRoot = (Get-Location).Path
    $AppName = "SkyScribe"
    $ConfigFile = Join-Path $ScriptRoot "SkyScribe.ini"

    # --- 2. CONFIGURATION ---
    $Config = @{
        SkipSeconds       = 15
        WindowSeconds     = 90
        FrameCount        = 15
        VideoExtensions   = ".mp4,.mov,.3gp,.m4v,.mkv,.avi"
        JumpGapMinutes    = 30
        MinFileSizeKB     = 100
        PreviewWidth      = 480
        MaxParallelFfmpeg = 4
        RecursivePeopleSearch = 0
        RecursiveJumpSearch   = 0
        RecursionDepth        = 1
        DefaultFolder         = ""
    }

    if (Test-Path $ConfigFile) {
        Get-Content $ConfigFile | ForEach-Object {
            if ($_ -match "^\s*(\w+)\s*=\s*(.*)$") {
                $Key = $matches[1].Trim()
                $Value = $matches[2].Trim()
                if ($Config.ContainsKey($Key)) { 
                    if ($Key -in @("VideoExtensions", "DefaultFolder")) { $Config[$Key] = $Value } 
                    elseif ($Value -match "^\d+$") { $Config[$Key] = [int]$Value }
                }
            }
        }
    } else {
        $Content = @(
            "[SkyScribe Settings]",
            "SkipSeconds=15",
            "WindowSeconds=90",
            "FrameCount=15",
            "JumpGapMinutes=30",
            "VideoExtensions=.mp4,.mov,.3gp,.m4v,.mkv,.avi",
            "MinFileSizeKB=100",
            "PreviewWidth=480",
            "MaxParallelFfmpeg=4",
            "RecursivePeopleSearch=0",
            "RecursiveJumpSearch=0",
            "RecursionDepth=1",
            "DefaultFolder="
        )
        $Content | Set-Content $ConfigFile
    }

    # --- 3. METADATA ENGINE (FFPROBE EDITION) ---
    function Get-MediaMetadata {
        param($FilePath, $ProbePath)
        
        $Item = Get-Item $FilePath
        $Result = [PSCustomObject]@{
            Date = $Item.LastWriteTime
            Duration = "" 
        }

        if ($ProbePath -and (Test-Path $ProbePath)) {
            try {
                $json = & $ProbePath -v quiet -print_format json -show_entries format=duration:format_tags=creation_time -i $FilePath | Out-String | ConvertFrom-Json
                
                if ($json.format.duration) {
                    $ts = [TimeSpan]::FromSeconds([double]$json.format.duration)
                    $Result.Duration = $ts.ToString("hh\:mm\:ss")
                }

                if ($json.format.tags.creation_time) {
                    try {
                        $Result.Date = [DateTime]$json.format.tags.creation_time
                    } catch {
                        Write-Host "[WARN]  Failed to parse creation_time '$($json.format.tags.creation_time)' for $([System.IO.Path]::GetFileName($FilePath)). Using LastWriteTime." -ForegroundColor Yellow
                    }
                }
            } catch {}
        }
        return $Result
    }

    # --- 4. PREFETCH ENGINE ---
    $PreviewJobScript = {
        param($FFmpegPath, $InputFile, $DurationStr, $BaseTempPath, $UniqueId, $CfgSkip, $CfgWindow, $CfgFrames, $CfgWidth, $MaxParallel)
        
        $TotalSecs = 0
        if ($DurationStr -match "(\d+):(\d+):(\d+)") { 
            $TotalSecs = ([int]$matches[1] * 3600) + ([int]$matches[2] * 60) + [int]$matches[3] 
        } elseif ($DurationStr -match "^\d+(\.\d+)?$") {
            $TotalSecs = [int][double]$DurationStr
        } else {
             $TotalSecs = 3600 
        }

        $OutDir = Join-Path $BaseTempPath $UniqueId
        if (Test-Path $OutDir) { Remove-Item $OutDir -Recurse -Force }
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

        $StartTime = $CfgSkip; $EndTime = $CfgSkip + $CfgWindow
        if ($TotalSecs -lt $CfgWindow) { $StartTime = 0; $EndTime = $TotalSecs } 
        elseif ($TotalSecs -lt $CfgSkip) { $StartTime = 0; $EndTime = $TotalSecs } 
        elseif ($TotalSecs -lt $EndTime) { $EndTime = $TotalSecs }

        $TimeWindow = $EndTime - $StartTime; if ($TimeWindow -le 0) { $TimeWindow = 1 }
        $Interval = $TimeWindow / ($CfgFrames + 1)
        
        $RunningProcs = @()
        for ($i=1; $i -le $CfgFrames; $i++) {
            while (($RunningProcs | Where-Object { try { -not $_.HasExited } catch { $false } }).Count -ge $MaxParallel) { Start-Sleep -Milliseconds 50 }
            $PadNum = $i.ToString("00"); $OutFile = Join-Path $OutDir "frame_$PadNum.jpg"
            $Offset = [math]::Round($Interval * $i); $FinalTime = $StartTime + $Offset
            $Args = "-ss $FinalTime -i `"$InputFile`" -frames:v 1 -vf scale=${CfgWidth}:-1 -q:v 5 -y `"$OutFile`""
            $p = Start-Process -FilePath $FFmpegPath -ArgumentList $Args -WindowStyle Hidden -PassThru
            $RunningProcs += $p
        }
        while (($RunningProcs | Where-Object { try { -not $_.HasExited } catch { $false } }).Count -gt 0) { Start-Sleep -Milliseconds 50 }
        $FailCount = 0
        foreach ($p in $RunningProcs) { 
            try { if ($p.ExitCode -ne 0) { $FailCount++ } } catch {} 
            try { $p.Dispose() } catch {} 
        }
        if ($FailCount -gt 0) { Write-Warning "FFmpeg failed on $FailCount of $CfgFrames frames." }
        return $OutDir
    }

    function Load-ImagesFromFolder {
        param($FolderPath)
        $Loaded = @()
        if ($FolderPath -and (Test-Path $FolderPath)) {
            $Files = Get-ChildItem -Path $FolderPath -Filter "*.jpg" | Sort-Object Name
            foreach ($f in $Files) {
                try {
                    $Bytes = [System.IO.File]::ReadAllBytes($f.FullName)
                    $Stream = New-Object System.IO.MemoryStream(,$Bytes)
                    $Loaded += [System.Drawing.Image]::FromStream($Stream)
                } catch {}
            }
        }
        return $Loaded
    }

    # --- 5. SETTINGS FORM ---
    function Show-SettingsForm {
        param($Cfg, $Path, $ParentForm)
        $SetForm = New-Object System.Windows.Forms.Form
        $SetForm.Text = "Settings"
        $SetForm.Size = New-Object System.Drawing.Size(500, 650)
        $SetForm.StartPosition = "CenterParent"
        $SetForm.FormBorderStyle = "FixedDialog"
        $SetForm.MaximizeBox = $false
        $SetForm.MinimizeBox = $false
        $SetForm.TopMost = $true

        $FontStd = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
        $Layout = @{ Top = 20 }
        
        $AddNum = { param($Lbl, $Key, $Min, $Max) 
            $l = New-Object System.Windows.Forms.Label; $l.Text=$Lbl; $l.Top=$Layout.Top+3; $l.Left=30; $l.AutoSize=$true; $l.Font=$FontStd; $SetForm.Controls.Add($l)
            $n = New-Object System.Windows.Forms.NumericUpDown; $n.Top=$Layout.Top; $n.Left=300; $n.Width=120; $n.Minimum=$Min; $n.Maximum=$Max; $n.Value=$Cfg[$Key]; $n.Font=$FontStd; $SetForm.Controls.Add($n)
            $Layout.Top += 45
            return $n
        }
        
        $AddChk = { param($Lbl, $Key)
            $c = New-Object System.Windows.Forms.CheckBox; $c.Text=$Lbl; $c.Top=$Layout.Top; $c.Left=30; $c.Width=400; $c.Checked=($Cfg[$Key] -eq 1); $c.Font=$FontStd; $SetForm.Controls.Add($c)
            $Layout.Top += 45
            return $c
        }

        $AddFolderBrowser = { param($Lbl, $Key)
            $l = New-Object System.Windows.Forms.Label; $l.Text=$Lbl; $l.Top=$Layout.Top+3; $l.Left=30; $l.AutoSize=$true; $l.Font=$FontStd; $SetForm.Controls.Add($l)
            $t = New-Object System.Windows.Forms.TextBox; $t.Top=$Layout.Top; $t.Left=250; $t.Width=140; $t.Text=$Cfg[$Key]; $t.Font=$FontStd; $SetForm.Controls.Add($t)
            
            $b = New-Object System.Windows.Forms.Button; $b.Text="..."; $b.Top=$Layout.Top-1; $b.Left=395; $b.Width=30; $b.Height=26; $SetForm.Controls.Add($b)
            $b.Tag = $t 
            
            $b.Add_Click({
                $txtBox = $this.Tag 
                $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
                if ($txtBox.Text -and (Test-Path $txtBox.Text)) { $fbd.SelectedPath = $txtBox.Text }
                if ($fbd.ShowDialog() -eq "OK") { $txtBox.Text = $fbd.SelectedPath }
                $fbd.Dispose()
            })
            $Layout.Top += 45
            return $t
        }

        $cRecP  = &$AddChk "Recursive People Search" "RecursivePeopleSearch"
        $cRecJ  = &$AddChk "Recursive Jump Search" "RecursiveJumpSearch"
        $nDepth = &$AddNum "Max Recursion Depth (Layers)" "RecursionDepth" 0 20
        $nSkip  = &$AddNum "Skip Start (seconds)" "SkipSeconds" 0 300
        $nWind  = &$AddNum "Window Duration (seconds)" "WindowSeconds" 10 600
        $nFrame = &$AddNum "Thumbnail Count" "FrameCount" 1 50
        $nGap   = &$AddNum "Jump Gap (minutes)" "JumpGapMinutes" 1 120
        $nPar   = &$AddNum "Parallel FFmpeg Threads" "MaxParallelFfmpeg" 1 16
        $nPrevW = &$AddNum "Preview Width (px)" "PreviewWidth" 100 1920
        $tDefF  = &$AddFolderBrowser "Default Folder" "DefaultFolder"

        $BtnSave = New-Object System.Windows.Forms.Button; $BtnSave.Text="Save"; $BtnSave.Top=$Layout.Top+20; $BtnSave.Left=120; $BtnSave.Width=100; $BtnSave.Height=35; $BtnSave.DialogResult="OK"; $SetForm.Controls.Add($BtnSave)
        $BtnCancel = New-Object System.Windows.Forms.Button; $BtnCancel.Text="Cancel"; $BtnCancel.Top=$Layout.Top+20; $BtnCancel.Left=240; $BtnCancel.Width=100; $BtnCancel.Height=35; $BtnCancel.DialogResult="Cancel"; $SetForm.Controls.Add($BtnCancel)
        $SetForm.AcceptButton = $BtnSave
        $SetForm.CancelButton = $BtnCancel

        $Result = $SetForm.ShowDialog($ParentForm)
        if ($Result -eq "OK") {
            $Cfg["RecursivePeopleSearch"] = if ($cRecP.Checked) { 1 } else { 0 }
            $Cfg["RecursiveJumpSearch"]   = if ($cRecJ.Checked) { 1 } else { 0 }
            $Cfg["RecursionDepth"]        = [int]$nDepth.Value
            $Cfg["SkipSeconds"]           = [int]$nSkip.Value
            $Cfg["WindowSeconds"]         = [int]$nWind.Value
            $Cfg["FrameCount"]            = [int]$nFrame.Value
            $Cfg["JumpGapMinutes"]        = [int]$nGap.Value
            $Cfg["MaxParallelFfmpeg"]     = [int]$nPar.Value
            $Cfg["PreviewWidth"]          = [int]$nPrevW.Value
            $Cfg["DefaultFolder"]         = $tDefF.Text.Trim()
            
            $IniKeyOrder = @(
                "RecursivePeopleSearch", "RecursiveJumpSearch", "RecursionDepth",
                "SkipSeconds", "WindowSeconds", "FrameCount", "JumpGapMinutes",
                "VideoExtensions", "MinFileSizeKB", "PreviewWidth", "MaxParallelFfmpeg",
                "DefaultFolder"
            )
            $NewContent = @("[SkyScribe Settings]")
            foreach ($k in $IniKeyOrder) { if ($Cfg.ContainsKey($k)) { $NewContent += "$k=$($Cfg[$k])" } }
            $NewContent | Set-Content $Path -Force
        }
        $SetForm.Dispose()
        return ($Result -eq "OK")
    }

    # --- 6. MAIN RE-NAMING FORM ---
    function Show-SkydiveForm {
        param($FileName, $FullName, $FileTime, $Duration, $SuggestedDate, $SuggestedJump, $SuggestedClip, $SuggestedPeople, $SuggestedDesc, $TargetFolder, $OriginalExt, $PreloadedImages, $Config)
        
        $Form = New-Object System.Windows.Forms.Form
        $Form.Text = "$AppName - $FileName"
        $Form.Size = New-Object System.Drawing.Size(760, 930) 
        $Form.StartPosition = "CenterScreen"
        $Form.Topmost = $true 
        $Form.FormBorderStyle = "Sizable"
        $Form.MaximizeBox = $true
        $Form.MinimumSize = New-Object System.Drawing.Size(760, 600)

        # --- MENU STRIP ---
        $MenuStrip = New-Object System.Windows.Forms.MenuStrip
        $FileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("File")
        $SettingsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Settings...")
        $ExitItem = New-Object System.Windows.Forms.ToolStripMenuItem("Exit Batch")
        
        $ExitItem.Add_Click({ $Form.DialogResult = [System.Windows.Forms.DialogResult]::Abort; $Form.Close() })
        $FileMenu.DropDownItems.Add($SettingsItem)
        $FileMenu.DropDownItems.Add("-")
        $FileMenu.DropDownItems.Add($ExitItem)
        $MenuStrip.Items.Add($FileMenu)
        $Form.MainMenuStrip = $MenuStrip
        $Form.Controls.Add($MenuStrip)

        $FontStd  = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
        $FontBold = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $FontPrev = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold) 

        $YOffset = 30
        
        $AddLabel = { param($txt, $top, $left=25) $l = New-Object System.Windows.Forms.Label; $l.Text = $txt; $l.Top = $top + $YOffset; $l.Left = $left; $l.AutoSize = $true; $l.Font = $FontBold; $Form.Controls.Add($l) }
        $AddValue = { param($txt, $top, $left=130) $l = New-Object System.Windows.Forms.Label; $l.Text = $txt; $l.Top = $top + $YOffset; $l.Left = $left; $l.AutoSize = $true; $l.Font = $FontStd; $Form.Controls.Add($l) }

        # --- CONTROLS ---
        &$AddLabel "LOCATION:" 15; $ShortLoc = if ($TargetFolder.Length -gt 60) { "..." + $TargetFolder.Substring($TargetFolder.Length - 60) } else { $TargetFolder }; &$AddValue $ShortLoc 15
        &$AddLabel "FILE:" 45; &$AddValue $FileName 45
        &$AddLabel "TIMESTAMP:" 75; &$AddValue $FileTime 75
        &$AddLabel "LENGTH:" 105; $DurText = if ($Duration) { $Duration } else { "---" }; &$AddValue $DurText 105

        &$AddLabel "Date (YYYY_MM_DD):" 150; $DateIn = New-Object System.Windows.Forms.TextBox; $DateIn.Top = 175 + $YOffset; $DateIn.Left = 30; $DateIn.Width = 380; $DateIn.Text = $SuggestedDate; $DateIn.Font = $FontStd
        $DateIn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $Form.Controls.Add($DateIn)

        &$AddLabel "Jump Number:" 220; $JumpIn = New-Object System.Windows.Forms.TextBox; $JumpIn.Top = 245 + $YOffset; $JumpIn.Left = 30; $JumpIn.Width = 180; $JumpIn.Text = $SuggestedJump; $JumpIn.Font = $FontStd; $Form.Controls.Add($JumpIn)
        &$AddLabel "Clip Number:" 220 230; $ClipIn = New-Object System.Windows.Forms.TextBox; $ClipIn.Top = 245 + $YOffset; $ClipIn.Left = 230; $ClipIn.Width = 180; $ClipIn.Text = $SuggestedClip; $ClipIn.Font = $FontStd; $Form.Controls.Add($ClipIn)

        &$AddLabel "People:" 280; $PeopleIn = New-Object System.Windows.Forms.TextBox; $PeopleIn.Top = 305 + $YOffset; $PeopleIn.Left = 30; $PeopleIn.Width = 380; $PeopleIn.Text = $SuggestedPeople; $PeopleIn.Font = $FontStd
        $PeopleIn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $Form.Controls.Add($PeopleIn)

        &$AddLabel "Description:" 340; $DescIn = New-Object System.Windows.Forms.TextBox; $DescIn.Top = 365 + $YOffset; $DescIn.Left = 30; $DescIn.Width = 380; $DescIn.Text = $SuggestedDesc; $DescIn.Font = $FontStd
        $DescIn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $Form.Controls.Add($DescIn)

        # Fill From File button
        $FillBtn = New-Object System.Windows.Forms.Button
        $FillBtn.Text = "Fill From File..."
        $FillBtn.Top = 400 + $YOffset
        $FillBtn.Left = 30
        $FillBtn.Width = 180
        $FillBtn.Height = 30
        $FillBtn.Font = $FontStd
        $FillBtn.Add_Click({
            $PickDlg = New-Object System.Windows.Forms.OpenFileDialog
            $PickDlg.Title = "Select a labeled video to copy info from"
            $PickDlg.InitialDirectory = $TargetFolder
            $PickDlg.Multiselect = $false
            $FilterExts = $Config.VideoExtensions -replace ",", ";" -replace "\.", "*."
            $PickDlg.Filter = "Video Files ($FilterExts)|$FilterExts|All Files (*.*)|*.*"
            
            if ($PickDlg.ShowDialog($Form) -eq "OK") {
                $PickedName = [System.IO.Path]::GetFileName($PickDlg.FileName)
                
                $ParsedJump = ""; $ParsedDate = ""; $ParsedPeople = ""; $ParsedDesc = ""
                
                if ($PickedName -match "^#(\d+)(?:-\d+)?\s+(\d{4}_\d{2}_\d{2})\s+(.*?)(?:\s+-(.*))?\.") {
                    $ParsedJump = $matches[1]
                    $ParsedDate = $matches[2]
                    $ParsedPeople = $matches[3].Trim()
                    if ($matches.Count -gt 4 -and $matches[4]) { $ParsedDesc = $matches[4].Trim() }
                } elseif ($PickedName -match "^#(\d+)(?:-\d+)?\s+(\d{4}_\d{2}_\d{2})") {
                    $ParsedJump = $matches[1]
                    $ParsedDate = $matches[2]
                }
                
                if ($ParsedJump) {
                    $DateIn.Text = $ParsedDate
                    $JumpIn.Text = $ParsedJump
                    $PeopleIn.Text = $ParsedPeople
                    $DescIn.Text = $ParsedDesc
                    
                    # Calculate next clip number
                    $SearchArgs = @{ LiteralPath = $TargetFolder; File = $true }
                    if ($Config.RecursiveJumpSearch -eq 1) { $SearchArgs["Recurse"] = $true; $SearchArgs["Depth"] = $Config.RecursionDepth }
                    
                    $TargetJump = $ParsedJump.Trim()
                    $existing = Get-ChildItem @SearchArgs | Where-Object { $_.Name -like "#$TargetJump*" }
                    $max = 0; $FoundAny = $false
                    $EscapedJump = [regex]::Escape($TargetJump)
                    
                    foreach ($ex in $existing) {
                        $FoundAny = $true
                        if ($ex.Name -match "^#$EscapedJump-(\d+)") { 
                            $val = [int]$matches[1]; if ($val -gt $max) { $max = $val } 
                        }
                    }
                    
                    if ($max -gt 0) { $ClipIn.Text = ($max + 1).ToString() } 
                    elseif ($FoundAny) { $ClipIn.Text = "2" } 
                    else { $ClipIn.Text = "1" }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Could not parse naming info from:`n$PickedName`n`nExpected format: #Jump Date People -Desc.ext", "Parse Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            }
            $PickDlg.Dispose()
        })
        $Form.Controls.Add($FillBtn)

        # Recent People
        &$AddLabel "RECENT PEOPLE (Double-Click):" 20 460
        $Form.Controls[$Form.Controls.Count-1].Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        
        $SortDrop = New-Object System.Windows.Forms.ComboBox; $SortDrop.Top = 45 + $YOffset; $SortDrop.Left = 460; $SortDrop.Width = 150; $SortDrop.Font = $FontStd
        $SortDrop.Items.Add("Sort: Frequency")
        $SortDrop.Items.Add("Sort: A-Z")
        $SortDrop.SelectedIndex = 0
        $SortDrop.DropDownStyle = "DropDownList"
        $SortDrop.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        $Form.Controls.Add($SortDrop)

        $PeopleList = New-Object System.Windows.Forms.ListBox; $PeopleList.Top = 75 + $YOffset; $PeopleList.Left = 460; $PeopleList.Width = 240; $PeopleList.Height = 350; $PeopleList.Font = $FontStd
        $PeopleList.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        $NameCounts = @{} 
        $Form.Controls.Add($PeopleList) 
        
        $RefreshPeopleList = {
            $PeopleList.Items.Clear()
            $NameCounts.Clear()
            $SearchArgs = @{ LiteralPath = $TargetFolder; File = $true }
            if ($Config.RecursivePeopleSearch -eq 1) { 
                $SearchArgs["Recurse"] = $true 
                $SearchArgs["Depth"] = $Config.RecursionDepth
            }
            Get-ChildItem @SearchArgs | Where-Object { $_.Name -match "^#\d+" } | ForEach-Object { 
                $clean = $_.BaseName
                if ($clean -match " -") { $clean = $clean.Substring(0, $clean.IndexOf(" -")) }
                $clean = $clean -replace "^#\d+\s+\d{4}_\d{2}_\d{2}", ""
                $clean.Trim().Split(" ") | ForEach-Object { 
                    $n = $_.Trim()
                    if ($n -and $n -notmatch "\d" -and $n -notmatch "-") { 
                        if (-not $NameCounts.ContainsKey($n)) { $NameCounts[$n] = 0 }
                        $NameCounts[$n]++ 
                    } 
                } 
            }
            $SortedNames = if ($SortDrop.SelectedIndex -eq 0) {
                $NameCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -ExpandProperty Key
            } else {
                $NameCounts.Keys | Sort-Object
            }
            foreach ($n in $SortedNames) { [void]$PeopleList.Items.Add($n) }
        }
        
        &$RefreshPeopleList
        $SortDrop.Add_SelectedIndexChanged({ &$RefreshPeopleList })

        $SettingsItem.Add_Click({ 
            if (Show-SettingsForm -Cfg $Config -Path $ConfigFile -ParentForm $Form) { &$RefreshPeopleList }
        })

        $PeopleList.Add_MouseDoubleClick({ if ($PeopleList.SelectedItem) { $current = $PeopleIn.Text.Trim(); if ($current -eq "") { $PeopleIn.Text = $PeopleList.SelectedItem } elseif ($current -notmatch "\b$([regex]::Escape($PeopleList.SelectedItem))\b") { $PeopleIn.Text = "$current $($PeopleList.SelectedItem)" } }})

        &$AddLabel "LIVE PREVIEW:" 430; $PreviewBox = New-Object System.Windows.Forms.Label; $PreviewBox.Top = 455 + $YOffset; $PreviewBox.Left = 30; $PreviewBox.Width = 670; $PreviewBox.Height = 50; $PreviewBox.ForeColor = "Blue"; $PreviewBox.Font = $FontPrev; $PreviewBox.BorderStyle = "FixedSingle"; $PreviewBox.TextAlign = "MiddleLeft"; $PreviewBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $Form.Controls.Add($PreviewBox)
        $UpdateBlock = { $j = $JumpIn.Text.Trim(); $c = $ClipIn.Text.Trim(); $suffix = if ($c) { "-$c" } else { "" }; $JumpStr = if ($j) { "#$j$suffix" } else { "" }; $DescStr = if ($DescIn.Text.Trim()) { "-$($DescIn.Text.Trim())" } else { "" }; $raw = "$JumpStr $($DateIn.Text) $($PeopleIn.Text) $DescStr$OriginalExt"; $PreviewBox.Text = ($raw -replace '\s+', ' ' -replace '\s+\.', '.').Trim() }
        $DateIn.Add_TextChanged($UpdateBlock); $JumpIn.Add_TextChanged($UpdateBlock); $ClipIn.Add_TextChanged($UpdateBlock); $PeopleIn.Add_TextChanged($UpdateBlock); $DescIn.Add_TextChanged($UpdateBlock); &$UpdateBlock

        $Footer = New-Object System.Windows.Forms.Panel; $Footer.Dock = [System.Windows.Forms.DockStyle]::Bottom; $Footer.Height = 80; $Form.Controls.Add($Footer)
        
        $SkipBtn = New-Object System.Windows.Forms.Button; $SkipBtn.Text = "SKIP"; $SkipBtn.Top = 15; $SkipBtn.Left = 30; $SkipBtn.Width = 120; $SkipBtn.Height = 50; $SkipBtn.DialogResult = [System.Windows.Forms.DialogResult]::Ignore; $Footer.Controls.Add($SkipBtn)
        $OkBtn = New-Object System.Windows.Forms.Button; $OkBtn.Text = "RENAME"; $OkBtn.Top = 15; $OkBtn.Left = 160; $OkBtn.Width = 540; $OkBtn.Height = 50; $OkBtn.BackColor = "LightGreen"; $OkBtn.Font = $FontBold; $OkBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.AcceptButton = $OkBtn; $OkBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $Footer.Controls.Add($OkBtn)

        $StartLbl = $Config.SkipSeconds; $EndLbl = $Config.SkipSeconds + $Config.WindowSeconds
        &$AddLabel "VIDEO FRAMES (${StartLbl}s to ${EndLbl}s):" 530
        $FlowPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $FlowPanel.Top = 560 + $YOffset; $FlowPanel.Left = 30; $FlowPanel.Width = 670; $FlowPanel.Height = $Form.ClientSize.Height - $Footer.Height - ($FlowPanel.Top) - 10
        $FlowPanel.WrapContents = $false; $FlowPanel.AutoScroll = $true; $FlowPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $Form.Controls.Add($FlowPanel)

        $ResizeImages = {
            $HScrollHeight = 25; $TargetH = $FlowPanel.Height - $HScrollHeight; if ($TargetH -lt 50) { $TargetH = 50 } 
            $TargetW = [int]($TargetH * (16/9))
            foreach ($ctrl in $FlowPanel.Controls) {
                if ($ctrl -is [System.Windows.Forms.PictureBox]) { if ([math]::Abs($ctrl.Height - $TargetH) -gt 2) { $ctrl.Size = New-Object System.Drawing.Size($TargetW, $TargetH) } }
            }
        }
        $FlowPanel.Add_Resize({ &$ResizeImages })

        if ($PreloadedImages.Count -eq 0) {
            $NoImg = New-Object System.Windows.Forms.Label; $NoImg.Text = "No Previews available"; $NoImg.AutoSize = $true; $NoImg.ForeColor = "Gray"; $FlowPanel.Controls.Add($NoImg)
        } else {
            $InitH = $FlowPanel.Height - 25; $InitW = [int]($InitH * (16/9))
            foreach ($img in $PreloadedImages) {
                $Pb = New-Object System.Windows.Forms.PictureBox; $Pb.Size = New-Object System.Drawing.Size($InitW, $InitH); $Pb.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage; $Pb.Image = $img; $Pb.BorderStyle = "FixedSingle"; $Pb.Margin = New-Object System.Windows.Forms.Padding(0,0,10,0); $FlowPanel.Controls.Add($Pb)
            }
        }

        $Result = $Form.ShowDialog()
        $OutData = if ($Result -eq "OK") { @{ Status="RENAME"; FinalName=$PreviewBox.Text; Date=$DateIn.Text; Jump=$JumpIn.Text; Clip=$ClipIn.Text; People=$PeopleIn.Text; Desc=$DescIn.Text } }
                    elseif ($Result -eq "Ignore") { @{ Status="SKIP" } }
                    else { @{ Status="ABORT" } }

        $Form.Dispose()
        return $OutData
    }

    # --- 7. BATCH PROCESSING WRAPPER ---
    function Start-SkyScribeBatch {
        
        # --- FILE SELECTION ---
        $OpenDlg = New-Object System.Windows.Forms.OpenFileDialog
        $OpenDlg.Title = "Select videos to process (Hold Ctrl/Shift to select multiple)"
        $OpenDlg.Multiselect = $true
        
        if ($Config.DefaultFolder -and (Test-Path $Config.DefaultFolder)) {
            $OpenDlg.InitialDirectory = $Config.DefaultFolder
        } else {
            $OpenDlg.InitialDirectory = $ScriptRoot
        }
        
        $FilterExts = $Config.VideoExtensions -replace ",", ";" -replace "\.", "*."
        $OpenDlg.Filter = "Video Files ($FilterExts)|$FilterExts|All Files (*.*)|*.*"

        if ($OpenDlg.ShowDialog() -eq "OK") {
            $RawFiles = $OpenDlg.FileNames | Get-Item
            $TargetFolder = $RawFiles[0].DirectoryName
            Log-Info "Selected $($RawFiles.Count) files."
        } else { 
            return # User canceled, return to Hub
        }

        # --- CHECK FFMPEG & FFPROBE ---
        $FFmpegPath = Join-Path $TargetFolder "ffmpeg.exe"
        if (-not (Test-Path $FFmpegPath)) { $FFmpegPath = Join-Path $ScriptRoot "ffmpeg.exe" }
        
        if (-not (Test-Path $FFmpegPath)) { 
            if (Get-Command "ffmpeg" -ErrorAction SilentlyContinue) { $FFmpegPath = (Get-Command "ffmpeg").Source } 
            else { 
                [System.Windows.Forms.MessageBox]::Show("FFmpeg not found! Please ensure ffmpeg.exe is in the script folder.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return 
            }
        }

        $FFprobePath = $FFmpegPath -ireplace 'ffmpeg\.exe$', 'ffprobe.exe'
        if (-not (Test-Path $FFprobePath)) {
            if (Get-Command "ffprobe" -ErrorAction SilentlyContinue) { $FFprobePath = (Get-Command "ffprobe").Source }
            else { Log-Warn "FFprobe not found. Metadata reading will be limited." }
        }

        # --- PROCESS LOOP ---
        $Sw = [System.Diagnostics.Stopwatch]::StartNew()
        Log-Info "Analyzing file metadata with ffprobe..."
        
        $FilesWithDates = @()
        foreach ($File in $RawFiles) {
            $Meta = Get-MediaMetadata -FilePath $File.FullName -ProbePath $FFprobePath
            $FilesWithDates += [PSCustomObject]@{ FileObject = $File; SortDate = $Meta.Date; Duration = $Meta.Duration }
        }
        
        $SortedQueue = @($FilesWithDates | Sort-Object SortDate)
        $LastJump = ""; $LastJumpTime = $null; $LastPeople = ""; $LastDesc = ""
        $NextJob = $null; $BaseTempPath = Join-Path $env:TEMP "SkydivePreviews"
        
        if (Test-Path $BaseTempPath) { Remove-Item $BaseTempPath -Recurse -Force -ErrorAction SilentlyContinue }
        New-Item -ItemType Directory -Path $BaseTempPath -Force | Out-Null

        for ($i = 0; $i -lt $SortedQueue.Count; $i++) {
            $QueueItem = $SortedQueue[$i]
            $File = $QueueItem.FileObject
            
            Write-Host "----------------------------------------------------" -ForegroundColor Gray
            Log-Info "Processing File [$($i+1)/$($SortedQueue.Count)]: $($File.Name)"
            $Sw.Restart()
            
            $CurrentMediaTime = $QueueItem.SortDate
            $Duration = $QueueItem.Duration
            $SuggestedDate = $CurrentMediaTime.ToString("yyyy_MM_dd")
            
            Log-Time "Metadata Read" $Sw
            $Images = @()

            if ($FFmpegPath) {
                if ($i -eq 0) {
                    Write-Host "      [SYNC] Generating initial thumbnails..." -ForegroundColor Yellow
                    $Job = Start-Job -ScriptBlock $PreviewJobScript -ArgumentList $FFmpegPath, $File.FullName, $Duration, $BaseTempPath, "0", $Config.SkipSeconds, $Config.WindowSeconds, $Config.FrameCount, $Config.PreviewWidth, $Config.MaxParallelFfmpeg
                    $ResultDir = $Job | Receive-Job -Wait -AutoRemoveJob
                    Log-Time "Thumbnail Gen" $Sw
                    $Images = Load-ImagesFromFolder $ResultDir
                    Log-Time "Image Load" $Sw
                } else {
                    if ($NextJob) {
                        Write-Host "      [ASYNC] Retrieving background job..." -ForegroundColor Gray
                        $ResultDir = $NextJob | Receive-Job -Wait -AutoRemoveJob
                        Log-Time "Retrieve Job" $Sw
                        $Images = Load-ImagesFromFolder $ResultDir
                        Log-Time "Image Load" $Sw
                    }
                }
            }

            if (($i + 1) -lt $SortedQueue.Count -and $FFmpegPath) {
                $NextItem = $SortedQueue[$i+1]
                $NextId = ($i + 1).ToString()
                $NextJob = Start-Job -ScriptBlock $PreviewJobScript -ArgumentList $FFmpegPath, $NextItem.FileObject.FullName, $NextItem.Duration, $BaseTempPath, $NextId, $Config.SkipSeconds, $Config.WindowSeconds, $Config.FrameCount, $Config.PreviewWidth, $Config.MaxParallelFfmpeg
                Write-Host "      [ASYNC] Prefetch started for next file." -ForegroundColor DarkGray
            } else { $NextJob = $null }

            $SuggestedJump = $LastJump; $SuggestedPeople = $LastPeople; $SuggestedDesc = $LastDesc; $SuggestedClip = ""

            $JumpFoundInSession = $false
            if ($null -ne $LastJumpTime) {
                if (($CurrentMediaTime - $LastJumpTime).TotalMinutes -le $Config.JumpGapMinutes) {
                    $JumpFoundInSession = $true
                }
            }

            if (-not $JumpFoundInSession -and $Config.RecursiveJumpSearch -eq 1) {
                 Log-Info "Scanning folder for existing jumps..."
                 $SearchArgs = @{ LiteralPath = $TargetFolder; File = $true }
                 if ($Config.RecursiveJumpSearch -eq 1) { 
                     $SearchArgs["Recurse"] = $true
                     $SearchArgs["Depth"] = $Config.RecursionDepth 
                 }

                 $Candidates = Get-ChildItem @SearchArgs | Where-Object { $_.Name -match "^#\d+" }
                 $BestMatch = $null; $SmallestGap = [double]::MaxValue
                 
                 foreach ($c in $Candidates) {
                    $NeighborTime = $c.LastWriteTime
                    $Diff = [math]::Abs(($NeighborTime - $CurrentMediaTime).TotalMinutes)
                    if ($Diff -le $Config.JumpGapMinutes -and $Diff -lt $SmallestGap) {
                        $SmallestGap = $Diff; $BestMatch = $c
                    }
                 }

                 if ($BestMatch) {
                    if ($BestMatch.Name -match "^#(\d+)(?:-\d+)?\s+\d{4}_\d{2}_\d{2}\s+(.*?)(?:\s+-(.*))?\.") {
                        $SuggestedJump = $matches[1]; $SuggestedPeople = $matches[2].Trim()
                        if ($matches.Count -gt 3) { $SuggestedDesc = $matches[3].Trim() }
                        Log-Info "Found neighbor: $($BestMatch.Name) (Diff: $([math]::Round($SmallestGap,1)) min)"
                        $JumpFoundInSession = $true 
                    }
                 }
            }

            if ($JumpFoundInSession) {
                if ($SuggestedJump -match "^\d+$") {
                    $SearchArgs = @{ LiteralPath = $TargetFolder; File = $true }
                    if ($Config.RecursiveJumpSearch -eq 1) { $SearchArgs["Recurse"] = $true; $SearchArgs["Depth"] = $Config.RecursionDepth }
                  
                    $TargetJump = $SuggestedJump.Trim()
                    $existing = Get-ChildItem @SearchArgs | Where-Object { $_.Name -like "#$TargetJump*" }
                    $max = 0; $FoundAny = $false
                    $EscapedJump = [regex]::Escape($TargetJump)
                    
                    foreach ($ex in $existing) {
                        $FoundAny = $true
                        if ($ex.Name -match "^#$EscapedJump-(\d+)") { 
                            $val = [int]$matches[1]; if ($val -gt $max) { $max = $val } 
                        }
                    }
                    
                    if ($max -gt 0) { $SuggestedClip = ($max + 1).ToString() } 
                    elseif ($FoundAny) { $SuggestedClip = "2" } 
                    else { $SuggestedClip = "1" }
                }
            } else {
                if ($LastJump -match "^\d+$") { $SuggestedJump = [int]$LastJump + 1 }
                $SuggestedPeople = ""; $SuggestedDesc = ""; $SuggestedClip = ""
            }

            Log-Info "Waiting for user input..."
            $Data = Show-SkydiveForm -FileName $File.Name -FullName $File.FullName -FileTime $CurrentMediaTime.ToString("MMM dd, yyyy @ HH:mm:ss") -Duration $Duration -SuggestedDate $SuggestedDate -SuggestedJump $SuggestedJump -SuggestedClip $SuggestedClip -SuggestedPeople $SuggestedPeople -SuggestedDesc $SuggestedDesc -TargetFolder $TargetFolder -OriginalExt $File.Extension -PreloadedImages $Images -Config $Config
            Log-Time "User Action" $Sw

            if ($Images) { foreach ($img in $Images) { $img.Dispose() } }
            $Images = $null
            [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

            if ($null -eq $Data) { Log-Info "Skipped (Null Data)."; continue }
            
            if ($Data.Status -eq "ABORT") { 
                if ($NextJob) { 
                    try { Stop-Job $NextJob -ErrorAction SilentlyContinue; Remove-Job $NextJob -Force -ErrorAction SilentlyContinue } catch {} 
                    $NextJob = $null
                }
                Log-Info "Exited processing early. Returning to Hub."; break 
            }
            if ($Data.Status -eq "SKIP") { Log-Info "Skipped."; continue }

            $LastJump = $Data.Jump; $LastPeople = $Data.People; $LastDesc = $Data.Desc; $LastJumpTime = $CurrentMediaTime
            
            $SanitizedName = Clean-FileName $Data.FinalName
            $NewPath = Join-Path $TargetFolder $SanitizedName
            
            if (Test-Path $NewPath) {
                Log-Warn "File exists! Appending ID to prevent overwrite."
                $Salt = (Get-Random -Minimum 100 -Maximum 999).ToString()
                $SanitizedName = $SanitizedName -replace "(\.[^.]+)$", "-$Salt`$1"
                $NewPath = Join-Path $TargetFolder $SanitizedName
            }
            
            try {
                Rename-Item -Path $File.FullName -NewName $SanitizedName -ErrorAction Stop
                Log-Info "Renamed to: $SanitizedName"
            } catch {
                Log-Warn "Rename failed: $($_.Exception.Message)"
            }
        }

        Remove-Item $BaseTempPath -Recurse -Force -ErrorAction SilentlyContinue
        Log-Info "Batch complete! Returning to Hub..."
    }

    # --- 8. MAIN HUB WINDOW ---
    $HubForm = New-Object System.Windows.Forms.Form
    $HubForm.Text = "$AppName - Hub"
    $HubForm.Size = New-Object System.Drawing.Size(500, 400)
    $HubForm.StartPosition = "CenterScreen"
    $HubForm.FormBorderStyle = "FixedDialog"
    $HubForm.MaximizeBox = $false
    $HubForm.BackColor = [System.Drawing.Color]::White

    # Menu Strip
    $HubMenu = New-Object System.Windows.Forms.MenuStrip
    $FileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("File")
    $OpenItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Videos...")
    $SettingsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Settings...")
    $ExitItem = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")

    $ExitItem.Add_Click({ $HubForm.Close() })
    $SettingsItem.Add_Click({ Show-SettingsForm -Cfg $Config -Path $ConfigFile -ParentForm $HubForm | Out-Null })
    $OpenItem.Add_Click({ 
        $HubForm.Hide()         
        Start-SkyScribeBatch    
        $HubForm.Show()         
    })

    $FileMenu.DropDownItems.Add($OpenItem)
    $FileMenu.DropDownItems.Add($SettingsItem)
    $FileMenu.DropDownItems.Add("-")
    $FileMenu.DropDownItems.Add($ExitItem)
    $HubMenu.Items.Add($FileMenu)
    $HubForm.MainMenuStrip = $HubMenu
    $HubForm.Controls.Add($HubMenu)

    # Logo Display
    $LogoBox = New-Object System.Windows.Forms.PictureBox
    $LogoBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $LogoBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    
    $LogoPath = Join-Path $ScriptRoot "SkyScribeLogo.jpg"
    
    if (Test-Path $LogoPath) {
        $LogoBox.Image = [System.Drawing.Image]::FromFile($LogoPath)
    } else {
        $FallbackLabel = New-Object System.Windows.Forms.Label
        $FallbackLabel.Text = "To display your logo, save the image as `n`n'SkyScribeLogo.jpg'`n`nin the same folder as this script."
        $FallbackLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
        $FallbackLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $HubForm.Controls.Add($FallbackLabel)
    }
    $HubForm.Controls.Add($LogoBox)
    $LogoBox.BringToFront()

    # Launch the Hub!
    [void]$HubForm.ShowDialog()

} catch {
    Write-Host "CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Error Details: $($_.ScriptStackTrace)" -ForegroundColor Yellow
    pause
}
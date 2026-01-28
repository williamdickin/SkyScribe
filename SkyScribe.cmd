<# : batch portion
@echo off
cd /d "%~dp0"
mode con: cols=100 lines=30
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
        return $Name -replace '[\\/:*?"<>|]', '-'
    }

    Write-Host "`n=== SKYCRIBE v17 (CLEAN) STARTED ===" -ForegroundColor Yellow
    Write-Host "Waiting for folder selection...`n" -ForegroundColor DarkGray

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $ScriptRoot = (Get-Location).Path
    $AppName = "SkyScribe"
    $ConfigFile = Join-Path $ScriptRoot "SkyScribe.ini"

    # --- 2. CONFIGURATION ---
    $Config = @{
        SkipSeconds       = 15
        WindowSeconds     = 60
        FrameCount        = 10
        VideoExtensions   = ".mp4,.mov,.3gp,.m4v,.mkv,.avi"
        JumpGapMinutes    = 30
        MinFileSizeKB     = 100
        PreviewWidth      = 480
    }

    if (Test-Path $ConfigFile) {
        Get-Content $ConfigFile | ForEach-Object {
            if ($_ -match "^\s*(\w+)\s*=\s*(.*)$") {
                $Key = $matches[1].Trim()
                $Value = $matches[2].Trim()
                if ($Config.ContainsKey($Key)) { 
                    if ($Key -eq "VideoExtensions") { $Config[$Key] = $Value } 
                    elseif ($Value -match "^\d+$") { $Config[$Key] = [int]$Value }
                }
            }
        }
    } else {
        $Content = @(
            "[SkyScribe Settings]",
            "SkipSeconds=15",
            "WindowSeconds=60",
            "FrameCount=10",
            "JumpGapMinutes=30",
            "VideoExtensions=.mp4,.mov,.3gp,.m4v,.mkv,.avi",
            "MinFileSizeKB=100",
            "PreviewWidth=480"
        )
        $Content | Set-Content $ConfigFile
    }
    $VideoExtArray = $Config.VideoExtensions -split "," | ForEach-Object { $_.Trim().ToLower() }

    # --- 3. FILE SELECTION ---
    $OpenDlg = New-Object System.Windows.Forms.OpenFileDialog
    $OpenDlg.Title = "Select videos to process (Hold Ctrl/Shift to select multiple)"
    $OpenDlg.Multiselect = $true
    $OpenDlg.InitialDirectory = $ScriptRoot
    $FilterExts = $Config.VideoExtensions -replace ",", ";" -replace "\.", "*."
    $OpenDlg.Filter = "Video Files ($FilterExts)|$FilterExts|All Files (*.*)|*.*"

    if ($OpenDlg.ShowDialog() -eq "OK") {
        $RawFiles = $OpenDlg.FileNames | Get-Item
        $TargetFolder = $RawFiles[0].DirectoryName
        Log-Info "Selected $($RawFiles.Count) files."
    } else { exit }

    # --- 4. CHECK FFMPEG ---
    $FFmpegPath = Join-Path $TargetFolder "ffmpeg.exe"
    if (-not (Test-Path $FFmpegPath)) { $FFmpegPath = Join-Path $ScriptRoot "ffmpeg.exe" }
    $FFmpegAvailable = Test-Path $FFmpegPath
    if (-not $FFmpegAvailable) { 
        if (Get-Command "ffmpeg" -ErrorAction SilentlyContinue) { $FFmpegPath = (Get-Command "ffmpeg").Source; $FFmpegAvailable = $true } 
        else { Write-Host "[ERROR] FFmpeg not found!" -ForegroundColor Red }
    }

    # --- 5. METADATA ENGINE ---
    $Shell = New-Object -ComObject Shell.Application
    $FolderObj = $Shell.NameSpace($TargetFolder)
    $DateIdx = 0; $DurIdx = 0
    for ($i = 0; $i -lt 320; $i++) {
        $name = $FolderObj.GetDetailsOf($null, $i)
        if ($name -match "^Media created$|^Date taken$") { $DateIdx = $i }
        if ($name -eq "Length") { $DurIdx = $i }
    }
    if ($DateIdx -eq 0) { $DateIdx = 4 }; if ($DurIdx -eq 0) { $DurIdx = 27 }

    # --- 6. PREFETCH ENGINE ---
    $PreviewJobScript = {
        param($FFmpegPath, $InputFile, $DurationStr, $BaseTempPath, $UniqueId, $CfgSkip, $CfgWindow, $CfgFrames, $CfgWidth)
        if ($DurationStr -match "(\d+):(\d+):(\d+)") { $TotalSecs = ([int]$matches[1] * 3600) + ([int]$matches[2] * 60) + [int]$matches[3] } else { return $null }
        $OutDir = Join-Path $BaseTempPath $UniqueId
        if (Test-Path $OutDir) { Remove-Item $OutDir -Recurse -Force }
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

        $StartTime = $CfgSkip; $EndTime = $CfgSkip + $CfgWindow
        if ($TotalSecs -lt $CfgSkip) { $StartTime = 0; $EndTime = $TotalSecs } else { if ($TotalSecs -lt $EndTime) { $EndTime = $TotalSecs } }
        $TimeWindow = $EndTime - $StartTime; if ($TimeWindow -le 0) { $TimeWindow = 1 }
        $Interval = $TimeWindow / ($CfgFrames + 1)
        
        $RunningProcs = @()
        for ($i=1; $i -le $CfgFrames; $i++) {
            $PadNum = $i.ToString("00")
            $OutFile = Join-Path $OutDir "frame_$PadNum.jpg"
            $Offset = [math]::Round($Interval * $i); $FinalTime = $StartTime + $Offset
            $Args = "-ss $FinalTime -i `"$InputFile`" -frames:v 1 -vf scale=${CfgWidth}:-1 -q:v 5 -y `"$OutFile`""
            $p = Start-Process -FilePath $FFmpegPath -ArgumentList $Args -WindowStyle Hidden -PassThru
            $RunningProcs += $p
        }
        if ($RunningProcs.Count -gt 0) { Wait-Process -InputObject $RunningProcs }
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
                    $Stream = New-Object System.IO.MemoryStream($Bytes, 0, $Bytes.Length)
                    $Loaded += [System.Drawing.Image]::FromStream($Stream)
                } catch {}
            }
        }
        return $Loaded
    }

    function Show-SkydiveForm {
        param($FileName, $FullName, $FileTime, $Duration, $SuggestedDate, $SuggestedJump, $SuggestedClip, $SuggestedPeople, $SuggestedDesc, $TargetFolder, $OriginalExt, $PreloadedImages, $Config)
        
        $Form = New-Object System.Windows.Forms.Form
        $Form.Text = "$AppName - $FileName"
        $Form.Size = New-Object System.Drawing.Size(760, 900)
        $Form.StartPosition = "CenterScreen"
        $Form.Topmost = $true 
        $Form.FormBorderStyle = "Sizable"
        $Form.MaximizeBox = $true
        $Form.MinimumSize = New-Object System.Drawing.Size(760, 600)

        $FontStd  = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
        $FontBold = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $FontPrev = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold) 

        $AddLabel = { param($txt, $top, $left=25) $l = New-Object System.Windows.Forms.Label; $l.Text = $txt; $l.Top = $top; $l.Left = $left; $l.AutoSize = $true; $l.Font = $FontBold; $Form.Controls.Add($l) }
        $AddValue = { param($txt, $top, $left=130) $l = New-Object System.Windows.Forms.Label; $l.Text = $txt; $l.Top = $top; $l.Left = $left; $l.AutoSize = $true; $l.Font = $FontStd; $Form.Controls.Add($l) }

        # --- CONTROLS ---
        &$AddLabel "LOCATION:" 15; $ShortLoc = if ($TargetFolder.Length -gt 60) { "..." + $TargetFolder.Substring($TargetFolder.Length - 60) } else { $TargetFolder }; &$AddValue $ShortLoc 15
        &$AddLabel "FILE:" 45; &$AddValue $FileName 45
        &$AddLabel "TIMESTAMP:" 75; &$AddValue $FileTime 75
        &$AddLabel "LENGTH:" 105; $DurText = if ($Duration) { $Duration } else { "---" }; &$AddValue $DurText 105

        &$AddLabel "Date (YYYY_MM_DD):" 150; $DateIn = New-Object System.Windows.Forms.TextBox; $DateIn.Top = 175; $DateIn.Left = 30; $DateIn.Width = 380; $DateIn.Text = $SuggestedDate; $DateIn.Font = $FontStd
        $DateIn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $Form.Controls.Add($DateIn)

        &$AddLabel "Jump Number:" 220; $JumpIn = New-Object System.Windows.Forms.TextBox; $JumpIn.Top = 245; $JumpIn.Left = 30; $JumpIn.Width = 180; $JumpIn.Text = $SuggestedJump; $JumpIn.Font = $FontStd; $Form.Controls.Add($JumpIn)
        &$AddLabel "Clip Number:" 220 230; $ClipIn = New-Object System.Windows.Forms.TextBox; $ClipIn.Top = 245; $ClipIn.Left = 230; $ClipIn.Width = 180; $ClipIn.Text = $SuggestedClip; $ClipIn.Font = $FontStd; $Form.Controls.Add($ClipIn)

        &$AddLabel "People:" 280; $PeopleIn = New-Object System.Windows.Forms.TextBox; $PeopleIn.Top = 305; $PeopleIn.Left = 30; $PeopleIn.Width = 380; $PeopleIn.Text = $SuggestedPeople; $PeopleIn.Font = $FontStd
        $PeopleIn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $Form.Controls.Add($PeopleIn)

        &$AddLabel "Description:" 340; $DescIn = New-Object System.Windows.Forms.TextBox; $DescIn.Top = 365; $DescIn.Left = 30; $DescIn.Width = 380; $DescIn.Text = $SuggestedDesc; $DescIn.Font = $FontStd
        $DescIn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $Form.Controls.Add($DescIn)

        # Recent People
        &$AddLabel "RECENT PEOPLE (Double-Click):" 20 460
        $Form.Controls[$Form.Controls.Count-1].Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        $PeopleList = New-Object System.Windows.Forms.ListBox; $PeopleList.Top = 45; $PeopleList.Left = 460; $PeopleList.Width = 240; $PeopleList.Height = 350; $PeopleList.Font = $FontStd
        $PeopleList.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        $NamesFound = New-Object System.Collections.Generic.HashSet[string]; Get-ChildItem -Path $TargetFolder | Where-Object { $_.Name -match "^#\d+" } | ForEach-Object { $clean = $_.BaseName; if ($clean -match " -") { $clean = $clean.Substring(0, $clean.IndexOf(" -")) }; $clean = $clean -replace "^#\d+\s+\d{4}_\d{2}_\d{2}", ""; $clean.Trim().Split(" ") | ForEach-Object { $n = $_.Trim(); if ($n -and $n -notmatch "\d" -and $n -notmatch "-") { [void]$NamesFound.Add($n) } } }; foreach ($n in ($NamesFound | Sort-Object)) { [void]$PeopleList.Items.Add($n) }
        $Form.Controls.Add($PeopleList)
        $PeopleList.Add_MouseDoubleClick({ if ($PeopleList.SelectedItem) { $current = $PeopleIn.Text.Trim(); if ($current -eq "") { $PeopleIn.Text = $PeopleList.SelectedItem } elseif ($current -notmatch "\b$([regex]::Escape($PeopleList.SelectedItem))\b") { $PeopleIn.Text = "$current $($PeopleList.SelectedItem)" } }})

        # Preview Label
        &$AddLabel "LIVE PREVIEW:" 430; $PreviewBox = New-Object System.Windows.Forms.Label; $PreviewBox.Top = 455; $PreviewBox.Left = 30; $PreviewBox.Width = 670; $PreviewBox.Height = 50; $PreviewBox.ForeColor = "Blue"; $PreviewBox.Font = $FontPrev; $PreviewBox.BorderStyle = "FixedSingle"; $PreviewBox.TextAlign = "MiddleLeft"; $PreviewBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $Form.Controls.Add($PreviewBox)
        $UpdateBlock = { $j = $JumpIn.Text.Trim(); $c = $ClipIn.Text.Trim(); $suffix = if ($c) { "-$c" } else { "" }; $JumpStr = if ($j) { "#$j$suffix" } else { "" }; $DescStr = if ($DescIn.Text.Trim()) { "-$($DescIn.Text.Trim())" } else { "" }; $raw = "$JumpStr $($DateIn.Text) $($PeopleIn.Text) $DescStr$OriginalExt"; $PreviewBox.Text = ($raw -replace '\s+', ' ' -replace '\s+\.', '.').Trim() }
        $DateIn.Add_TextChanged($UpdateBlock); $JumpIn.Add_TextChanged($UpdateBlock); $ClipIn.Add_TextChanged($UpdateBlock); $PeopleIn.Add_TextChanged($UpdateBlock); $DescIn.Add_TextChanged($UpdateBlock); &$UpdateBlock

        # --- DOCKED FOOTER ---
        $Footer = New-Object System.Windows.Forms.Panel; $Footer.Dock = [System.Windows.Forms.DockStyle]::Bottom; $Footer.Height = 80; $Form.Controls.Add($Footer)
        $SkipBtn = New-Object System.Windows.Forms.Button; $SkipBtn.Text = "SKIP"; $SkipBtn.Top = 15; $SkipBtn.Left = 30; $SkipBtn.Width = 120; $SkipBtn.Height = 50; $SkipBtn.DialogResult = [System.Windows.Forms.DialogResult]::Ignore; $Footer.Controls.Add($SkipBtn)
        $OkBtn = New-Object System.Windows.Forms.Button; $OkBtn.Text = "RENAME"; $OkBtn.Top = 15; $OkBtn.Left = 160; $OkBtn.Width = 540; $OkBtn.Height = 50; $OkBtn.BackColor = "LightGreen"; $OkBtn.Font = $FontBold; $OkBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.AcceptButton = $OkBtn; $OkBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $Footer.Controls.Add($OkBtn)

        # --- SMART FILMSTRIP ---
        $StartLbl = $Config.SkipSeconds; $EndLbl = $Config.SkipSeconds + $Config.WindowSeconds
        &$AddLabel "VIDEO FRAMES (${StartLbl}s to ${EndLbl}s):" 530
        $FlowPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $FlowPanel.Top = 560; $FlowPanel.Left = 30; $FlowPanel.Width = 670; $FlowPanel.Height = $Form.ClientSize.Height - $Footer.Height - 560 - 10
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
        if ($Result -eq "OK") { return @{ Status="RENAME"; FinalName=$PreviewBox.Text; Date=$DateIn.Text; Jump=$JumpIn.Text; Clip=$ClipIn.Text; People=$PeopleIn.Text; Desc=$DescIn.Text } }
        if ($Result -eq "Ignore") { return @{ Status="SKIP" } }
        return $null
    }

    # --- 7. PROCESS LOOP ---
    $Sw = [System.Diagnostics.Stopwatch]::StartNew()
    
    Log-Info "Sorting selected files by Media Created date..."
    $FilesWithDates = @()
    foreach ($File in $RawFiles) {
        $ShellFile = $FolderObj.ParseName($File.Name)
        $RawDate = $FolderObj.GetDetailsOf($ShellFile, $DateIdx) -replace '[^0-9/ :APM]', ''
        $RealDate = if ($RawDate -as [DateTime]) { [DateTime]$RawDate } else { $File.LastWriteTime }
        $FilesWithDates += [PSCustomObject]@{ FileObject = $File; SortDate = $RealDate }
    }
    $Files = $FilesWithDates | Sort-Object SortDate | Select-Object -ExpandProperty FileObject
    
    $LastJump = ""; $LastJumpTime = $null; $LastPeople = ""; $LastDesc = ""
    $NextJob = $null; $BaseTempPath = Join-Path $env:TEMP "SkydivePreviews"
    if (Test-Path $BaseTempPath) { Remove-Item $BaseTempPath -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $BaseTempPath -Force | Out-Null

    for ($i = 0; $i -lt $Files.Count; $i++) {
        $File = $Files[$i]
        Write-Host "----------------------------------------------------" -ForegroundColor Gray
        Log-Info "Processing File [$($i+1)/$($Files.Count)]: $($File.Name)"
        $Sw.Restart()
        $ShellFile = $FolderObj.ParseName($File.Name)
        $RawDate = $FolderObj.GetDetailsOf($ShellFile, $DateIdx) -replace '[^0-9/ :APM]', ''
        $Duration = $FolderObj.GetDetailsOf($ShellFile, $DurIdx)
        
        if ($RawDate -as [DateTime]) {
             $CurrentMediaTime = [DateTime]$RawDate
             $SuggestedDate = $CurrentMediaTime.ToString("yyyy_MM_dd")
        } else {
             $CurrentMediaTime = $File.LastWriteTime 
             $SuggestedDate = ""
        }

        Log-Time "Metadata Read" $Sw
        $Images = @()

        if ($FFmpegAvailable) {
            if ($i -eq 0) {
                Write-Host "      [SYNC] Generating initial thumbnails..." -ForegroundColor Yellow
                $Job = Start-Job -ScriptBlock $PreviewJobScript -ArgumentList $FFmpegPath, $File.FullName, $Duration, $BaseTempPath, "0", $Config.SkipSeconds, $Config.WindowSeconds, $Config.FrameCount, $Config.PreviewWidth
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

        if (($i + 1) -lt $Files.Count -and $FFmpegAvailable) {
            $NextFile = $Files[$i+1]
            $NextShell = $FolderObj.ParseName($NextFile.Name)
            $NextDur = $FolderObj.GetDetailsOf($NextShell, $DurIdx)
            $NextId = ($i + 1).ToString()
            $NextJob = Start-Job -ScriptBlock $PreviewJobScript -ArgumentList $FFmpegPath, $NextFile.FullName, $NextDur, $BaseTempPath, $NextId, $Config.SkipSeconds, $Config.WindowSeconds, $Config.FrameCount, $Config.PreviewWidth
            Write-Host "      [ASYNC] Prefetch started for next file." -ForegroundColor DarkGray
        } else { $NextJob = $null }

        $SuggestedJump = $LastJump; $SuggestedPeople = $LastPeople; $SuggestedDesc = $LastDesc; $SuggestedClip = ""

        if ($null -ne $LastJumpTime) {
            if (($CurrentMediaTime - $LastJumpTime).TotalMinutes -le $Config.JumpGapMinutes) {
                $existing = Get-ChildItem -Path $TargetFolder | Where-Object { $_.Name -like "#$SuggestedJump*" }
                $max = 1
                foreach ($ex in $existing) { if ($ex.Name -match "#$SuggestedJump-(\d+)") { $val = [int]$matches[1]; if ($val -gt $max) { $max = $val } } }
                $SuggestedClip = ($max + 1).ToString()
            } else {
                if ($LastJump -match "^\d+$") { $SuggestedJump = [int]$LastJump + 1 }
                $SuggestedPeople = ""; $SuggestedDesc = ""; $SuggestedClip = ""
            }
        }

        Log-Info "Waiting for user input..."
        $Data = Show-SkydiveForm -FileName $File.Name -FullName $File.FullName -FileTime $CurrentMediaTime.ToString("MMM dd, yyyy @ HH:mm:ss") -Duration $Duration -SuggestedDate $SuggestedDate -SuggestedJump $SuggestedJump -SuggestedClip $SuggestedClip -SuggestedPeople $SuggestedPeople -SuggestedDesc $SuggestedDesc -TargetFolder $TargetFolder -OriginalExt $File.Extension -PreloadedImages $Images -Config $Config
        Log-Time "User Action" $Sw

        if ($Images) { foreach ($img in $Images) { $img.Dispose() } }
        if ($null -eq $Data) { break }
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
    Log-Info "All done. Closing..."
    Start-Sleep -Seconds 1

} catch {
    Write-Host "CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Error Details: $($_.ScriptStackTrace)" -ForegroundColor Yellow
    pause
}

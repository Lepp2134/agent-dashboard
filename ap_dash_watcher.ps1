# AP Dashboard Auto-Updater
# Watches C:\Users\chpace\Documents\AP Dash for new Week *.xlsx files
# Extracts data and updates the dashboard + pushes to GitHub

$watchFolder = "C:\Users\chpace\Documents\AP Dash"
$workspaceFolder = "C:\Users\chpace\Documents\Kiro"
$logFile = "$workspaceFolder\ap_dash_watcher.log"

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts - $msg" | Tee-Object -FilePath $logFile -Append
}

function Get-WeekNumber($filename) {
    if ($filename -match 'Week\s+(\d+)') { return [int]$Matches[1] }
    return $null
}

function Process-WeekFile($filePath) {
    $weekNum = Get-WeekNumber (Split-Path $filePath -Leaf)
    if (-not $weekNum) { Log "Could not determine week number from: $filePath"; return }
    
    $dashFile = "$workspaceFolder\weekly_dashboard.html"
    $text = [System.IO.File]::ReadAllText($dashFile, [System.Text.UTF8Encoding]::new($false))
    
    # Check if this week already exists
    if ($text.Contains("const W$weekNum=[")) {
        Log "Week $weekNum already exists in dashboard, skipping"
        return
    }
    
    Log "Processing Week $weekNum from $filePath"
    
    # Copy file to workspace
    Copy-Item $filePath "$workspaceFolder\Week $weekNum Variable Scorecard.xlsx" -Force
    
    # Use ImportExcel to read the data
    $sheetName = "Week $weekNum"
    try {
        $data = Import-Excel -Path $filePath -WorksheetName $sheetName -StartRow 7 -NoHeader
    } catch {
        # Try without specific sheet name (use first sheet)
        try {
            $sheets = Get-ExcelSheetInfo -Path $filePath
            $sheetName = ($sheets | Where-Object { $_.Name -match "Week" } | Select-Object -First 1).Name
            if (-not $sheetName) { $sheetName = $sheets[0].Name }
            $data = Import-Excel -Path $filePath -WorksheetName $sheetName -StartRow 7 -NoHeader
        } catch {
            Log "ERROR: Could not read Excel file: $_"
            return
        }
    }
    
    # Build W{N} array, PRECALL, POSTCALL, CHC2C entries
    $agentRows = [System.Collections.ArrayList]@()
    $precallEntries = [System.Collections.ArrayList]@()
    $postcallEntries = [System.Collections.ArrayList]@()
    $chc2cEntries = [System.Collections.ArrayList]@()
    
    foreach ($row in $data) {
        $props = $row.PSObject.Properties | Select-Object -ExpandProperty Value
        if ($props.Count -lt 18) { continue }
        
        $login = "$($props[0])".Trim()
        $shift = "$($props[1])".Trim()
        $manager = "$($props[2])".Trim()
        
        if (-not $login -or -not $shift -or $shift -eq ' ') { continue }
        
        # Helper to clean numeric values
        function CleanNum($v) {
            $s = "$v".Trim()
            if ($s -eq '' -or $s -eq ' ' -or $s -eq $null) { return 0 }
            try { return [double]$s } catch { return 0 }
        }
        
        $aht = CleanNum $props[15]      # Col P - AHT
        $wimsHr = CleanNum $props[6]    # Col G - C2C WIMs/HR
        $qs = CleanNum $props[21]       # Col V - Quality Score
        $atrDef = CleanNum $props[18]   # Col S - ATR Defect %
        $pctAvg = CleanNum $props[20]   # Col U - Percent to AVG
        $precall = CleanNum $props[16]  # Col Q - Pre-Call
        $postcall = CleanNum $props[17] # Col R - Post-Call
        $chc2c = CleanNum $props[4]     # Col E - WIMS/HR CH/C2C
        
        $feedback = "$($props[22])".Trim()
        if ($feedback -eq ' ' -or $feedback -eq $null) { $feedback = '' }
        $feedback = $feedback.TrimEnd(' ,')
        $feedback = $feedback -replace '"', "'"
        
        [void]$agentRows.Add("[`"$login`",`"$shift`",`"$manager`",$aht,$wimsHr,$qs,$atrDef,$pctAvg,`"$feedback`"]")
        [void]$precallEntries.Add("`"$login`":$precall")
        [void]$postcallEntries.Add("`"$login`":$postcall")
        [void]$chc2cEntries.Add("`"$login`":$chc2c")
    }
    
    if ($agentRows.Count -eq 0) {
        Log "ERROR: No valid agent data found in Week $weekNum"
        return
    }
    
    Log "Extracted $($agentRows.Count) agents for Week $weekNum"
    
    # Insert W{N} array before "const WEEKS="
    $wArray = "const W$weekNum=[`r`n" + ($agentRows -join ",`r`n") + "`r`n];`r`n`r`n"
    $text = $text.Replace("const WEEKS=", "$wArray`const WEEKS=")
    
    # Update WEEKS object
    $text = $text -replace "(const WEEKS=\{[^}]+)\}", "`$1,$($weekNum):W$weekNum}"
    
    # Update WEEK_LABELS
    $text = $text -replace "(const WEEK_LABELS=\{[^}]+)\}", "`$1,$($weekNum):'Week $weekNum'}"
    
    # Add to PRECALL
    $precallStr = "$($weekNum):{" + ($precallEntries -join ',') + "}"
    $text = $text -replace "(const PRECALL=\{.*?)(};)", "`$1,$precallStr`$2"
    
    # Add to POSTCALL
    $postcallStr = "$($weekNum):{" + ($postcallEntries -join ',') + "}"
    $text = $text -replace "(const POSTCALL=\{.*?)(};)", "`$1,$postcallStr`$2"
    
    # Add to CHC2C
    $chc2cStr = "$($weekNum):{" + ($chc2cEntries -join ',') + "}"
    $text = $text -replace "(const CHC2C=\{.*?)(};)", "`$1,$chc2cStr`$2"
    
    # Write dashboard
    [System.IO.File]::WriteAllText($dashFile, $text, [System.Text.UTF8Encoding]::new($false))
    
    # Rebuild index.html
    $gate = @'
<div id="authGate" style="position:fixed;top:0;left:0;right:0;bottom:0;background:#0f172a;z-index:9999;display:flex;align-items:center;justify-content:center">
<div style="background:#1e293b;border:1px solid #334155;border-radius:16px;padding:40px;text-align:center;max-width:360px;width:90%">
<div style="font-size:20px;font-weight:700;color:#f1f5f9;margin-bottom:6px">Agent Performance Dashboard</div>
<div style="font-size:12px;color:#64748b;margin-bottom:24px">Enter password to continue</div>
<input type="password" id="authPw" placeholder="Password" style="width:100%;padding:10px 14px;background:#0f172a;border:1px solid #334155;border-radius:8px;color:#f1f5f9;font-size:14px;outline:none;margin-bottom:12px;box-sizing:border-box" onkeydown="if(event.key==='Enter')checkAuth()">
<button onclick="checkAuth()" style="width:100%;padding:10px;background:#3b82f6;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer">Enter</button>
<div id="authErr" style="color:#ef4444;font-size:12px;margin-top:10px;display:none">Incorrect password</div>
</div>
</div>
<script>
function checkAuth(){var p=document.getElementById('authPw').value;if(p==='DSS-PHX10'){document.getElementById('authGate').remove();sessionStorage.setItem('authed','1')}else{document.getElementById('authErr').style.display='block'}}
if(sessionStorage.getItem('authed')==='1'){document.addEventListener('DOMContentLoaded',function(){var g=document.getElementById('authGate');if(g)g.remove()})}
</script>
'@
    $idxText = $text.Replace('<body>', "<body>`n$gate")
    [System.IO.File]::WriteAllText("$workspaceFolder\index.html", $idxText, [System.Text.UTF8Encoding]::new($false))
    
    # Git commit and push
    Push-Location $workspaceFolder
    git add weekly_dashboard.html index.html
    git commit -m "Auto-add Week $weekNum data"
    git push
    Pop-Location
    
    Log "Week $weekNum successfully added and pushed to GitHub"
}

# --- Main watcher loop ---
Log "AP Dashboard Watcher started. Watching: $watchFolder"
Log "Press Ctrl+C to stop"

$processed = @{}

while ($true) {
    $files = Get-ChildItem "$watchFolder\Week*.xlsx" -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        if (-not $processed.ContainsKey($f.FullName)) {
            # Wait a moment for file to finish writing
            Start-Sleep -Seconds 3
            Log "New file detected: $($f.Name)"
            Process-WeekFile $f.FullName
            $processed[$f.FullName] = $true
        }
    }
    Start-Sleep -Seconds 10
}

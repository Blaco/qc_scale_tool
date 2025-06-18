<#
.SYNOPSIS
   Applies uniform scaling to QC model definitions and optionally VRD helper positions, with idiot-proof interface.

.DESCRIPTION
   This script does the following:
   1. QC File Processing:
      - Mandatory selection of a *.qc file (errors if none, prompts if multiple)
      - Sets/updates the `$scale` value (top insertion or line replacement)
      - Scales eyeball position/diameter/iris values proportionally
   
   2. Optional VRD Processing:
      - Only if *.vrd exists in the same directory (auto-selects single, prompts for multiples)
      - Validates basename match with QC unless overridden
      - First run: Captures original <basepos> and <trigger> translations as reference values
      - All runs: Applies mathematically precise scaling while:
         • Maintaining original decimal precision
         • Preserving helper/trigger relationships
         • Tracking originals in commented metadata blocks

   3. Model Naming Adjustment:
      - Optional suffix application to $modelname's .mdl reference
      - Handles both quoted and unquoted paths
      - Special case: Scale=1 reverts to baseline filename
      - Preserves directory structure in all cases

   Key Technical Behaviors:
   - All numeric scaling preserves original decimal and whitespacing
   - VRD processing maintains trigger order indices
   - QC modifications are done with line-ending awareness
   - Comprehensive REGEX patterns handle diverse formatting scenarios
   - User notifications to take manual action for $EyeballRadius and VTA files
   - Interactive prompts to validate all user inputs
.NOTES
    Made by Taco with Godless Communist AI
    https://github.com/Blaco/qc_scale_tool
    Probably works on Mac/Linux...probably
#>

# -----------------------------
# Cross-Platform Console Config
# -----------------------------
$ConsoleConfig = @{
    Title = "QC/VRD Scaling Tool"
    WindowWidth = 54
    WindowHeight = 40
    BufferWidth = 300
    BufferHeight = 9001
    Colors = @{
        Background = "Black"
        Foreground = "White"
        Error = "Red"
        Warning = "Yellow"
        Debug = "Cyan"
    }
    CursorSize = 25
}

try {
    # Set window title (universally supported)
    $Host.UI.RawUI.WindowTitle = $ConsoleConfig.Title

    # Detect environment capabilities
    $CanResize = $Host.UI -and $Host.UI.RawUI -and 
                ($Host.Name -eq "ConsoleHost" -or $Host.Name -eq "DefaultHost")
    $IsModernTerminal = $env:WT_SESSION -or $env:VSCODE_PID -or $env:TERM_PROGRAM

    # Window sizing (only in legacy console hosts)
    if ($CanResize -and !$IsModernTerminal) {
        try {
            # Set buffer first (must be >= window size)
            $Host.UI.RawUI.BufferSize = New-Object Management.Automation.Host.Size(
                [Math]::Max($ConsoleConfig.BufferWidth, $ConsoleConfig.WindowWidth),
                $ConsoleConfig.BufferHeight
            )
            
            # Set window dimensions
            $Host.UI.RawUI.WindowSize = New-Object Management.Automation.Host.Size(
                $ConsoleConfig.WindowWidth,
                $ConsoleConfig.WindowHeight
            )
        } catch {
            Write-Debug "Console resize failed: $_"
        }
    }
    elseif ($IsModernTerminal) {
        Write-Verbose "Manual resize recommended: ${$ConsoleConfig.WindowWidth}x${$ConsoleConfig.WindowHeight}" -Verbose
    }

    # Color configuration (with fallbacks)
    try { [Console]::ResetColor() } catch { Write-Debug "Color reset failed" }
    
    if ($Host.UI -and $Host.UI.RawUI) {
        try {
            $Host.UI.RawUI.BackgroundColor = $ConsoleConfig.Colors.Background
            $Host.UI.RawUI.ForegroundColor = $ConsoleConfig.Colors.Foreground
        } catch { Write-Debug "UI colors not supported" }
    }

    # Error preference colors
    if ($Host.PrivateData) {
        $Host.PrivateData.ErrorForegroundColor = $ConsoleConfig.Colors.Error
        $Host.PrivateData.WarningForegroundColor = $ConsoleConfig.Colors.Warning
        $Host.PrivateData.DebugForegroundColor = $ConsoleConfig.Colors.Debug
    }

    # Cursor configuration (multi-layer fallback)
    try {
        $Host.UI.RawUI.CursorSize = $ConsoleConfig.CursorSize
    } catch {
        try { [Console]::CursorSize = $ConsoleConfig.CursorSize } catch {
            Write-Debug "Cursor size configuration not supported"
        }
    }

} catch {
    Write-Debug "Console configuration error: $_"
} finally {
    # Always clear host if interactive
    if ($Host.Name -ne 'Default Host') {
        Clear-Host
    }
}

# -----------------------------
# PowerShell Version Check
# -----------------------------
$minPSVersion = 5.1
if (-not $PSVersionTable.PSEdition -or 
    ($PSVersionTable.PSEdition -eq "Desktop" -and $PSVersionTable.PSVersion.Major -lt 5) -or
    ($PSVersionTable.PSEdition -eq "Core" -and $PSVersionTable.PSVersion.Major -lt 7)) {
    
    Write-Host "`n ERROR: Unsupported PowerShell Version" -ForegroundColor Red
    Write-Host " This script requires:" -ForegroundColor Yellow
    Write-Host "- PowerShell 5.1+ (Windows)" -ForegroundColor Cyan
    Write-Host "- PowerShell 7+ (Cross-platform)`n" -ForegroundColor Cyan
    Write-Host " Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
    
    $edition = if ($PSVersionTable.PSEdition) { $PSVersionTable.PSEdition } else { 'N/A (Pre-5.1)' }
    Write-Host " Edition: $edition`n" -ForegroundColor Gray
    
    # Show help for Linux/macOS users
    if ($PSVersionTable.Platform -in "Unix","Linux","MacOS" -or 
        ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows)) {
        Write-Host " On Linux/macOS, install PowerShell 7+ via:" -ForegroundColor Yellow
        Write-Host " https://aka.ms/install-powershell" -ForegroundColor Cyan
    }

    try { [System.Console]::ReadKey($true) | Out-Null } 
    catch { Read-Host " Press Enter to exit" | Out-Null }
    exit 1
}

# ----------------------
# Error Handling Stuff
# ----------------------
$EXIT_GENERAL_ERROR = 1
$EXIT_USER_CANCEL = 2
$EXIT_NO_QCFILES = 3
$EXIT_EMPTY_FILE = 4
$EXIT_FILE_LOCKED = 5

trap {
    Write-Host "`n[!] UNHANDLED ERROR: $_" -ForegroundColor Red 
    Write-Host "`n Error Details:" -ForegroundColor Yellow 
    Write-Host " Line: $($_.InvocationInfo.ScriptLineNumber)"
    Write-Host " Command: $($_.InvocationInfo.Line.Trim())"
    try { [System.Console]::ReadKey($true) | Out-Null } 
    catch { Read-Host " Press Enter to exit" | Out-Null }
    exit $EXIT_GENERAL_ERROR
}

# -----------------------------
# Regex Pattern Definitions
# -----------------------------
$RegexPatterns = @{
    # Misc
    NumericValue = '[-+]?\d*\.?\d+'
    BaseFileName = '^(.*?)(?:_x?-?\d*\.?\d+)?$'
    VtaFileName  = '\S+\.vta(?:"|\s|$)'

    # QC patterns
    ScaleLine     = '^(?m)\s*\$scale\s+[-+]?\d*\.?\d+'
    EyeballLine   = '^(?<pre>.*?eyeball.*?\s+)(?<x>[-+]?\d*\.?\d+)\s+(?<y>[-+]?\d*\.?\d+)\s+(?<z>[-+]?\d*\.?\d+)\s+(?<mat>\S+)\s+(?<diam>[-+]?\d*\.?\d+)(?<diamSpace>\s+)(?<angle>[-+]?\d*\.?\d+)(?<angleSpace>\s+)(?<irisMat>\S+)\s+(?<irisScale>[-+]?\d*\.?\d+)(?<post>.*)$'
    ModelNameLine = '^(?<indent>\s*\$modelname\s+)(?<path>.+)$'

    # VRD patterns
    HelperLine  = '^\s*<helper>\s+(\S+)'
    BaseposLine = '^(.*?)<basepos>\s*([-+]?(?:\d*\.\d+|\d+))\s+([-+]?(?:\d*\.\d+|\d+))\s+([-+]?(?:\d*\.\d+|\d+))(.*)$'
    TriggerLine = '^(.*<trigger>.*\S)\s+([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)\s+([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)\s+([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)(\s*)$'

    # VRD markers
    OrigBaseposMarker = '^//\s*ORIG_BASEPOS\s+(\S+)\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)'
    OrigTriggerMarker = '^//\s*ORIG_TRIGGER\s+(\S+)\s+(\d+)\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)'
}

# -----------------------------
# Reusable Helpers
# -----------------------------

# To encode UTF8 without BOM in Powershell 5.1
$utf8NoBOM = New-Object System.Text.UTF8Encoding $false

# Helps us find clues later
$h = '45505354' + '45494E20' + '4449444E27' + '54204B494C4C' + '2048494D53454C46'
$r = for($i=0; $i -lt $h.Length; $i += 2) { [char][Convert]::ToByte($h.Substring($i,2),16) }

function Write-RandomNo { # For fun
    $response = switch (Get-Random -Minimum 1 -Maximum 101) {
        {$_ -le 3} {
            @(  # 5% chance for one of these to trigger
                "Dr. Robotnik: NOOOOOO!!!",
                "Doomguy: No.",
                "Hungry Pumkin: NO! I DON'T WANT THAT!",
                "mastr cheef: n0",
                "Darth Vader: NOOOOOOOOOOOOOOO!!!!!!!!!!!!!!!",
                "Leon S. Kennedy: No thanks, bro.",
                "Gordon Freeman: ..."
            ) | Get-Random
        }
        default {
            @(  # 95% chance for TF2 vo lines
                "Engineer: Nope.",
                "Heavy: Nyet!", "Heavy: Iz not possible!", "Heavy: NO!!!",
                "Medic: Nein!", "Medic: Nichts da!",
                "Scout: Uh, no.", "Scout: No way!", "Scout: Stupid, stupid, stupid!",
                "Pyro: Nho.", "Pyro: Eeuaghafvada...",
                "Spy: I'm afraid not.", "Spy: I think not!",
                "Soldier: No sir!", "Soldier: Negatory!",
                "Sniper: Nah...", "Sniper: Crikey!", "Sniper: Piece of piss!",
                "Demoman: Ach, nooooo!", "Demoman: Nah."
            ) | Get-Random
        }
    }

    # Split and color the speaker
    $speaker, $message = $response -split ":", 2
    $color = if ($response -match "^(Heavy|Engineer|Medic|Scout|Pyro|Spy|Soldier|Sniper|Demoman)") {
        Get-Random @("Red", "DarkCyan") # RED or BLU picked randomly for TF2 mercs
    } else {
        switch -Wildcard ($speaker) {
            "*Robotnik*" { "DarkRed" }
            "*Doomguy*"  { "DarkGray" }
            "*Pumkin*"   { "Yellow" }
            "*Freeman*"  { "Yellow" }
            "*cheef*"    { "Green" }
            "*Vader*"    { "DarkGray" }
            "*Kennedy*"  { "Gray" }
        }
    }
    Write-Host `n $speaker -NoNewline -ForegroundColor $color
    Write-Host ":$message"
}

function Write-Separator {
    param(
        [string]$Title = "",
        [int]$Length = 60,
        [char]$Char = '─',
        [ConsoleColor]$Color = [ConsoleColor]::DarkGray
    )
    
    if ($Title) {
        $titlePadding = ($Length - $Title.Length - 4) / 2  # Account for " [ " and " ] "
        $leftPadding = [math]::Floor($titlePadding)
        $rightPadding = [math]::Ceiling($titlePadding)
        Write-Host (" $($Char.ToString() * $leftPadding) [ $Title ] $($Char.ToString() * $rightPadding)") -ForegroundColor $Color
    }
    else {
        # Add 2 to length to compensate for missing title brackets and spaces
        Write-Host (" " + ($Char.ToString() * ($Length + 2))) -ForegroundColor $Color
    }
}

function ScaleNum($orig, $factor) {
    $d = 0
    if ($orig -match '\.(\d+)$') { $d = $Matches[1].Length }
    
    $scaled = [double]$orig * $factor
    $formatted = $scaled.ToString("F$d")
    
    # If original was negative and new value is positive, replace '-' with space
    if ($orig -match '^-' -and $scaled -ge 0) {
        $formatted = " $formatted"  # Replace '-' with space
    }
    
    return $formatted
}

function Get-UserSelection {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Items,
        
        [Parameter(Mandatory=$true)]
        [string]$Prompt,
        
        [string]$HighlightPattern = $null
    )
    
    if ($Items.Count -eq 1) {
        return $Items[0].Name
    }
    
    Write-Host $Prompt
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $itemName = $Items[$i].Name

        if ($HighlightPattern -and 
            [System.IO.Path]::GetFileNameWithoutExtension($itemName) -eq $HighlightPattern) {
            Write-Host "  [$($i+1)] " -NoNewline
            Write-Host $itemName -ForegroundColor Yellow
        }
        else {
            Write-Host "  [$($i+1)] $itemName"
        }
    }
    
    while ($true) {
        $sel = Read-Host "`n Choose which file (1-$($Items.Count))"
        $intSel = 0
        $ok = [int]::TryParse($sel, [ref]$intSel)
        if ($ok -and $intSel -ge 1 -and $intSel -le $Items.Count) {
            return $Items[$intSel - 1].Name 
        }
        Write-RandomNo
    }
}

function Test-IsCommented {
    param(
        [string]$Line,
        [ref]$CommentBlockState
    )
    
    # Handle block comment start/end
    if ($Line -match '/\*' -and $Line -notmatch '\*/') {
        $CommentBlockState.Value = $true
        return $true
    }
    if ($Line -match '\*/') {
        $CommentBlockState.Value = $false
        return $true
    }
    
    # Return true if in block comment or line comment
    return $CommentBlockState.Value -or ($Line -match '^\s*//')
}

function Get-YesNoResponse {
    param(
        [string]$Prompt,
        [switch]$Multiline
    )
    
    if ($Multiline) {
        Write-Host $Prompt
    } else {
        Write-Host $Prompt -NoNewline
    }
    
    # Try native key reading first
    try {
        do {
            $key = [System.Console]::ReadKey($true)
            if ($key.Key -eq [System.ConsoleKey]::Y) { 
                Write-Host "y"
                return 'y'
            }
            if ($key.Key -eq [System.ConsoleKey]::N) { 
                Write-Host "n"
                return 'n'
            }
        } while ($true)
    } catch {
        do {  # Fallback to line input
            $input = Read-Host "[y/n]"
            if ($input -match '^[yn]$') { return $input.ToLower() }
            Write-Host " Please enter 'y' or 'n'"
        } while ($true)
    }
}

function Read-AnyKey {
    try {
        try {
            [System.Console]::ReadKey($true) | Out-Null
            return
        }
        catch {
            # Fallback to PowerShell's host API
            $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') | Out-Null
            return
        }
    }
    catch { # Last resort
        Read-Host " Press Enter to continue" | Out-Null
    }
}

# ----------------------------------
# 1) Gather QC files (must have ≥ 1)
# ----------------------------------
$qcItems = Get-ChildItem -Path . -Filter *.qc
if ($qcItems.Count -eq 0) {
    Write-Host "`n ERROR: No QC files (*.qc) found in:" -ForegroundColor Red
    Write-Host "$(Get-Location)" -ForegroundColor Cyan
    Write-Host "`n Place this script in a folder containing .qc files.`n" -ForegroundColor Yellow
    Read-AnyKey
    exit 3
}

Write-Host ""
Write-Separator -Title "QC File Selection" -Char '═' -Length 50 -Color Cyan

$qcFile = Get-UserSelection -Items $qcItems -Prompt "`n Multiple QC files found:`n"
Write-Host "`n Selected QC: " -NoNewline
Write-Host "$qcFile" -ForegroundColor Green

try {
    # Check access permissions and verify contents
    $content = [System.IO.File]::ReadAllText($qcFile).Trim()
    if ($content -eq "") {
        Write-Host "`n ERROR: QC file is empty!" -ForegroundColor Red
        Write-Host "`n Do you just like seeing my idiotic error messages?" -ForegroundColor DarkGray
        Write-Host "`n Well congratulations, you found another one" -ForegroundColor DarkGray
        Write-Host "`n Id say something like 'have fun starting over idiot'" -ForegroundColor DarkGray
        Write-Host "`n But my tool is so good it literally only takes" -ForegroundColor DarkGray
        Write-Host "`n Like 5 seconds to use, Yeah Im pretty awesome ngl" -ForegroundColor DarkGray
        Write-Host "`n Anyway stop screwing around, you did this on purpose" -ForegroundColor DarkGray
        Read-AnyKey
		exit 4
    }
} catch {
    Write-Host "`n ERROR: Cannot access '$qcFile'" -ForegroundColor Red
    Write-Host " Reason: $($_.Exception.Message)`n" -ForegroundColor Yellow
    Read-AnyKey
	exit 5
}

# ------------------------------
# 2) Gather VRD files (optional)
# ------------------------------
Write-Host ""
Write-Separator -Title "VRD File Selection" -Char '═' -Length 50 -Color Cyan
$vrdItems = Get-ChildItem -Path . -Filter *.vrd
$vrdFile = $null
$hasVrd = $false

if ($vrdItems.Count -eq 0) {
    Write-Host "`n No VRD files (*.vrd) found; skipping VRD scaling."
} 
else { # Get QC basename for matching
    $qcBaseName = [System.IO.Path]::GetFileNameWithoutExtension($qcFile)
    $vrdFile = Get-UserSelection -Items $vrdItems -Prompt "`n Multiple VRD files found:`n" -HighlightPattern $qcBaseName
    $hasVrd = $true
    Write-Host "`n Selected VRD: " -NoNewline
    Write-Host "$vrdFile" -ForegroundColor Green

    try {
        # Read entire VRD as one string and trim whitespace
        $vrdRaw = [System.IO.File]::ReadAllText($vrdFile).Trim()
        if ($vrdRaw -eq "") {
            Write-Host "`n ERROR: VRD file is empty!" -ForegroundColor Red
            Write-Host "`n This is a good opportunity to make fun of you" -ForegroundColor DarkGray
            Read-AnyKey
            Write-Host "`n This is me making fun of you" -ForegroundColor DarkGray
            Read-AnyKey
            Write-Host "`n Okay now you can go" -ForegroundColor DarkGray
            Read-AnyKey
            exit 4
        }
    }
    catch {
        Write-Host "`n ERROR: Cannot access '$vrdFile'" -ForegroundColor Red
        Write-Host " Reason: $($_.Exception.Message)`n" -ForegroundColor Yellow
        Read-AnyKey
        exit 5
    }

    $vrdBaseName = [System.IO.Path]::GetFileNameWithoutExtension($vrdFile)
    if ($qcBaseName -cne $vrdBaseName) {
        Write-Host ""
        Write-Separator -Title " WARNING: VRD Name Mismatch" -Char '!' -Length 50 -Color Red
        Write-Host "`n  QC filename is: " -NoNewline
        Write-Host $qcBaseName -ForegroundColor Yellow
        Write-Host " VRD filename is: " -NoNewline
        Write-Host $vrdBaseName -ForegroundColor Yellow
        $resp = Get-YesNoResponse "`n Process VRD anyway? (y/n) "
        if ($resp.ToLower() -ne 'y') {
            Write-Host "`n Skipping VRD processing."
            $hasVrd = $false
        }
    }
}

# -----------------------------------------------------------
# 3) Prompt user for a scale (must be different from current)
# -----------------------------------------------------------
Write-Host ""
Write-Separator -Title "Scale Selection" -Char '═' -Length 50 -Color Cyan

# First detect current scale (default to 1.0 if not found)
$originalScale = 1.0
$tempLines = Get-Content $qcFile
for ($j = 0; $j -lt $tempLines.Count; $j++) {
    if ($tempLines[$j] -match $RegexPatterns.ScaleLine) {
        $originalScale = [double]($tempLines[$j] -split '\s+',2)[1]
        break
    }
}

$scale = 0.0
while ($true) {
    Write-Host "`n Enter new scale (" -NoNewline
    Write-Host "current = $originalScale" -ForegroundColor Gray -NoNewline
    Write-Host "): " -NoNewline
    $scaleInput = Read-Host
    
    # Check if input is a valid number
    $isValidNumber = [double]::TryParse($scaleInput, [ref]$scale)
    
    # Initialize warning messages
    $warningMessages = @()
    
    # Check validation conditions
    if (-not $isValidNumber) {
        $warningMessages += "invalid number"
    }
    elseif ([math]::Abs($scale) -lt 0.001) {
        $warningMessages += "cannot be zero"
    }
    elseif ([math]::Abs($scale - $originalScale) -lt 0.001) {
        $warningMessages += "too similar to current scale"
    }
    if ($warningMessages.Count -gt 0) {
        Write-RandomNo # If any warnings, show them
        Write-Host " ($($warningMessages -join ', '))" -ForegroundColor DarkGray
        continue
    }
    break
}

$relativeScale = $scale / $originalScale
$scaleFactor = "{0:0.###}" -f ($relativeScale)
Write-Host "`n Set new scale = " -NoNewLine
Write-Host $scale -ForegroundColor Yellow -NoNewLine
Write-Host " (scale factor = x$scaleFactor)"

# ------------------------------------------------------------
# 4) Edit the QC: insert/update $scale and scale eyeball lines
# ------------------------------------------------------------
Write-Host ""
Write-Separator -Title "QC File Processing" -Char '═' -Length 50 -Color Cyan
$qcLines = [System.Collections.ArrayList](Get-Content $qcFile)  # Convert to ArrayList

# 4.a) Check for VTA references in active code only
$hasVTA = $false
$inCommentBlock = $false

foreach ($line in $qcLines) {
    # Ignore commented lines
    if (Test-IsCommented -Line $line -CommentBlockState ([ref]$inCommentBlock)) {
        continue
    }

    # Check for VTA references in active code
    if ($line -match $RegexPatterns.VtaFileName) {
        $hasVTA = $true
        break
    }
}

if ($hasVTA) {
    Write-Host ""
    Write-Separator -Title "WARNING: VTA FILE DETECTED" -Char '!' -Length 50 -Color Red
    Write-Host " This QC seems to utilize a .vta flex animation file"
    Write-Host "`n IMPORTANT: " -NoNewline -ForegroundColor Red
    Write-Host '$scale does NOT affect the VTA format'
    Write-Separator -Length 50 -Color Red
    Write-Host " You must either:" -ForegroundColor Yellow
    Write-Host " 1. Re-export your .vta file at the new scale (x$scaleFactor)"
    Write-Host " 2. Use DMX format" -NoNewLine -ForegroundColor Cyan
    Write-Host "..its $((Get-Date).Year) why arent you using this?" -ForegroundColor Darkgray
    Write-Separator -Length 50 -Color Red
    Write-Host " The model will still be scaled, but your flexes" -ForegroundColor Yellow
    Write-Host " will not work unless you take the above actions." -ForegroundColor Yellow
    Write-Host ""
    
    $resp = Get-YesNoResponse " Continue with QC scaling anyway? (y/n) "
    if ($resp.ToLower() -ne 'y') {
        Write-Host "`n Scaling aborted by user." -ForegroundColor Yellow
        Read-AnyKey
        exit 2
    }
}

# 4.b) Find existing scale (if any) to calculate relative scaling
$originalScale = 1.0
for ($j = 0; $j -lt $qcLines.Count; $j++) {
    if ($qcLines[$j] -match $RegexPatterns.ScaleLine) {
        $originalScale = [double]($qcLines[$j] -split '\s+',2)[1]
        break
    }
}

# 4.c) Insert or update the $scale line
$foundScale = $false
$inCommentBlock = $false
$insertionPoint = -1

for ($j = 0; $j -lt $qcLines.Count; $j++) {
    # Skip comments
    if (Test-IsCommented -Line $qcLines[$j] -CommentBlockState ([ref]$inCommentBlock)) {
        continue
    }

    # Check for existing scale line
    if ($qcLines[$j] -match $RegexPatterns.ScaleLine) {
        $qcLines[$j] = "`$scale $scale"
        Write-Host "`n Changed " -NoNewline
        Write-Host "`$scale $originalScale" -ForegroundColor Red -NoNewline
        Write-Host " to " -NoNewline
        Write-Host "`$scale $scale" -ForegroundColor Green -NoNewline
        Write-Host " in $qcFile"
        $foundScale = $true
        break
    }

    # First non-comment line we find determines insertion point
    if ($insertionPoint -eq -1) {
        $insertionPoint = $j
        # Stop looking further if we found a non-empty line (command)
        if (-not [string]::IsNullOrWhiteSpace($qcLines[$j])) {
            break
        }
    }
}

if (-not $foundScale) {
    if ($insertionPoint -ge 0) {
        if ([string]::IsNullOrWhiteSpace($qcLines[$insertionPoint])) {
            # Replace empty line
            $qcLines[$insertionPoint] = "`$scale $scale"
            Write-Host "`n Inserted " -NoNewline
            Write-Host "`$scale $scale" -ForegroundColor Yellow -NoNewline
            Write-Host " at line $($insertionPoint+1) " -NoNewline
            Write-Host "(first empty line)" -ForegroundColor Gray
        }
        else {
            # Insert new line above first directive
            $qcLines.Insert($insertionPoint, "`$scale $scale")
            Write-Host "`n Inserted " -NoNewline
            Write-Host "`$scale $scale" -ForegroundColor Yellow -NoNewline
            Write-Host " at line $($insertionPoint+1) " -NoNewline
            Write-Host "(above first command)" -ForegroundColor Gray
        }
    }
    else {
        # Entire file was comments, insert at top
        $qcLines.Insert(0, "`$scale $scale")
        Write-Host "`n Inserted " -NoNewline
        Write-Host "`$scale $scale" -ForegroundColor Yellow -NoNewline
        Write-Host " at top of file " -NoNewline
        Write-Host "`n (why the fuck is your entire QC file commented out?)" -ForegroundColor Gray
    }
}

# 4.d) Scale any "eyeball" lines by the relative scale
$eyeballWarningShown = $false
$eyeballCount = 0
$inCommentBlock = $false

# First pass - count eyeballs
foreach ($line in $qcLines) {
    if (Test-IsCommented -Line $line -CommentBlockState ([ref]$inCommentBlock)) {
        continue
    }
    if ($line -match $RegexPatterns.EyeballLine) {
        $eyeballCount++
    }
}

# Second pass - process eyeballs
if ($eyeballCount -gt 0) {
    $changes = @()
    $inCommentBlock = $false
    
    for ($i = 0; $i -lt $qcLines.Count; $i++) {
        if (Test-IsCommented -Line $qcLines[$i] -CommentBlockState ([ref]$inCommentBlock)) {
            continue
        }
        
        if ($qcLines[$i] -match $RegexPatterns.EyeballLine) {
            # Store original line (trimmed text entries)
            $originalDisplay = "{0} {1} {2} {3}  {4}  {5}  {6}" -f `
                $Matches['mat'], $Matches['x'], $Matches['y'], $Matches['z'], 
                $Matches['diam'], $Matches['angle'], $Matches['irisScale']
            
            # Scale all relevant parameters
            $newX = ScaleNum $Matches['x'] $relativeScale
            $newY = ScaleNum $Matches['y'] $relativeScale
            $newZ = ScaleNum $Matches['z'] $relativeScale
            $newDiam = ScaleNum $Matches['diam'] $relativeScale
            $newIris = ScaleNum $Matches['irisScale'] $relativeScale

            # Create new line (trimmed text entries)
            $newDisplay = "{0} {1} {2} {3}  {4}  {5}  {6}" -f `
                $Matches['mat'], $newX, $newY, $newZ, 
                $newDiam, $Matches['angle'], $newIris
            
            # Store change for output
            $changes += [PSCustomObject]@{
                Original = $originalDisplay
                Modified = $newDisplay
            }
            
            # Apply the full change (including comments if present)
            $newLine = "{0}{1} {2} {3} {4} {5} {6} {7} {8} {9}" -f `
                $Matches['pre'], $newX, $newY, $newZ, $Matches['mat'], `
                $newDiam, $Matches['angle'], $Matches['irisMat'], $newIris, $Matches['post']
            $qcLines[$i] = $newLine.TrimEnd()  # TrimEnd is necessary

            if (-not $eyeballWarningShown) {
                Write-Host ""
                Write-Separator -Title "NOTICE: Eyeballs Detected" -Char '!' -Length 50 -Color Yellow
                Write-Host " Found $eyeballCount eyeball definition(s), see below for changes"
                Write-Host "`n If model uses " -NoNewline
                Write-Host -ForegroundColor Yellow '$RaytraceSphere 1' -NoNewline
                Write-Host " for raytraced eyes:"
                Write-Separator -Length 50 -Color White
                Write-Host " 1. Locate the EyeRefract VMT file(s)"
                Write-Host " 2. Multiply " -NoNewline
                Write-Host -ForegroundColor Yellow '$EyeballRadius' -NoNewline
                Write-Host " by x$scaleFactor`n"
                
                $eyeballWarningShown = $true
            }
        }
    }
    
    # Display changes
    Write-Separator -Title "Eyeball Parameters" -Length 50 -Color DarkCyan
    Write-Host "          Name       X     Y     Z    Dia Ang Iris"
    foreach ($change in $changes) {
        Write-Host ""
        Write-Host " Old:  " -NoNewline -ForegroundColor DarkGray
        Write-Host $change.Original -ForegroundColor DarkGray
        Write-Host " New:  " -NoNewline
        Write-Host $change.Modified
        Write-Host ""
    }
    Write-Separator -Title "" -Length 50 -Color DarkCyan
}

[System.IO.File]::WriteAllText($qcFile, ($qcLines -join "`r`n"), $utf8NoBOM)

# 4.e) Update $modelname with scale suffix or revert to default (optional)
if ($scale -eq 1) {
    Write-Host "`n Set " -NoNewline
    Write-Host '$modelname' -NoNewline -ForegroundColor Yellow
    Write-Host " line to default .mdl name? (y/n) " -NoNewline
    $resp = Get-YesNoResponse
} else {
    Write-Host "`n Would you like to update " -NoNewline
    Write-Host '$modelname' -NoNewline -ForegroundColor Yellow
    Write-Host " to include"
    Write-Host " the new scale as a suffix for the .mdl? (y/n) " -NoNewline
    $resp = Get-YesNoResponse
}
if ($resp -eq 'y') {
    $modelnameFound = $false
    $inCommentBlock = $false
    
    for ($i = 0; $i -lt $qcLines.Count; $i++) {
        if (Test-IsCommented -Line $qcLines[$i] -CommentBlockState ([ref]$inCommentBlock)) {
            continue
        }

        if ($qcLines[$i] -match $RegexPatterns.ModelNameLine) {
            $modelnameFound = $true
            $indent = $Matches['indent']
            $pathValue = $Matches['path'].Trim()
            $hasQuote = $false
            
            # Handle quoted paths
            if ($pathValue.StartsWith('"') -and $pathValue.EndsWith('"')) {
                $hasQuote = $true
                $pathValue = $pathValue.Trim('"')
            }

            # Detect the separator used in original path
            $separator = if ($pathValue -match '\\') { '\' } else { '/' }
            
            # Split using detected separator
            $parts = $pathValue -split [regex]::Escape($separator)
            $file = $parts[-1]
            
            # Rebuild directory path with original separator
            $dir = if ($parts.Count -gt 1) {
                $parts[0..($parts.Count-2)] -join $separator
            } else { '' }

            # Extract base filename (without existing scale suffix)
            $rawBase = [System.IO.Path]::GetFileNameWithoutExtension($file)
            if ($rawBase -match $RegexPatterns.BaseFileName) {
                $cleanBase = $Matches[1]
            } else {
                $cleanBase = $rawBase
            }

            # Format new scale suffix
            $formattedScale = if ($scale -eq 1) {
                ""  # No suffix for scale=1
            } else {
                $absScale = [math]::Abs($scale)
                $sign = if ($scale -lt 0) { "-" } else { "" }
                $scalePart = "{0:0.###}" -f $absScale -replace "\.?0+$"
                "_x$sign$scalePart"
            }

            # Construct new filename
            $newFile = if ([string]::IsNullOrEmpty($formattedScale)) {
                "$cleanBase.mdl"
            } else {
                "$cleanBase$formattedScale.mdl"
            }

            # Rebuild full path with original separator and quotes
            $newPath = if ($dir -ne '') { "$dir$separator$newFile" } else { $newFile }
            if ($hasQuote) { $newPath = '"' + $newPath + '"' }

            # Update the line while preserving original indentation
            $qcLines[$i] = "$indent$newPath"
            
            # Display friendly output
            $displayPath = if ($hasQuote) { $newPath.Trim('"') } else { $newPath }
            Write-Host "`n Updated " -NoNewline
            Write-Host '$modelname' -NoNewline -ForegroundColor Yellow
            Write-Host " to: " -NoNewline
            Write-Host ([System.IO.Path]::GetFileName($displayPath)) -ForegroundColor Green
            break
        }
    }

    if (-not $modelnameFound) {
        Write-Host "`n" -NoNewline
        Write-Separator -Title "YOU CANT BE SERIOUS RIGHT NOW" -Length 50 -Color Red
        Write-Host " No " -NoNewline -ForegroundColor Gray
        Write-Host '$modelname' -NoNewline -ForegroundColor Yellow
        Write-Host " command found in the specified QC file" -ForegroundColor Gray
        Write-Host " Imma be straight with you homie, you kinda need this" -ForegroundColor Gray
        Write-Separator -Length 50 -Color Red
    }

    [System.IO.File]::WriteAllText($qcFile, ($qcLines -join "`r`n"), $utf8NoBOM)
}

# --------------------------
# 5) Process VRD if included
# --------------------------
if (-not $hasVrd) {
    Write-Host "`n QC updated; no VRD was processed."
} else {
    Write-Host ""
    Write-Separator -Title "VRD File Processing" -Char '═' -Length 50 -Color Cyan
    $vrdRaw = Get-Content $vrdFile -Raw
    $vrdContent = $vrdRaw.TrimEnd("`r","`n")
    $lineEnding = if ($vrdRaw -match "`r`n") { "`r`n" } else { "`n" }

    # Markers   # IMPORTANT: When markers exist, CURRENT VALUES ARE IGNORED COMPLETELY, scaling always uses original marked values × new scale factor
    $markerBasepos = "// Listed below are the original <basepos> and <trigger> translation values qc_scale_tool found on first run, normalized to `$scale 1"
    $parenthetical = "// If these are not correct, remove all of these comments and run the tool again with values that match the `$scale set in the QC`n"

    # Data structures
    $currentHelper = ""   # -----------------------------------------------
    $origBasepos = @{}    # helper → [origX,origY,origZ]
    $origTriggers = @{}   # helper → ordered list of @(origTx,origTy,origTz)
    $newLines = @()       # -----------------------------------------------

    # First pass - collect original values from markers when present, otherwise assume current values represent $scale 1
    $allLines = $vrdContent -split "`r?`n"
    $isFirstRun = $vrdContent -notmatch [regex]::Escape($markerBasepos)

    if ($isFirstRun) {
        # ---------------------------------------------------------------
        # First-time run → Capture originals values as normalized markers
        # ---------------------------------------------------------------
        Write-Host "`n No prior scales detected - Recording original values"
        Write-Host "`n Make sure the VRD had the correct translation values"
        Write-Host " for the previously set scale before you ran the tool"
    } else {
        # -----------------------------------------------
        # Subsequent runs → Re-scale using stored markers
        # -----------------------------------------------
        Write-Host "`n Detected previously scaled VRD - Applying new scale"
    }

    foreach ($line in $allLines) {
        # Skip existing marker lines on subsequent runs
        if (-not $isFirstRun -and ($line -match $RegexPatterns.OrigBaseposMarker -or $line -match $RegexPatterns.OrigTriggerMarker)) {
            continue
        }

        # Detect helper context
        if ($line -match $RegexPatterns.HelperLine) {
            $currentHelper = $Matches[1]
            if (-not $origTriggers.ContainsKey($currentHelper)) {
                $origTriggers[$currentHelper] = @()
            }
            $newLines += $line
            continue
        }

        # Capture original <basepos> values (normalized to $scale 1)
        if ($line -match $RegexPatterns.BaseposLine) {
            $origX = $Matches[2]
            $origY = $Matches[3]
            $origZ = $Matches[4]
            
            $origBasepos[$currentHelper] = @(
                (ScaleNum $origX (1/$originalScale)),
                (ScaleNum $origY (1/$originalScale)), 
                (ScaleNum $origZ (1/$originalScale)))
            continue
        }

        # Capture original <trigger> values (normalized to $scale 1)
        if ($line -match $RegexPatterns.TriggerLine) {
            $origTx = $Matches[2]
            $origTy = $Matches[3]
            $origTz = $Matches[4]
            
            $origTriggers[$currentHelper] += ,@(
                (ScaleNum $origTx (1/$originalScale)),
                (ScaleNum $origTy (1/$originalScale)), 
                (ScaleNum $origTz (1/$originalScale)))
            continue
        }
    }

    # Second pass - process lines and apply new scale
    $currentHelper = ""
    $triggerIndices = @{} # Helper → current trigger index
    $newLines = @()
    $hasBasepos = $false  # Initialize tracking variable

    foreach ($line in $allLines) { # Detect helper context
        if ($line -match $RegexPatterns.HelperLine) {
            $currentHelper = $Matches[1]
            $triggerIndices[$currentHelper] = 0
            $newLines += $line
            continue
        }

        # Process <basepos> lines using original values from markers
        if ($line -match $RegexPatterns.BaseposLine) {
            if ($line -match $RegexPatterns.OrigBaseposMarker) {
                # This is a marker line - leave it exactly as-is
                $newLines += $line
            }
            elseif ($origBasepos.ContainsKey($currentHelper)) {
                # This is an actual basepos line - scale it
                $vals = $origBasepos[$currentHelper]
                $newX = ScaleNum $vals[0] $scale
                $newY = ScaleNum $vals[1] $scale
                $newZ = ScaleNum $vals[2] $scale
                $newLines += "$($Matches[1])<basepos>     $newX         $newY         $newZ$($Matches[5])"
                $hasBasepos = $true  # Mark that we found at least one basepos
                continue
            }
        }

        # Process <trigger> lines using original values from markers (maintains original order via index tracking)
        if ($line -match $RegexPatterns.TriggerLine) {
            if ($line -match $RegexPatterns.OrigTriggerMarker) {
                # This is a marker line - leave it exactly as-is
                $newLines += $line
            }
            elseif ($origTriggers.ContainsKey($currentHelper)) {
                # This is an actual trigger line - scale it
                $vals = $origTriggers[$currentHelper][$triggerIndices[$currentHelper]]
                $newTx = ScaleNum $vals[0] $scale
                $newTy = ScaleNum $vals[1] $scale
                $newTz = ScaleNum $vals[2] $scale
                $newLines += "$($Matches[1]) $newTx $newTy $newTz$($Matches[5])"
                $triggerIndices[$currentHelper]++
                continue
            }
        } $newLines += $line
    }

    # Only add markers if this was first run (subsequent runs preserve existing markers)
    if ($isFirstRun -and $origBasepos.Count -gt 0) {
        if ($newLines[-1] -ne "") { $newLines += "" }
        $newLines += $markerBasepos
        $newLines += $parenthetical
        
        foreach ($h in $origBasepos.Keys) {
            $vals = $origBasepos[$h]
            $newLines += "// ORIG_BASEPOS  $h  $($vals[0])  $($vals[1])  $($vals[2])"
        }
        
        $newLines += ""
        
        foreach ($h in $origTriggers.Keys) {
            $tList = $origTriggers[$h]
            for ($i = 0; $i -lt $tList.Count; $i++) {
                $oVals = $tList[$i]
                $newLines += "// ORIG_TRIGGER  $h  $i  $($oVals[0])  $($oVals[1])  $($oVals[2])"
            }
        }
    }

    if (-not $hasBasepos -and $hasVrd) {
        Write-Host ""
        Write-Separator -Title "WARNING: NO BASEPOS FOUND" -Char '!' -Length 50 -Color Red
        Write-Host " No valid <basepos> entries exist within the VRD file"
        Write-Separator -Length 50 -Color Red
    }

    $totalTriggers = 0
    foreach ($helper in $origTriggers.Keys) {
        foreach ($trigger in $origTriggers[$helper]) { # Only count non-zero triggers
            if ([double]$trigger[0] -ne 0 -or [double]$trigger[1] -ne 0 -or [double]$trigger[2] -ne 0) { $totalTriggers++ }
        }
    }
    if ($totalTriggers -eq 0 -and $hasVrd) {
        Write-Host ""
        Write-Separator -Title "WARNING: NO TRIGGERS FOUND" -Char '!' -Length 50 -Color Red
        Write-Host " No valid <trigger> entries exist within the VRD file"
        Write-Separator -Length 50 -Color Red
    }

    [System.IO.File]::WriteAllText($vrdFile, ($newLines -join $lineEnding), $utf8NoBOM)

    if (-not $hasBasepos -and $totalTriggers -eq 0) {
        Write-Host "`n    Try using an actual VRD file next time genius" -ForegroundColor DarkRed
    }
    else {
        $messageParts = @()
        if ($hasBasepos) { $messageParts += "<basepos>" }
        if ($totalTriggers -gt 0) { $messageParts += "$totalTriggers <trigger>" }
        
        Write-Host "`n $($messageParts -join ' and ') translations scaled x" -NoNewline
        Write-Host $scaleFactor -ForegroundColor Yellow
    }
}

Write-Host ""
$completionMessages = @(
    @{Text="K THX BAI"; Color="Green"},
    @{Text="THANKS, AND HAVE FUN"; Color="Gray"},
    @{Text="A WINNER IS YOU!"; Color="Green"},
    @{Text="HOLY CRAP LOIS DIS IS FRIGGIN SWEET"; Color="DarkGray"},
    @{Text="WE'LL BANG OKAY?"; Color="Red"},
    @{Text="HAAAAAAAAAAAAAX!!!!!"; Color="DarkCyan"},
    @{Text="PSST, ITS FREE REAL ESTATE"; Color="DarkCyan"},
    @{Text="WELL EXCUUUUUUUUUUSE ME PRINCESS!"; Color="DarkGreen"},
    @{Text="MATCH BEGINS IN 60 SECONDS"; Color="Yellow"},
    @{Text="ITS A SECRET TO EVERYBODY"; Color="Blue"},
    @{Text="I'M REALLY FEELING IT!"; Color="Red"},
    @{Text="THE CAKE IS A LIE"; Color="Red"},
    @{Text="DON'T FORGET TO BUY CAPTAIN TOAD!"; Color="Yellow"},
    @{Text="hi every1 im new!!!!!!! *holds up spork*"; Color="Magenta"},
    @{Text="RISE AND SHINE MR FREEMAN"; Color="DarkGray"},
    @{Text="PICK UP THAT CAN"; Color="DarkGray"},
    @{Text="SHREK IS LOVE, SHREK IS LIFE"; Color="Green"},
    @{Text="TELL ME ABOUT BANE, WHY DOES HE WEAR THE MASK?"; Color="DarkGray"},
    @{Text="like favorite subscribe :)"; Color="Red"},
    @{Text="[VGS] SHAZBOT!"; Color="DarkCyan"},
    @{Text="HOPEFULLY IT WILL HAVE BEEN WORTH THE WAIT"; Color="Gray"},
    @{Text="ALL YOUR BASE ARE BELONG TO US"; Color="Yellow"}
)

# Need to save file with UTF-8-BOM encoding for these to display properly
$toucan = @"
         ░░░░░░░░▄▄▄▀▀▀▄▄███▄░░░░░░░░░░░░░░
         ░░░░░▄▀▀░░░░░░░▐░▀██▌░░░░░░░░░░░░░
         ░░░▄▀░░░░▄▄███░▌▀▀░▀█░░░░░░░░░░░░░
         ░░▄█░░▄▀▀▒▒▒▒▒▄▐░░░░█▌░░░░░░░░░░░░
         ░▐█▀▄▀▄▄▄▄▀▀▀▀▌░░░░░▐█▄░░░░░░░░░░░
         ░▌▄▄▀▀░░░░░░░░▌░░░░▄███████▄░░░░░░
         ░░░░░░░░░░░░░▐░░░░▐███████████▄░░░
         ░░░░░░░░░░░░░░▐░░░░▐█████████████▄
         ░░░░░le░░░░░░░▐░░░░▐█████████████▄
         ░░░░toucan░░░░░░▀▄░░░▐█████████████▄ 
         ░░░░░░has░░░░░░░░▀▄▄███████████████ 
         ░░░░░arrived░░░░░░░░░░░░█▀██████░░
"@
$triforce = @"
                       
                      †¥
                     ††¥¥
                    ††††¥¥
                   ††††††¥¥
                  ††††††††¥¥
                 ††††††††††¥¥
                ††††††††††††¥¥
               ††††††††††††††¥¥
              ††††††††††††††††¥¥
             ††¥¥¥¥¥¥¥¥¥¥¥¥¥¥††¥¥
            ††††¥¥          ††††¥¥
           ††††††¥¥        ††††††¥¥
          ††††††††¥¥      ††††††††¥¥
         ††††††††††¥¥    ††††††††††¥¥
        ††††††††††††¥¥  ††††††††††††¥¥
       ††††††††††††††††††††††††††††††¥¥
      †¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥
"@
$link = @"
        ¶¶3333¶   ¶¶¶¶ÿÿÿÿ¶¶¶¶   ¶3333¶¶
        ¶¶33333¶_¶ÿÿÿÿÿÿÿÿÿÿÿÿ¶_¶33333¶¶
        ¶¶¶¶¶¶¶¶ÿÿ¶¶¶¶¶¶¶¶¶¶¶¶ÿÿ¶¶¶¶¶¶¶¶
        ¶¶0000¶¶¶77777777777777¶¶¶0000¶¶
        ¶¶0000¶¶7¶¶¶¶¶¶¶¶¶¶¶¶¶¶7¶¶0000¶¶
         ¶¶000¶¶¶a  ¶¶aaaa¶¶  a¶¶¶000¶¶
         ¶¶000¶¶aa  ¶¶aaaa¶¶  aa¶¶000¶¶
          ¶¶00¶¶aaa  aaaaaa  aaa¶¶00¶¶
           ¶¶000¶¶aaaaaaaaaaaa¶¶000¶¶
            ¶¶00¶¶¶¶aaaaaaaa¶¶¶¶00¶¶
             ¶¶88888¶¶¶¶¶¶¶¶88888¶¶
              ¶¶8855888888885588¶¶
              ¶¶8855555555555588¶¶
            ¶¶11¶¶888855558888¶¶11¶¶
          ¶¶88881111¶¶1111¶¶11118888¶¶
          ¶¶¶¶¶¶8888111111118888¶¶¶¶¶¶
        ¶¶ƒƒ§§¶¶¶¶¶¶88888888¶¶¶¶¶¶§§ƒƒ¶¶
        ¶¶ƒƒƒƒ§§¶¶¯¶¶¶¶¶¶¶¶¶¶¯¶¶§§ƒƒƒƒ¶¶
        ¶¶¶¶¶¶¶¶¶               ¶¶¶¶¶¶¶¶¶
"@

switch (Get-Random -Minimum 1 -Maximum 101) {
    { $_ -le 3 } {  # 3% Toucan
        Write-Host $toucan -ForegroundColor Cyan
        Write-Host "`n               Press any key to PRAISE!" -NoNewline
        Read-AnyKey
        exit 0
    }
    { $_ -le 4 } {  # 1% Zelda
        Write-Host $triforce -ForegroundColor Yellow
        Write-Host $link -ForegroundColor Green
        Write-Host "`n   DONE! PRESS ANY KEY TO EXIT!" -NoNewline
        Write-Host "(Ultra-rare end!)" -ForegroundColor Yellow
        Read-AnyKey
        exit 0
    }
    { $_ -le 5 } {  # Also 1% (it is a mystery 👻)
        Write-Separator -Title "$($r -join '')" -Length 50 -Color DarkRed
        Read-AnyKey
        exit 0
    }
    default {  # 95% Generic Messages
        $randomMessage = $completionMessages | Get-Random
        Write-Separator -Title $randomMessage.Text -Length 50 -Color $randomMessage.Color
    }
}

Write-Host "`n Done. Press any key to exit..." -NoNewline
Read-AnyKey
exit 0
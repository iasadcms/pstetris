<#
.SYNOPSIS
    PowerShell Tetris    
.DESCRIPTION
    Quick and dirty PowerShell Tetris
    Use arrow keys - left, right, down, up == rotate.
    Escape == quit
    No music
#>


$CONBUFF_WIDTH = 36
$CONBUFF_HEIGHT = 27


$FORECOLOUR_BACKGROUND = 'Cyan';        $BACKCOLOUR_BACKGROUND = 'DarkGray'
$FORECOLOUR_TITLE = 'White';            $BACKCOLOUR_TITLE = 'Black'
$FORECOLOUR_TITLESHADOW = 'DarkGray';   $BACKCOLOUR_TITLESHADOW = 'Black'
$FORECOLOUR_SCORE = 'White';            $BACKCOLOUR_SCORE = 'Black'
$FORECOLOUR_BORDER = 'Cyan';            $BACKCOLOUR_BORDER = 'Black'
$FORECOLOUR_MENU = 'White';             $BACKCOLOUR_MENU = 'Black'
$FORECOLOUR_GAMETEXT = 'White';         $BACKCOLOUR_GAMETEXT = 'Black'
$BACKCOLOUR_GAMEAREA = 'Black'


$MAINMENU_TOPSCORE_BLINKMS = 1000           # how fast top score blinks on the main menu


#######################################################################################################
# Misc
#######################################################################################################


$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()



#######################################################################################################
# Console
#######################################################################################################


$SAVED_FORECOLOUR = [console]::ForegroundColor
$SAVED_BACKCOLOUR = [console]::BackgroundColor 


$CHAR_SOLIDBLOCK = [char]9608
$CHAR_FADE1 = [char]9619
$CHAR_FADE2 = [char]9618
$CHAR_FADE3 = [char]9617
$CHAR_EMPTY = [char]32
$CHAR_ARROWLEFT = [char]9668
$CHAR_ARROWRIGHT = [char]9658
$CHAR_ARROWUP = [char]9650
$CHAR_ARROWDOWN = [char]9660


function ClearKeyboardBuffer {
<# clears the keyboard buffer #>
    while (([Console]::KeyAvailable)) {
        [void][Console]::ReadKey($false)
    }
}
    

# win32 type lib - helps remove blinking cursor
if ("Win32" -as [type]) { } 
else {
Add-Type -TypeDefinition @"
    using System;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
     
    public static class Win32
    {
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern IntPtr GetStdHandle(int nStdHandle);
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool GetConsoleCursorInfo(IntPtr hConsoleOutput, out CONSOLE_CURSOR_INFO lpConsoleCursorInfo);
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool SetConsoleCursorInfo(IntPtr hConsoleOutput, ref CONSOLE_CURSOR_INFO lpConsoleCursorInfo);
        
        [StructLayout(LayoutKind.Sequential)]
        public struct CONSOLE_CURSOR_INFO {
            public uint dwSize;
            public bool bVisible;
        }
    }
"@
}

if ("TrustAllCertsPolicy" -as [type]) {} else {
    Add-Type "using System.Net;using System.Security.Cryptography.X509Certificates;public class TrustAllCertsPolicy : ICertificatePolicy {public bool CheckValidationResult(ServicePoint srvPoint, X509Certificate certificate, WebRequest request, int certificateProblem) {return true;}}"
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}

$global:hStdOut = [win32]::GetStdHandle(-11)
$global:CursorInfo = New-Object Win32+CONSOLE_CURSOR_INFO


function GetCursorVisible {
    if ([win32]::GetConsoleCursorInfo($hStdOut, [ref]$global:CursorInfo)) {
        return [bool]($global:CursorInfo.bVisible)
    }
    else { return $true }
}


$SAVED_CURSORVISIBLE = GetCursorVisible


function SetCursorVisible {
param(
    [bool]$Visible = $true
)
    if ([win32]::GetConsoleCursorInfo($hStdOut, [ref]$global:CursorInfo)) {
        $global:CursorInfo.bVisible = $Visible
        [win32]::SetConsoleCursorInfo($hStdOut, [ref]$global:CursorInfo) | Out-Null
    }
}



#######################################################################################################
# console buffer
#######################################################################################################


$CONBUFF_SIZE = ($CONBUFF_WIDTH * $CONBUFF_HEIGHT)                      # size of the virtual console arrays (w x h)

function NewConBuffObject {
<# creates a new object to store console buffer info #>
    $Result = [pscustomobject]@{
        Buffer = (New-Object int[]($CONBUFF_SIZE))                       # the current buffer we'll print
        ForeColours = (New-Object int[]($CONBUFF_SIZE))                  # forground colours
        BackColours = (New-Object int[]($CONBUFF_SIZE))                  # background colours
    }
    return $Result
}


# our virtual console vars
$ConBuff = NewConBuffObject                                             # our virtual console
$LastConBuff = NewConBuffObject                                         # the last console printed on screen


function ResetLastConBuff {
<# sets LastConBuff to values that likely force a print during the next call to PrintConBuff #>
    for ($idx = 0; $idx -lt $LastConBuff.Buffer.Length; $idx++) { $LastConBuff.Buffer[$idx] = 365 }
}


function ClearConBuff {
<# clears ConBuff, by setting every char to "space", with saved foreground and background colours #>
    for ($idx = 0; $idx -lt $LastConBuff.Buffer.Length; $idx++) { 
        $ConBuff.Buffer[$idx] = 32
        $ConBuff.ForeColours[$idx] = $SAVED_FORECOLOUR
        $ConBuff.BackColours[$idx] = $SAVED_BACKCOLOUR
    }
}


function CopyToConBuff($Source) {
<# copies source buff on to $ConBuff #>
    for ($idx = 0; $idx -lt $CONBUFF_SIZE; $idx++) {
        $ConBuff.Buffer[$idx] = $Source.Buffer[$idx]
        $ConBuff.ForeColours[$idx] = $Source.ForeColours[$idx]
        $ConBuff.BackColours[$idx] = $Source.BackColours[$idx]
    }
}
    
    
function ConBuffWrite {
<# writes text to $ConBuff, starting at position x, y. #>
param(
    [int]$x,
    [int]$y,
    [string]$Text,
    [ConsoleColor]$ForeColour,
    [ConsoleColor]$BackColour
)    
    if (($y -lt 0) -or ($y -ge $CONBUFF_HEIGHT)) { return }
    $ypos = $y * $CONBUFF_WIDTH
    for ($dx = 0; $dx -lt $Text.Length; $dx++) {
        $idx = $ypos + $x + $dx
        if ($idx -lt 0) { continue }
        if ($idx -ge $CONBUFF_SIZE) { break }
        $ConBuff.Buffer[$idx] = $Text[$dx]
        $ConBuff.ForeColours[$idx] = $ForeColour
        $ConBuff.BackColours[$idx] = $BackColour
    }
}


function ConBuffDrawFromMask {
<# draws char on conbuff, where the value in Mask isn't whitespace #>
param(
    [int]$x,
    [int]$y,
    [string[]]$Mask,
    [char]$Char,
    [ConsoleColor]$ForeColour,
    [ConsoleColor]$BackColour
)
    for ($dy = 0; $dy -lt $Mask.Length; $dy++) {
        for ($dx = 0; $dx -lt $Mask[$dy].Length; $dx++) {
            if (-not [string]::IsNullOrWhiteSpace($Mask[$dy][$dx])) {
                ConBuffWrite -x ($x + $dx) -y ($y + $dy) -Text $Char -ForeColour $ForeColour -BackColour $BackColour
            }
        }
    }
}   
    

function ConBuffDrawBox {
<# writes box of a char to $global:ConBuff, starting at position x, y, with length, height (down and right) #>
param(
    [char]$Char,
    [int]$x,
    [int]$y,
    [int]$Width,
    [int]$Height,
    [ConsoleColor]$ForeColour,
    [ConsoleColor]$BackColour
)
    for ($_x = $x; $_x -lt ($x + $Width); $_x++) {
        for ($_y = $y; $_y -lt ($y + $Height); $_y++) {
            ConBuffWrite -Text $Char -x $_x -y $_y -ForeColour $ForeColour -BackColour $BackColour
        }
    }
}
    

function ConBuffDrawFrame {
<# draws a box, starting at position x, y, with length, height (down and right).#>
param(
    [int]$x,
    [int]$y,
    [int]$Width,
    [int]$Height,
    [ConsoleColor]$ForeColour,
    [ConsoleColor]$BackColour
)
    # draw top and bottom lines
    for ($_x = $x + 1; $_x -lt ($x + $Width) - 1; $_x++) { 
        ConBuffWrite -Text ([char]0x2550) -x $_x -y $y -ForeColour $ForeColour -BackColour $BackColour
        ConBuffWrite -Text ([char]0x2550) -x $_x -y ($y + $Height - 1) -ForeColour $ForeColour -BackColour $BackColour
    }
    # draw left and right lines
    for ($_y = $y + 1; $_y -lt ($y + $Height) - 1; $_y++) { 
        ConBuffWrite -Text ([char]0x2551) -x $x -y $_Y  -ForeColour $ForeColour -BackColour $BackColour 
        ConBuffWrite -Text ([char]0x2551) -x ($x + $Width - 1) -y $_y  -ForeColour $ForeColour -BackColour $BackColour
    }
    # draw corners
    ConBuffWrite -Text ([char]0x2554) -x $x -y $y -ForeColour $ForeColour -BackColour $BackColour 
    ConBuffWrite -Text ([char]0x2557) -x ($x + $Width - 1) -y $y -ForeColour $ForeColour -BackColour $BackColour
    ConBuffWrite -Text ([char]0x255A) -x $x -y ($y + $Height - 1) -ForeColour $ForeColour -BackColour $BackColour 
    ConBuffWrite -Text ([char]0x255D) -x ($x + $Width - 1) -y ($y + $Height - 1) -ForeColour $ForeColour -BackColour $BackColour
}
   

function ConBuffDrawFramedBox {
<# draws a framed box in conbuff #>
param(
    [char]$Char,
    [int]$x,
    [int]$y,
    [int]$Width,
    [int]$Height,
    [ConsoleColor]$FrameForeColour,
    [ConsoleColor]$FrameBackColour,
    [ConsoleColor]$BoxForeColour,
    [ConsoleColor]$BoxBackColour
)
    $Width = [System.Math]::Max($Width, 2)
    $Height = [System.Math]::Max($Height, 2)
    ConBuffDrawFrame -ForeColour $FrameForeColour -BackColour $FrameBackColour -x $x -y $y -Width $Width -Height $Height
    ConBuffDrawBox -Char $Char -ForeColour $BoxForeColour -BackColour $BoxBackColour -x ($x + 1) -y ($y + 1) -Width ($Width - 2) -Height ($Height - 2)
}
    

function PrintConBuff {
<# displays $global:ConBuff in the console window #>
param(
    [switch]$Force                  # if not set, only changes are updated
)
    # loop through the screen buffer and write out any changes - drawing bottom up appears smoother
    for ($y = ($CONBUFF_HEIGHT - 1); $y -ge 0; $y--) {
        $ypos = $y * $CONBUFF_WIDTH                         # new base for index calcs
        $lastx = [int]::MinValue;                           # if x is $lastx + 1, we don't need to SetCursor
        for ($x = 0; $x -lt $CONBUFF_WIDTH; $x++) {
            # get values from conbuff
            $idx = $ypos + $x
            $new_forecol = $ConBuff.ForeColours[$idx]
            $new_backcol = $ConBuff.BackColours[$idx]
            $new_value = $ConBuff.Buffer[$idx]
            # update if required             
            [bool]$_updatereq = $Force.IsPresent -or (($new_value -ne $LastConBuff.Buffer[$idx]) -or ($new_forecol -ne $LastConBuff.ForeColours[$idx]) -or ($new_backcol -ne $LastConBuff.BackColours[$idx]))
            if ($true -eq $_updatereq) {                
                if ($x -ne ($lastx + 1)) {
                    [Console]::SetCursorPosition($x, $y);
                    $lastx = $x
                }                
                [Console]::BackgroundColor = $new_backcol
                [Console]::ForegroundColor = $new_forecol
                [Console]::Write([char]($new_value))
                $LastConBuff.Buffer[$idx] = $new_value
                $LastConBuff.ForeColours[$idx] = $new_forecol
                $LastConBuff.BackColours[$idx] = $new_backcol
            }
        }
    }
    #######################
    [console]::ForegroundColor = $SAVED_FORECOLOUR
    [console]::BackgroundColor = $SAVED_BACKCOLOUR
}





#######################################################################################################
# Tetromino town
#######################################################################################################


# Define the TetrominoType enum
Add-Type -TypeDefinition @"
public enum TetrominoType {
    I = 1,
    J = 2,
    L = 3,
    O = 4,
    S = 5,
    T = 6,
    Z = 7
}
"@
$TETROMINOTYPE_MAXINTVALUE = [TetrominoType].GetEnumValues() | ForEach-Object { [int]$_ } | Sort-Object | Select-Object -Last 1
$TETROMINOTYPE_MININTVALUE = [TetrominoType].GetEnumValues() | ForEach-Object { [int]$_ } | Sort-Object | Select-Object -First 1
$TETROMINOTYPE_NUMVALUES = [TetrominoType].GetEnumValues().Count

$TETROMINO_LENGTH = 4                     # x and y size of a tetromino
$TETROMINO_CHAR = $CHAR_SOLIDBLOCK

# Define the tetromino shapes as a dictionary of 4x4 multi-dimensional arrays indexed by the TetrominoType enum
$TetrominoShapes = @{
    [TetrominoType]::I = @(
        @(0,0,0,0),
        @(1,1,1,1),
        @(0,0,0,0),
        @(0,0,0,0)
    )
    [TetrominoType]::J = @(
        @(0,0,0,0),
        @(1,1,1,0),
        @(0,0,1,0),
        @(0,0,0,0)
    )
    [TetrominoType]::L = @(
        @(0,0,0,0),
        @(1,1,1,0),
        @(1,0,0,0),
        @(0,0,0,0)
    )
    [TetrominoType]::O = @(
        @(0,0,0,0),
        @(0,1,1,0),
        @(0,1,1,0),
        @(0,0,0,0)
    )
    [TetrominoType]::S = @(
        @(0,0,0,0),
        @(0,1,1,0),
        @(1,1,0,0),
        @(0,0,0,0)
    )
    [TetrominoType]::T = @(
        @(0,0,0,0),
        @(1,1,1,0),
        @(0,1,0,0),
        @(0,0,0,0)
    )
    [TetrominoType]::Z = @(
        @(0,0,0,0),
        @(1,1,0,0),
        @(0,1,1,0),
        @(0,0,0,0)
    )
}


function NewTetrominoObject {
<# generates a tetromino object. If type is defined, the tetomino shape is copied to this #>
param(
    [TetrominoType]$Type
)
    $Result = [pscustomobject]@{
        Type = $Type                            # the type of tetromino defined by shape
        Shape = $null                           # shape of the tetromino (copied from $TetrominoShapes)
        Row = 0                                 # current row in the game area
        Column = 0                              # current column in the game area
        ShapeStartRow = 0                       # initial row to start drawing tetromino when it enters the board.
        ShapeEndRow = 0                         # last row where the tetromino shape ends
    }
    return $Result
}


function GetRandomTetrominoType {   
<# returns a random tetromino type value #> 
    return [TetrominoType](Get-Random -Minimum $TETROMINOTYPE_MININTVALUE -Maximum $TETROMINOTYPE_MAXINTVALUE)
}


function ConBuffDrawTetromino {
<# draws a tetromino on conbuff at x, y#>
param(
    [TetrominoType]$Type,
    $Shape,
    [int]$x,
    [int]$y,
    [ConsoleColor]$BackColour
)
    if ($null -eq $Shape) {        
        $Shape = GetTetrominoPrintShape -Shape ($TetrominoShapes[$Type])
    }
    $backcol = if ($null -eq $BackColour) { $TetrominoBackColours[$Type] } else { $BackColour }
    for ($dx = 0; $dx -lt $TETROMINO_LENGTH; $dx++) {
        for ($dy = 0; $dy -lt $TETROMINO_LENGTH; $dy++) {
            if ($Shape[$dy][$dx]) { 
                $_char = $Shape[$dy][$dx]
                if ($_char -eq 1) { $_char = $TETROMINO_CHAR }                
                ConBuffWrite -Text $_char -x ($x + $dx) -y ($y + $dy) -ForeColour $TetrominoForeColours[$Type] -BackColour $backcol
            }             
        }
    }
}







#######################################################################################################
# Game Data
#######################################################################################################


$GAMEBOARD_WIDTH = 10                # width of the tetris game area
$GAMEBOARD_HEIGHT = 18               # height of the tetris game area

[int]$GAMEBOARD_EMPTYVAL = 0                                       # the value representing an empty game board cell
[int]$GAMEBOARD_EXPLODEVAL = $TETROMINOTYPE_MAXINTVALUE + 1         # the int value representing an exploding cell


# game types
Enum GameType {
    AGame
}

function NewGameDataObject {
    $Result = [PSCustomObject]@{
        Score = 0
        Level = 0
        Lines = 0
        TopScore = 0
        GameType = $null
        Name = $null        
        Statistics = @{}
        CurrentTetromino = NewTetrominoObject
        NextTetrominoType = GetRandomTetrominoType
        Board = [int[,]]::new($GAMEBOARD_HEIGHT, $GAMEBOARD_WIDTH);             # this holds data for placement of pieces
        BoardChars = [int[,]]::new($GAMEBOARD_HEIGHT, $GAMEBOARD_WIDTH);        # this holds the character to print
        StartTime = [datetime]::MinValue
        EndTime = [datetime]::MinValue
    }
    [Enum]::GetValues([TetrominoType]) | ForEach-Object { $Result.Statistics[$_] = 0 }
    return $Result
}


# store data for each game here
$AllGameData = @{}
[enum]::GetValues([GameType]) | % {
    $rec = NewGameDataObject
    $rec.GameType = $_
    $rec.Name = $_.ToString()
    $AllGameData[$_] = $rec    
}
$AllGameData[[GameType]::AGame].Name = "A-Type"

# current game data stored here
#$GameData = $AllGameData[[GameType]::AGame]

function ResetGameData {
<# resets game data #>
    $GameData = $AllGameData[[GameType]::AGame]
    $GameData.Score = 0
    $GameData.Level = 1
    $GameData.Lines = 0
    [Enum]::GetValues([TetrominoType]) | ForEach-Object { $GameData.Statistics[$_] = 0 }
    $GameData.CurrentTetromino = NewTetrominoObject
    $GameData.NextTetrominoType = GetRandomTetrominoType    
    for ($y = 0; $y -lt $GAMEBOARD_HEIGHT; $y++) {
        for ($x = 0; $x -lt $GAMEBOARD_WIDTH; $x++) {
            $GameData.Board[$y, $x] = $GAMEBOARD_EMPTYVAL
            $GameData.BoardChars[$y, $x] = $CHAR_EMPTY
        }
    }
    $GameData.StartTime = [datetime]::MinValue
    $GameData.EndTime = [datetime]::MinValue
}





# number of points for each line cleared, per level
$PointsPerLineCleared = @(0, 40, 100, 300, 1200)
$PointsPerNewOnKeyDown = 10

# max number of rows that can be cleared
$MAXROWSCLEARED = 4

function GetNumberOfPointsForLinesCleared {
param(
    [int]$NumLinesCleared
)
<# returns the number of points for clearing lines for a level#>
    $NumLinesCleared = [System.Math]::Max(0, [System.Math]::Min(4, $NumLinesCleared))
    $points = ($PointsPerLineCleared[$NumLinesCleared] * ($GameData.Level + 1))
    return $points
}


function AddPointsToScore {
param (
    [int]$PointToAdd
)
    $GameData.Score += $PointToAdd
    $GameData.TopScore = [System.Math]::Max($GameData.TopScore, $GameData.Score)
}


# speed of fall in ms per level. Anything above max index will use the last value.
$DelaymsPerLevel = @(1000, 800, 757, 714, 671, 628, 585, 542, 499, 456, 420, 399, 378, 357, 336, 315, 294, 273, 252, 231)

function GetDelaymsForCurrentLevel {   
<# returns the current ms delay for tetrino falls used in the main game loop #>
    $idx = if ($GameData.Level -lt $DelaymsPerLevel.Length) { $GameData.Level } else { $DelaymsPerLevel.Length }
    return $DelaymsPerLevel[$idx]
}




#######################################################################################################
# Game Area
#######################################################################################################


function ConvertScoreToString {
<# Converts a score value in to a string of numdigits long #>
param(
    [int]$Value = 0,
    $NumDigits = 6
)    
    # is value -lt max int val
    $maxIntVal = [math]::Pow(10, $NumDigits) - 1
    if ($Value -le $maxIntVal) {
        return "{0:d$NumDigits}" -f $Value
    }
    # Check if the value is greater than the maximum hexadecimal value
    $maxHexValue = [Convert]::ToString([Math]::Pow(16, $NumDigits) - 1, 16).ToUpper()
    if ($Value -le [Convert]::ToInt32($maxHexValue, 16)) {
        return [Convert]::ToString($Value, 16).ToUpper().PadLeft($NumDigits, '0')
    }
    # too big
    return "e".ToString().PadLeft($NumDigits, '#')
}
        

# colours used to draw tetrominos
#$TetrominoForeColours = @{
$TetrominoBackColours = @{
    [TetrominoType]::I = [ConsoleColor]::Cyan
    [TetrominoType]::J = [ConsoleColor]::Blue
    [TetrominoType]::L = [ConsoleColor]::DarkYellow
    [TetrominoType]::O = [ConsoleColor]::Gray
    [TetrominoType]::S = [ConsoleColor]::Green
    [TetrominoType]::T = [ConsoleColor]::Magenta
    [TetrominoType]::Z = [ConsoleColor]::Red
}  
#$TetrominoBackColours = @{
$TetrominoForeColours = @{
    [TetrominoType]::I = [ConsoleColor]::DarkCyan
    [TetrominoType]::J = [ConsoleColor]::DarkBlue
    [TetrominoType]::L = [ConsoleColor]::Yellow
    [TetrominoType]::O = [ConsoleColor]::DarkGray
    [TetrominoType]::S = [ConsoleColor]::DarkGreen
    [TetrominoType]::T = [ConsoleColor]::DarkMagenta
    [TetrominoType]::Z = [ConsoleColor]::DarkRed
} 


#$borderChars = "─", "│", "┌", "┐","└","┘","├","┤","┬","┴","┼"
#$borderChars = [char]0x2500, [char]0x2502, [char]0x250C, [char]0x2510, [char]0x2514, [char]0x2518, [char]0x251C, [char]0x2524, [char]0x252C, [char]0x2534, [char]0x253C

$CHAR_NESW = [char]0x253C
$CHAR_NSW = [char]0x2524
$CHAR_NES = [char]0x251C
$CHAR_NS = [char]0x2502
$CHAR_EW = [char]0x2500
$CHAR_NEW = [char]0x2534
$CHAR_ESW = [char]0x252C
$CHAR_NE = [char]0x2514
$CHAR_ES = [char]0x250C
$CHAR_SW = [char]0x2510
$CHAR_WN = [char]0x2518


function GetTetrominoPrintShape {
<# returns the tetromino an array of chars to print #>
param(
    [array]$Shape
)
    $newshape = @(@(0, 0, 0, 0), @(0, 0, 0, 0), @(0, 0, 0, 0), @(0, 0, 0, 0))
    for ($r = 0; $r -lt $TETROMINO_LENGTH; $r++) {
        for ($c = 0; $c -lt $TETROMINO_LENGTH; $c++) {
            if (-not ($Shape[$r][$c])) {
                continue
            }
            $n = [int[]]@(0, 0, 0, 0)
            if ($r -gt 0) { $n[0] = $Shape[($r - 1)][$c] }                      # above
            if ($r -lt $TETROMINO_LENGTH - 1) { $n[1] = $Shape[($r + 1)][$c] }  # below
            if ($c -gt 0) { $n[2] = $Shape[$r][($c - 1)] }                      # left
            if ($c -lt $TETROMINO_LENGTH - 1) { $n[3] = $Shape[$r][($c + 1)] }  # right
                                
            # if has north and south
            $newshape[$r][$c] = if ($n[0] -and $n[1]) { 
                # if has east and west
                if ($n[2] -and $n[3]) { $CHAR_NESW }
                # if west
                elseif ($n[2]) { $CHAR_NSW }
                # if east
                elseif ($n[3]) { $CHAR_NES }
                # else north/south
                else { $CHAR_NS}
            }
            # if has west and east
            elseif ($n[2] -and $n[3]) {
                # has north ?
                if ($n[0]) { $CHAR_NEW }
                # has south ?
                elseif ($n[1]) { $CHAR_ESW }
                # else east/west
                else { $CHAR_EW }
            } 
            else {
                # if north
                if ($n[0]) {
                    if ($n[2]) { $CHAR_WN }                            
                    elseif ($n[3]) { $CHAR_NE }
                    else { $CHAR_NS }
                }
                # if south
                elseif ($n[1]) {
                    if ($n[2]) { $CHAR_SW }                               
                    elseif ($n[3]) { $CHAR_ES }
                    else { $CHAR_NS }
                }
                elseif ($n[2]) {
                    if ($n[0]) { $CHAR_WN }
                    elseif ($n[1]) { $CHAR_SW }
                    else { $CHAR_EW }
                }
                elseif ($n[3]) {
                    if ($n[0]) { $CHAR_NE }
                    elseif ($n[1]) { $CHAR_ES }
                    else { $CHAR_EW }
                }
                else {
                    $CHAR_SOLIDBLOCK
                }
            }
        }
    }
    return $newshape
}


# some vars to help line things
$GAMEAREA_COL1 = 2                                      # game area column a
$GAMEAREA_COL2 = 14                                     # game area column b
$GAMEAREA_COL3 = $GAMEAREA_COL2 + $GAMEBOARD_WIDTH + 2   # game area column c
$GAMEAREA_ROW1 = 5                                      # start row of game area border


function DrawGameBackground {
<# local paint, to paint the currently displayed menu #>
    # draw the background and title
    ClearConBuff
    CopyToConBuff -Source $FancyBackground
    # draw a-type game area
    ConBuffDrawFramedBox -Char $CHAR_EMPTY -x ($GAMEAREA_COL1 + 1) -y 3 -Width 8 -Height 3 -FrameForeColour $FORECOLOUR_BORDER -FrameBackColour $BACKCOLOUR_BORDER -BoxForeColour $FORECOLOUR_GAMETEXT -BoxBackColour $BACKCOLOUR_GAMETEXT
    # draw statistics
    ConBuffDrawFramedBox -Char $CHAR_EMPTY -x $GAMEAREA_COL1 -y 7 -Width 12 -Height 18 -FrameForeColour $FORECOLOUR_BORDER -FrameBackColour $BACKCOLOUR_BORDER -BoxForeColour $FORECOLOUR_GAMETEXT -BoxBackColour $BACKCOLOUR_GAMETEXT
    # draw lines game area
    ConBuffDrawFramedBox -Char $CHAR_EMPTY -x $GAMEAREA_COL2 -y 2 -Width 12 -Height 3 -FrameForeColour $FORECOLOUR_BORDER -FrameBackColour $BACKCOLOUR_BORDER -BoxForeColour $FORECOLOUR_GAMETEXT -BoxBackColour $BACKCOLOUR_GAMETEXT
    # draw main game area
    ConBuffDrawFramedBox -Char $CHAR_EMPTY -x $GAMEAREA_COL2 -y $GAMEAREA_ROW1 -Width ($GAMEBOARD_WIDTH + 2) -Height ($GAMEBOARD_HEIGHT + 2) -FrameForeColour $FORECOLOUR_BORDER -FrameBackColour $BACKCOLOUR_BORDER -BoxForeColour $FORECOLOUR_GAMETEXT -BoxBackColour $BACKCOLOUR_GAMETEXT
    # draw scores area
    ConBuffDrawFramedBox -Char $CHAR_EMPTY -x $GAMEAREA_COL3 -y 2 -Width 8 -Height 9 -FrameForeColour $FORECOLOUR_BORDER -FrameBackColour $BACKCOLOUR_BORDER -BoxForeColour $FORECOLOUR_GAMETEXT -BoxBackColour $BACKCOLOUR_GAMETEXT
    # draw level
    ConBuffDrawFramedBox -Char $CHAR_EMPTY -x $GAMEAREA_COL3 -y 18 -Width 7 -Height 4 -FrameForeColour $FORECOLOUR_BORDER -FrameBackColour $BACKCOLOUR_BORDER -BoxForeColour $FORECOLOUR_GAMETEXT -BoxBackColour $BACKCOLOUR_GAMETEXT
}


function DrawGameNextTetromino {
    ConBuffDrawFramedBox -Char $CHAR_EMPTY -x $GAMEAREA_COL3 -y 11 -Width 6 -Height 7 -FrameForeColour $FORECOLOUR_BORDER -FrameBackColour $BACKCOLOUR_BORDER -BoxForeColour $FORECOLOUR_SCORE -BoxBackColour $BACKCOLOUR_SCORE
    ConBuffWrite -Text "NEXT" -x ($GAMEAREA_COL3 + 1) -y 12 -ForeColour $FORECOLOUR_GAMETEXT -BackColour $BACKCOLOUR_GAMETEXT
    ConBuffDrawTetromino -Type $GameData.NextTetrominoType -x ($GAMEAREA_COL3 + 1)  -y 13 #-BackColour $BACKCOLOUR_GAMETEXT
}


function DrawGameText {
<# draws game text data in conbuff #>
    # display game type 
    ConBuffWrite -Text ($GameData.Name) -x ($GAMEAREA_COL1 + 2) -y 4 -ForeColour $FORECOLOUR_GAMETEXT -BackColour $BACKCOLOUR_GAMETEXT     
    # display statistics
    ConBuffWrite -Text "STATISTICS" -x ($GAMEAREA_COL1 + 1) -y 8 -ForeColour $FORECOLOUR_GAMETEXT -BackColour $BACKCOLOUR_GAMETEXT
    $y = 7
    foreach ($_type in $GameData.Statistics.Keys) {
        $_val = $GameData.Statistics[$_type]
        $_valstr = ConvertScoreToString -Value $_val -NumDigits 3
        $y += 2
        ConBuffDrawTetromino -Type $_type -x ($GAMEAREA_COL1 + 2) -y $y #-BackColour $BACKCOLOUR_GAMETEXT
        ConBuffWrite -Text $_valstr -x ($GAMEAREA_COL1 + 7) -y ($y + 1) -ForeColour $FORECOLOUR_GAMETEXT -BackColour $BACKCOLOUR_GAMETEXT
    }    
    # display lines
    [string]$LinesStr = ConvertScoreToString -Value $GameData.Lines -NumDigits 3
    ConBuffWrite -Text " Lines-$LinesStr" -x ($GAMEAREA_COL2 + 1) -y 3 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    # display score
    ConBuffWrite -Text "TOP" -x ($GAMEAREA_COL3 + 1) -y 4 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    [string]$TopScoreStr = ConvertScoreToString -Value $GameData.TopScore -NumDigits 6
    ConBuffWrite -Text $TopScoreStr -x ($GAMEAREA_COL3 + 1) -y 5 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    ConBuffWrite -Text "SCORE" -x ($GAMEAREA_COL3 + 1) -y 7 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    [string]$ScoreStr = ConvertScoreToString -Value $GameData.Score -NumDigits 6
    ConBuffWrite -Text $ScoreStr -x ($GAMEAREA_COL3 + 1) -y 8 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    # display next
    DrawGameNextTetromino
    # display level
    ConBuffWrite -Text "LEVEL" -x ($GAMEAREA_COL3 + 1) -y 19 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    [string]$LevelStr = ConvertScoreToString -Value $GameData.Level -NumDigits 3
    ConBuffWrite -Text " $LevelStr " -x ($GAMEAREA_COL3 + 1) -y 20 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
}


function DrawGameBoard() {
<# draws the game board in conbuff #>
    for ($boardy = 0; $boardy -lt $GAMEBOARD_HEIGHT; $boardy++) {
        for ($boardx = 0; $boardx -lt $GAMEBOARD_WIDTH; $boardx++) {
            [int]$piece = $GameData.Board[$boardy, $boardx]
            [char]$char = $GameData.BoardChars[$boardy, $boardx]
            if ($piece -in $TETROMINOTYPE_MININTVALUE..$TETROMINOTYPE_MAXINTVALUE) {
                $forecolour = $TetrominoForeColours[[TetrominoType]$piece]
                $backcolour = $TetrominoBackColours[[TetrominoType]$piece]
            }
            elseif ($piece -eq $GAMEBOARD_EXPLODEVAL) {
                $forecolour = [ConsoleColor]::White
                $backcolour = $BACKCOLOUR_GAMEAREA            
                $char = $CHAR_FADE2
            }
            else {
                $forecolour = $BACKCOLOUR_GAMEAREA
                $backcolour = $BACKCOLOUR_GAMEAREA
            }
            ConBuffWrite -Text $char -x ($GAMEAREA_COL2 + 1 + $boardx) -y ($GAMEAREA_ROW1 + 1 + $boardy) -ForeColour $forecolour -BackColour $backcolour
        }       
    }
}


function GetFullBoardRows {
<# returns all rows full #>
    $fullrows = New-Object System.Collections.ArrayList    
    for ($y = $GAMEBOARD_HEIGHT - 1; $y -ge 0; $y--) {
        $xpos = 0
        for ($x = 0; $x -lt $GAMEBOARD_WIDTH; $x++) {
            if ($GameData.Board[$y, $x] -ne $GAMEBOARD_EMPTYVAL) {
                $xpos++
            }
        }
        # is full row ?
        if ($xpos -eq $GAMEBOARD_WIDTH) {
            [void]$fullrows.Add($y)
        }
    }
    return $fullrows
}



#######################################################################################################
# Menus
#######################################################################################################


function NewMenuObject {
    <# creates an object for us to store menu data #>
    param(
        [string[]]$Items    
    )
        $Result = [pscustomobject]@{
            Items = $Items
            Index = 0        
        }
        return $Result
    }
    
    
    function GetNextMenuIndex {
    <# returns the next non empty menu item index #>
    param(
        $MenuObject
    )
        if ($MenuObject.Items.Count -le 0) { return -1 }
        if ($MenuObject.Items.Count -eq 1) { return 0 }
        $index = $MenuObject.Index
        do {
            $index++
            if ($index -ge $MenuObject.items.Count) { $index = 0 }
        } while ([string]::IsNullOrWhiteSpace($MenuObject.Items[$index]))
        return $index
    }
    
    
    function GetLastMenuIndex {
    <# returns the last non empty menu item index #>
    param(
        $MenuObject
    )
        if ($MenuObject.Items.Count -le 0) { return -1 }
        if ($MenuObject.Items.Count -eq 1) { return 0 }
        $index = $MenuObject.Index
        do {
            $index--
            if ($index -lt 0) { $index = $MenuObject.Items.Count - 1 }
        } while ([string]::IsNullOrWhiteSpace($MenuObject.Items[$index]))
        return $index
    }
    
    
    function ConBuffDrawMenu {
    <# draws a menu in to conbuff #>
    param(
        $Menu,
        [int]$x,
        [int]$y,
        [int]$Width,
        [switch]$DrawSelected,
        [switch]$CenterAligned
    )
        $wd2 = $Width / 2    
        for ($menuy = 0; $menuy -lt $Menu.Items.Count; $menuy++) {
            $xpos = if ($true -eq $CenterAligned.IsPresent) { $x + $wd2 - ($Menu.Items[$menuy].Length / 2) } else { $x }
            ConBuffWrite -Text $Menu.Items[$menuy] -x $xpos -y ($y + $menuy) -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU
        }
        if ($true -eq $DrawSelected.IsPresent) {
            ConBuffWrite -Text $CHAR_ARROWRIGHT -x $x -y ($y + $Menu.Index) -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
            ConBuffWrite -Text $CHAR_ARROWLEFT -x ($x + $Width - 1) -y ($y + $Menu.Index) -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
        }
    }
    


#######################################################################################################
# Game Tetromino Stuff
#######################################################################################################


# Will the current tetromino fit at a given row and column on the game board
function WillCurrentTetrominoFit {
param(
    [int]$Row,
    [int]$Column
)
    for ($r = 0; $r -lt $TETROMINO_LENGTH; $r++) {
        for ($c = 0; $c -lt $TETROMINO_LENGTH; $c++) {
            # skip if empty part of shape
            if (-not $GameData.CurrentTetromino.Shape[$r][$c]) {
                continue
            }
            # check row within bounds
            $newr = $Row + $r
            if ($newr -lt 0) {
                continue
            }
            if ($newr -ge $GAMEBOARD_HEIGHT) {
                return $false
            }
            # check colum within bounds
            $newc = $Column + $c
            if (-not ($newc -in 0..($GAMEBOARD_WIDTH - 1))) {
                return $false
            }
            # is gameboard not empty at that spot - boom!
            if ($GameData.Board[$newr, $newc] -ne $GAMEBOARD_EMPTYVAL) {
                return $false
            }
        }
    }
    return $true
}


function Debug_PrintGameBoard {
<# prints the game board by tetromino type #>
    0..($GAMEBOARD_HEIGHT - 1) | ForEach-Object { @(for($x = 0; $x -lt $GAMEBOARD_WIDTH; $x++) { $GameData.Board[$_, $x] }) -join '' }
}


function PlaceCurrentTetromino {
<# paints or removes the current tetromino on the board #>
param(
    [switch]$Remove                 # remove tetromino from board. 
)
    if (($null -eq $GameData.CurrentTetromino) -or ($null -eq $GameData.CurrentTetromino.Shape)) {
        return
    }
    $Shape = GetTetrominoPrintShape -Shape $GameData.CurrentTetromino.Shape
    for ($r = 0; $r -lt $TETROMINO_LENGTH; $r++) {
        for ($c = 0; $c -lt $TETROMINO_LENGTH; $c++) {
            if (-not $GameData.CurrentTetromino.Shape[$r][$c]) {
                continue 
            }
            $_row = $GameData.CurrentTetromino.Row + $r
            if (-not ($_row -in 0..($GAMEBOARD_HEIGHT -1))) {
                continue
            }
            $_col = $GameData.CurrentTetromino.Column + $c
            if (-not ($_col -in 0..($GAMEBOARD_WIDTH - 1))) {
                continue
            }
            if ($Remove.IsPresent) {
                $GameData.Board[$_row, $_col] = $GAMEBOARD_EMPTYVAL
                $GameData.BoardChars[$_row, $_col] = $CHAR_EMPTY
            }
            else {
                $GameData.Board[$_row, $_col] = [int]$GameData.CurrentTetromino.Type
                $GameData.BoardChars[$_row, $_col] = $Shape[$r][$c]
            }
        }
    }
}


function NewCurrentTetromino {
<# creates a new tetromino by changing CurTetromino to use the type defined by NextTetrominoType #>
    if ($null -eq $GameData.NextTetrominoType) { $GameData.NextTetrominoType = GetRandomTetrominoType }
    $GameData.CurrentTetromino.Type = $GameData.NextTetrominoType
    $GameData.CurrentTetromino.Shape = $TetrominoShapes[$GameData.NextTetrominoType]
    $GameData.CurrentTetromino.Column = ($GAMEBOARD_WIDTH - $TETROMINO_LENGTH) / 2
    $GameData.NextTetrominoType = GetRandomTetrominoType
    # get the first bottom row that has something in it - uses ShapeStartRow as temp var before setting it properly
    $GameData.CurrentTetromino.ShapeStartRow = -1
    $GameData.CurrentTetromino.ShapeEndRow = -1
    for ($y = 0; $y -lt $TETROMINO_LENGTH; $y++) {
        for ($x = 0; $x -lt $TETROMINO_LENGTH; $x++) {
            if ($GameData.CurrentTetromino.Shape[$y][$x]) {
                if ($GameData.CurrentTetromino.ShapeStartRow -eq -1) {
                    $GameData.CurrentTetromino.ShapeStartRow = $y
                }
                $GameData.CurrentTetromino.ShapeEndRow = $y
            }
        }
    }
    $GameData.CurrentTetromino.ShapeStartRow = [System.Math]::Max(0, $GameData.CurrentTetromino.ShapeStartRow)
    $GameData.CurrentTetromino.ShapeEndRow = [System.Math]::Max(0, $GameData.CurrentTetromino.ShapeEndRow)
    # set row 
    $GameData.CurrentTetromino.Row = (-$TETROMINO_LENGTH) + $GameData.CurrentTetromino.ShapeEndRow
}


function ProcessFullRows {
param(
    [int[]]$FullRows
)
<# animates the removal of the lowest 4 complete lines, then returns the number of lines cleared. should never be called with more than 4 complete rows#>
    if ($FullRows.Length -eq 0) {
        return
    }
    # update score and other stats
    $pointsForClearedRows = GetNumberOfPointsForLinesCleared -NumLinesCleared ([System.Math]::Min($MAXROWSCLEARED, $FullRows.Length)) 
    AddPointsToScore -AddPoints $pointsForClearedRows
    $GameData.Lines += ($fullrows.Count)
    $GameData.Level = [Math]::Floor($GameData.Lines / 10) + 1
    # fade out the each row
    foreach ($fullrow in $FullRows) {
        for ($x = 0; $x -lt $GAMEBOARD_WIDTH; $x++) {
            $GameData.BoardChars[$fullrow, $x] = $GAMEBOARD_EXPLODEVAL
        }
    }
    DrawGameBoard
    PrintConBuff
    Start-Sleep -Milliseconds 300
    foreach ($fullrow in $FullRows) {
        for ($x = 0; $x -lt $GAMEBOARD_WIDTH; $x++) {
            $GameData.BoardChars[$fullrow, $x] = $GAMEBOARD_EMPTYVAL
        }
    }
    DrawGameBoard
    PrintConBuff
    Start-Sleep -Milliseconds 300
    # remove the complete rows and reprint - copy non-empty rows in to temp array, then reassign
    $desty = $GAMEBOARD_HEIGHT - 1
    for ($y = ($GAMEBOARD_HEIGHT - 1); $y -ge 0; $y--) {
        if ($fullrows.Contains($y)) { continue }                # skip copy if it's a full row
        for ($x = 0; $x -lt $GAMEBOARD_WIDTH; $x++) {
            $GameData.Board[$desty, $x] = $GameData.Board[$y, $x]
            $GameData.BoardChars[$desty, $x] = $GameData.BoardChars[$y, $x]
        }
        $desty--
    }
    while ($desty -ge 0) {                                       # fill rows we didn't copy to with empty spaces
        for ($x = 0; $x -lt $GAMEBOARD_WIDTH; $x++) {
            $GameData.Board[$desty, $x] = $GAMEBOARD_EMPTYVAL
            $GameData.BoardChars[$desty, $x] = $CHAR_EMPTY
        }
        $desty--
    }
    DrawGameBoard    
    PrintConBuff
}



function MoveTetrominoDown {  
<# moves the tetromino down a row. returns true if successful, false if it's game over#>
param(
    [switch]$CalledByDropTimer           # true if this function was called by the drop timer
) 
    # remove tetromino from current position, then check if will fit in it's new proposed position.
    PlaceCurrentTetromino -Remove
    $willfit = WillCurrentTetrominoFit -Row ($GameData.CurrentTetromino.Row + 1) -Column $GameData.CurrentTetromino.Column
    if ($true -eq $willfit) {
        $GameData.CurrentTetromino.Row++
        PlaceCurrentTetromino
        return $true
    }
    # we didn't fit, so we rest the shape on the bottom and draw a new one.    
    PlaceCurrentTetromino            
    # if we're on the top row then we can't add any more tetrominos - that's game over.
    if ($GameData.CurrentTetromino.Row + $GameData.CurrentTetromino.ShapeStartRow -le 0) {
        return $false
    }
    # add points if this was a non-timer drop
    if (-not $CalledByDropTimer.IsPresent) {
        AddPointsToScore -PointToAdd ($PointsPerNewOnKeyDown * $GameData.Level)        
    }
    # remove full lines
    $fullrows = GetFullBoardRows
    if ($fullrows.Count -gt 0) {
        ProcessFullRows -FullRows $fullrows
    }
    # create new tetromino
    NewCurrentTetromino
    $willfit = WillCurrentTetrominoFit -Row $GameData.CurrentTetromino.Row -Column $GameData.CurrentTetromino.Column
    if ($false -eq $willfit) {
        PlaceCurrentTetromino
        return $false
    }    
    PlaceCurrentTetromino
    $GameData.Statistics[$GameData.CurrentTetromino.Type]++
    return $true
}


function MoveTetrominoLeft {
<# slide to the left #>
PlaceCurrentTetromino -Remove
    $willfit = WillCurrentTetrominoFit -Row $GameData.CurrentTetromino.Row -Column ($GameData.CurrentTetromino.Column  - 1)
    if ($willfit) {
        $GameData.CurrentTetromino.Column--
    }
    PlaceCurrentTetromino
    return [bool]($willfit)
}


function MoveTetrominoRight {
<# slide to the right #>
    PlaceCurrentTetromino -Remove
    $willfit = WillCurrentTetrominoFit -Row $GameData.CurrentTetromino.Row -Column ($GameData.CurrentTetromino.Column  + 1)
    if ($willfit) {
        $GameData.CurrentTetromino.Column++
    }
    PlaceCurrentTetromino
    return [bool]($willfit)
}


function RotateTetrominoClockwise {
<# xris xross - rotates the shape of the current tetromino clockwise (if it fits) #>
    PlaceCurrentTetromino -Remove
    $old_shape = $GameData.CurrentTetromino.Shape
    $new_shape = @((New-Object int[](4)), (New-Object int[](4)), (New-Object int[](4)), (New-Object int[](4)))
    # rotate your owl!
    for ($r = 0; $r -lt $TETROMINO_LENGTH; $r++) {
        for ($c = 0; $c -lt $TETROMINO_LENGTH; $c++) {
            $new_shape[$c][($TETROMINO_LENGTH - 1 - $r)] = $old_shape[$r][$c]            
        }
    }
    $GameData.CurrentTetromino.Shape = $new_shape
    # if we don't fit, revert back to old shape
    $willfit = WillCurrentTetrominoFit -Row $GameData.CurrentTetromino.Row -Column $GameData.CurrentTetromino.Column
    if (-not $willfit) {
        $GameData.CurrentTetromino.Shape = $old_Shape
    }
    PlaceCurrentTetromino
    return [bool]($willfit)
}




#######################################################################################################
# Games
#######################################################################################################



function GameOver {
    # end the current game
    $GameData.EndTime = [datetime]::Now
    DrawGameText
    DrawGameBoard
    # print game over and wait for a second
    [console]::SetCursorPosition(0, 0)
    ConBuffDrawFramedBox -Char $CHAR_EMPTY -x ($GAMEAREA_COL2 - 1) -y ($GAMEAREA_ROW1 + ($GAMEBOARD_HEIGHT / 2) - 1)  -Height 3 -Width 14 -BoxForeColour $FORECOLOUR_GAMETEXT -BoxBackColour $BACKCOLOUR_GAMEAREA -FrameForeColour $FORECOLOUR_BORDER -FrameBackColour $BACKCOLOUR_BORDER
    ConBuffWrite -Text "Game Over!" -x ($GAMEAREA_COL2 + 1) -y ($GAMEAREA_ROW1 + ($GAMEBOARD_HEIGHT / 2)) -ForeColour $FORECOLOUR_GAMETEXT -BackColour $BACKCOLOUR_GAMETEXT
    PrintConBuff
    Start-Sleep -Seconds 1;
    # clear kb buffer a couple of times
    ClearKeyboardBuffer 
    Start-Sleep -Milliseconds 10    
    ClearKeyboardBuffer
    # wait for any key to return to main menu
    while ($true) {
        if ([Console]::KeyAvailable) {
            [Console]::ReadKey($false) | Out-Null
            break;
        }
        Start-Sleep -Milliseconds 10;
    }
    throw New-Object Exception("Game Over")
}


function HandleGameKey {
<# handles a game key, and returns true if handled #>
param(
    [ConsoleKeyInfo]$KeyInfo
)
    switch ($KeyInfo.Key) {
        'DownArrow' {
            if (-not (MoveTetrominoDown)) {
                GameOver
            }
            return $true
        }
        'LeftArrow' {
            MoveTetrominoLeft
            return $true
        }
        'RightArrow' {
            MoveTetrominoRight
            return $true
        }
        'UpArrow' {
            RotateTetrominoClockwise
            return $true
        }
        'Escape' {
            GameOver
            return $false
        }
        default {
            return $false
        }
    }
}



function DoAGame {

    ## begin DoAGame
    ResetGameData
    $GameData.StartTime = [datetime]::Now
    DrawGameBackground

    NewCurrentTetromino
    PlaceCurrentTetromino

    $last_tetrominofallms = $Stopwatch.ElapsedMilliseconds
    $cur_tetrominofalldelayms = GetDelaymsForCurrentLevel

    [bool]$updatereq = $true
    do {

        # update the screen ?
        if ($true -eq $updatereq) {
            $updatereq = $false
            DrawGameBoard
            DrawGameText
            PrintConBuff
            [console]::SetCursorPosition(0, 0)
        }
        Start-Sleep -Milliseconds 5

        $millis_now = $Stopwatch.ElapsedMilliseconds

        if ([Console]::KeyAvailable) {
            $_key = [Console]::ReadKey($false)
            if ($true -eq (HandleGameKey -KeyInfo $_key)) {            
                ClearKeyboardBuffer
                $updatereq = $true
            }
        }

        # tetromino fall ?
        $tetrominofallms = $millis_now
        if ($tetrominofallms - $last_tetrominofallms -ge $cur_tetrominofalldelayms) {
            $last_tetrominofallms = $tetrominofallms
            if (-not (MoveTetrominoDown -CalledByDropTimer)) {
                GameOver
            }
            $updatereq = $true
        }

    } while ($true)

}




#######################################################################################################
# Game - Main Menu
#######################################################################################################


# http://patorjk.com/software/taag/
$TITLEMASK = @(
    "           xxxx   xxx"
    "           x   x x"
    "           xxxx   xxx"
    "           x         x"
    "           x     xxxx"
    ""	   
    "xxxxxx xxxx xxxxxx xxxx  xxx  xxx "
    "  xx   x      xx   x   x  x  x    "
    "  xx   xxx    xx   xxxx   x   xxx "
    "  xx   x      xx   x x    x      x"
    "  xx   xxxx   xx   x  xx xxx xxxx " 
)


# our main menu options
Enum MainMenuOption {
    NewAGame
    NewBGame
    Controls
    Quit
}

# main menu value to name lookup
$MainMenuOptionNames =@{
    [MainMenuOption]::NewAGame = "New A Game"
    [MainMenuOption]::Controls = "Controls"
    [MainMenuOption]::Quit = "Quit"
}

# our main menu object - this is kinda how it looks on screen
$MainMenu = NewMenuObject -Items @(
    $MainMenuOptionNames[[MainMenuOption]::NewAGame], 
    "", 
    $MainMenuOptionNames[[MainMenuOption]::Controls], 
    "", 
    $MainMenuOptionNames[[MainMenuOption]::Quit]
)



function DoMainMenu {
<# This does the Menu you see before playing any games #>

    # the different types of menu displayed
    Enum MenuScreenType {
        Main
        Controls
    }    
    # the currently displayed menu    
    $CurrentMenuDisplayed = [MenuScreenType]::Main
    # a menu object for displaying our controls - not really a menu, but works as one.
    $ControlsMenu = NewMenuObject -Items @(
        ("   {0}    Rotate" -f $CHAR_ARROWUP)
        ("   {0}    Down" -f $CHAR_ARROWDOWN)
        ("   {0}    Right" -f $CHAR_ARROWRIGHT)
        ("   {0}    Left" -f $CHAR_ARROWLEFT)
        "  ESC   Quit"
    )
    # should the top score currently be displayed on screen. 
    [bool]$DisplayTopScore = $false


    function Paint {
    <# local paint, to paint the currently displayed menu #>
        # draw the background and title
        ClearConBuff
        CopyToConBuff -Source $FancyBackground
        ConBuffDrawFromMask -x 0 -y 2 -Char $CHAR_FADE2 -Mask $TITLEMASK -ForeColour $FORECOLOUR_TITLESHADOW -BackColour $BACKCOLOUR_TITLESHADOW
        ConBuffDrawFromMask -x 1 -y 1 -Char $CHAR_SOLIDBLOCK -Mask $TITLEMASK -ForeColour $FORECOLOUR_TITLE -BackColour $BACKCOLOUR_TITLE
        # display top score, if there is one
        if ($true -eq $DisplayTopScore) {
            $_score_str = "Top Score: " + [string]::Format('{0:N0}', $GameData.TopScore)
            ConBuffWrite -Text $_score_str -x (($CONBUFF_WIDTH / 2) - ($_score_str.Length / 2)) -y 14 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_BACKGROUND
        }
        # draw menu and content
        ConBuffDrawFramedBox -Char $CHAR_EMPTY -x 8 -y 16 -Width 20 -Height 9 -BoxForeColour $FORECOLOUR_BORDER -BoxBackColour $BACKCOLOUR_BORDER -FrameForeColour $FORECOLOUR_BACKGROUND -FrameBackColour $BACKCOLOUR_BORDER
        switch ($CurrentMenuDisplayed) {
            'Main' {
                ConBuffDrawMenu -Menu $MainMenu -x 10 -y 18 -Width 16 -DrawSelected -CenterAligned
                break
            }
            'Controls' {
                ConBuffDrawMenu -Menu $ControlsMenu -x 10 -y 18 -Width 16
                break
            }
        }
        # write to screen
        PrintConBuff
    }

    ## begin DoMainMenu

    [bool]$updatereq = $true                            # is a screen update required
    [bool]$kbHandled = $false
    $millis_scoreblink = $Stopwatch.ElapsedMilliseconds
    while ($true) {

        $millis_now = $Stopwatch.ElapsedMilliseconds
        
        # blink the top score?
        if ($GameData.TopScore -gt 0) {
            if ($millis_now - $millis_scoreblink -ge $MAINMENU_TOPSCORE_BLINKMS) {
                $millis_scoreblink = $millis_now
                $DisplayTopScore = -not $DisplayTopScore
                $updatereq = $true
            }
        }

        # handle key press
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($false)
            switch ($CurrentMenuDisplayed) {
                'Controls' {
                    # escape out to main menu
                    if ($key.Key -eq [consolekey]::Enter) {
                        $CurrentMenuDisplayed = [MenuScreenType]::Main
                        $kbHandled = $true
                        break                     
                    }
                }
                'Main' {
                    # select menu item below
                    if ($key.Key -eq [consolekey]::DownArrow) {
                        $MainMenu.Index = GetNextMenuIndex -MenuObject $MainMenu
                        $kbHandled = $true
                        break
                    }
                    # select menu item above
                    elseif ($key.Key -eq [consolekey]::UpArrow) {
                        $MainMenu.Index = GetLastMenuIndex -MenuObject $MainMenu
                        $kbHandled = $true
                        break
                    }
                    # select the menu item
                    elseif ($key.Key -eq [consolekey]::Enter) {                        
                        switch ($MainMenu.Index) {
                            0 { return "AGame" }
                            2 {
                                $CurrentMenuDisplayed = [MenuScreenType]::Controls
                                $kbHandled = $true
                                break
                            }
                            4 { return "Quit" }
                        }
                    }
                }
            }
        }

        if ($true -eq $kbHandled) {
            ClearKeyboardBuffer
            $kbHandled = $false
            $updatereq = $true
        }
        
        # update the screen ?
        if ($true -eq $updatereq) {
            $updatereq = $false
            Paint
        }

        Start-Sleep -Milliseconds 5
    }

}




#######################################################################################################
# Fancy Background
#######################################################################################################


# this stores the game background screen
$FancyBackground = NewConBuffObject

$fancy_background_chars = $CHAR_EMPTY, $CHAR_FADE3, $CHAR_FADE2, $CHAR_FADE1, $CHAR_SOLIDBLOCK
for ($idx = 0; $idx -lt $CONBUFF_SIZE; $idx++) {
    $_randpos = Get-Random -Minimum 0 -Maximum $fancy_background_chars.Length
    $FancyBackground.Buffer[$idx] = [int]($fancy_background_chars[$_randpos])
    $FancyBackground.ForeColours[$idx] = [ConsoleColor]$FORECOLOUR_BACKGROUND
    $FancyBackground.BackColours[$idx] = [ConsoleColor]$BACKCOLOUR_BACKGROUND
}




#######################################################################################################
# BEGIN
#######################################################################################################


try {
    ## init screen
    SetCursorVisible -Visible $false
    Clear-Host
    [console]::SetCursorPosition(0, 0)    
    ResetLastConBuff        # init last print buffer to a char that's unliklely to be printed

    # init game stuff
    ResetGameData

    # start main game
    while ($true) {
        $menuOption = DoMainMenu
        switch ($menuOption) {
            'AGame' {
                $GameData = $AllGameData[[GameType]::AGame]
                try {
                    DoAGame
                }
                catch {
                    # if the game didnt' end cleanly, rethrow the exception                    
                    if ($GameData.EndTime -eq [datetime]::MinValue) {
                        throw $_
                    }
                }
                break
            }
            'Quit' {
                return
            }
        }
    }

}
catch {    
    $caught_thing = $_   
    throw $caught_thing 
}
finally {
    # reset console
    [console]::ForegroundColor = $SAVED_FORECOLOUR
    [console]::BackgroundColor = $SAVED_BACKCOLOUR
    [console]::SetCursorPosition($CONBUFF_WIDTH, $CONBUFF_HEIGHT);
    SetCursorVisible -Visible $SAVED_CURSORVISIBLE 
    "" | Out-Host
}

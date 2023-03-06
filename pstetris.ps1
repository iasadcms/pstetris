<#
.SYNOPSIS
    PowerShell Tetris

.DESCRIPTION
    Quick and dirty PowerShell Tetris

    Use arrow keys - left, right, down, up == rotate.
    Escape == quit

    No music
#>


$CONBUFF_WIDTH = 36                 # width of the console buffer
$CONBUFF_HEIGHT = 27                # height of the console buffer
$GAMEAREA_WIDTH = 10                # width of the tetris game area
$GAMEAREA_HEIGHT = 18               # height of the tetris game area


$FORECOLOUR_BACKGROUND = 'Cyan';        $BACKCOLOUR_BACKGROUND = 'DarkGray'
$FORECOLOUR_BORDER = 'Cyan';            $BACKCOLOUR_BORDER = 'Black'
$FORECOLOUR_GAMETEXT = 'White';         $BACKCOLOUR_GAMETEXT = 'Black'
$FORECOLOUR_TITLE = 'White';            $BACKCOLOUR_TITLE = 'Black'
$FORECOLOUR_TITLESHADOW = 'DarkGray';   $BACKCOLOUR_TITLESHADOW = 'Black'
$FORECOLOUR_SCORE = 'White';            $BACKCOLOUR_SCORE = 'Black'
$FORECOLOUR_MENU = 'White';             $BACKCOLOUR_MENU = 'Black'
$BACKCOLOUR_GAMEAREA = 'Black'


$CHAR_SOLIDBLOCK = [char]9608
$CHAR_FADE1 = [char]9619
$CHAR_FADE2 = [char]9618
$CHAR_FADE3 = [char]9617
$CHAR_EMPTY = [char]32
$CHAR_ARROWLEFT = [char]9668
$CHAR_ARROWRIGHT = [char]9658
$CHAR_ARROWUP = [char]9650
$CHAR_ARROWDOWN = [char]9660


#######################################################################################################
# Cursor
#######################################################################################################


# win32 type lib - helps remove blinking cursor
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
$global:hStdOut = [win32]::GetStdHandle(-11)
$global:CursorInfo = New-Object Win32+CONSOLE_CURSOR_INFO


function GetCursorVisible {
    if ([win32]::GetConsoleCursorInfo($hStdOut, [ref]$global:CursorInfo)) {
        return [bool]($global:CursorInfo.bVisible)
    }
    else { return $true }
}


function SetCursorVisible {
param(
    [bool]$Visible = $true
)
    if ([win32]::GetConsoleCursorInfo($hStdOut, [ref]$global:CursorInfo)) {
        $global:CursorInfo.bVisible = $Visible
        [win32]::SetConsoleCursorInfo($hStdOut, [ref]$global:CursorInfo) | Out-Null
    }
}


function SetCursorPos {
    <# sets the console cursor at x, y #>
    param(
        [parameter(Mandatory=$true)]$x, 
        [parameter(Mandatory=$true)]$y
    )    
    $x = [System.Math]::Max([System.Math]::Min($x, [console]::WindowWidth), 0)
    $y = [System.Math]::Max([System.Math]::Min($y, [console]::WindowHeight), 0)
    [Console]::SetCursorPosition($x, $y);
}




#######################################################################################################
# Keyboard
#######################################################################################################


function ClearKeyboardBuffer {
<# clears the keyboard buffer #>
    while (([Console]::KeyAvailable)) {
        [void][Console]::ReadKey($false)
    }
}




#######################################################################################################
# Screen Updates
#######################################################################################################


function ConsoleWriteAt {
<# writes text to the console screen in using forecolour, and backcolour, at position x, y #>
param(
    [parameter(Mandatory=$true)][string]$Text,
    [ConsoleColor]$ForeColour = $SavedForeColour,
    [ConsoleColor]$BackColour = $SavedBackColour,
    [parameter(Mandatory=$true)][int]$x,
    [parameter(Mandatory=$true)][int]$y
)
    SetCursorPos -x $x -y $y
    [Console]::BackgroundColor = $BackColour
    [Console]::ForegroundColor = $ForeColour        
    [Console]::Write($Text)
}




#######################################################################################################
# Virtual screen
#######################################################################################################


$CONBUFF_IDX_VALUE = 0
$CONBUFF_IDX_FORECOL = 1
$CONBUFF_IDX_BACKCOL = 2
$CONBUFF_IDX_LENGTH = ($CONBUFF_IDX_BACKCOL + 1)

# this is our virtual screen, and a copy of it 
$global:ConBuff = New-Object 'System.Object[,,]'($CONBUFF_WIDTH, $CONBUFF_HEIGHT, $CONBUFF_IDX_LENGTH)
$global:LastPrintedConBuff = New-Object 'System.Object[,,]'($CONBUFF_WIDTH, $CONBUFF_HEIGHT, $CONBUFF_IDX_LENGTH)
for ($x = 0; $x -lt $CONBUFF_WIDTH; $x++) {
    for ($y = 0; $y -lt $CONBUFF_HEIGHT; $y++) {
        $global:LastPrintedConBuff[$x, $y, $CONBUFF_IDX_VALUE] = [char]365   # random char chosen to invalidate LastConBuff
    }
}


function UpdateLastConBuff() {
<# Copies $global:ConBuff in to $global:LastPrintedConBuff - PrintConBuff only updates differences between the two, unless forced #>
    for ($x = 0; $x -lt $CONBUFF_WIDTH; $x++) {
        for ($y = 0; $y -lt $CONBUFF_HEIGHT; $y++) {
            $global:LastPrintedConBuff[$x, $y, $CONBUFF_IDX_VALUE] = $global:ConBuff[$x, $y, $CONBUFF_IDX_VALUE]
            $global:LastPrintedConBuff[$x, $y, $CONBUFF_IDX_FORECOL] = $global:ConBuff[$x, $y, $CONBUFF_IDX_FORECOL]
            $global:LastPrintedConBuff[$x, $y, $CONBUFF_IDX_BACKCOL] = $global:ConBuff[$x, $y, $CONBUFF_IDX_BACKCOL]
        }
    }
}


#function ConBuffWriteAt {
<# writes text to $global:ConBuff, starting at position x, y, using gamepart colours #>
<#param(
    [parameter(Mandatory=$true)][string]$Text,
    [parameter(Mandatory=$true)][GamePart]$GamePart,
    [parameter(Mandatory=$true)][int]$x,
    [parameter(Mandatory=$true)][int]$y
)
    for ($pos = 0; $pos -lt $Text.Length; $pos++) {
        $xpos = $x + $pos
        if ($xpos -gt $CONBUFF_WIDTH) { break }
        $global:ConBuff[$xpos, $y, $CONBUFF_IDX_VALUE] = $Text[$pos]
        $global:ConBuff[$xpos, $y, $CONBUFF_IDX_FORECOL] = $gamePartDefs[$GamePart]['ForeColour']
        $global:ConBuff[$xpos, $y, $CONBUFF_IDX_BACKCOL] = $gamePartDefs[$GamePart]['BackColour']
    }
}#>


function ConBuffWrite {
<# writes text to $global:ConBuff, starting at position x, y #>
param(
    [parameter(Mandatory=$true)][string]$Text,
    [parameter(Mandatory=$true)][int]$x,
    [parameter(Mandatory=$true)][int]$y,
    [ConsoleColor]$ForeColour = $SavedForeColour,
    [ConsoleColor]$BackColour = $SavedBackColour
)    
if ($y -lt 0) { return }
if ($y -ge $CONBUFF_HEIGHT) { return }
for ($pos = 0; $pos -lt $Text.Length; $pos++) {
        $xpos = $x + $pos
        if ($xpos -lt 0) { break }
        if ($xpos -ge $CONBUFF_WIDTH) { break }
        $global:ConBuff[$xpos, $y, $CONBUFF_IDX_VALUE] = $Text[$pos]
        $global:ConBuff[$xpos, $y, $CONBUFF_IDX_FORECOL] = $ForeColour
        $global:ConBuff[$xpos, $y, $CONBUFF_IDX_BACKCOL] = $BackColour
    }
}


function ClearConBuff() {
<# CLS, but for $global:ConBuff #>
    for ($x = 0; $x -lt $CONBUFF_WIDTH; $x++) {
        for ($y = 0; $y -lt $CONBUFF_HEIGHT; $y++) {
            $global:ConBuff[$x, $y, $CONBUFF_IDX_VALUE] = " "
            $global:ConBuff[$x, $y, $CONBUFF_IDX_FORECOL] = $SavedForeColour
            $global:ConBuff[$x, $y, $CONBUFF_IDX_BACKCOL] = $SavedBackColour
        }
    }    
}


function PrintConBuff {
<# displays $global:ConBuff in the console window #>
param(
    [switch]$Force,                   # if not set, only changes are updated
    [switch]$NoCursorReset            # cursor is not reset to pos it was at entry 
)

    # save cursor position for later. we'll reset it back to this on our way out.
    $cursorX = [Console]::CursorLeft
    $cursorY = [Console]::CursorTop

    # loop through the screen buffer and write out any changes
    for ($y = 0; $y -lt $CONBUFF_HEIGHT; $y++) {
        #if ($y -ge $CONBUFF_HEIGHT) { break }
        for ($x = 0; $x -lt $CONBUFF_WIDTH; $x++) {
            #if ($x -ge $CONBUFF_WIDTH) { break }

            $new_forecol = $global:ConBuff[$x, $y, $CONBUFF_IDX_FORECOL]
            $new_backcol = $global:ConBuff[$x, $y, $CONBUFF_IDX_BACKCOL]
            $new_value = $global:ConBuff[$x, $y, $CONBUFF_IDX_VALUE]
            $new_value = if ([string]::IsNullOrEmpty($new_value)) { " " } else { $new_value.ToString() }

            # update if required 
            $_updatereq = $Force.IsPresent -or (($new_forecol -ne $global:LastPrintedConBuff[$x, $y, $CONBUFF_IDX_FORECOL]) -or ($new_backcol -ne $global:LastPrintedConBuff[$x, $y, $CONBUFF_IDX_BACKCOL]) -or ($new_value -ne $global:LastPrintedConBuff[$x, $y, $CONBUFF_IDX_VALUE]))
            if ($_updatereq -eq $true) {
                ConsoleWriteAt -x:($cursorX + $x) -y:($cursorY + $y) -Text:$new_value -ForeColour:$new_forecol -BackColour:$new_backcol
            }
        }
    }
    
    # confirm changes written to screen - use as compare next time
    for ($x = 0; $x -lt $CONBUFF_WIDTH; $x++) {
        for ($y = 0; $y -lt $CONBUFF_HEIGHT; $y++) {
            for ($z = 0; $z -lt $CONBUFF_IDX_LENGTH; $z++) {
                $global:LastPrintedConBuff[$x, $y, $z] = $global:ConBuff[$x, $y, $z]
            }
        }
    }    
    
    # reset cursor, if it's been moved
    if (($cursorX -ne ([Console]::CursorLeft)) -or (($cursorY -ne ([Console]::CursorTop)))) {
        if (-not $NoCursorReset.IsPresent) {
            SetCursorPos -x $cursorX -y $cursorY
        }
    }
    

}




#######################################################################################################
# Virtual screen drawing funcs
#######################################################################################################


function ConBuffCopyFrom($Source) {
<# copies source buff on to $global:ConBuff #>
    for ($x = 0; $x -lt $CONBUFF_WIDTH; $x++) {
        for ($y = 0; $y -lt $CONBUFF_HEIGHT; $y++) {                    
            $global:ConBuff[$x, $y, $CONBUFF_IDX_VALUE] = $Source[$x, $y, $CONBUFF_IDX_VALUE]
            $global:ConBuff[$x, $y, $CONBUFF_IDX_FORECOL] = $Source[$x, $y, $CONBUFF_IDX_FORECOL]
            $global:ConBuff[$x, $y, $CONBUFF_IDX_BACKCOL] = $Source[$x, $y, $CONBUFF_IDX_BACKCOL]
        }
    }
}


function ConBuffFillBox {
<# writes box of a char to $global:ConBuff, starting at position x, y, with length, height (down and right) #>
param(
    [parameter(Mandatory=$true)][char]$Char,
    [parameter(Mandatory=$true)][int]$x,
    [parameter(Mandatory=$true)][int]$y,
    [parameter(Mandatory=$true)][int]$Width,
    [parameter(Mandatory=$true)][int]$Height,
    [parameter(Mandatory=$true)][ConsoleColor]$ForeColour,
    [parameter(Mandatory=$true)][ConsoleColor]$BackColour
)
    for ($_x = $x; $_x -lt ($x + $Width); $_x++) {
        for ($_y = $y; $_y -lt ($y + $Height); $_y++) {
            ConBuffWrite -Text:$Char -x $_x -y $_y -ForeColour:$ForeColour -BackColour:$BackColour
        }
    }
}


function ConBuffDrawBox {
<# draws a box, starting at position x, y, with length, height (down and right).#>
param(
    [parameter(Mandatory=$true)][int]$x,
    [parameter(Mandatory=$true)][int]$y,
    [parameter(Mandatory=$true)][int]$Width,
    [parameter(Mandatory=$true)][int]$Height,
    [parameter(Mandatory=$true)][ConsoleColor]$ForeColour,
    [parameter(Mandatory=$true)][ConsoleColor]$BackColour
)
    # draw top and bottom lines
    for ($_x = $x + 1; $_x -lt ($x + $Width) - 1; $_x++) { 
        ConBuffWrite -Text ([char]0x2550)  -x $_x -y $y -ForeColour:$ForeColour -BackColour:$BackColour
        ConBuffWrite -Text ([char]0x2550)  -x $_x -y ($y + $Height - 1)  -ForeColour:$ForeColour -BackColour:$BackColour
    }
    # draw left and right lines
    for ($_y = $y + 1; $_y -lt ($y + $Height) - 1; $_y++) { 
        ConBuffWrite -Text ([char]0x2551) -x $x -y $_Y  -ForeColour:$ForeColour -BackColour:$BackColour 
        ConBuffWrite -Text ([char]0x2551) -x ($x + $Width - 1) -y $_y  -ForeColour:$ForeColour -BackColour:$BackColour
    }
    # draw corners
    ConBuffWrite -Text ([char]0x2554) -x $x -y $y  -ForeColour:$ForeColour -BackColour:$BackColour 
    ConBuffWrite -Text ([char]0x2557) -x ($x + $Width - 1) -y $y  -ForeColour:$ForeColour -BackColour:$BackColour
    ConBuffWrite -Text ([char]0x255A) -x $x -y ($y + $Height - 1)  -ForeColour:$ForeColour -BackColour:$BackColour 
    ConBuffWrite -Text ([char]0x255D) -x ($x + $Width - 1) -y ($y + $Height - 1)  -ForeColour:$ForeColour -BackColour:$BackColour
}


function ConBuffDrawBorderexFillBox {
param(
    [parameter(Mandatory=$true)][char]$Char,
    [parameter(Mandatory=$true)][int]$x,
    [parameter(Mandatory=$true)][int]$y,
    [parameter(Mandatory=$true)][int]$Width,
    [parameter(Mandatory=$true)][int]$Height,
    [parameter(Mandatory=$true)][ConsoleColor]$ForeColour,
    [parameter(Mandatory=$true)][ConsoleColor]$BackColour
)
    $Width = [System.Math]::Max($Width, 2)
    $Height = [System.Math]::Max($Height, 2)
    ConBuffFillBox -Char:$Char  -ForeColour:$ForeColour -BackColour:$BackColour -x ($x + 1) -y ($y + 1) -Width ($Width - 1) -Height ($Height - 1)
    ConBuffDrawBox -ForeColour:$ForeColour -BackColour:$BackColour -x:$x -y:$y -Width:$Width -Height:$Height
}




#######################################################################################################
# Tetromino town
#######################################################################################################


# Define the TetrominoType enum
Add-Type -TypeDefinition @"
public enum TetrominoType {
    I = 0,
    J = 1,
    L = 2,
    O = 3,
    S = 4,
    T = 5,
    Z = 6
}
"@


$TETROMINO_SIZE = 4                     # size x, y size of a tetromino
$TETROMINO_CHAR = $CHAR_SOLIDBLOCK      # the character to use for a tetromino

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


$TETROMINOTYPE_MAXINTVALUE = [TetrominoType].GetEnumValues() | ForEach-Object { [int]$_ } | Sort-Object | Select-Object -Last 1

function GetRandomTetrinoType {   
<# returns a random tetromino type value #> 
    return [TetrominoType](Get-Random -Minimum 0 -Maximum $TETROMINOTYPE_MAXINTVALUE)
}
 


# the current tetromino
$global:CurTetromino = [pscustomobject]@{
    Type = $null
    Shape = $null
    Row = 0
    Column = 0
    StartRow = 0
}
# the type of the next tetromino
$global:NextTetrominoType = GetRandomTetrinoType



$TetrominoColours = @{
    [TetrominoType]::I = [ConsoleColor]::Cyan
    [TetrominoType]::J = [ConsoleColor]::Blue
    [TetrominoType]::L = [ConsoleColor]::DarkYellow
    [TetrominoType]::O = [ConsoleColor]::Gray
    [TetrominoType]::S = [ConsoleColor]::Green
    [TetrominoType]::T = [ConsoleColor]::Magenta
    [TetrominoType]::Z = [ConsoleColor]::Red
}  


function GetTetrominoColour {
param(
    [parameter(Mandatory=$true)]$Type    
)    
<# returns a colour for each teromino shape #>   
    return $TetrominoColours[$Type]
}


function ConBuffDrawTetromino {
<# draws a tetromino on conbuff at x, y#>
param(
    [parameter(Mandatory=$true)]$Type,
    [parameter(Mandatory=$true)][int]$x,
    [parameter(Mandatory=$true)][int]$y
)
    $colour = GetTetrominoColour -Type $Type
    $_shape = $TetrominoShapes[$Type]
    for ($_x = 0; $_x -lt $TETROMINO_SIZE; $_x++) {
        for ($_y = 0; $_y -lt $TETROMINO_SIZE; $_y++) {
            if ($_shape[$_y][$_x]) { 
                $_text = $TETROMINO_CHAR
                ConBuffWrite -Text $_text -x ($x + $_x) -y ($y + $_y) -ForeColour $colour -BackColour 'Black'
            }             
        }
    }
}  




#######################################################################################################
# Board
#######################################################################################################


# game board - the value in each cell is the int value of TetrominoType. -1 is blank
$global:GameBoard = [int[,]]::new($GAMEAREA_WIDTH, $GAMEAREA_HEIGHT)

[int]$GAMEBOARD_EMPTYVAL = -1                                       # the value representing an empty game board cell
[int]$GAMEBOARD_EXPLODEVAL = $TETROMINOTYPE_MAXINTVALUE + 1         # the int value representing an exploding cell


# hashtable of board values > char to print on the board. Exception is tetromino.
$TetrominoBoardChars = @{} 
$TetrominoBoardChars[$GAMEBOARD_EMPTYVAL] = ' '
[TetrominoType].GetEnumValues() | ForEach-Object { $TetrominoBoardChars[([int]$_)] = $TETROMINO_CHAR }
$TetrominoBoardChars[$GAMEBOARD_EXPLODEVAL] = $CHAR_FADE2


function ClearGameBoard {
<# clears the game board - fills with empty values#>
    for ($i = 0; $i -lt $GAMEAREA_WIDTH; $i++) {
        for ($j = 0; $j -lt $GAMEAREA_HEIGHT; $j++) {
            $global:GameBoard[$i, $j] = $GAMEBOARD_EMPTYVAL
        }
    }
}


function GetBoardObjectForeColour {
<# returns the fore colour of a the value of a board cell#>
param(
    [parameter(Mandatory=$true)]$Object
)
    switch ($Object) {
        $GAMEBOARD_EMPTYVAL { return [ConsoleColor]::Black }
        $GAMEBOARD_EXPLODEVAL { return [ConsoleColor]::Gray }
    }    
    return GetTetrominoColour -Type ([Enum]::ToObject([TetrominoType], $Object))
}


function PlaceCurTetromino {
<# paints or removes the current tetromino on the board #>
param(
    [switch]$Remove                 # remove tetromino from board. 
)
    if (($null -eq $global:CurTetromino) -or ($null -eq $global:CurTetromino.Shape)) {
        return
    }
    for ($r = 0; $r -lt $TETROMINO_SIZE; $r++) {
        for ($c = 0; $c -lt $TETROMINO_SIZE; $c++) {
            if ($global:CurTetromino.Shape[$r][$c]) {
                $_row = $global:CurTetromino.Row + $r
                $_col = $global:CurTetromino.Column + $c
                if (($_row -ge 0) -and ($_row -lt $GAMEAREA_HEIGHT) -and ($_col -ge 0) -and ($_col -lt $GAMEAREA_WIDTH)) {
                    [int]$val = if ($Remove.IsPresent) { $GAMEBOARD_EMPTYVAL } else { [int]$global:CurTetromino.Type }
                    $global:GameBoard[$_col, $_row] = $val
                }
            }
        }
    }
}


function NewTetromino {
<# creates a new tetromino by changing CurTetromino to use the type defined by NextTetrominoType #>
    if ($null -eq $global:NextTetrominoType) { 
        $global:NextTetrominoType = GetRandomTetrinoType 
    }
    $global:CurTetromino.Type = $global:NextTetrominoType
    $global:CurTetromino.Shape = $TetrominoShapes[$global:NextTetrominoType]
    $global:CurTetromino.Column = ($GAMEAREA_WIDTH - $TETROMINO_SIZE) / 2
    $global:NextTetrominoType = GetRandomTetrinoType
    # get the first line that has something in it. We want that to be the first thing that displays on row 0.
    $global:CurTetromino.StartRow = -1;   # default to -1
    for ($y = 0; $y -lt $TETROMINO_SIZE; $y++) {
        for ($x = 0; $x -lt $TETROMINO_SIZE; $x++) {
            if ($global:CurTetromino.Shape[$y][$x]) {                
                $global:CurTetromino.StartRow = $y
                break
            }
        }
        if ($global:CurTetromino.StartRow -ge 0) { 
            break 
        }
    }
    if ($StartRow -eq -1) { 
        $StartRow = 0
    }
    $global:CurTetromino.Row = (0 - $global:CurTetromino.StartRow)
    # will it fit ?
    $willfit = WillCurTetrominoFit -Row $global:CurTetromino.Row -Column $global:CurTetromino.Column
    if (-not $willfit) {
        GameOver
    }
    # update stats
    $global:GameData_ObjectStats[([int]$global:CurTetromino.Type)]++
}


# Will the current tetromino fit at a given row and column on the game board
function WillCurTetrominoFit {
param(
    [parameter(Mandatory=$true)][int]$Row,
    [parameter(Mandatory=$true)][int]$Column
)
    for ($r = 0; $r -lt $TETROMINO_SIZE; $r++) {
        for ($c = 0; $c -lt $TETROMINO_SIZE; $c++) {
            if (-not $global:CurTetromino.Shape[$r][$c]) { 
                continue 
            }
            # is row out of bounds
            if (($Row + $r -lt 0) -or ($Row + $r -ge $GAMEAREA_HEIGHT)) { 
                return $false 
            }
            # is column out of bounds
            if (($Column + $c -lt 0) -or ($Column + $c -ge $GAMEAREA_WIDTH)) { 
                return $false 
            }
            # is gameboard empty at spot
            if ($global:GameBoard[($Column + $c),($Row + $r)] -gt $GAMEBOARD_EMPTYVAL) { 
                return $false 
            }
        }
    }
    return $true
}





#######################################################################################################
# Game Data
#######################################################################################################


[bool]$global:GracefulGameFinish = $false       # this is set to true when the current game ended gracefully.

[int]$global:GameData_TopScore = 0              # topscore is not reset during 'ResetGame'
[int]$global:GameData_Score = 0                 # score
[int]$global:GameData_Lines = 0                 # number of lines cleared
[int]$global:GameData_Level = 0                 # current level

# game stats - how many of each object
$global:GameData_ObjectStats = New-Object int[] ($TETROMINOTYPE_MAXINTVALUE + 1)


$TOPSCORE_FILENAME = 'pstetris.topscore.txt'

# read-in topscore from file (if exists)
if ((Test-Path $TOPSCORE_FILENAME -ErrorAction SilentlyContinue)) {
    $line = Get-Content 'pstetris.topscore' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not [string]::IsNullOrWhiteSpace($line)) {
        [int]::TryParse($line, [ref]$global:GameData_TopScore) | Out-Null
    }
}


function ResetGame {
<# clears the game board and resets the game #>
    for ($x = 0; $x -lt $GAMEAREA_WIDTH; $x++) {
        for ($y = 0; $y -lt $GAMEAREA_HEIGHT; $y++) {
            $global:GameBoard[$x, $y] = $GAMEBOARD_EMPTYVAL
        }
    }
    $global:GameData_Score, $global:GameData_Lines = 0;
    $global:GameData_Level = 1
    $global:CurTetromino.Type   = $null
    $global:NextTetrominoType = GetRandomTetrinoType    
    for ($i = 0; $i -lt $global:GameData_ObjectStats.Length; $i++) {                 # reset stats
        $global:GameData_ObjectStats[$i] = 0
    }
    $global:GracefulGameFinish = $false
}


# speed of fall in ms per level. Anything above max index will use the last value.
$DelaymsPerLevel = @(1000, 800, 757, 714, 671, 628, 585, 542, 499, 456, 420, 399, 378, 357, 336, 315, 294, 273, 252, 231)

function GetDelaymsForCurrentLevel {   
<# returns the current ms delay for tetrino falls used in the main game loop #>
    $idx = if ($global:GameData_Level -lt $DelaymsPerLevel.Length) { $global:GameData_Level } else { $DelaymsPerLevel.Length }
    return $DelaymsPerLevel[$idx]
}


# number of points for each line cleared, per level
$global:PointsPerLineCleared = @(0, 40, 100, 300, 1200)

function GetNumberOfPointsForLinesCleared {
param(
    [parameter(Mandatory=$true)][int]$NumLinesCleared
)
<# returns the number of points for clearing lines for a level#>
    $NumLinesCleared = [System.Math]::Max(0, [System.Math]::Min(4, $NumLinesCleared))
    $points = ($global:PointsPerLineCleared[$NumLinesCleared] * ($global:GameData_Level + 1))
    return $points
}


function UpdateScore($AddPoints) {
    $global:GameData_Score += $AddPoints
    $global:GameData_TopScore = [System.Math]::Max($global:GameData_TopScore, $global:GameData_Score)
}


function ConvertScoreToString {
<# Converts a score value in to a string of numdigits long #>
param(
    [parameter(Mandatory=$true)]$Value,
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




#######################################################################################################
# Game Control
#######################################################################################################


function CheckAndProcessFullLines {    
<# animates the removal of the lowest 4 complete lines, then returns the number of lines cleared. #>
    # get first 4 full lines
    $completerows = New-Object System.Collections.ArrayList    
    for ($y = $GAMEAREA_HEIGHT - 1; $y -ge 0; $y--) {
        $xcounter = 0
        for ($x = 0; $x -lt $GAMEAREA_WIDTH; $x++) {
            if ($global:GameBoard[$x, $y] -ne $GAMEBOARD_EMPTYVAL) {
                $xcounter++
            }
        }
        # is full row ?
        if ($xcounter -eq $GAMEAREA_WIDTH) {
            [void]$completerows.Add($y)
            if ($completerows.Count -eq 4) { 
                break
            }
        }
    }
    # shortcut break
    if ($completerows.Count -eq 0) {
        return 0
    }    
    # update score
    $points = GetNumberOfPointsForLinesCleared -NumLinesCleared $completerows.Count 
    UpdateScore -AddPoints $points
    $global:GameData_Lines += ($completerows.Count)
    $global:GameData_Level = [Math]::Floor($global:GameData_Lines / 10) + 1
    # fade out the each row
    foreach ($completerow in $completerows) {
        for ($x = 0; $x -lt $GAMEAREA_WIDTH; $x++) {
            $global:GameBoard[$x, $completerow] = $GAMEBOARD_EXPLODEVAL
        }
    }
    PrintGameBoard
    Start-Sleep -Milliseconds 5
    foreach ($completerow in $completerows) {
        for ($x = 0; $x -lt $GAMEAREA_WIDTH; $x++) {
            $global:GameBoard[$x, $completerow] = $GAMEBOARD_EMPTYVAL
        }
    }
    PrintGameBoard
    Start-Sleep -Milliseconds 5
    # remove the complete rows and reprint
    $tempBoard = [int[,]]::new($GAMEAREA_WIDTH, $GAMEAREA_HEIGHT)
    $desty = $GAMEAREA_HEIGHT - 1
    for ($y = ($GAMEAREA_HEIGHT - 1); $y -ge 0; $y--) {
        if ($completerows.Contains($y)) { continue }             # skip copy if it's a full row
        for ($x = 0; $x -lt $GAMEAREA_WIDTH; $x++) {
            $tempBoard[$x, $desty] = $global:GameBoard[$x, $y]
        }
        $desty--
    }
    while ($desty -ge 0) {                                       # fill rows we didn't copy to with empty spaces
        for ($x = 0; $x -lt $GAMEAREA_WIDTH; $x++) {
            $tempBoard[$x, $desty] = $GAMEBOARD_EMPTYVAL
        }
        $desty--
    }
    $global:GameBoard = $tempBoard
    PrintGameBoard    
    return $completerows.length                                  # make sure we returns the num rows cleared
}


function MoveTetrominoDown {  
param(
    [switch]$KeyboardUsed           #set if called by a keyboard handler - will give an extra 10 points * cur level.
) 
    PlaceCurTetromino -Remove
    $willfit = WillCurTetrominoFit -Row ($global:CurTetromino.Row + 1) -Column $global:CurTetromino.Column 
    if (-not $willfit) {        
        PlaceCurTetromino                           # redraw old shape on board
        # if we can't move, and we're at the top - game over man!
        if ($global:CurTetromino.Row + $CurTetromino.StartRow -eq 0) {
            GameOver
        }    
        # if we're creating a new tetromino after the last piece was placed by a keydown, then give some points
        if ($KeyboardUsed.IsPresent) {
            UpdateScore -AddPoints (10 * $global:GameData_Level)
        }
        # remove full lines
        while ((CheckAndProcessFullLines -ne 0)) {         # keep checking for complete rows unti all clear
            # blank
            Start-Sleep -Milliseconds 100
        }
        # create new tetromino
        NewTetromino                                
    }
    else {
        $global:CurTetromino.Row++
    }
    # update screen
    PlaceCurTetromino
    PrintGameBoard    
}


function MoveTetrominoLeft {
<# slide to the left #>
    PlaceCurTetromino -Remove
    $willfit = WillCurTetrominoFit -Row $global:CurTetromino.Row -Column ($global:CurTetromino.Column  - 1)
    if ($willfit) {
        $global:CurTetromino.Column--
    }
    PlaceCurTetromino
    PrintGameBoard    
}


function MoveTetrominoRight {
<# slide to the right #>
    PlaceCurTetromino -Remove
    $willfit = WillCurTetrominoFit -Row $global:CurTetromino.Row -Column ($global:CurTetromino.Column  + 1)
    if ($willfit) {
        $global:CurTetromino.Column++
    }
    PlaceCurTetromino
    PrintGameBoard    
}


function RotateTetrominoClockwise {
<# xris xross - rotates the shape of the current tetromino clockwise (if it fits) #>
    PlaceCurTetromino -Remove
    $old_shape = $CurTetromino.Shape
    $new_shape = @((New-Object int[](4)), (New-Object int[](4)), (New-Object int[](4)), (New-Object int[](4)))
    # rotate your owl!
    for ($r = 0; $r -lt $TETROMINO_SIZE; $r++) {
        for ($c = 0; $c -lt $TETROMINO_SIZE; $c++) {
            $new_shape[$c][($TETROMINO_SIZE - 1 - $r)] = $old_shape[$r][$c]            
        }
    }
    $CurTetromino.Shape = $new_shape
    # if we don't fit, revert back to old shape
    $willfit = WillCurTetrominoFit -Row $global:CurTetromino.Row -Column $global:CurTetromino.Column
    if (-not $willfit) {
        $global:CurTetromino.Shape = $old_Shape
    }
    PlaceCurTetromino
    PrintGameBoard    
}




#######################################################################################################
# Fancy Background
#######################################################################################################


# this stores the game background screen
$FancyBackground = $global:ConBuff.Clone()

$fancy_background_chars = $CHAR_EMPTY, $CHAR_FADE3, $CHAR_FADE2, $CHAR_FADE1, $CHAR_SOLIDBLOCK
for ($x = 0; $x -lt $CONBUFF_WIDTH; $x++) {
    for ($y = 0; $y -lt $CONBUFF_HEIGHT; $y++) {        
        $_randpos = Get-Random -Minimum 0 -Maximum $fancy_background_chars.Length
        $FancyBackground[$x, $y, $CONBUFF_IDX_VALUE] = $fancy_background_chars[$_randpos]
        $FancyBackground[$x, $y, $CONBUFF_IDX_FORECOL] = $FORECOLOUR_BACKGROUND
        $FancyBackground[$x, $y, $CONBUFF_IDX_BACKCOL] = $BACKCOLOUR_BACKGROUND
    }
}
    



#######################################################################################################
# Menu Screen
#######################################################################################################


# http://patorjk.com/software/taag/
$title = @(
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


function DrawTitle {
param(
    [int]$titley = 2
)
<# draws the title, with shadow #>    
    $titlex = 1
    # draw title shadow
    for ($y = 0; $y -lt $title.Count; $y++) {
        $line = $title[$y]
        for ($x = 0; $x -lt $line.Length; $x++) {
            if ((($x + 1) -lt $line.Length) -and (-not ([string]::IsNullOrWhiteSpace($line[($x + 1)])))) {
                $char = $CHAR_FADE2
                ConBuffWrite -Text $char -x ($titlex + $x) -y  ($titley + $y + 1) -ForeColour $FORECOLOUR_TITLESHADOW -BackColour $BACKCOLOUR_TITLESHADOW
            }
        }
    }
    # draw title white
    for ($y = 0; $y -lt $title.Count; $y++) {
        $line = $title[$y]
        for ($x = 0; $x -lt $line.Length; $x++) {
            $char = $CHAR_SOLIDBLOCK
            if (-not ([string]::IsNullOrWhiteSpace($line[$x]))) {
                ConBuffWrite -Text $char -x ($titlex + $x) -y  ($titley + $y) -ForeColour $FORECOLOUR_TITLE -BackColour $BACKCOLOUR_TITLE
            }
        }
    }
}


# write the title to the console
function DoMainMenu() {

    $menux = 10
    $menuy = 18
    $cury = $menuy
    $maxmenuy = $menuy + 4

    $mode = 'Main'       # current menu mode/option.

    $lastpaintmode = ""
    [bool]$allowtopscoreonscreen = $true
   
    function Paint {
    param(
        [switch]$NoPrint        # only draw on conbuff, don't print at the end.
    )
        # background and title
        ClearConBuff
        ConBuffCopyFrom -Source $FancyBackground
       
        # title
        DrawTitle
        
        # menu
        ConBuffDrawBorderexFillBox -Char ' ' -x 8 -y 16 -Width 20 -Height 9 -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU
        if ($Mode -eq 'Main') {
            # draw topscore
            if ($allowtopscoreonscreen -and ($global:GameData_TopScore -gt 0)) {
                [string]$scoreStr = $global:GameData_TopScore.ToString()
                $scoreStr = "TopScore: " + $scoreStr
                ConBuffWrite -Text $scoreStr -x (($CONBUFF_WIDTH / 2) - ($scoreStr.Length / 2)) -y 14 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
            }
            # draw menu items
            ConBuffWrite -Text "   New A Game" -x $menux -y $menuy -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
            ConBuffWrite -Text "    Controls" -x $menux -y ($menuy + 2) -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
            ConBuffWrite -Text "     Quit!" -x $menux -y ($menuy + 4) -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
            # draw current menu item
            ConBuffWrite -Text $CHAR_ARROWRIGHT -x $menux -y $cury -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
            ConBuffWrite -Text $CHAR_ARROWLEFT -x ($menux + 15) -y $cury -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
        }
        elseif ($Mode -eq 'Controls') {
            ConBuffWrite -Text (" {0}    Rotate" -f $CHAR_ARROWUP) -x ($menux + 1) -y ($menuy) -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
            ConBuffWrite -Text (" {0}    Down" -f $CHAR_ARROWDOWN) -x ($menux + 1) -y ($menuy + 1) -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
            ConBuffWrite -Text (" {0}    Right" -f $CHAR_ARROWRIGHT) -x ($menux + 1) -y ($menuy + 2) -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
            ConBuffWrite -Text (" {0}    Left" -f $CHAR_ARROWLEFT) -x ($menux + 1) -y ($menuy + 3) -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
            ConBuffWrite -Text "ESC   Quit" -x ($menux + 1) -y ($menuy + 4) -ForeColour $FORECOLOUR_MENU -BackColour $BACKCOLOUR_MENU 
        }

        if (-not $NoPrint.IsPresent) {
            if ($lastpaintmode -ne $Mode) { 
                PrintConBuff -Force
                $lastpaintmode = $Mode
            }
            else {
                PrintConBuff
            }
        }
    }

    Paint
    $lasttopscorems = $stopwatch.ElapsedMilliseconds
    while ($true) {
        $topscorems = $stopwatch.ElapsedMilliseconds
        if ($topscorems - $lasttopscorems -gt 1000) {
            $lasttopscorems = $topscorems
            $allowtopscoreonscreen = !$allowtopscoreonscreen
            Paint -NoPrint
            PrintConBuff -Force
        }
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($false)
            if ($mode -eq 'Main') {
                if ($key.key -eq [ConsoleKey]::DownArrow) {
                    $cury = $cury + 2;
                    if ($cury -gt $maxmenuy) {
                        $cury = $menuy;
                    }
                    $updatereq = $true;
                }
                elseif ($key.Key -eq [ConsoleKey]::UpArrow) {
                    $cury = $cury - 2;
                    if ($cury -lt $menuy) {
                        $cury = $maxmenuy
                    }
                    $updatereq = $true
                }
                elseif ($key.Key -eq [ConsoleKey]::Enter) {
                    switch ($cury) {
                        $menuy { return "AGame" }
                        ($menuy + 2) { 
                            $mode = 'Controls' 
                            $updatereq = $true
                            break
                        }  
                        ($menuy + 4) { return "Quit" }  
                    }
                }
            }
            elseif ($mode -eq 'Controls') {
                if ($Key.Key -eq 'Escape') {
                    $mode = 'Main'
                    $updatereq = $true                
                }
            }
        }
        # update screen
        if ($updatereq -eq $true) {
            Paint
            $updatereq = $false      
            ClearKeyboardBuffer            
        }
        # sleep
        Start-Sleep -Milliseconds 5
    }
    return "Quit"
}




#######################################################################################################
# MAIN GAME
#######################################################################################################


# some vars to help line things
$gacolx1 = 2              # game area column a
$gacolx2 = 14             # game area column b
$garow = 5                # start row of game area border
$gacolx3 = $gacolx2 + $GAMEAREA_WIDTH + 2 # game area column c


function DrawGameBackground() {

    ConBuffCopyFrom -Source $FancyBackground
 
     # draw a-type game area
    ConBuffDrawBorderexFillBox -Char $CHAR_EMPTY -x ($gacolx1 + 1) -y 3 -Width 8 -Height 3 -ForeColour $FORECOLOUR_BORDER -BackColour $BACKCOLOUR_BORDER
    # draw statistics
    ConBuffDrawBorderexFillBox -Char $CHAR_EMPTY -x $gacolx1 -y 7 -Width 12 -Height 18 -ForeColour $FORECOLOUR_BORDER -BackColour $BACKCOLOUR_BORDER

    # draw lines game area
    ConBuffDrawBorderexFillBox -Char $CHAR_EMPTY -x $gacolx2 -y 2 -Width 12 -Height 3 -ForeColour $FORECOLOUR_BORDER -BackColour $BACKCOLOUR_BORDER
    # draw main game area
    ConBuffDrawBorderexFillBox -Char $CHAR_EMPTY -x $gacolx2 -y $garow -Width ($GAMEAREA_WIDTH + 2) -Height ($GAMEAREA_HEIGHT + 2) -ForeColour $FORECOLOUR_BORDER -BackColour $BACKCOLOUR_BORDER

    # draw scores area
    ConBuffDrawBorderexFillBox -Char $CHAR_EMPTY -x $gacolx3 -y 2 -Width 8 -Height 9 -ForeColour $FORECOLOUR_BORDER -BackColour $BACKCOLOUR_BORDER
    # draw next item
    # draw level
    ConBuffDrawBorderexFillBox -Char $CHAR_EMPTY -x $gacolx3 -y 18 -Width 7 -Height 4 -ForeColour $FORECOLOUR_BORDER -BackColour $BACKCOLOUR_BORDER

}


function DrawScores() {
    # display type
    ConBuffWrite -Text "A-Type" -x ($gacolx1 + 2) -y 4 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    
    # display statistics
    ConBuffWrite -Text "STATISTICS" -x 3 -y 8 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    $_y = 9
    for ($i = 0; $i -le $TETROMINOTYPE_MAXINTVALUE; $i++) {
        $statsval = $global:GameData_ObjectStats[$i]
        $statsvalstr = ConvertScoreToString -Value $statsval -NumDigits 3
        $y = $_y + ($i * 2)
        ConBuffDrawTetromino -Type ([TetrominoType]$i) -x ($gacolx1 + 2) -y $y
        ConBuffWrite -Text $statsvalstr -x ($gacolx1 + 7) -y ($y + 1) -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    }
    
    # display lines
    [string]$LinesStr = ConvertScoreToString -Value $global:GameData_Lines -NumDigits 3
    ConBuffWrite -Text " Lines-$LinesStr" -x ($gacolx2 + 1) -y 3 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    
    # display score
    ConBuffWrite -Text "TOP" -x ($gacolx3 + 1) -y 4 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    [string]$TopScoreStr = ConvertScoreToString -Value $global:GameData_TopScore -NumDigits 6
    ConBuffWrite -Text $TopScoreStr -x ($gacolx3 + 1) -y 5 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    ConBuffWrite -Text "SCORE" -x ($gacolx3 + 1) -y 7 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    [string]$ScoreStr = ConvertScoreToString -Value $global:GameData_Score -NumDigits 6
    ConBuffWrite -Text $ScoreStr -x ($gacolx3 + 1) -y 8 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE

    # display next
    ConBuffDrawBorderexFillBox -Char $CHAR_EMPTY -x $gacolx3 -y 11 -Width 6 -Height 7 -ForeColour $FORECOLOUR_BORDER -BackColour $BACKCOLOUR_BORDER
    ConBuffWrite -Text "NEXT" -x ($gacolx3 + 1) -y 12 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    ConBuffDrawTetromino -Type $global:NextTetrominoType -x ($gacolx3 + 1)  -y 13

    # display level
    ConBuffWrite -Text "LEVEL" -x ($gacolx3 + 1) -y 19 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
    [string]$LevelStr = ConvertScoreToString -Value $global:GameData_Level -NumDigits 3
    ConBuffWrite -Text " $LevelStr " -x ($gacolx3 + 1) -y 20 -ForeColour $FORECOLOUR_SCORE -BackColour $BACKCOLOUR_SCORE
}


function DrawGameBoard() {
<# draws the game board in conbuff #>
    for ($boardy = 0; $boardy -lt $GAMEAREA_HEIGHT; $boardy++) {
        for ($boardx = 0; $boardx -lt $GAMEAREA_WIDTH; $boardx++) {
            $val = $global:GameBoard[$boardx, $boardy]
            $char = $TetrominoBoardChars[$val]
            $forecolour = GetBoardObjectForeColour -Object $val
            ConBuffWrite -Text $char -x ($gacolx2 + 1 + $boardx) -y ($garow + 1 + $boardy) -ForeColour $forecolour -BackColour $BACKCOLOUR_GAMEAREA
        }
    }
}


function PrintGameBoard {
<# prints the game board on screen - quicker than a complete screen write#>

    # save cursor position for later. we'll reset it back to this on our way out.
    $cursorX = [Console]::CursorLeft
    $cursorY = [Console]::CursorTop

    for ($boardy = 0; $boardy -lt $GAMEAREA_HEIGHT; $boardy++) {
        for ($boardx = 0; $boardx -lt $GAMEAREA_WIDTH; $boardx++) {
            # update conbuff with updated char
            $val = $global:GameBoard[$boardx, $boardy]
            $char = $TetrominoBoardChars[$val]
            $forecolour = GetBoardObjectForeColour -Object $val
            $cx = $gacolx2 + 1 + $boardx
            $cy = $garow + 1 + $boardy
            ConBuffWrite -Text $char -x $cx -y $cy -ForeColour $forecolour -BackColour $BACKCOLOUR_GAMEAREA

            # last printed val is diff to this, then print it
            $old_forecol = $global:LastPrintedConBuff[$cx, $cy, $CONBUFF_IDX_FORECOL]
            $old_backcol = $global:LastPrintedConBuff[$cx, $cy, $CONBUFF_IDX_BACKCOL]
            $old_value = $global:LastPrintedConBuff[$cx, $cy, $CONBUFF_IDX_VALUE]
            $old_value = if ([string]::IsNullOrEmpty($old_value)) { " " } else { $old_value.ToString() }
            $_updatereq = ($old_forecol -ne $forecolour) -or ($old_backcol -ne ([ConsoleColor]::Black)) -or ($old_value -ne $char)
            if ($_updatereq -eq $true) {
                ConsoleWriteAt -x $cx -y $cy -Text $char -ForeColour $forecolour -BackColour $BACKCOLOUR_GAMEAREA
            }

            # directly update lastconbuff
            $global:LastPrintedConBuff[$cx, $cy, $CONBUFF_IDX_VALUE] = $global:ConBuff[$cx, $cy, $CONBUFF_IDX_VALUE]
            $global:LastPrintedConBuff[$cx, $cy, $CONBUFF_IDX_FORECOL] = $global:ConBuff[$cx, $cy, $CONBUFF_IDX_FORECOL]
            $global:LastPrintedConBuff[$cx, $cy, $CONBUFF_IDX_BACKCOL] = $global:ConBuff[$cx, $cy, $CONBUFF_IDX_BACKCOL]        
        }
    }  
    
    # reset cursor, if it's been moved
    if (($cursorX -ne ([Console]::CursorLeft)) -or (($cursorY -ne ([Console]::CursorTop)))) {
        SetCursorPos -x $cursorX -y $cursorY
    }    
}


function HandleGameKeyChar {
<# handles game keychars #>
param(
    [parameter(Mandatory=$true)][ConsoleKeyInfo]$Key
)
    if ($Key.Key -eq [ConsoleKey]::DownArrow) {     
        MoveTetrominoDown -KeyboardUsed      
    }
    # up arrow
    elseif ($Key.Key -eq [ConsoleKey]::LeftArrow) {  
        MoveTetrominoLeft          
    }
    # right arrow
    elseif ($Key.Key -eq [ConsoleKey]::UpArrow) {     
        RotateTetrominoClockwise               
    }
    # left arrow
    elseif ($Key.Key -eq [ConsoleKey]::RightArrow) {
        MoveTetrominoRight
    }
    elseif ($Key.Key -eq 'Escape') {
        GameOver
    }
    ClearKeyboardBuffer                         
}


function DoAGame {

    $GAMELOOP_CONUPDATEMS = 1000;

    # new-game
    ResetGame
    ClearGameBoard
    
    DrawGameBackground
    DrawScores
    DrawGameBoard

    NewTetromino
    PlaceCurTetromino
    
    PrintConBuff -Force

    $_last_conupdatems = 0
    $_last_tetrinoupdatems = $global:stopwatch.ElapsedMilliseconds    

    $_last_score = 0    

    while ($true) {

        [bool]$doscreenupdate = $false
    
        # capture key press
        if ([Console]::KeyAvailable) {           # always waits for next possible key, and then pauses for key wait
            $key = [Console]::ReadKey($false)
            HandleGameKeyChar -Key $Key    
        }

        # update tetrino fall
        $tetrinoms = $global:stopwatch.ElapsedMilliseconds;
        if ($tetrinoms - $_last_tetrinoupdatems -ge (GetDelaymsForCurrentLevel)) {
            $_last_tetrinoupdatems = $tetrinoms
            MoveTetrominoDown
        }

        # score updates
        if ($_last_score -ne $global:GameData_Score) {
            $_last_score = $global:GameData_Score
            DrawScores            
            $doscreenupdate = $true
        }

        # console screen updates
        $conupdatems = $global:stopwatch.ElapsedMilliseconds;
        if ($conupdatems - $_last_conupdatems -ge $GAMELOOP_CONUPDATEMS) {
            $_last_conupdatems = $conupdatems
            DrawGameBackground
            DrawScores
            DrawGameBoard
            $doscreenupdate = $true
        }

        if ($doscreenupdate -eq $true) {
            PrintConBuff
        }

        # sleep per loop (?)
        Start-Sleep -Milliseconds 1;
    }    
}


function GameOver {
    $global:GracefulGameFinish = $true  
    # print game over and wait for a second
    SetCursorPos -x 0 -y 0
    ConBuffDrawBorderexFillBox -Char $CHAR_EMPTY -x ($gacolx2 - 1) -y ($garow + ($GAMEAREA_HEIGHT / 2) - 1)  -Height 3 -Width 14 -ForeColour $FORECOLOUR_BORDER -BackColour $BACKCOLOUR_BORDER 
    ConBuffWrite -Text "Game Over!" -x ($gacolx2 + 1) -y ($garow + ($GAMEAREA_HEIGHT / 2)) -ForeColour $FORECOLOUR_GAMETEXT -BackColour $BACKCOLOUR_GAMETEXT
    PrintConBuff
    Start-Sleep -Seconds 1;
    $c = 3
    while ($c--) { 
        Start-Sleep -Milliseconds 10
        ClearKeyboardBuffer 
    }
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




#######################################################################################################
# START OF SCRIPT
#######################################################################################################


$global:stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Clear-Host

$SavedForeColour = [Console]::ForegroundColor
$SavedBackColour = [Console]::BackgroundColor
$SavedCursorVisible = GetCursorVisible

$ErrorActionPreference = 'Stop'
try {    
    # init
    
    SetCursorVisible -Visible $false
    ClearConBuff
    UpdateLastConBuff
    PrintConBuff

    while ($true) {
        # game main menu
        $menuoption = DoMainMenu
        if ($menuoption -eq 'Quit') {
            $global:GracefulGameFinish = $true
            break
        }       
        # main game
        if ($menuoption -eq 'AGame') {
            try {
                DoAGame
            }
            catch {
                if ($_.ToString() -eq 'Game Over') {     
                    $global:GracefulGameFinish = $true   
                }            
            }
        }
    }
          
}
catch {
    $caught_thing = $_
    throw $caught_thing
}
finally {    

    # clean-up the console if we finished cleanly
    if ($global:GracefulGameFinish -eq $true) {
        # clear screen
        ClearConBuff    
        PrintConBuff -Force -NoCursorReset
        [Console]::WriteLine()        
    }

    # restore screen and cursor settings    
    [Console]::BackgroundColor = $SavedBackColour
    [Console]::ForegroundColor = $SavedForeColour
    SetCursorVisible -Visible $SavedCursorVisible

    # write top score to file
    if ($global:GracefulGameFinish -eq $true) {
        $global:GameData_TopScore | Out-File -FilePath $TOPSCORE_FILENAME -Force -ErrorAction SilentlyContinue
    }

}

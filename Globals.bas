Attribute VB_Name = "Globals"
Option Explicit
Option Base 0

'Windows API functions
Public Declare Function GetTickCount Lib "kernel32" () As Long

'colors
Public Const C_BLACK As Long = &H0

'customize the program here
Public Const FULLSCREEN As Boolean = False
Public Const SCREENWIDTH As Long = 240
Public Const SCREENHEIGHT As Long = 160
Public Const STEP As Integer = 1


'tile and game world constants
Public Const TILEWIDTH As Long = 16
Public Const TILEHEIGHT As Long = 16
Public Const MAPWIDTH As Long = 256
Public Const MAPHEIGHT As Long = 192
Public Const GAMEWORLDWIDTH As Long = TILEWIDTH * MAPWIDTH
Public Const GAMEWORLDHEIGHT As Long = TILEHEIGHT * MAPHEIGHT

'scrolling window size
Public Const WINDOWWIDTH As Integer = (SCREENWIDTH \ TILEWIDTH) * TILEWIDTH
Public Const WINDOWHEIGHT As Integer = (SCREENHEIGHT \ TILEHEIGHT) * TILEHEIGHT

'scroll buffer size
Public Const SCROLLBUFFERWIDTH As Integer = SCREENWIDTH + TILEWIDTH
Public Const SCROLLBUFFERHEIGHT As Integer = SCREENHEIGHT + TILEHEIGHT

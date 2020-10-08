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

'game world size
Public Const GAMEWORLDWIDTH As Long = 4096
Public Const GAMEWORLDHEIGHT As Long = 3072

'tile size
Public Const TILEWIDTH As Integer = 16
Public Const TILEHEIGHT As Integer = 16

'scrolling  window size
Public Const WINDOWWIDTH  As Integer = (SCREENWIDTH \ TILEWIDTH) * TILEWIDTH
Public Const WINDOWHEIGHT As Integer = (SCREENHEIGHT \ TILEHEIGHT) * TILEHEIGHT

Public Const SCROLLBUFFERWIDTH As Integer = SCREENWIDTH + TILEWIDTH
Public Const SCROLLBUFFERHEIGHT As Integer = SCREENHEIGHT + TILEHEIGHT

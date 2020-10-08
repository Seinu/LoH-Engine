VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Shutdown
End Sub
Private Sub Form_Load()
    'set up the main form
    Form1.Caption = "Game"
    Form1.ScaleMode = 3
    
    Form1.width = Screen.TwipsPerPixelX * (SCREENWIDTH + 12)
    Form1.height = Screen.TwipsPerPixelY * (SCREENHEIGHT + 30)
    Form1.Show
    
    'initialize Direct3D
    InitDirect3D Me.hWnd, SCREENWIDTH, SCREENHEIGHT, FULLSCREEN
    
    'get reference to the back buffer
    Set backbuffer = d3ddev.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
    
    'load bitmap file
    'Set tiles = LoadSurface(App.Path & "\assets\graphics\blank.BMP", 320, 96)
    Set tiles = LoadSurface(App.Path & "\assets\graphics\map-tilesets\map1.bmp", 320, 80)
    
    'load the map data from the Mappy export file
    'LoadMap App.Path & "\assets\graphics\blank.CSV"
    LoadBinaryMap App.Path & "\assets\binary-data\binmaps\map1.mar", 33, 33
    
    'create the small scroll buffer surface
     Set scrollbuffer = d3ddev.CreateImageSurface( _
        SCROLLBUFFERWIDTH, _
        SCROLLBUFFERHEIGHT, _
        dispmode.Format)
    
    'this helps to keep a steady framerate
    Dim start As Long
    start = GetTickCount()
    
    'clear the screen to black
    d3ddev.Clear 0, ByVal 0, D3DCLEAR_TARGET, C_BLACK, 1, 0
    
    'main loop
    Do While (True)
        'update the scroll position
        UpdateScrollPosition
        
        'draw tiles onto the scroll buffer
        DrawTiles
        
        'draw the scroll window onto the screen
        DrawScrollWindow
        
        'set the screen refresh to about 40 fps
        If GetTickCount - start > 25 Then
            d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0
            start = GetTickCount
            DoEvents
        End If
    Loop
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
        
        'move mouse on left side to scroll left
        If X < SCREENWIDTH / 2 Then SpeedX = -STEP
        
        'move mouse on right side to scroll right
        If X > SCREENWIDTH / 2 Then SpeedX = STEP
        
        'move mouse on top half to scroll up
        If Y < SCREENHEIGHT / 2 Then SpeedY = -STEP
        
        'move mouse on bottom half to scroll down
        If Y > SCREENHEIGHT / 2 Then SpeedY = STEP
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, Unload As Integer)
    Shutdown
End Sub

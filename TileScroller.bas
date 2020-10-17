Attribute VB_Name = "TileScroller"
'---------------------------------------------------------------
' Visual Basic Game Programming for Teens
' Tile Scrolling Support File
'---------------------------------------------------------------

Option Explicit
Option Base 0


'tile scroller surfaces
Public scrollbuffer As Direct3DSurface8
Public tiles As Direct3DSurface8

'map data
Public mapdata() As Integer

'scrolling values
Public ScrollX As Long
Public ScrollY As Long
Public SpeedX As Long
Public SpeedY As Long


Public Sub UpdateScrollPosition()
    'update horizontal scrolling position and speed
    ScrollX = ScrollX + SpeedX
    If (ScrollX < 0) Then
        ScrollX = 0
        SpeedX = 0
    ElseIf ScrollX > GAMEWORLDWIDTH - WINDOWWIDTH Then
        ScrollX = GAMEWORLDWIDTH - WINDOWWIDTH
        SpeedX = 0
    End If

    'update vertical scrolling position and speed
    ScrollY = ScrollY + SpeedY
    If ScrollY < 0 Then
        ScrollY = 0
        SpeedY = 0
    ElseIf ScrollY > GAMEWORLDHEIGHT - WINDOWHEIGHT Then
        ScrollY = GAMEWORLDHEIGHT - WINDOWHEIGHT
        SpeedY = 0
    End If
End Sub


Public Sub DrawTiles()
    Dim tilex As Integer
    Dim tiley As Integer
    Dim x As Integer
    Dim y As Integer
    Dim tilenum As Integer
    
    'calculate starting tile position
    'integer division drops the remainder
    tilex = ScrollX \ TILEWIDTH
    tiley = ScrollY \ TILEHEIGHT
    
    'Debug.Print "Scroll: " & ScrollX & "," & ScrollY & "; Tile: " & tilex & "," & tiley
    
    Dim columns As Long
    Dim rows As Long
    columns = WINDOWWIDTH \ TILEWIDTH
    rows = WINDOWHEIGHT \ TILEHEIGHT
    
    'draw tiles onto the scroll buffer surface
    For y = 0 To rows
        For x = 0 To columns
            
            '*** This condition shouldn't be necessary. I will try to
            '*** resolve this problem and make the change during AR.
            If tiley + y = MAPHEIGHT Then tiley = tiley - 1
            
            tilenum = mapdata((tiley + y) * MAPWIDTH + (tilex + x))
            DrawTile tiles, tilenum, TILEWIDTH, TILEHEIGHT, 20, scrollbuffer, _
                x * TILEWIDTH, y * TILEHEIGHT
        Next x
    Next y
End Sub

Public Sub DrawScrollWindow()
    Dim r As DxVBLibA.RECT
    Dim point As DxVBLibA.point
    Dim partialx As Integer
    Dim partialy As Integer

    'calculate the partial sub-tile lines to draw
    partialx = ScrollX Mod TILEWIDTH
    partialy = ScrollY Mod TILEHEIGHT
    
    'set dimensions of the source image
    r.Left = partialx
    r.Top = partialy
    r.Right = partialx + WINDOWWIDTH
    r.bottom = partialy + WINDOWHEIGHT
        
    'set the destination point
    point.x = 0
    point.y = 0
    
    'draw the scroll window
    d3ddev.CopyRects scrollbuffer, r, 1, backbuffer, point
End Sub

Public Sub DrawTile( _
    ByRef source As Direct3DSurface8, _
    ByVal tilenum As Long, _
    ByVal width As Long, _
    ByVal height As Long, _
    ByVal columns As Long, _
    ByVal dest As Direct3DSurface8, _
    ByVal destx As Long, _
    ByVal desty As Long)
    
    'create a RECT to describe the source image
    Dim r As DxVBLibA.RECT
    
    'set the upper left corner of the source image
    r.Left = (tilenum Mod columns) * width
    r.Top = (tilenum \ columns) * height
    
    'set the bottom right corner of the source image
    r.Right = r.Left + width
    r.bottom = r.Top + height
    
    'create a POINT to define the destination
    Dim point1 As DxVBLibA.point
    
    'set the upper left corner of where to draw the image
    point1.x = destx
    point1.y = desty
    
    'draw the source bitmap tile image
    d3ddev.CopyRects source, r, 1, dest, point1

End Sub

Public Sub LoadBinaryMap(ByVal filename As String, ByVal lWidth As Long, ByVal lHeight As Long)
    Dim filenum As Integer
    Dim n As Long
    Dim i As Integer

    'open the map file
    filenum = FreeFile()
    Open filename For Binary As filenum

    'prepare the array for the map data
    ReDim mapdata(lWidth * lHeight)
    
    Dim time1 As Long
    time1 = GetTickCount()
    
    'read the map data
    For n = 0 To lWidth * lHeight - 1
        Get filenum, , i
        mapdata(n) = i / 32 - 1
    Next n
    
    Debug.Print "Map loaded in " & GetTickCount() - time1 & " ms"
    
    'close the file
    Close filenum


End Sub






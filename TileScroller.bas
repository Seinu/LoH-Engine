Attribute VB_Name = "TileScroller"
Option Explicit
Option Base 0


'tile scroller surfaces
Public scrollbuffer As Direct3DSurface8
Public tiles As Direct3DSurface8

'map data
Public mapdata() As Integer
Public mapwidth As Long
Public mapheight As Long

'scrolling values
Public ScrollX As Long
Public ScrollY As Long
Public SpeedX As Integer
Public SpeedY As Integer
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
    If (ScrollY < 0) Then
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
    Dim columns As Integer
    Dim rows As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim tilenum As Integer
    
    'calculate starting tile position
    'integer division drops the remainder
    tilex = ScrollX \ TILEWIDTH
    tiley = ScrollY \ TILEHEIGHT
    
    'calculate the number of columns and rows
    'integer divsion drops the remainder
    columns = WINDOWWIDTH \ TILEWIDTH
    rows = WINDOWHEIGHT \ TILEHEIGHT
    
    'draw tiles onto the scroll buffer surface
     For Y = 0 To rows
        For X = 0 To columns
        
        '*** this condition shouldn't be necessary. i will try to
        '*** resolve this problem and make the change during AR.
        If tiley + Y = mapheight Then tiley = tiley - 1
        
            tilenum = mapdata((tiley + Y) * mapwidth + (tilex + X))
            DrawTile tiles, tilenum, TILEWIDTH, TILEHEIGHT, 20, scrollbuffer, _
                X * TILEWIDTH, Y * TILEHEIGHT
                
        Next X
    Next Y
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
    point.X = 0
    point.Y = 0
    
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
    
    'create a POINT tp define the destination
    Dim point As DxVBLibA.point
    
    'set the upper left corner of where to draw the image
    point.X = destx
    point.Y = desty
    
    'draw the source bitmap tile image
    d3ddev.CopyRects source, r, 1, dest, point
        
End Sub
Public Sub LoadMap(ByVal filename As String)
    Dim num As Integer
    Dim line As String
    Dim buffer As String
    Dim s As String
    Dim value As String
    Dim index As Long
    Dim pos As Long
    Dim buflen As Long
    
    'open the map file
    num = FreeFile()
    Open filename For Input As num
    
    'read the width and height
    Input #num, mapwidth, mapheight
    
    'read the map data
    While Not EOF(num)
        Line Input #num, line
        buffer = buffer & line
    Wend
    
    'close the file
    Close num
    
    'prepare the array for the map data
    ReDim mapdata(mapwidth * mapheight)
    index = 0
    buflen = Len(buffer)
    
    'convert the text data to an array
    For pos = 1 To buflen
    
        'get next character
        s = Mid$(buffer, pos, 1)
        
        'tiles are seperated by commas
        If s = "," Then
            If Len(value) > 0 Then
                
                'store tile # in array
                mapdata(index) = CInt(value - 1)
                index = index + 1
            End If
            
            'get ready for next #
            value = ""
            s = ""
        Else
            value = value & s
        End If
    Next pos
    
    'save last item to array
    mapdata(index) = CInt(value - 1)
                
End Sub
Public Sub LoadBinaryMap( _
    ByVal filename As String, _
    ByVal lWidth As Long, _
    ByVal lHeight As Long)
    
    Dim filenum As Integer
    Dim n As Long
    Dim i As Integer
    
    'open the map file
    filenum = FreeFile()
    Open filename For Binary As filenum
    
    'prepare the array for the map data
    ReDim mapdata(lWidth * lHeight)
    
    'read the map data
    For n = 0 To lWidth * lHeight - 1
        Get filenum, , i
        mapdata(n) = i / 32 - 1
    Next n
    
    'close the file
    Close filenum
    
End Sub


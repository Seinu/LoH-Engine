Attribute VB_Name = "Direct3D"
Public dx As DirectX8
Public d3d As Direct3D8
Public d3dx As New D3DX8
Public dispmode As D3DDISPLAYMODE
Public d3dpp As D3DPRESENT_PARAMETERS
Public d3ddev As Direct3DDevice8
Public backbuffer As Direct3DSurface8
Public Sub InitDirect3D( _
    ByVal hWnd As Long, _
    ByVal lWidth As Long, _
    ByVal lHeight As Long, _
    ByVal bFullscreen As Boolean)
    
    'catch any errors here
    On Local Error GoTo fatal_error
    
    'create the directX Oject
    Set dx = New DirectX8
    
    'create the Direct3D object
    Set d3d = dx.Direct3DCreate()
    If d3d Is Nothing Then
        MsgBox "Error initializing Direct3D!"
        Shutdown
    End If
    
    'tell D3D to use the current color depth
    d3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, dispmode
    
    'set the display settings used to create the device
    Dim d3dpp As D3DPRESENT_PARAMETERS
    d3dpp.hDeviceWindow = hWnd
    d3dpp.BackBufferCount = 1
    d3dpp.BackBufferWidth = lWidth
    d3dpp.BackBufferHeight = lHeight
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.BackBufferFormat = dispmode.Format
    
    'set windowsed or fullscreen mode
    If bFullscreen Then
        d3dpp.Windowed = 0
    Else
        d3dpp.Windowed = 1
    End If
    
    d3dpp.MultiSampleType = D3DMULTISAMPLE_NONE
    d3dpp.AutoDepthStencilFormat = D3DFMT_D32
    
    'create the D3D primary device
    Set d3ddev = d3d.CreateDevice( _
        D3DADAPTER_DEFAULT, _
        D3DDEVTYPE_HAL, _
        hWnd, _
        D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
        d3dpp)
        
    If d3ddev Is Nothing Then
        MsgBox "Error creating the Direct3D device!"
        Shutdown
    End If
    Exit Sub
    
fatal_error:
    MsgBox "Critical error in Start_Direct3D!"
    Shutdown
End Sub
Public Function LoadSurface( _
    ByVal filename As String, _
    ByVal width As Long, _
    ByVal height As Long) _
    As Direct3DSurface8
    
    On Local Error GoTo fatal_error
    Dim Surf As Direct3DSurface8
    
    'return error by default
    Set LoadSurface = Nothing
    
    'create the new surface
    Set Surf = d3ddev.CreateImageSurface(width, height, dispmode.Format)
    If Surf Is Nothing Then
        MsgBox "Error creating Surface!"
        Exit Function
    End If
        
    'load surface from file
    d3dx.LoadSurfaceFromFile Surf, ByVal 0, ByVal 0, filename, _
    ByVal 0, D3DX_DEFAULT, 0, ByVal 0
    
    If Surf Is Nothing Then
        MsgBox "Error loading " & filename & "!"
        Exit Function
    End If
    
    'return the new surface
    Set LoadSurface = Surf
    
fatal_error:
    Exit Function
End Function
Public Sub Shutdown()
    Set gameworld = Nothing
    Set d3ddev = Nothing
    Set d3d = Nothing
    Set dx = Nothing
    End
End Sub


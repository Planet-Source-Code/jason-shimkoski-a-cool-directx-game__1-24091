Attribute VB_Name = "modDD"
'many of the things below are self explanitory so I will not go into detail. If you need
'detail, check out the source for Mr. Eat Gum Guy 1.

'the main direct draw object
Public dd As DirectDraw7

'this is for testing if the program still has focus
Public bRestore As Boolean

'these are the front and backbuffers as well as there descriptions
Public ddsFrontBuf As DirectDrawSurface7
Public ddsdFrontBuf As DDSURFACEDESC2
Public ddsBackBuf As DirectDrawSurface7
Public ddsdBackBuf As DDSURFACEDESC2

'this creates the front and back buffer
'we are using one backbuffer, if you want 2, simply put 2 instead of 1 as the lBackBufferCount value
Sub CreateFrontBackBuf()
On Error GoTo errCreateFrontBackBuf

    ddsdBackBuf.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsdBackBuf.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
    ddsdBackBuf.lBackBufferCount = 1
    Set ddsFrontBuf = dd.CreateSurface(ddsdBackBuf)

    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set ddsBackBuf = ddsFrontBuf.GetAttachedSurface(caps)
    Exit Sub

errCreateFrontBackBuf:
    ErrorFound "Create Front and Back Buffer"
End Sub

Sub InitGraphics()
    CreateGraphicsFromFile "bg", ddsBG, ddsdBG, 16, 16
End Sub

Sub CreateGraphicsFromFile(fName As String, dds As DirectDrawSurface7, ddsd As DDSURFACEDESC2, ddsdWidth As Integer, ddsdHeight As Integer)
On Error GoTo errCreateGraphicsFromFile

    fName = App.Path & "\images\" & fName & ".bmp"

    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd.lHeight = ddsdHeight
    ddsd.lWidth = ddsdWidth

    Set dds = dd.CreateSurfaceFromFile(fName, ddsd)

    'this is for transparency
    Dim ColorKey As DDCOLORKEY
    ColorKey.high = 0
    ColorKey.low = 0
    Call dds.SetColorKey(DDCKEY_SRCBLT, ColorKey)

    Exit Sub

errCreateGraphicsFromFile:
    ErrorFound "Create a Graphic From its File"
End Sub

'this is the "painting of the sprites" area
Sub BltFast(rTop As Integer, rLeft As Integer, Width As Integer, Height As Integer, dds As DirectDrawSurface7, srcRect As RECT, X As Integer, Y As Integer, Transparency As Boolean)
On Error GoTo errBltFast

    DoUntilReady

    srcRect.Top = rTop
    srcRect.Left = rLeft
    srcRect.Right = Width
    srcRect.Bottom = Height

    If Transparency = False Then
        Call ddsBackBuf.BltFast(X, Y, dds, srcRect, DDBLTFAST_WAIT)
    Else
        Call ddsBackBuf.BltFast(X, Y, dds, srcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    Exit Sub

errBltFast:
    ErrorFound "Blit The Graphics"
End Sub

'this takes the 16x16 tile I made and tiles it across the screen
Sub CreateBG(TileWidth As Integer, TileHeight As Integer, dds As DirectDrawSurface7, CurrentResolution As String)
On Error GoTo errRenderMap

Dim X As Integer, Y As Integer, r As RECT, retVal As Long
Dim NumTilesX As Integer, NumTilesY As Integer

    Select Case CurrentResolution
        Case "640x480"
            Select Case TileWidth
                Case 10
                    NumTilesX = 64
                    NumTilesY = 48
                Case 16
                    NumTilesX = 40
                    NumTilesY = 30
                Case Else
                    ErrorFound "Render The Map Due To Invalid Tile Size"
            End Select
        Case "800x600"
            Select Case TileWidth
                Case 10
                    NumTilesX = 80
                    NumTilesY = 60
                Case 16
                    NumTilesX = 50
                    NumTilesY = 38
                Case Else
                    ErrorFound "Render The Map Due To Invalid Tile Size"
            End Select
        Case "1024x768"
            Select Case TileWidth
                Case 16
                    NumTilesX = 64
                    NumTilesY = 48
                Case Else
                    ErrorFound "Render The Map Due To Invalid Tile Size"
            End Select
        Case Else
            ErrorFound "Render The Map Due To Invalid Screen Resolution"
    End Select

For X = 0 To NumTilesX
For Y = 0 To NumTilesY

r.Right = r.Left + TileWidth
r.Bottom = r.Top + TileHeight

retVal = ddsBackBuf.BltFast(Int(X * TileWidth), Int(Y * TileHeight), dds, r, DDBLTFAST_WAIT)
Next Y
Next X
Exit Sub

errRenderMap:
    ErrorFound "Render The Map"
End Sub

'this is to check to see if we still have the focus
Sub DoUntilReady()
Dim bRest As Boolean

    bRest = False
    Do Until InDxMode
        DoEvents
        bRest = True
    Loop

    DoEvents
    If bRest Then
        bRest = False
        dd.RestoreAllSurfaces
    End If
End Sub

Function InDxMode() As Boolean
Dim TestCoopLevel As Long

    TestCoopLevel = dd.TestCooperativeLevel
    If (TestCoopLevel = DD_OK) Then
        InDxMode = True
    Else
        InDxMode = False
    End If
End Function

'this sets the font options
Sub SetDDFontOptions(size As Integer, Color As ColorConstants, Optional name As String = "Arial")
On Error GoTo errDDFontOptions

Dim ddFont As New StdFont

        ddFont.name = name
        ddFont.size = size

        ddsBackBuf.SetFont ddFont
        ddsBackBuf.SetForeColor Color
        ddsBackBuf.SetFontTransparency True
        Exit Sub

errDDFontOptions:
    ErrorFound "Set Font Options!"
End Sub

Attribute VB_Name = "modDX"
'many of the things below are self explanitory, so I will no do comments for most lines

'this is the main directx object
Public dx As New DirectX7

Sub InitDX()
On Error GoTo errInitDX

    Set dd = dx.DirectDrawCreate("")
    Set ds = dx.DirectSoundCreate("")

    frmMain.Show
    Exit Sub

errInitDX:
    ErrorFound "Initiate DirectX"
End Sub

'this sets the cooperative level of direct draw and direct sound to the highest priority
Sub SetCoopLevel(fhWnd As Form)
On Error GoTo errSetCoopLevel

    Call dd.SetCooperativeLevel(fhWnd.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX)
    Call ds.SetCooperativeLevel(fhWnd.hWnd, DSSCL_PRIORITY)
    Exit Sub

errSetCoopLevel:
    ErrorFound "Set Cooperative Level"
End Sub

'This sets the display mode
Sub SetDisplay(w As Integer, h As Integer, bpp As Integer)
On Error GoTo errSetDisplay

    Call dd.SetDisplayMode(w, h, bpp, 0, DDSDM_DEFAULT)
    Exit Sub

errSetDisplay:
    ErrorFound "Set the Display Mode"
End Sub

'This is for ending the application. cleans up directx
Sub DestroyDX()
    Set ddsBG = Nothing
    Set ddsGreeny = Nothing
    Set ddsYellow = Nothing
    Set ddsLogo = Nothing
    Set ddsMainMenu = Nothing
    Set ddsGums = Nothing
    Set diKeyBoard = Nothing
    UnloadSounds

    Set dmSegment = Nothing
    Set dmPerformance = Nothing
    Set dmLoader = Nothing
    Set ds = Nothing
    Set di = Nothing
    Set dd = Nothing
    Set dx = Nothing
End Sub

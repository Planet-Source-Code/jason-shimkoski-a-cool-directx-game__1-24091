Attribute VB_Name = "modMain"
'Well I have finally completed Mr. Eat Gum Guy 2. This game has many upgraded features
'from the last one. It has a two player mode, has a new musical piece, made personally
'by me, and direct input is used. To top it all off, this version is probably 1/3 the
'size on disk as the last version. Despite a few coding problems I realize
'I have, I could have reused so much of the code, I feel this is a damn neato little
'app.
'
'Thanks for your Time and your download,
'Jason Shimkoski  master@mastercodes.com


'this is for adjusting the frames per second
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'This is used for Showing and Hiding the cursor
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'these is a variable to see what resolution we are at
Public ChosenRes As String
Public ScreenCurW As Integer, ScreenCurH As Integer, ScreenCurBPP As Integer
'this is for checking if gum is eaten as well as checking collision
Public gumEaten As Boolean, gumX As Integer, gumY As Integer
Public BgumEaten As Boolean, BgumX As Integer, BgumY As Integer

'this is for animation
Public AnimationTimer As Integer

'this is for movement of the guys
Public p1CurX As Integer, p1CurY As Integer, p1Speed As Integer
Public p2CurX As Integer, p2CurY As Integer, p2Speed As Integer

'this keeps tracks of scores and the rank of the player
Public p1Score As Integer, p2Score As Integer
Public p1Rank As String

'this is to see if the player is going up or down
Public p1Up As Boolean, p1Down As Boolean, p1Left As Boolean, p1Right As Boolean
Public p2Up As Boolean, p2Down As Boolean, p2Left As Boolean, p2Right As Boolean

'this is to see if the bad gum should be set down or not
Public BgumSet As Boolean
'duh
Public exitprogram As Boolean
'duh
Public gamerun As Boolean

'these are to..you guessed it..the see if the player checked off if the want music and sound or not
Public AllowMusic As Boolean
Public AllowSound As Boolean

'this stuff is pretty self explanitory, just read each line and you'll understand
Sub MainLoop()

    exitprogram = True

    InitDX
    InitDI
    LoadSounds
    CreateLoaderPerformance frmMain.hWnd
    SetCoopLevel frmMain
    SetDisplay ScreenCurW, ScreenCurH, ScreenCurBPP

    CreateFrontBackBuf
    modSurfaces.InitGraphics

    'hides the cursor
    ShowCursor False
    exitprogram = False
    Do
        DoEvents
        MainMenuScreen
    Loop Until exitprogram = True
    StopSounds
    EndGame
End Sub

Sub EndGame()
    DestroyDX

    'shows the cursor
    ShowCursor True
    Unload frmMain
    End
End Sub

'this draws the good gum in a random area of the screen
Sub DrawGGum()
    If gumEaten = True Then
        Randomize
        gumX = Int((ScreenCurW - 63) * Rnd)
        gumY = Int((ScreenCurH - 36) * Rnd)
        BltFast 0, 0, 63, 36, ddsGums, rectGums, gumX, gumY, True
    Else
        BltFast 0, 0, 63, 36, ddsGums, rectGums, gumX, gumY, True
    End If
    gumEaten = False
End Sub

'this checks to see if the good gum has been eaten
Sub CheckGGumEat(guyX As Integer, guyY As Integer, Speed As Integer, tScore As Integer)
        Select Case guyX
            Case gumX - 63 To gumX + 63
            Select Case guyY
                Case gumY - 63 To gumY + 36
                gumEaten = True
            End Select
        End Select

        If gumEaten = True Then
            Speed = Speed + 1
            tScore = tScore + 1
            If AllowSound = True Then dsPlay WooHooBuffer, False
            DrawGGum
        Else
            Speed = Speed
        End If
End Sub

'this draws the bad gum based on score
Sub DrawBGum(tScore)
    If BgumSet = True Then
        DrawBGumNoPlacement p1Score
        Exit Sub
    End If

    If tScore = 5 Or tScore = 10 Or tScore = 15 Or tScore = 20 Or tScore = 25 Or tScore = 30 Or tScore = 35 Or tScore = 40 Or tScore = 45 Or tScore = 50 Or tScore = 55 Then
        DrawBGumPlacement
    Else
        Exit Sub
    End If
End Sub

'this draws the bad gum with placement
Sub DrawBGumPlacement()
    setBGumPos
    BltFast 36, 0, 63, 36 * 2, ddsGums, rectGums, BgumX, BgumY, True
    BgumSet = True
    BgumEaten = False
End Sub

'this draws the bad gum without placement. this prevents the bad gum from flying around
'all over the place.
Sub DrawBGumNoPlacement(tScore As Integer)
    If tScore = 5 Or tScore = 10 Or tScore = 15 Or tScore = 20 Or tScore = 25 Or tScore = 30 Or tScore = 35 Or tScore = 40 Or tScore = 45 Or tScore = 50 Or tScore = 55 Then
            BltFast 36, 0, 63, 36 * 2, ddsGums, rectGums, BgumX, BgumY, True
        Else
            Exit Sub
    End If
End Sub

'this randomly generates a position for the bad gum
Sub setBGumPos()
    Randomize
    BgumX = Int((ScreenCurW - 63) * Rnd)
    BgumY = Int((ScreenCurH - 36) * Rnd)
End Sub

'checks to see if the bad gum was eaten, if it was it does certain tasks
Sub CheckBGumEat(guyX As Integer, guyY As Integer, Speed As Integer, tScore As Integer)
    If tScore = 5 Or tScore = 10 Or tScore = 15 Or tScore = 20 Or tScore = 25 Or tScore = 30 Or tScore = 35 Or tScore = 40 Or tScore = 45 Or tScore = 50 Or tScore = 55 Then
        Select Case guyX
            Case BgumX - 63 To BgumX + 63
                Select Case guyY
                    Case BgumY - 63 To BgumY + 36
                    BgumEaten = True
                End Select
        End Select

        If BgumEaten = True Then
            Speed = Speed + 10
            tScore = tScore - 5
            If AllowSound = True Then dsPlay CrapBuffer, False
            BgumSet = False
            BgumEaten = False
        Else
            Speed = Speed
        End If
    Else
        setBGumPos
        Exit Sub
    End If
End Sub

'this draws the guy, either Yellow or Greeny
Sub DrawGuy(guy As String)
Dim dds As DirectDrawSurface7, r As RECT
Dim guyX As Integer, guyY As Integer
Dim tUp As Boolean, tDown As Boolean, tLeft As Boolean, tRight As Boolean

    Set dds = Nothing

    If guy = "Yellow" Then
        Set dds = ddsYellow
        r = rectYellow
        guyX = p1CurX
        guyY = p1CurY
        tUp = p1Up
        tDown = p1Down
        tLeft = p1Left
        tRight = p1Right
    ElseIf guy = "Greeny" Then
        Set dds = ddsGreeny
        r = rectGreeny
        guyX = p2CurX
        guyY = p2CurY
        tUp = p2Up
        tDown = p2Down
        tLeft = p2Left
        tRight = pRight
    End If

    'this draws the guy based on which direction he is going
    If tUp = False And tDown = False And tLeft = False And tRight = False Then
        BltFast 0, 0, 64, 64, dds, r, guyX, guyY, True
    ElseIf tUp = True Then
        Select Case AnimationTimer
            Case 0 To 10
            BltFast 64, 0, 64, 64 * 2, dds, r, guyX, guyY, True
            Case 11 To 20
            BltFast 64, 64 * 2, 64 * 3, 64 * 2, dds, r, guyX, guyY, True
            End Select
    ElseIf tDown = True Then
        Select Case AnimationTimer
            Case 0 To 10
            BltFast 64, 64, 64 * 2, 64 * 2, dds, r, guyX, guyY, True
            Case 11 To 20
            BltFast 64, 64 * 3, 64 * 4, 64 * 2, dds, r, guyX, guyY, True
        End Select
    ElseIf tLeft = True Then
        Select Case AnimationTimer
            Case 0 To 10
            BltFast 0, 64, 64 * 2, 64, dds, r, guyX, guyY, True
            Case 11 To 20
            BltFast 0, 64 * 3, 64 * 4, 64, dds, r, guyX, guyY, True
        End Select
    ElseIf tRight = True Then
        Select Case AnimationTimer
            Case 0 To 10
            BltFast 0, 0, 64, 64, dds, r, guyX, guyY, True
            Case 11 To 20
            BltFast 0, 64 * 2, 64 * 3, 64, dds, r, guyX, guyY, True
        End Select
    End If
End Sub

'for error checking
Sub ErrorFound(ErrText As String)
    Debug.Print "Could Not " & ErrText & "! Sorry You Cannot Play!"
    EndGame
End Sub

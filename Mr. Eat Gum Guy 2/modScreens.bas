Attribute VB_Name = "modScreens"
'The following code is very simple to understand. therefore I am not going to write to
'many comments. If you just read through everything, you will understand everything you
'need to.

'this is for the main menu
Public mListItem As Integer

Sub MainMenuScreen()
On Error GoTo errMainMenuScreen

Dim mGoUp As Boolean, mGoDown As Boolean
Dim EnterPress As Boolean

    DoEvents

    diKeyBoard.Acquire
    Call dmPerformance.Stop(Nothing, Nothing, 0, 0)
    diKeyBoard.GetDeviceStateKeyboard KeyboardState
    If KeyboardState.Key(DIK_ESCAPE) <> 0 Then
        If mListItem = 2 Then
            exitprogram = True
        Else
            mListItem = 2
        End If
    End If

    diKeyBoard.Unacquire
    diKeyBoard.Acquire

    If KeyboardState.Key(DIK_RETURN) <> 0 Then EnterPress = True
    If KeyboardState.Key(DIK_UP) <> 0 Then mGoDown = False: mGoUp = True
    If KeyboardState.Key(DIK_DOWN) <> 0 Then mGoUp = False: mGoDown = True

    DoEvents

    CreateBG 16, 16, ddsBG, ChosenRes
    BltFast 0, 0, 400, 128, ddsLogo, rectLogo, (ScreenCurW \ 2) - 200, 10, True
    BltFast 0, (357 \ 2), 357, 177, ddsMainMenu, rectMainMenu, (ScreenCurW \ 2) - (177 \ 2), 175, True

    If mGoUp = True Then
        mListItem = mListItem - 1
        If mListItem < 0 Then mListItem = 2
        If mListItem > 2 Then mListItem = 0
    ElseIf mGoDown = True Then
        mListItem = mListItem + 1
        If mListItem > 2 Then mListItem = 0
        If mListItem < 0 Then mListItem = 2
    End If

    DoEvents

    Select Case mListItem
        Case 0
            BltFast 0, 0, (357 \ 2), (177 \ 3), ddsMainMenu, rectMainMenu, (ScreenCurW \ 2) - (178 \ 2), 175, True
            If EnterPress = True Then OnePGame
        Case 1
            BltFast (177 \ 3), 0, (357 \ 2), ((177 \ 3) + (177 \ 3)), ddsMainMenu, rectMainMenu, (ScreenCurW \ 2) - (178 \ 2), (175 + (177 \ 3)), True
            If EnterPress = True Then TwoPGame
        Case 2
            BltFast ((177 \ 3) * 2), 0, (357 \ 2), ((177 \ 3) + ((177 \ 3) * 2)), ddsMainMenu, rectMainMenu, (ScreenCurW \ 2) - (178 \ 2), (175 + ((177 \ 3) * 2)), True
            If EnterPress = True Then exitprogram = True
        Case Else
            ddsBackBuf.DrawText 10, 30, "The Main Menu Has a Problem!", False
    End Select

    EnterPress = False

    SetDDFontOptions 9, vbWhite
    ddsBackBuf.DrawText ScreenCurW - 250, ScreenCurH - 30, "Created & Developed by: Jason Shimkoski", False
    ddsFrontBuf.Flip Nothing, DDFLIP_WAIT

    Exit Sub
errMainMenuScreen:
    ErrorFound "Access The Main Menu"
End Sub

Sub OnePGame()
On Error GoTo errOnePGame

Dim TempTime As Long

StartOver:
    gamerun = True
    
    diKeyBoard.Unacquire

    p1Up = False
    p1Down = False
    p1Left = False
    p1Right = False

    p1Speed = 5
    p1CurX = 10
    p1CurY = 10
    p1Score = 0
    gumEaten = True
    BgumEaten = False
    BgumSet = False

    If AllowMusic = True Then Call LoadPlayMidi("scary.mid")

    Do
        TempTime = timeGetTime
        AnimationTimer = AnimationTimer + 1
        If AnimationTimer = 20 Then AnimationTimer = 0

        diKeyBoard.Acquire
        diKeyBoard.GetDeviceStateKeyboard KeyboardState
        If KeyboardState.Key(DIK_ESCAPE) <> 0 Then gamerun = False
        If KeyboardState.Key(DIK_UP) <> 0 Then p1Up = True: p1Down = False: p1Left = False: p1Right = False
        If KeyboardState.Key(DIK_DOWN) <> 0 Then p1Down = True: p1Up = False: p1Left = False: p1Right = False
        If KeyboardState.Key(DIK_LEFT) <> 0 Then p1Left = True: p1Up = False: p1Down = False: p1Right = False
        If KeyboardState.Key(DIK_RIGHT) <> 0 Then p1Right = True: p1Up = False: p1Down = False: p1Left = False

        DoEvents

        ddsBackBuf.SetFillColor vbBlack
        ddsBackBuf.SetForeColor vbBlack
        ddsBackBuf.DrawBox 0, 0, ScreenCurW, ScreenCurH
        CreateBG 16, 16, ddsBG, ChosenRes

        DrawGGum
        DrawBGum p1Score

        DrawGuy "Yellow"

        If p1Up = False And p1Down = False And p1Left = False And p1Right = False Then
            p1CurY = p1CurY
            p1CurX = p1CurX
        ElseIf p1Up = True Then
            p1CurY = p1CurY - p1Speed
            p1CurX = p1CurX
        ElseIf p1Down = True Then
            p1CurY = p1CurY + p1Speed
            p1CurX = p1CurX
        ElseIf p1Left = True Then
            p1CurX = p1CurX - p1Speed
            p1CurY = p1CurY
        ElseIf p1Right = True Then
            p1CurX = p1CurX + p1Speed
            p1CurY = p1CurY
        End If

        CheckGGumEat p1CurX, p1CurY, p1Speed, p1Score
        CheckBGumEat p1CurX, p1CurY, p1Speed, p1Score

        If p1CurX < 0 Then p1CurX = 0: gamerun = False
        If p1CurY < 0 Then p1CurY = 0: gamerun = False
        If p1CurX + 64 > ScreenCurW Then p1CurX = ScreenCurW - 64: gamerun = False
        If p1CurY + 64 > ScreenCurH Then p1CurY = ScreenCurH - 64: gamerun = False

        SetDDFontOptions 9, vbWhite
        ddsBackBuf.DrawText 10, 10, "Score: " & p1Score, False

        DoEvents

        ddsFrontBuf.Flip Nothing, DDFLIP_WAIT

        Do Until timeGetTime >= TempTime + 16.6666666666667
            'loops until frame rate met
        Loop
    Loop Until gamerun = False

'this is the rank screen, it seems like a whole lot but it really isn't
    'this draws the basic background and logo
    ddsBackBuf.SetFillColor vbBlack
    ddsBackBuf.SetForeColor vbBlack
    ddsBackBuf.DrawBox 0, 0, ScreenCurW, ScreenCurH
    CreateBG 16, 16, ddsBG, ChosenRes
    BltFast 0, 0, 400, 128, ddsLogo, rectLogo, (ScreenCurW \ 2) - 200, 10, True
    GetRank p1Score, p1Rank

    'this draws the shadow of the text
    SetDDFontOptions 16, vbBlack
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 155, "Your Score: " & p1Score, False
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 195, "Your Rank: " & p1Rank & "!", False
    SetDDFontOptions 12, vbBlack
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 255, "Press Enter To Start A New Game!", False
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 285, "Press Escape To Return to the Main Menu!", False

    'this draws the rank text
    SetDDFontOptions 16, vbWhite
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 150, "Your Score: " & p1Score, False
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 190, "Your Rank: " & p1Rank & "!", False
    SetDDFontOptions 12, vbWhite
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 250, "Press Enter To Start A New Game!", False
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 280, "Press Escape To Return to the Main Menu!", False
    ddsFrontBuf.Flip Nothing, DDFLIP_WAIT

RankScreen:
    'this loops until the certain button is pressed
    diKeyBoard.Acquire
    diKeyBoard.GetDeviceStateKeyboard KeyboardState
    Call dmPerformance.Stop(Nothing, Nothing, 0, 0)
    If KeyboardState.Key(DIK_ESCAPE) <> 0 Then
        diKeyBoard.Unacquire
        GoTo MainMenu
    ElseIf KeyboardState.Key(DIK_RETURN) <> 0 Then
        diKeyBoard.Unacquire
        GoTo StartOver
    Else
        GoTo RankScreen
    End If

MainMenu:
    MainMenuScreen
    mListItem = 0

    Exit Sub
errOnePGame:
    ErrorFound "Continue Play In 1 Player Mode"
End Sub

'this is basically the same as the one player mode only the screen collision has been
'changed, another player has been added, and the bad gum doesn't show up.
Sub TwoPGame()
On Error GoTo errTwoPGame

Dim TempTime As Long

StartOver:
    gamerun = True
    
    diKeyBoard.Unacquire

    p1Up = False
    p1Down = False
    p1Left = False
    p1Right = True
    p2Up = False
    p2Down = False
    p2Left = True
    p2Right = False

    p1Speed = 5
    p1CurX = 10
    p1CurY = 10
    p1Score = 0
    p2Speed = 5
    p2CurX = ScreenCurW - 74
    p2CurY = ScreenCurH - 74
    p2Score = 0

    gumEaten = True

    If AllowMusic = True Then Call LoadPlayMidi("scary.mid")

    Do
        TempTime = timeGetTime
        AnimationTimer = AnimationTimer + 1
        If AnimationTimer = 20 Then AnimationTimer = 0

        diKeyBoard.Acquire
        diKeyBoard.GetDeviceStateKeyboard KeyboardState
        If KeyboardState.Key(DIK_ESCAPE) <> 0 Then gamerun = False
        If KeyboardState.Key(DIK_W) <> 0 Then p1Up = True: p1Down = False: p1Left = False: p1Right = False
        If KeyboardState.Key(DIK_S) <> 0 Then p1Down = True: p1Up = False: p1Left = False: p1Right = False
        If KeyboardState.Key(DIK_A) <> 0 Then p1Left = True: p1Up = False: p1Down = False: p1Right = False
        If KeyboardState.Key(DIK_D) <> 0 Then p1Right = True: p1Up = False: p1Down = False: p1Left = False

        If KeyboardState.Key(DIK_UP) <> 0 Then p2Up = True: p2Down = False: p2Left = False: p2Right = False
        If KeyboardState.Key(DIK_DOWN) <> 0 Then p2Down = True: p2Up = False: p2Left = False: p2Right = False
        If KeyboardState.Key(DIK_LEFT) <> 0 Then p2Left = True: p2Up = False: p2Down = False: p2Right = False
        If KeyboardState.Key(DIK_RIGHT) <> 0 Then p2Right = True: p2Up = False: p2Down = False: p2Left = False

        DoEvents

        ddsBackBuf.SetFillColor vbBlack
        ddsBackBuf.SetForeColor vbBlack
        ddsBackBuf.DrawBox 0, 0, ScreenCurW, ScreenCurH
        CreateBG 16, 16, ddsBG, ChosenRes

        DrawGGum

        DrawGuy "Yellow"
        DrawGuy "Greeny"

        If p1Up = False And p1Down = False And p1Left = False And p1Right = False Then
            p1CurY = p1CurY
            p1CurX = p1CurX
        ElseIf p1Up = True Then
            p1CurY = p1CurY - p1Speed
            p1CurX = p1CurX
        ElseIf p1Down = True Then
            p1CurY = p1CurY + p1Speed
            p1CurX = p1CurX
        ElseIf p1Left = True Then
            p1CurX = p1CurX - p1Speed
            p1CurY = p1CurY
        ElseIf p1Right = True Then
            p1CurX = p1CurX + p1Speed
            p1CurY = p1CurY
        End If

        If p2Up = False And p2Down = False And p2Left = False And p2Right = False Then
            p2CurY = p2CurY
            p2CurX = p2CurX
        ElseIf p2Up = True Then
            p2CurY = p2CurY - p2Speed
            p2CurX = p2CurX
        ElseIf p2Down = True Then
            p2CurY = p2CurY + p2Speed
            p2CurX = p2CurX
        ElseIf p2Left = True Then
            p2CurX = p2CurX - p2Speed
            p2CurY = p2CurY
        ElseIf p2Right = True Then
            p2CurX = p2CurX + p2Speed
            p2CurY = p2CurY
        End If

        CheckGGumEat p1CurX, p1CurY, p1Speed, p1Score
        CheckGGumEat p2CurX, p2CurY, p2Speed, p2Score

        If p1CurX < 0 Then p1CurX = ScreenCurW - 64: p1Speed = p1Speed - 0.5: p2Speed = p2Speed + 0.25
        If p1CurY < 0 Then p1CurY = ScreenCurH - 64: p1Speed = p1Speed - 0.5: p2Speed = p2Speed + 0.25
        If p1CurX + 64 > ScreenCurW Then p1CurX = 0: p1Speed = p1Speed - 0.5: p2Speed = p2Speed + 0.25
        If p1CurY + 64 > ScreenCurH Then p1CurY = 0: p1Speed = p1Speed - 0.5: p2Speed = p2Speed + 0.25

        If p2CurX < 0 Then p2CurX = ScreenCurW - 64: p2Speed = p2Speed - 0.5: p1Speed = p1Speed + 0.25
        If p2CurY < 0 Then p2CurY = ScreenCurH - 64: p2Speed = p2Speed - 0.5: p1Speed = p1Speed + 0.25
        If p2CurX + 64 > ScreenCurW Then p2CurX = 0: p2Speed = p2Speed - 0.5: p1Speed = p1Speed + 0.25
        If p2CurY + 64 > ScreenCurH Then p2CurY = 0: p2Speed = p2Speed - 0.5: p1Speed = p1Speed + 0.25

        If p1Speed = 0 Then p1Speed = 1
        If p2Speed = 0 Then p2Speed = 1

        If p1Score >= 30 Then
            gamerun = False
        ElseIf p2Score >= 30 Then
            gamerun = False
        End If

        SetDDFontOptions 9, vbWhite
        ddsBackBuf.DrawText 10, 10, " P1 Score: " & p1Score, False
        ddsBackBuf.DrawText ScreenCurW - 74, ScreenCurH - 20, "P2 Score: " & p2Score, False

        DoEvents

        ddsFrontBuf.Flip Nothing, DDFLIP_WAIT

        Do Until timeGetTime >= TempTime + 16.6666666666667
            'loops until frame rate met
        Loop
    Loop Until gamerun = False

    ddsBackBuf.SetFillColor vbBlack
    ddsBackBuf.SetForeColor vbBlack
    ddsBackBuf.DrawBox 0, 0, ScreenCurW, ScreenCurH
    CreateBG 16, 16, ddsBG, ChosenRes
    BltFast 0, 0, 400, 128, ddsLogo, rectLogo, (ScreenCurW \ 2) - 200, 10, True
    GetRank p1Score, p1Rank

    SetDDFontOptions 16, vbBlack
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 155, "Player 1 Score: " & p1Score, False
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 195, "Player 2 Score: " & p2Score, False
    If p1Score > p2Score Then
        ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 235, "Player 1 Wins!", False
    ElseIf p2Score > p1Score Then
        ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 235, "Player 2 Wins!", False
    ElseIf p2Score = p1Score Then
        ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 235, "No Winners!", False
    End If
    SetDDFontOptions 12, vbBlack
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 275, "Press Enter To Start A New 2 Player Game!", False
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 195, 305, "Press Escape To Return to the Main Menu!", False

    SetDDFontOptions 16, vbWhite
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 150, "Player 1 Score: " & p1Score, False
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 190, "Player 2 Score: " & p2Score, False
    If p1Score > p2Score Then
        ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 230, "Player 1 Wins!", False
    ElseIf p2Score > p1Score Then
        ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 230, "Player 2 Wins!", False
    ElseIf p2Score = p1Score Then
        ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 230, "No Winners!", False
    End If
    SetDDFontOptions 12, vbWhite
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 270, "Press Enter To Start A New 2 Player Game!", False
    ddsBackBuf.DrawText (ScreenCurW \ 2) - 200, 300, "Press Escape To Return to the Main Menu!", False
    ddsFrontBuf.Flip Nothing, DDFLIP_WAIT

RankScreen:
    diKeyBoard.Acquire
    diKeyBoard.GetDeviceStateKeyboard KeyboardState
    Call dmPerformance.Stop(Nothing, Nothing, 0, 0)
    If KeyboardState.Key(DIK_ESCAPE) <> 0 Then
        diKeyBoard.Unacquire
        GoTo MainMenu
    ElseIf KeyboardState.Key(DIK_RETURN) <> 0 Then
        diKeyBoard.Unacquire
        GoTo StartOver
    Else
        GoTo RankScreen
    End If

MainMenu:
    MainMenuScreen
    mListItem = 0

    Exit Sub
errTwoPGame:
    ErrorFound "Continue Play In Two Player Mode"
End Sub

'this checks the score and gives the player a rank based on it
Sub GetRank(pScore As Integer, tRank As String)
    Select Case pScore
        Case -100 To 1
            tRank = "Sludge Worm"
        Case 2 To 3
            tRank = "Puke Face"
        Case 4 To 5
            tRank = "Gigantic Loser"
        Case 6 To 7
            tRank = "Big Loser"
        Case 8 To 9
            tRank = "Loser"
        Case 10 To 11
            tRank = "Pathetic"
        Case 12 To 14
            tRank = "Bad"
        Case 15 To 18
            tRank = "Fair"
        Case 19 To 21
            tRank = "Average"
        Case 22 To 24
            tRank = "Pretty Good"
        Case 25 To 27
            tRank = "Good"
        Case 28 To 30
            tRank = "Very Good"
        Case 31 To 33
            tRank = "Excellent"
        Case 34 To 38
            tRank = "Master of Excellence"
        Case 39 To 49
            tRank = "Unbelievable Talent"
        Case 50 To 60
            tRank = "Almost the Champion"
        Case Else
            If p1Score > -101 Then
                tRank = "Champion"
            Else
                tRank = "Impossible Score, Cheater"
            End If
    End Select
End Sub

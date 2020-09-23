Attribute VB_Name = "modDI"
'this stuff is pretty easy to understand

Public di As DirectInput
Public diKeyBoard As DirectInputDevice
Public KeyboardState As DIKEYBOARDSTATE

Sub InitDI()
    Set di = dx.DirectInputCreate
    Set diKeyBoard = di.CreateDevice("GUID_SysKeyboard")
    diKeyBoard.SetCommonDataFormat DIFORMAT_KEYBOARD
    diKeyBoard.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
End Sub

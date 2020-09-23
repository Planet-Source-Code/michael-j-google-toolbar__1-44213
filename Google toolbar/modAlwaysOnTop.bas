Attribute VB_Name = "modAlwaysOnTop"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE As Long = &H2
Const SWP_NOSIZE As Long = &H1
Public Sub MakeAlwaysOnTop(TheForm As Form, SetOnTop As Boolean)
    Dim lflag
    If SetOnTop Then
        lflag = HWND_TOPMOST
    Else
        lflag = HWND_NOTOPMOST
    End If
    SetWindowPos TheForm.hwnd, lflag, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Attribute VB_Name = "Timer"
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Sub RefreshState(ByVal hwnd As Long, ByVal uint1 As Long, ByVal nEventId As Long, ByVal dwParam As Long)
    Hidden.DetSS
End Sub

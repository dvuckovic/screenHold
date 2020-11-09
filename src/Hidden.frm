VERSION 5.00
Begin VB.Form Hidden 
   BorderStyle     =   0  'None
   ClientHeight    =   570
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   1155
   ControlBox      =   0   'False
   Icon            =   "Hidden.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   1155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHook 
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Image ChgIco 
      Height          =   195
      Left            =   15
      Picture         =   "Hidden.frx":0E42
      Top             =   315
      Width           =   195
   End
   Begin VB.Image ExitIco 
      Height          =   195
      Left            =   495
      Picture         =   "Hidden.frx":1354
      Top             =   315
      Width           =   195
   End
   Begin VB.Image AboutIco 
      Height          =   195
      Left            =   255
      Picture         =   "Hidden.frx":1478
      Top             =   315
      Width           =   195
   End
   Begin VB.Image InvalidSS 
      Height          =   240
      Left            =   480
      Picture         =   "Hidden.frx":1598
      Top             =   0
      Width           =   240
   End
   Begin VB.Image InactiveSS 
      Height          =   240
      Left            =   240
      Picture         =   "Hidden.frx":16E2
      Top             =   0
      Width           =   240
   End
   Begin VB.Image ActiveSS 
      Height          =   240
      Left            =   0
      Picture         =   "Hidden.frx":182C
      Top             =   0
      Width           =   240
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Begin VB.Menu mnuPopupToggle 
         Caption         =   ""
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuPopupExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Hidden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const SPI_GETSCREENSAVEACTIVE = 16
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Dim TrayIcon As NOTIFYICONDATA
Private Sub Form_Load()
    App.TaskVisible = False
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = picHook.hwnd
    TrayIcon.uId = 1&
    TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayIcon.ucallbackMessage = WM_MOUSEMOVE
    DetSS
    Shell_NotifyIcon NIM_ADD, TrayIcon
    SetTimer TrayIcon.hwnd, 1, 1000, AddressOf RefreshState
    SetMenuItemBitmaps GetMenu(Me.hwnd), 2, 0, ChgIco.Picture, 0&
    SetMenuItemBitmaps GetMenu(Me.hwnd), 4, 0, AboutIco.Picture, 0&
    SetMenuItemBitmaps GetMenu(Me.hwnd), 5, 0, ExitIco.Picture, 0&
    Me.Hide
End Sub
Private Function SSState() As Integer
    Dim APIChk As Boolean
    Call SystemParametersInfo(SPI_GETSCREENSAVEACTIVE, 0, APIChk, 0)
    RegChk = regQuery_A_Key(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE")
    If APIChk And Len(RegChk) > 0 Then
        SSState = -1
    ElseIf Not APIChk And Len(RegChk) > 0 Then
        SSState = 1
    ElseIf Not APIChk And Len(RegChk) = 0 Then
        SSState = 0
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    KillTimer TrayIcon.hwnd, 1
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = picHook.hwnd
    TrayIcon.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayIcon
    End
End Sub
Private Sub mnuPopupAbout_Click()
    MsgBox "  (d) by dUcA 2oo2." & Chr(10) _
         & "http://cuzcko.cjb.net", , "screenHold v2.0"
End Sub
Private Sub mnuPopUpExit_Click()
    Unload Me
End Sub
Private Sub mnuPopupToggle_Click()
    ChgSS
End Sub
Public Sub DetSS()
    Select Case SSState
        Case -1
            TrayIcon.hIcon = ActiveSS.Picture
            TrayIcon.szTip = "Active" & Chr$(0)
            mnuPopupToggle.Caption = "Inactivate"
            mnuPopupToggle.Enabled = True
        Case 1
            TrayIcon.hIcon = InactiveSS.Picture
            TrayIcon.szTip = "Inactive" & Chr$(0)
            mnuPopupToggle.Caption = "Activate"
            mnuPopupToggle.Enabled = True
        Case 0
            TrayIcon.hIcon = InvalidSS.Picture
            TrayIcon.szTip = "Disabled" & Chr$(0)
            mnuPopupToggle.Caption = "Disabled"
            mnuPopupToggle.Enabled = False
    End Select
    Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub
Private Sub ChgSS()
    Select Case SSState
        Case -1
            ToggleScreenSaver False
            TrayIcon.hIcon = InactiveSS.Picture
            TrayIcon.szTip = "Inactive" & Chr$(0)
            mnuPopupToggle.Caption = "Activate"
            mnuPopupToggle.Enabled = True
        Case 1
            ToggleScreenSaver True
            TrayIcon.hIcon = ActiveSS.Picture
            TrayIcon.szTip = "Active" & Chr$(0)
            mnuPopupToggle.Caption = "Inactivate"
            mnuPopupToggle.Enabled = True
        Case 0
            TrayIcon.hIcon = InvalidSS.Picture
            TrayIcon.szTip = "Disabled" & Chr$(0)
            mnuPopupToggle.Caption = "Disabled"
            mnuPopupToggle.Enabled = False
    End Select
    Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub
Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    Select Case Msg
        Case WM_LBUTTONDOWN And WM_LBUTTONDBLCLK
            ChgSS
        Case WM_RBUTTONUP
            SetForegroundWindow Me.hwnd
            PopupMenu mnuPopup, 8, , , mnuPopupToggle
    End Select
End Sub
Private Function ToggleScreenSaver(Active As Boolean) As Boolean
    Dim Val As Long
    Dim ActiveFlag As Long
    ActiveFlag = IIf(Active, 1, 0)
    Val = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, ActiveFlag, 0, 0)
    ToggleScreenSaver = Val > 0
End Function

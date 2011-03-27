VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   1395
   ClientTop       =   975
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   955
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ctlKeyWatcher 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   6390
      TabIndex        =   1
      Top             =   990
      Width           =   375
   End
   Begin VB.Timer ctlTimer 
      Interval        =   10
      Left            =   540
      Top             =   2700
   End
   Begin floTa.MorphDisplay ctlLcd 
      Height          =   1755
      Left            =   495
      TabIndex        =   0
      Top             =   90
      Width           =   5685
      _extentx        =   10028
      _extenty        =   3096
      burnincolor     =   0
      burnincolorneg  =   0
      interdigitgap   =   20
      interdigitgapexp=   4
      intersegmentgap =   2
      intersegmentgapexp=   1
      numdigits       =   4
      numdigitsexp    =   2
      segmentlitcolor =   65535
      segmentlitcolorneg=   65535
      segmentheight   =   40
      segmentheightexp=   20
      segmentstyle    =   0
      segmentstyleexp =   0
      segmentwidth    =   12
      segmentwidthexp =   6
      xoffset         =   10
      xoffsetexp      =   305
      yoffset         =   8
      yoffsetexp      =   58
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intHiddenTop As Integer
Dim intNormalTop As Integer
Dim intToTop As Integer
Dim intNormalHeight As Integer
Dim intTransparentPct As Integer
Dim bolMoving As Boolean
Dim timTo As Long
Dim timFrom As Long
Dim bolTimer As Boolean


Private Sub initialization()
    intHiddenTop = -1800
    intNormalTop = 0
    intNormalHeight = 2000
    intToTop = intHiddenTop
    bolMoving = False
    
    Me.Left = 0
    Me.Top = intToTop
    Me.Width = Screen.Width
    Me.Height = intNormalHeight
    
    setTransparent 210

End Sub


Private Function getTimeInMs() As Long
    Dim timTime As SYSTEMTIME
    GetLocalTime timTime
    getTimeInMs = ((CLng(timTime.wHour) * 60 + timTime.wMinute) * 60 + timTime.wSecond) * 100 + Round(timTime.wMilliseconds / 10)
End Function


Private Sub countDown(intMin)
    timFrom = getTimeInMs
    timTo = timFrom + intMin * 60 * 100
End Sub





Private Sub setTransparent(intPct)
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, intPct, LWA_ALPHA
End Sub


Private Sub hidden(bolFlag)
    If bolMoving = True Then
        Exit Sub
    End If
    bolMoving = True
    If bolFlag = True Then
        intToTop = intHiddenTop
    Else
        intToTop = intNormalTop
    End If
End Sub


Private Sub changeHeight()
    If Me.Top > intToTop Then
        Me.Top = Me.Top - 50
    ElseIf Me.Top < intToTop Then
        Me.Top = Me.Top + 50
    ElseIf Me.Top = intToTop Then
        bolMoving = False
    End If
End Sub


Private Sub ctlKeyWatcher_KeyPress(KeyAscii As Integer)
Debug.Print KeyAscii
    Select Case KeyAscii
        Case "8"
            MsgBox 8
        Case "13"
            MsgBox 13
    End Select
    
End Sub

Private Sub ctlTimer_Timer()
ctlKeyWatcher.SetFocus
    changeHeight
    'Debug.Print getTimeInMs
    Dim s As String
    s = Format(Now, "hh:mm:ss ampm")
    If Left(s, 1) = "0" Then
      s = Right(s, Len(s) - 1)
      ctlLcd.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
    Else
      ctlLcd.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
    End If
End Sub


Private Sub Form_DblClick()
    hidden True
End Sub


Private Sub Form_Load()
    initialization
    
    
    
    
    
     Dim ret As Long
    '记录原来的window程序地址
    preWinProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    '用自定义程序代替原来的window程序
    ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf wndproc)
    idHotKey = 1 'in the range ＆h0000 through ＆hBFFF
    Modifiers = MOD_ALT '辅助键为Alt
    uVirtKey1 = vbKeyQ '注册的热键为Alt+Q
    '注册热键
    ret = RegisterHotKey(Me.hwnd, idHotKey, Modifiers, uVirtKey1)
    If ret = 0 Then
        MsgBox "注册热键失败,请使用其它热键!", vbCritical, "错误"
    End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    hidden False
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ret As Long
    '取消Message的截取，使之送往原来的window程序
    ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, preWinProc)
    Call UnregisterHotKey(Me.hwnd, uVirtKey1)
End Sub

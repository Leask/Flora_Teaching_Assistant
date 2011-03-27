VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   1395
   ClientTop       =   975
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ShowInTaskbar   =   0   'False
   Begin VB.Timer ctlTimer 
      Interval        =   10
      Left            =   5715
      Top             =   2520
   End
   Begin floTa.MorphDisplay ctlLcd 
      Height          =   1755
      Left            =   495
      TabIndex        =   0
      Top             =   90
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   3096
      BurnInColor     =   0
      BurnInColorNeg  =   0
      InterDigitGap   =   20
      InterDigitGapExp=   4
      InterSegmentGap =   2
      InterSegmentGapExp=   1
      NumDigits       =   4
      NumDigitsExp    =   2
      SegmentHeight   =   40
      SegmentHeightExp=   20
      SegmentLitColor =   65535
      SegmentLitColorNeg=   65535
      SegmentStyle    =   0
      SegmentStyleExp =   0
      SegmentWidth    =   12
      SegmentWidthExp =   6
      XOffset         =   10
      XOffsetExp      =   305
      YOffset         =   8
      YOffsetExp      =   58
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
Dim timTo As Date


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


Private Function getTimeInMs()
    Dim timTime As SYSTEMTIME
    GetLocalTime timTime

s = CStr(sys.wYear) + ", " + CStr(sys.wMonth) + ", " + CStr(sys.wDay)
     '   Label1.Caption = s
        
    s1 = CStr(sys.wHour) + ": " + CStr(sys.wMinute) + ": " + CStr(sys.wSecond) + ": " + CStr(sys.wMilliseconds)
'   »òÕßs   =   Time$   +   "   :   "   +   CStr(sys.wMilliseconds)

 Timer1_Timer = s1
        
End Function


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


Private Sub ctlTimer_Timer()

    changeHeight
    //Debug.Print Timer1_Timer
    Dim s As String
    s = Format(Now, "hh:mm:ss ampm")
    If Left(s, 1) = "0" Then
      s = Right(s, Len(s) - 1)
      ctlLcd.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
    Else
      ctlLcd.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
    End If
End Sub

Private Sub Form_Click()
    'MsgBox "ok"
    
    hidden True
End Function

Private Sub Form_Load()
    initialization
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    hidden False
End Sub


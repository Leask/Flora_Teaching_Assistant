VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   1395
   ClientTop       =   975
   ClientWidth     =   3180
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   ShowInTaskbar   =   0   'False
   Begin VB.Timer ctlTimer 
      Interval        =   10
      Left            =   2655
      Top             =   540
   End
   Begin FloTa.MorphDisplay ctlLcd 
      Height          =   945
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   1667
      BurnInColor     =   0
      BurnInColorNeg  =   0
      InterDigitGap   =   10
      InterDigitGapExp=   4
      InterSegmentGap =   2
      InterSegmentGapExp=   1
      NumDigits       =   4
      NumDigitsExp    =   2
      SegmentHeight   =   20
      SegmentHeightExp=   10
      SegmentLitColor =   65535
      SegmentLitColorNeg=   65535
      SegmentStyle    =   0
      SegmentStyleExp =   0
      SegmentWidth    =   6
      Value           =   "1126"
      XOffsetExp      =   165
      YOffset         =   4
      YOffsetExp      =   31
   End
   Begin VB.Label ctlPctLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   1080
      Width           =   2985
   End
   Begin VB.Shape ctlPctShape 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   330
      Left            =   45
      Top             =   1035
      Width           =   3075
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
Dim intTsnMax As Integer
Dim intTsnCur As Integer
Dim intTsnTo As Integer
Dim timFrom As Long
Dim timTo As Long
Dim timALl As Long
Dim timPause As Long
Dim intPctWidth As Integer
Dim bolFlash As Boolean
Dim lngFlashTimer As Long


Private Sub initialization()
    If App.PrevInstance Then End
    intNormalHeight = 1400
    intNormalTop = 0
    intHiddenTop = intNormalTop - intNormalHeight
    intToTop = intHiddenTop
    intPctWidth = 199
    intTsnMax = 210
    timPause = 0
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = intToTop
    Me.Height = intNormalHeight
    
    setFormOnTop Me
    
    setTransparent 0
End Sub


Private Function getTimeInMs() As Long
    Dim timTime As SYSTEMTIME
    GetLocalTime timTime
    getTimeInMs = ((CLng(timTime.wHour) * 60 + timTime.wMinute) * 60 + timTime.wSecond) * 100 + Round(timTime.wMilliseconds / 10)
End Function


Public Sub countDown(intMin)
    If intMin = 0 Then
        timFrom = 0
        timTo = 0
        hidden True
        Exit Sub
    Else
        hidden False
    End If
    bolFlash = False
    timPause = 0
    timFrom = getTimeInMs
    timALl = intMin * 60 * 100
    timTo = timFrom + timALl
End Sub


Private Sub setTransparent(intPct)
    intTsnCur = intPct
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, intPct, LWA_ALPHA
End Sub


Private Sub doFlash()
    If bolFlash Then
        lngFlashTimer = lngFlashTimer + 1
        If lngFlashTimer Mod 10 = 0 Then
            If ctlLcd.Visible Then
                ctlLcd.Visible = False
            Else
                ctlLcd.Visible = True
            End If
        End If
        If lngFlashTimer >= 1000 Then
            countDown 0
        End If
    Else
        lngFlashTimer = 0
        ctlLcd.Visible = True
    End If
End Sub


Private Sub hidden(bolFlag)
    If bolFlag = True Then
        intToTop = intHiddenTop
        intTsnTo = 0
    Else
        intToTop = intNormalTop
        intTsnTo = intTsnMax
    End If
End Sub


Private Sub changeHeight()
    If Me.Top > intToTop Then
        Me.Top = Me.Top - 50
    ElseIf Me.Top < intToTop Then
        Me.Top = Me.Top + 50
    End If
End Sub


Private Sub changeTransparent()
    If intTsnCur > intTsnTo Then
        setTransparent intTsnCur - 10
    ElseIf intTsnCur < intTsnTo Then
        setTransparent intTsnCur + 10
    End If
End Sub


Public Sub chHide()
    If intToTop = intHiddenTop Then
        hidden False
    Else
        hidden True
    End If
End Sub


Public Sub pauseTimer()
    If timPause Then
        Dim lngTimPausePast As Long
        lngTimPausePast = getTimeInMs - timPause
        timFrom = timFrom + lngTimPausePast
        timTo = timTo + lngTimPausePast
        timPause = 0
    Else
        timPause = getTimeInMs
    End If
    hidden False
End Sub


Private Sub ctlTimer_Timer()
    changeHeight
    changeTransparent
    doFlash
    If timTo And timPause = 0 Then
        Dim lessTime As Long
        Dim lngMin As Long
        Dim lngSec As Long
        Dim lngMsc As Long
        Dim sngPct As Single
        lessTime = timTo - getTimeInMs
        If lessTime >= 0 Then
            lngMin = lessTime \ 6000
            lngSec = (lessTime - (lngMin * 6000)) \ 100
            lngMsc = lessTime - (lngMin * 6000) - (lngSec * 100)
            ctlLcd.Value = Format(lngMin, "00") & ":" & Format(lngSec, "00") & "E+" & Format(lngMsc, "00")
            sngPct = (getTimeInMs - timFrom) / timALl
            ctlPctLabel.Width = intPctWidth * sngPct
            If ctlPctLabel.Width > 25 Then
                ctlPctLabel.Caption = Int(sngPct * 100) & "% "
            Else
                ctlPctLabel.Caption = ""
            End If
            If Abs(lessTime - 500) <= 10 Then
                PlaySound App.Path & "\rsCountDown.wav"
            End If
        Else
            ctlLcd.Value = "00:00E+00"
            ctlPctLabel.Caption = "100% "
            bolFlash = True
        End If
    End If
End Sub


Private Sub Form_Load()
    initialization
    SetHotkey 0, "Alt,48", "Add"
    SetHotkey 1, "Alt,49", "Add"
    SetHotkey 2, "Alt,50", "Add"
    SetHotkey 3, "Alt,51", "Add"
    SetHotkey 4, "Alt,52", "Add"
    SetHotkey 5, "Alt,53", "Add"
    SetHotkey 6, "Alt,54", "Add"
    SetHotkey 7, "Alt,55", "Add"
    SetHotkey 8, "Alt,56", "Add"
    SetHotkey 9, "Alt,57", "Add"
    SetHotkey 10, "Alt,27", "Add"
    SetHotkey 11, "Alt,80", "Add"
    SetHotkey 12, "Alt,72", "Add"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetHotkey 0, "Alt,48", "Del"
    SetHotkey 1, "Alt,49", "Del"
    SetHotkey 2, "Alt,50", "Del"
    SetHotkey 3, "Alt,51", "Del"
    SetHotkey 4, "Alt,52", "Del"
    SetHotkey 5, "Alt,53", "Del"
    SetHotkey 6, "Alt,54", "Del"
    SetHotkey 7, "Alt,55", "Del"
    SetHotkey 8, "Alt,56", "Del"
    SetHotkey 9, "Alt,57", "Del"
    SetHotkey 10, "Alt,27", "Del"
    SetHotkey 11, "Alt,80", "Del"
    SetHotkey 12, "Alt,72", "Del"
End Sub

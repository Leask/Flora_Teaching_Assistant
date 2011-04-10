VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   1395
   ClientTop       =   975
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   ShowInTaskbar   =   0   'False
   Begin VB.Timer ctlTimer 
      Interval        =   10
      Left            =   5220
      Top             =   1305
   End
   Begin floTa.MorphDisplay ctlLcd 
      Height          =   1755
      Left            =   45
      TabIndex        =   0
      Top             =   45
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
      TabIndex        =   1
      Top             =   1890
      Width           =   5595
   End
   Begin VB.Shape ctlPctShape 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   330
      Left            =   45
      Top             =   1845
      Width           =   5685
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
Dim timTo As Long
Dim timFrom As Long
Dim timALl As Long
Dim intPctWidth As Integer
Dim bolFlash As Boolean
Dim infFlashTimer As Integer


Private Sub initialization()
    intNormalHeight = 2200
    intNormalTop = 0
    intHiddenTop = intNormalTop - intNormalHeight
    intToTop = intHiddenTop
    intPctWidth = 373
    intTsnMax = 210
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = intToTop
    Me.Height = intNormalHeight
    
    SetFormOnTop Me
    
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
        infFlashTimer = infFlashTimer + 1
        If infFlashTimer Mod 10 = 0 Then
            If ctlLcd.Visible Then
                ctlLcd.Visible = False
            Else
                ctlLcd.Visible = True
            End If
        End If
    Else
        infFlashTimer = 0
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


Private Sub ctlTimer_Timer()
    changeHeight
    changeTransparent
    doFlash
    If timTo Then
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
            ctlPctLabel.Width = 373 * sngPct
            If ctlPctLabel.Width > 25 Then
                ctlPctLabel.Caption = Int(sngPct * 100) & "% "
            Else
                ctlPctLabel.Caption = ""
            End If
        Else
            ctlLcd.Value = "00:00E+00"
            ctlPctLabel.Caption = "101% "
            bolFlash = True
        End If
    End If
End Sub


Private Sub Form_DblClick()
    hidden True
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
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetHotkey 0, "Alt,40", "Del"
    SetHotkey 1, "Alt,49", "Del"
    SetHotkey 2, "Alt,50", "Del"
    SetHotkey 3, "Alt,51", "Del"
    SetHotkey 4, "Alt,52", "Del"
    SetHotkey 5, "Alt,53", "Del"
    SetHotkey 6, "Alt,54", "Del"
    SetHotkey 7, "Alt,55", "Del"
    SetHotkey 8, "Alt,56", "Del"
    SetHotkey 9, "Alt,57", "Del"
End Sub

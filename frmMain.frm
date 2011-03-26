VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6840
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer ctlTimer 
      Interval        =   500
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
      InterDigitGap   =   20
      InterDigitGapExp=   4
      InterSegmentGap =   2
      InterSegmentGapExp=   1
      NumDigits       =   4
      NumDigitsExp    =   2
      SegmentHeight   =   40
      SegmentHeightExp=   20
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
Dim

Private Sub ctlTimer_Timer()
   Dim s As String
   s = Format(Now, "hh:mm:ss ampm")
   If Left(s, 1) = "0" Then
      s = Right(s, Len(s) - 1)
      ctlLcd.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
   Else
      ctlLcd.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
   End If
End Sub


VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "数字时钟演示"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   6615
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   6000
      Top             =   240
   End
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   630
      Left            =   1200
      TabIndex        =   0
      Top             =   3000
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1111
      BurnInColor     =   96
      BurnInColorNeg  =   96
      InterSegmentGap =   1
      NumDigits       =   4
      NumDigitsExp    =   2
      SegmentHeight   =   12
      SegmentHeightExp=   8
      SegmentLitColor =   255
      SegmentLitColorNeg=   255
      SegmentStyle    =   0
      SegmentWidth    =   4
      SegmentWidthExp =   4
      XOffsetExp      =   105
      YOffsetExp      =   14
   End
   Begin Project1.MorphDisplay MorphLCD2 
      Height          =   1755
      Left            =   495
      TabIndex        =   1
      Top             =   90
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   3096
      BurnInColor     =   4210688
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
   Begin Project1.MorphDisplay MorphLCD3 
      Height          =   2235
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   3942
      BorderColor     =   255
      BurnInColor     =   4210688
      InterDigitGap   =   20
      InterDigitGapExp=   6
      InterSegmentGap =   2
      InterSegmentGapExp=   1
      NumDigits       =   4
      NumDigitsExp    =   2
      Picture         =   "frmTimeDemo.frx":0000
      PictureMode     =   2
      SegmentFillStyle=   0
      SegmentHeight   =   60
      SegmentHeightExp=   30
      SegmentStyle    =   1
      SegmentStyleExp =   1
      SegmentWidth    =   12
      SegmentWidthExp =   6
      ShowBurnIn      =   0   'False
      XOffset         =   10
      XOffsetExp      =   345
      YOffset         =   10
      YOffsetExp      =   74
   End
   Begin Project1.MorphDisplay MorphLCD4 
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
      BackColor1      =   4210752
      BackColor2      =   12632256
      BorderColor     =   0
      BurnInColor     =   14737632
      BurnInColorNeg  =   96
      InterDigitGap   =   8
      InterSegmentGap =   1
      InterSegmentGapExp=   1
      NumDigits       =   4
      NumDigitsExp    =   2
      SegmentHeight   =   24
      SegmentHeightExp=   12
      SegmentLitColor =   0
      SegmentLitColorNeg=   255
      SegmentStyle    =   1
      SegmentStyleExp =   1
      SegmentWidth    =   6
      SegmentWidthExp =   4
      ShowBurnIn      =   0   'False
      XOffsetExp      =   140
      YOffset         =   7
      YOffsetExp      =   31
   End
   Begin Project1.MorphDisplay MorphLCD5 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      BurnInColor     =   20560
      BurnInColorNeg  =   20560
      InterDigitGap   =   4
      InterDigitGapExp=   3
      NumDigits       =   4
      NumDigitsExp    =   2
      SegmentHeight   =   4
      SegmentHeightExp=   4
      SegmentLitColor =   65535
      SegmentLitColorNeg=   65535
      SegmentWidth    =   2
      SegmentWidthExp =   2
      Theme           =   6
      XOffsetExp      =   50
   End
   Begin Project1.MorphDisplay MorphLCD6 
      Height          =   825
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1455
      BorderColor     =   49152
      BurnInColor     =   16384
      BurnInColorNeg  =   96
      NumDigits       =   4
      NumDigitsExp    =   2
      SegmentHeight   =   16
      SegmentHeightExp=   8
      SegmentLitColor =   8454016
      SegmentLitColorNeg=   255
      SegmentWidth    =   4
      XOffsetExp      =   115
      YOffset         =   7
      YOffsetExp      =   26
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTimeDemo.frx":025E
      Height          =   5775
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
' ：） 人人为我，我为人人 ：）
'枕善居汉化收藏整理
'发布日期：06/02/15
'描    述：LED/LCD数字显示控件
'网    站：http://www.mndsoft.com/
'e-mail  ：mnd@mndsoft.com
'OICQ    ：88382850
'****************************************************************************
Private Sub Timer1_Timer()
   Dim s As String
   s = Format(Now, "hh:mm:ss ampm")
   If Left(s, 1) = "0" Then
      s = Right(s, Len(s) - 1)
      MorphLCD1.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
   Else
      MorphLCD1.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
   End If
   MorphLCD2.Value = MorphLCD1.Value
   MorphLCD3.Value = MorphLCD1.Value
   MorphLCD4.Value = MorphLCD1.Value
   MorphLCD5.Value = MorphLCD1.Value
   MorphLCD6.Value = MorphLCD1.Value
End Sub


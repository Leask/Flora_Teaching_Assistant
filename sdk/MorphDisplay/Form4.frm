VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "使用图片示例"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form4"
   ScaleHeight     =   8040
   ScaleWidth      =   9420
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   6000
      Top             =   1440
   End
   Begin Project1.MorphDisplay MorphLCD2 
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
   Begin Project1.MorphDisplay MorphLCD4 
      Height          =   975
      Left            =   6600
      TabIndex        =   1
      Top             =   120
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
   Begin Project1.MorphDisplay MorphLCD10 
      Height          =   630
      Left            =   7200
      TabIndex        =   2
      Top             =   1680
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
   Begin Project1.MorphDisplay MorphLCD5 
      Height          =   300
      Left            =   8280
      TabIndex        =   3
      Top             =   3840
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
      Left            =   7080
      TabIndex        =   4
      Top             =   2760
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
   Begin Project1.MorphDisplay MorphLCD3 
      Height          =   2235
      Left            =   120
      TabIndex        =   5
      Top             =   1920
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
      Picture         =   "Form4.frx":0000
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
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   450
      Index           =   3
      Left            =   3720
      TabIndex        =   6
      Top             =   4320
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   794
      BurnInColor     =   96
      BurnInColorNeg  =   96
      SegmentHeightExp=   5
      SegmentLitColor =   255
      SegmentLitColorNeg=   255
      SegmentStyle    =   0
      SegmentWidth    =   4
      ShowExponent    =   0   'False
      Theme           =   1
      Value           =   "12345"
      XOffsetExp      =   375
   End
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   570
      Index           =   4
      Left            =   1560
      TabIndex        =   7
      Top             =   5520
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   1005
      BackColor1      =   14737632
      BackColor2      =   14737632
      BurnInColor     =   13684944
      BurnInColorNeg  =   13684944
      InterDigitGap   =   4
      InterSegmentGap =   1
      SegmentHeight   =   12
      SegmentLitColor =   0
      SegmentLitColorNeg=   0
      SegmentStyle    =   1
      SegmentWidth    =   4
      Theme           =   3
      Value           =   "2.71828182845905"
      XOffsetExp      =   370
   End
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   735
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   6960
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   1296
      BurnInColor     =   96
      BurnInColorNeg  =   96
      InterSegmentGap =   1
      SegmentHeight   =   12
      SegmentHeightExp=   8
      SegmentLitColor =   255
      SegmentLitColorNeg=   255
      SegmentStyle    =   0
      SegmentWidth    =   4
      SegmentWidthExp =   4
      ShowThousandsSeparator=   -1  'True
      Theme           =   2
      Value           =   "3.1415926535897932"
      XOffsetExp      =   490
   End
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   529
      BurnInColor     =   20560
      BurnInColorNeg  =   20560
      InterDigitGap   =   4
      InterDigitGapExp=   3
      SegmentHeight   =   4
      SegmentHeightExp=   4
      SegmentLitColor =   65535
      SegmentLitColorNeg=   65535
      SegmentWidth    =   2
      SegmentWidthExp =   2
      Theme           =   6
      Value           =   "12345"
      XOffsetExp      =   195
   End
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   630
      Index           =   6
      Left            =   840
      TabIndex        =   10
      Top             =   6240
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   1111
      BackColor1      =   4194304
      BackColor2      =   16711680
      BorderColor     =   0
      BurnInColor     =   13684944
      BurnInColorNeg  =   13684944
      InterSegmentGap =   1
      SegmentHeight   =   14
      SegmentHeightExp=   8
      SegmentLitColor =   0
      SegmentStyle    =   1
      SegmentWidth    =   4
      SegmentWidthExp =   4
      ShowBurnIn      =   0   'False
      Theme           =   4
      Value           =   "-1.23456789E+123"
      XOffsetExp      =   455
   End
   Begin Project1.MorphDisplay MorphDisplay1 
      Height          =   585
      Left            =   1680
      TabIndex        =   11
      Top             =   4800
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   1032
      ShowThousandsSeparator=   -1  'True
      Value           =   "12333"
      YOffset         =   6
      YOffsetExp      =   6
   End
End
Attribute VB_Name = "Form4"
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
Private Sub Form_Load()
MorphDisplay1.Value = 123456789.123
MorphLCD1(4).Value = 2.71828182845905
MorphLCD1(3).Value = 987654321
MorphLCD1(6).Value = -9.087654321693E+72
MorphLCD1(5).Value = 1.70908765432169E+75
MorphLCD1(2).Value = 987654321.98765
End Sub
Private Sub Timer1_Timer()
   Dim s As String
   s = Format(Now, "hh:mm:ss ampm")
   If Left(s, 1) = "0" Then
      s = Right(s, Len(s) - 1)
      MorphLCD10.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
   Else
      MorphLCD10.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
   End If
   MorphLCD2.Value = MorphLCD10.Value
   MorphLCD3.Value = MorphLCD10.Value
   MorphLCD4.Value = MorphLCD10.Value
   MorphLCD5.Value = MorphLCD10.Value
   MorphLCD6.Value = MorphLCD10.Value
End Sub


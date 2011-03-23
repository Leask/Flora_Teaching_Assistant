VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "使用分隔符演示"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form3"
   ScaleHeight     =   4935
   ScaleWidth      =   8535
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option1 
      Caption         =   "三位"
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   8
      Top             =   720
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "2位"
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin Project1.MorphDisplay MorphDisplay3 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   1296
      BackColor1      =   14737632
      BackColor2      =   14737632
      BurnInColor     =   13684944
      BurnInColorNeg  =   13684944
      InterSegmentGap =   1
      SegmentHeight   =   14
      SegmentHeightExp=   8
      SegmentLitColor =   0
      SegmentLitColorNeg=   0
      SegmentStyle    =   1
      SegmentWidth    =   4
      SegmentWidthExp =   4
      ShowThousandsSeparator=   -1  'True
      Theme           =   4
      XOffsetExp      =   455
   End
   Begin Project1.MorphDisplay MorphDisplay2 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1560
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
      XOffsetExp      =   490
   End
   Begin Project1.MorphDisplay MorphDisplay1 
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   1032
      ShowThousandsSeparator=   -1  'True
      Value           =   "123456789"
      YOffset         =   6
      YOffsetExp      =   6
   End
   Begin Project1.MorphDisplay MorphDisplay4 
      Height          =   585
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   1032
      DecimalSeparator=   0
      ShowThousandsSeparator=   -1  'True
      ThousandsSeparator=   1
      Value           =   "123456789"
      YOffset         =   6
      YOffsetExp      =   6
   End
   Begin Project1.MorphDisplay MorphDisplay5 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   1296
      BurnInColor     =   96
      BurnInColorNeg  =   96
      DecimalSeparator=   0
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
      ThousandsSeparator=   1
      XOffsetExp      =   490
   End
   Begin Project1.MorphDisplay MorphDisplay6 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   1296
      BackColor1      =   14737632
      BackColor2      =   14737632
      BurnInColor     =   13684944
      BurnInColorNeg  =   13684944
      DecimalSeparator=   0
      InterSegmentGap =   1
      SegmentHeight   =   14
      SegmentHeightExp=   8
      SegmentLitColor =   0
      SegmentLitColorNeg=   0
      SegmentStyle    =   1
      SegmentWidth    =   4
      SegmentWidthExp =   4
      ShowThousandsSeparator=   -1  'True
      Theme           =   4
      ThousandsSeparator=   1
      XOffsetExp      =   455
   End
   Begin VB.Label Label1 
      Caption         =   "千位分组"
      Height          =   255
      Left            =   6360
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
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
   MorphDisplay1.Value = -1234567890.123
   MorphDisplay2.Value = 1234567890.123
   MorphDisplay3.Value = 1234567890.123
   MorphDisplay4.Value = -1234567890.123
   MorphDisplay5.Value = 1234567890.123
   MorphDisplay6.Value = 1234567890.123
End Sub

Private Sub Option1_Click(Index As Integer)

   Select Case Index
      Case 0
         MorphDisplay1.ThousandsGrouping = 2
         MorphDisplay2.ThousandsGrouping = 2
         MorphDisplay3.ThousandsGrouping = 2
         MorphDisplay4.ThousandsGrouping = 2
         MorphDisplay5.ThousandsGrouping = 2
         MorphDisplay6.ThousandsGrouping = 2
      Case 1
         MorphDisplay1.ThousandsGrouping = 3
         MorphDisplay2.ThousandsGrouping = 3
         MorphDisplay3.ThousandsGrouping = 3
         MorphDisplay4.ThousandsGrouping = 3
         MorphDisplay5.ThousandsGrouping = 3
         MorphDisplay6.ThousandsGrouping = 3
   End Select
End Sub

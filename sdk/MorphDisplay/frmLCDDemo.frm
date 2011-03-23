VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "默认主题示例"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   8655
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Text            =   "500"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "自动"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "-500"
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "12345.6789"
      Top             =   5880
      Width           =   3855
   End
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   1085
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
      Theme           =   2
      Value           =   "12345"
      XOffsetExp      =   490
   End
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   465
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   820
      Value           =   "12345"
   End
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2280
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
      Height          =   450
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   794
      BurnInColor     =   96
      BurnInColorNeg  =   96
      InterDigitGap   =   4
      SegmentHeightExp=   5
      SegmentLitColor =   255
      SegmentLitColorNeg=   255
      SegmentStyle    =   0
      SegmentWidth    =   4
      Theme           =   1
      Value           =   "12345"
      XOffsetExp      =   335
   End
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   570
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   3840
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
      Value           =   "12345"
      XOffsetExp      =   370
   End
   Begin Project1.MorphDisplay MorphLCD1 
      Height          =   615
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   1085
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
      Theme           =   4
      Value           =   "12345"
      XOffsetExp      =   455
   End
   Begin VB.Label Label6 
      Caption         =   "LCD Trap Medium"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "LCD Trap Small"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Rectangular Medium"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Rectangular Small"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "LED Hex Medium"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "LED Hex Small"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
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

Option Explicit
Private i As Long, j As Long

Private Sub Command1_Click()
For i = 0 To 5
   MorphLCD1(i).Value = Text1.Text
Next i
End Sub

Private Sub Command2_Click()
Dim n1 As Long, n2 As Long, i As Long
n1 = Val(Text2.Text)
n2 = Val(Text3.Text)
For i = n1 To n2
   MorphLCD1(0).Value = i
   MorphLCD1(1).Value = i
   MorphLCD1(2).Value = i
   MorphLCD1(3).Value = i
   MorphLCD1(4).Value = i
   MorphLCD1(5).Value = i
Next i
End Sub

Private Sub Form_Load()
For i = 0 To 5
   MorphLCD1(i).Value = Text1.Text
Next i

Form2.Show
Form3.Show

End Sub

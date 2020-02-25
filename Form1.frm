VERSION 5.00
Object = "*\AaxLabelPlus.vbp"
Begin VB.Form Form1 
   Caption         =   "New Properties axLabelPlus (Mod Version of Great LabelPlus from Leandro Ascierto)"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGlowing 
      Caption         =   "Glowing"
      Height          =   360
      Left            =   4575
      TabIndex        =   38
      Top             =   2310
      Width           =   1305
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   7875
      TabIndex        =   36
      Text            =   "Caption2"
      Top             =   5415
      Width           =   2160
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   7875
      TabIndex        =   35
      Text            =   "Caption1"
      Top             =   5010
      Width           =   2160
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   3225
      TabIndex        =   32
      Top             =   4650
      Width           =   1860
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Cross Visible"
      Height          =   210
      Left            =   1815
      TabIndex        =   31
      Top             =   4605
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Visible=True"
      Height          =   360
      Left            =   1710
      TabIndex        =   30
      Top             =   5880
      Width           =   1350
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   210
      Left            =   7095
      Max             =   50
      TabIndex        =   25
      Top             =   2445
      Width           =   2625
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1725
      Left            =   6660
      Max             =   50
      TabIndex        =   24
      Top             =   3045
      Width           =   210
   End
   Begin VB.TextBox Y2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9060
      TabIndex        =   20
      Text            =   "00"
      Top             =   4560
      Width           =   405
   End
   Begin VB.TextBox Y1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9060
      TabIndex        =   18
      Text            =   "00"
      Top             =   4170
      Width           =   405
   End
   Begin VB.TextBox X2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8190
      TabIndex        =   16
      Text            =   "00"
      Top             =   4560
      Width           =   405
   End
   Begin VB.TextBox X1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8205
      TabIndex        =   14
      Text            =   "00"
      Top             =   4170
      Width           =   405
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1725
      Left            =   6375
      Max             =   50
      TabIndex        =   13
      Top             =   3045
      Width           =   210
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   210
      Left            =   7095
      Max             =   50
      TabIndex        =   12
      Top             =   2190
      Width           =   2625
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Gradient"
      Height          =   225
      Left            =   3360
      TabIndex        =   9
      Top             =   3615
      Width           =   1950
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeHotlineColor"
      Height          =   195
      Index           =   2
      Left            =   3360
      TabIndex        =   3
      Top             =   1050
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeBorderColor"
      Height          =   195
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   780
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeNone"
      Height          =   195
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   510
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   " Glowing "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   3735
      TabIndex        =   39
      Top             =   1815
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1050
      Left            =   3630
      Top             =   1950
      Width           =   2520
   End
   Begin AXLPCTRL.axLabelPlus axLPGlow 
      Height          =   555
      Left            =   3810
      TabIndex        =   37
      Top             =   2190
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BackColor       =   65535
      Border          =   -1  'True
      BorderColor     =   8438015
      BorderColorOpacity=   30
      BorderCornerLeftTop=   50
      BorderCornerRightTop=   50
      BorderCornerBottomRight=   50
      BorderCornerBottomLeft=   50
      BorderPosition  =   0
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption1        =   "Form1.frx":0000
      Caption2        =   "Form1.frx":0022
      CaptionBorderColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ChangeOnMouseOver=   0
      GradientColorP1 =   0
      GradientColorP1Opacity=   0
      GradientColorP2 =   0
      GradientColorP2Opacity=   0
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      HotLineColor    =   0
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureArr      =   0
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption1"
      Height          =   195
      Index           =   1
      Left            =   7140
      TabIndex        =   34
      Top             =   5070
      Width           =   645
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption2"
      Height          =   195
      Index           =   1
      Left            =   7140
      TabIndex        =   33
      Top             =   5490
      Width           =   645
   End
   Begin AXLPCTRL.axLabelPlus axLPCross 
      Height          =   915
      Left            =   345
      TabIndex        =   29
      Top             =   4890
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1614
      BackColor       =   12648447
      BackColorPress  =   8438015
      Border          =   -1  'True
      BorderColor     =   16711680
      ColorOnMouseOver=   255
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption1        =   "Form1.frx":0042
      Caption2        =   "Form1.frx":007E
      Caption2PaddingX=   10
      Caption2PaddingY=   10
      CrossPosition   =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOnPress=   16777215
      ChangeColorOnClick=   -1  'True
      ChangeOnMouseOver=   0
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      HotLine         =   -1  'True
      HotLineWidth    =   8
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CrossClose"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   360
      TabIndex        =   28
      Top             =   4545
      Width           =   2925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   6480
      TabIndex        =   27
      Top             =   135
      Width           =   2925
   End
   Begin AXLPCTRL.axLabelPlus axLPValue 
      Height          =   990
      Left            =   6480
      TabIndex        =   26
      Top             =   525
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1746
      BackColor       =   12648447
      BackColorPress  =   8438015
      Border          =   -1  'True
      BorderColor     =   -2147483635
      ColorOnMouseOver=   255
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      CaptionAlignmentH=   1
      Caption1        =   "Form1.frx":009E
      Caption2        =   "Form1.frx":00D0
      Caption2SizeMinus=   1
      Caption2PaddingX=   50
      Caption2PaddingY=   12
      CrossPosition   =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOnPress=   16777215
      ChangeColorOnClick=   -1  'True
      ChangeOnMouseOver=   0
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      HotLine         =   -1  'True
      HotLineWidth    =   45
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption2"
      Height          =   195
      Index           =   0
      Left            =   7170
      TabIndex        =   23
      Top             =   4590
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption1"
      Height          =   195
      Index           =   0
      Left            =   7170
      TabIndex        =   22
      Top             =   4230
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y :"
      Height          =   195
      Left            =   8820
      TabIndex        =   21
      Top             =   4605
      Width           =   195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y :"
      Height          =   195
      Left            =   8820
      TabIndex        =   19
      Top             =   4215
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X :"
      Height          =   195
      Left            =   7950
      TabIndex        =   17
      Top             =   4605
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X :"
      Height          =   195
      Left            =   7965
      TabIndex        =   15
      Top             =   4215
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DualCaption"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   6435
      TabIndex        =   11
      Top             =   1845
      Width           =   2925
   End
   Begin AXLPCTRL.axLabelPlus axLPdc 
      Height          =   1245
      Left            =   7020
      TabIndex        =   10
      Top             =   2790
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2196
      BackColor       =   12648447
      BackColorPress  =   8438015
      Border          =   -1  'True
      BorderColor     =   16711680
      ColorOnMouseOver=   255
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":00FA
      Caption2        =   "Form1.frx":012A
      CrossPosition   =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOnPress=   16777215
      ChangeColorOnClick=   -1  'True
      ChangeOnMouseOver=   0
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      HotLine         =   -1  'True
      HotLineWidth    =   8
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ChangeColorOnClick"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   300
      TabIndex        =   8
      Top             =   3135
      Width           =   2925
   End
   Begin AXLPCTRL.axLabelPlus axLPccc 
      Height          =   795
      Left            =   285
      TabIndex        =   7
      Top             =   3480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1402
      BackColor       =   12648447
      BackColorPress  =   8438015
      Border          =   -1  'True
      BorderColor     =   16711680
      ColorOnMouseOver=   255
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption1        =   "Form1.frx":015A
      Caption2        =   "Form1.frx":018C
      Caption2PaddingX=   10
      Caption2PaddingY=   10
      CrossPosition   =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOnPress=   16777215
      ChangeColorOnClick=   -1  'True
      ChangeOnMouseOver=   0
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      HotLine         =   -1  'True
      HotLineWidth    =   8
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLPbc 
      Height          =   795
      Index           =   2
      Left            =   300
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1402
      BackColor       =   12648447
      Border          =   -1  'True
      BorderColor     =   16711680
      ColorOnMouseOver=   255
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":01AC
      Caption2        =   "Form1.frx":01CC
      Caption1PaddingX=   10
      Caption2PaddingX=   10
      Caption2PaddingY=   10
      CrossPosition   =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ChangeOnMouseOver=   0
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      HotLine         =   -1  'True
      HotLineWidth    =   8
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLPbc 
      Height          =   795
      Index           =   1
      Left            =   300
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1402
      BackColor       =   12648447
      Border          =   -1  'True
      BorderColor     =   16711680
      ColorOnMouseOver=   255
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":01EC
      Caption2        =   "Form1.frx":020C
      Caption1PaddingX=   10
      Caption2PaddingX=   10
      Caption2PaddingY=   10
      CrossPosition   =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ChangeOnMouseOver=   0
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      HotLine         =   -1  'True
      HotLineWidth    =   8
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ChangeColorOnMouseOver"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   255
      TabIndex        =   4
      Top             =   150
      Width           =   2925
   End
   Begin AXLPCTRL.axLabelPlus axLPbc 
      Height          =   795
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1402
      BackColor       =   12648447
      Border          =   -1  'True
      BorderColor     =   16711680
      ColorOnMouseOver=   255
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":022C
      Caption2        =   "Form1.frx":024C
      Caption1PaddingX=   10
      Caption2PaddingX=   10
      Caption2PaddingY=   10
      CrossPosition   =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ChangeOnMouseOver=   0
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      HotLine         =   -1  'True
      HotLineWidth    =   8
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer


Private Sub axLPValue_Click()
axLPValue.Caption2 = axLPValue.Value
If axLPValue.Value = False Then
  axLPValue.HotLineColor = &HFF&
  axLPValue.BorderColor = &HFF&
Else
  axLPValue.HotLineColor = &H8000&
  axLPValue.BorderColor = &H8000&
End If
End Sub

Private Sub Check1_Click()
axLPccc.Gradient = Check1.Value
End Sub

Private Sub Check2_Click()
axLPCross.CrossVisible = Check2.Value
End Sub

Private Sub cmdGlowing_Click()
axLPGlow.Glowing = Not axLPGlow.Glowing
End Sub

Private Sub Command1_Click()
axLPCross.Visible = True
End Sub

Private Sub Form_Load()
With List1
  .AddItem "cTopRight"
  .AddItem "cMiddleRight"
  .AddItem "cBottomRight"
  .AddItem "cTopLeft"
  .AddItem "cMiddleLeft"
  .AddItem "cBottomLeft"
  .AddItem "cMiddleTop"
  .AddItem "cMiddleBottom"
End With

End Sub

Private Sub HScroll1_Change()
axLPdc.Caption1PaddingX = HScroll1.Value
X1.Text = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
axLPdc.Caption2PaddingX = HScroll2.Value
X2.Text = HScroll2.Value
End Sub

Private Sub List1_Click()
axLPCross.CrossPosition = List1.ListIndex
End Sub

Private Sub Option1_Click(Index As Integer)
For I = 0 To 2
  axLPbc(I).ChangeOnMouseOver = Index
Next I
End Sub

Private Sub Text1_Change()
axLPdc.Caption1 = Text1.Text
End Sub

Private Sub Text2_Change()
axLPdc.Caption2 = Text2.Text
End Sub

Private Sub VScroll1_Change()
axLPdc.Caption1PaddingY = VScroll1.Value
Y1.Text = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
axLPdc.Caption2PaddingY = VScroll2.Value
Y2.Text = VScroll2.Value
End Sub

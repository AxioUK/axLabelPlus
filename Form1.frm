VERSION 5.00
Object = "*\AaxLabelPlus.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "New Properties axLabelPlus (Mod Version of Great LabelPlus from Leandro Ascierto)"
   ClientHeight    =   6555
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
   ScaleHeight     =   6555
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   6450
      Top             =   5610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccionar Fuente"
   End
   Begin VB.CommandButton cmdFont2 
      Caption         =   "..."
      Height          =   330
      Left            =   10080
      TabIndex        =   49
      Top             =   6165
      Width           =   345
   End
   Begin VB.CommandButton cmfFont1 
      Caption         =   "..."
      Height          =   330
      Left            =   10080
      TabIndex        =   48
      Top             =   5745
      Width           =   345
   End
   Begin VB.TextBox txtFont1 
      BackColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   7875
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "Font1"
      Top             =   5745
      Width           =   2160
   End
   Begin VB.TextBox txtFont2 
      BackColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   7875
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "Font2"
      Top             =   6150
      Width           =   2160
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Shadow"
      Height          =   210
      Left            =   3120
      TabIndex        =   0
      Top             =   4905
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox SW 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4515
      TabIndex        =   5
      Text            =   "2"
      Top             =   4890
      Width           =   405
   End
   Begin VB.TextBox HW 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4515
      TabIndex        =   6
      Text            =   "7"
      Top             =   4590
      Width           =   405
   End
   Begin VB.CheckBox Check3 
      Caption         =   "HotLine Visible"
      Height          =   210
      Left            =   3120
      TabIndex        =   7
      Top             =   4605
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton cmdGlowing 
      Caption         =   "Glowing"
      Height          =   360
      Left            =   4575
      TabIndex        =   10
      Top             =   2310
      Width           =   1305
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   7875
      TabIndex        =   26
      Text            =   "axLabelPlus2"
      Top             =   5340
      Width           =   2160
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   7875
      TabIndex        =   35
      Text            =   "axLabelPlus1"
      Top             =   4935
      Width           =   2160
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   3150
      TabIndex        =   32
      Top             =   5190
      Width           =   1860
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Cross Visible"
      Height          =   210
      Left            =   1815
      TabIndex        =   31
      Top             =   4605
      Width           =   1200
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   210
      Left            =   7095
      Max             =   50
      TabIndex        =   25
      Top             =   2430
      Value           =   10
      Width           =   2625
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1725
      Left            =   6660
      Max             =   50
      TabIndex        =   24
      Top             =   3045
      Value           =   15
      Width           =   210
   End
   Begin VB.TextBox Y2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9060
      TabIndex        =   20
      Text            =   "15"
      Top             =   4485
      Width           =   405
   End
   Begin VB.TextBox Y1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9060
      TabIndex        =   18
      Text            =   "00"
      Top             =   4095
      Width           =   405
   End
   Begin VB.TextBox X2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8190
      TabIndex        =   16
      Text            =   "10"
      Top             =   4485
      Width           =   405
   End
   Begin VB.TextBox X1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8205
      TabIndex        =   14
      Text            =   "10"
      Top             =   4095
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
      Value           =   10
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
   Begin VB.CommandButton Command1 
      Caption         =   "Visible=True"
      Height          =   360
      Left            =   345
      TabIndex        =   30
      Top             =   4815
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font 2"
      Height          =   195
      Index           =   2
      Left            =   7275
      TabIndex        =   47
      Top             =   6225
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font 1"
      Height          =   195
      Index           =   2
      Left            =   7275
      TabIndex        =   46
      Top             =   5805
      Width           =   645
   End
   Begin AXLPCTRL.axLabelPlus axLPGlow 
      Height          =   705
      Left            =   3780
      TabIndex        =   43
      Top             =   2130
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1244
      Border          =   -1  'True
      BorderCornerLeftTop=   30
      BorderCornerRightTop=   30
      BorderCornerBottomRight=   30
      BorderCornerBottomLeft=   30
      BorderWidth     =   10
      Caption1        =   "Form1.frx":0000
      Caption2        =   "Form1.frx":0022
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Caption1WordWrap=   0   'False
      Caption2WordWrap=   0   'False
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
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
   Begin AXLPCTRL.axLabelPlus axLPdc 
      Height          =   1260
      Left            =   7005
      TabIndex        =   42
      Top             =   2730
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   2223
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":0044
      Caption2        =   "Form1.frx":007C
      Caption1PaddingX=   10
      Caption2PaddingX=   10
      Caption2PaddingY=   15
      CaptionShadow   =   -1  'True
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption1WordWrap=   0   'False
      Caption2WordWrap=   0   'False
      HotLine         =   -1  'True
      HotLineWidth    =   7
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61384
      IconPaddingX    =   15
      IconAlignmentH  =   2
      IconAlignmentV  =   1
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLPValue 
      Height          =   1125
      Left            =   6525
      TabIndex        =   41
      Top             =   495
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   1984
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":00B4
      Caption2        =   "Form1.frx":00EC
      Caption1PaddingX=   20
      Caption1PaddingY=   15
      Caption2PaddingX=   20
      Caption2PaddingY=   35
      CaptionShadow   =   -1  'True
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption1WordWrap=   0   'False
      Caption2WordWrap=   0   'False
      HotLine         =   -1  'True
      HotLineWidth    =   7
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61384
      IconPaddingX    =   15
      IconAlignmentH  =   2
      IconAlignmentV  =   1
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLPCross 
      Height          =   1125
      Left            =   240
      TabIndex        =   40
      Top             =   5250
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   1984
      BackShadow      =   -1  'True
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":0124
      Caption2        =   "Form1.frx":015C
      Caption1PaddingX=   42
      Caption1PaddingY=   15
      Caption2PaddingX=   20
      Caption2PaddingY=   35
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ChangeOnMouseOver=   0
      Caption1WordWrap=   0   'False
      Caption2WordWrap=   0   'False
      ShadowSize      =   2
      ShadowColor     =   12582912
      ShadowOffsetX   =   3
      ShadowOffsetY   =   3
      ShadowColorOpacity=   90
      HotLine         =   -1  'True
      HotLineWidth    =   7
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61384
      IconPaddingX    =   7
      IconAlignmentV  =   1
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLPccc 
      Height          =   810
      Left            =   270
      TabIndex        =   39
      Top             =   3420
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1429
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":0194
      Caption2        =   "Form1.frx":01CC
      Caption1PaddingX=   10
      Caption2PaddingX=   10
      Caption2PaddingY=   20
      CaptionShadow   =   -1  'True
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ChangeOnMouseOver=   0
      Caption1WordWrap=   0   'False
      Caption2WordWrap=   0   'False
      HotLine         =   -1  'True
      HotLineWidth    =   7
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61384
      IconPaddingX    =   10
      IconAlignmentH  =   2
      IconAlignmentV  =   1
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLabelPlus1 
      Height          =   810
      Index           =   2
      Left            =   255
      TabIndex        =   38
      Top             =   2175
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1429
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":0204
      Caption2        =   "Form1.frx":023C
      Caption1PaddingX=   10
      Caption2PaddingX=   10
      Caption2PaddingY=   20
      CaptionShadow   =   -1  'True
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ChangeOnMouseOver=   0
      Caption1WordWrap=   0   'False
      Caption2WordWrap=   0   'False
      HotLine         =   -1  'True
      HotLineWidth    =   7
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61384
      IconPaddingX    =   10
      IconAlignmentH  =   2
      IconAlignmentV  =   1
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLabelPlus1 
      Height          =   810
      Index           =   1
      Left            =   255
      TabIndex        =   37
      Top             =   1320
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1429
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":0274
      Caption2        =   "Form1.frx":02AC
      Caption1PaddingX=   10
      Caption2PaddingX=   10
      Caption2PaddingY=   20
      CaptionShadow   =   -1  'True
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ChangeOnMouseOver=   0
      Caption1WordWrap=   0   'False
      Caption2WordWrap=   0   'False
      HotLine         =   -1  'True
      HotLineWidth    =   7
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61384
      IconPaddingX    =   10
      IconAlignmentH  =   2
      IconAlignmentV  =   1
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLabelPlus1 
      Height          =   810
      Index           =   0
      Left            =   255
      TabIndex        =   36
      Top             =   465
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1429
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":02E4
      Caption2        =   "Form1.frx":031C
      Caption1PaddingX=   10
      Caption2PaddingX=   10
      Caption2PaddingY=   20
      CaptionShadow   =   -1  'True
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ChangeOnMouseOver=   0
      Caption1WordWrap=   0   'False
      Caption2WordWrap=   0   'False
      HotLine         =   -1  'True
      HotLineWidth    =   7
      HotLinePosition =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61384
      IconPaddingX    =   10
      IconAlignmentH  =   2
      IconAlignmentV  =   1
      PictureArr      =   0
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
      TabIndex        =   29
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption1"
      Height          =   195
      Index           =   1
      Left            =   7140
      TabIndex        =   34
      Top             =   4995
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
      Top             =   5415
      Width           =   645
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption2"
      Height          =   195
      Index           =   0
      Left            =   7170
      TabIndex        =   23
      Top             =   4515
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
      Top             =   4155
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y :"
      Height          =   195
      Left            =   8820
      TabIndex        =   21
      Top             =   4530
      Width           =   195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y :"
      Height          =   195
      Left            =   8820
      TabIndex        =   19
      Top             =   4140
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X :"
      Height          =   195
      Left            =   7950
      TabIndex        =   17
      Top             =   4530
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X :"
      Height          =   195
      Left            =   7965
      TabIndex        =   15
      Top             =   4140
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I As Integer

Dim mFont As StdFont

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

Private Sub Check3_Click()
axLPCross.HotLine = Check3.Value
End Sub

Private Sub Check4_Click()
axLPCross.BackShadow = Check4.Value
End Sub

Private Sub cmdFont2_Click()

With cmDlg
  .DialogTitle = "Seleccionar Fuente Caption2"
  .ShowFont
  txtFont2.Text = .FontName
  mFont.Name = .FontName
  mFont.Bold = .FontBold
  mFont.Italic = .FontItalic
  mFont.Size = .FontSize
  Set axLPdc.Caption2Font = mFont
End With
End Sub

Private Sub cmdGlowing_Click()
axLPGlow.Glowing = Not axLPGlow.Glowing
End Sub

Private Sub cmfFont1_Click()
With cmDlg
  .DialogTitle = "Seleccionar Fuente Caption1"
  .ShowFont
  txtFont1.Text = .FontName
  mFont.Name = .FontName
  mFont.Bold = .FontBold
  mFont.Italic = .FontItalic
  mFont.Size = .FontSize
  Set axLPdc.Caption1Font = mFont
End With
End Sub

Private Sub Command1_Click()
axLPCross.Visible = True
End Sub

Private Sub Form_Load()

Set mFont = New StdFont

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

Private Sub HW_Change()
On Error Resume Next
axLPCross.HotLineWidth = CInt(HW.Text)
End Sub

Private Sub List1_Click()
axLPCross.CrossPosition = List1.ListIndex
End Sub

Private Sub Option1_Click(Index As Integer)
For I = 0 To 2
  axLabelPlus1(I).ChangeOnMouseOver = Index
Next I
End Sub

Private Sub SW_Change()
On Error Resume Next
axLPCross.ShadowSize = CInt(SW.Text)
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


VERSION 5.00
Object = "*\AaxLabelPlus.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   ClientHeight    =   8145
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
   ScaleHeight     =   8145
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Form2 Sample"
      Height          =   360
      Left            =   5055
      TabIndex        =   68
      Top             =   7410
      Width           =   1485
   End
   Begin VB.TextBox OP2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8835
      TabIndex        =   65
      Text            =   "50"
      Top             =   7005
      Width           =   405
   End
   Begin VB.TextBox OP1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8835
      TabIndex        =   64
      Text            =   "50"
      Top             =   6600
      Width           =   405
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   330
      Left            =   7890
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   7005
      Width           =   345
   End
   Begin VB.PictureBox Picture1 
      Height          =   330
      Left            =   7890
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   6570
      Width           =   345
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   330
      Left            =   8325
      TabIndex        =   59
      Top             =   6585
      Width           =   345
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   330
      Left            =   8325
      TabIndex        =   58
      Top             =   7005
      Width           =   345
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeCaptionHotLine"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   3555
      TabIndex        =   57
      Top             =   2850
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeCaptionBorder"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   3555
      TabIndex        =   56
      Top             =   2604
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeCaptionIcon"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   3555
      TabIndex        =   55
      Top             =   2358
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeIconBorder"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   3555
      TabIndex        =   54
      Top             =   2112
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeIconOnly"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   3555
      TabIndex        =   53
      Top             =   1866
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeCaptions"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   3555
      TabIndex        =   52
      Top             =   1620
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeCaption2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3555
      TabIndex        =   51
      Top             =   1374
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeCaption1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   3555
      TabIndex        =   50
      Top             =   1128
      Width           =   2040
   End
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
      ForeColor       =   &H00404040&
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
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   7875
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "Font2"
      Top             =   6150
      Width           =   2160
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00000000&
      Caption         =   "Shadow"
      ForeColor       =   &H00FFFFFF&
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
      BackColor       =   &H00000000&
      Caption         =   "HotLine Visible"
      ForeColor       =   &H00FFFFFF&
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
      Left            =   2190
      TabIndex        =   10
      Top             =   7440
      Width           =   1305
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   7875
      TabIndex        =   26
      Text            =   "axLabelPlus2"
      Top             =   5340
      Width           =   2160
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
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
      BackColor       =   &H00000000&
      Caption         =   "Cross Visible"
      ForeColor       =   &H00FFFFFF&
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
      Value           =   20
      Width           =   210
   End
   Begin VB.TextBox Y2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9060
      TabIndex        =   20
      Text            =   "20"
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
      BackColor       =   &H00000000&
      Caption         =   "Gradient"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3360
      TabIndex        =   9
      Top             =   3615
      Width           =   1950
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeHotlineColor"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   3555
      TabIndex        =   3
      Top             =   882
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeBorderColor"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   3555
      TabIndex        =   2
      Top             =   636
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeNone"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3555
      TabIndex        =   1
      Top             =   390
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
      Caption         =   "% Opacity"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   9270
      TabIndex        =   67
      Top             =   7050
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Opacity"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   9270
      TabIndex        =   66
      Top             =   6645
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   7275
      TabIndex        =   61
      Top             =   6645
      Width           =   645
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   7275
      TabIndex        =   60
      Top             =   7065
      Width           =   645
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font 2"
      ForeColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   7275
      TabIndex        =   46
      Top             =   5805
      Width           =   645
   End
   Begin AXLPCTRL.axLabelPlus axLPGlow 
      Height          =   705
      Left            =   525
      TabIndex        =   43
      Top             =   7080
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   1244
      BackColor       =   -2147483633
      Border          =   -1  'True
      BorderColor     =   255
      BorderColorOpacity=   75
      BorderCornerLeftTop=   30
      BorderCornerRightTop=   30
      BorderCornerBottomRight=   30
      BorderCornerBottomLeft=   30
      BorderWidth     =   15
      Caption1        =   "Form1.frx":0000
      Caption2        =   "Form1.frx":0020
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
      Caption1        =   "Form1.frx":0040
      Caption2        =   "Form1.frx":0078
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
      Caption1ForeColor=   16777215
      Caption1ForeColorOpacity=   50
      Caption2ForeColor=   65535
      Caption2ForeColorOpacity=   50
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
      IconForeColor   =   16777215
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
      Caption1        =   "Form1.frx":00B0
      Caption2        =   "Form1.frx":00E8
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
      Caption1ForeColor=   16777215
      Caption2ForeColor=   16777215
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
      IconForeColor   =   192
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
      BackColor       =   -2147483633
      BackShadow      =   -1  'True
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":0120
      Caption2        =   "Form1.frx":0158
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
      BackColor       =   -2147483633
      BackColorPress  =   8421504
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":0190
      Caption2        =   "Form1.frx":01C8
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
      ChangeColorOnClick=   -1  'True
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
      BackColor       =   12648384
      Border          =   -1  'True
      BorderColor     =   16448
      ColorOnMouseOver=   65280
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":0200
      Caption2        =   "Form1.frx":0238
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
      BackColor       =   12648384
      Border          =   -1  'True
      BorderColor     =   16448
      ColorOnMouseOver=   33023
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":0270
      Caption2        =   "Form1.frx":02A8
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
      BackColor       =   12648384
      Border          =   -1  'True
      BorderColor     =   16448
      ColorOnMouseOver=   16711935
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":02E0
      Caption2        =   "Form1.frx":0318
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
      BackColor       =   &H00000000&
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   480
      TabIndex        =   29
      Top             =   6765
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1050
      Left            =   375
      Top             =   6900
      Width           =   3360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption1"
      ForeColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
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
      ForeColor       =   &H0000FF00&
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
      ForeColor       =   &H0000C000&
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
      ForeColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
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
      ForeColor       =   &H0000FFFF&
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
      ForeColor       =   &H0000FFFF&
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
      ForeColor       =   &H0000FFFF&
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
      ForeColor       =   &H0000FFFF&
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
      ForeColor       =   &H0000C000&
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
      ForeColor       =   &H0000FF00&
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
      ForeColor       =   &H0000FF00&
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

Private Sub Command2_Click()
With cmDlg
  .DialogTitle = "Seleccionar Color Caption2"
  .ShowColor
  Picture2.BackColor = .Color
  axLPdc.Caption2Forecolor = .Color
End With
End Sub

Private Sub Command3_Click()
With cmDlg
  .DialogTitle = "Seleccionar Color Caption1"
  .ShowColor
  Picture1.BackColor = .Color
  axLPdc.Caption1Forecolor = .Color
End With
End Sub

Private Sub Command4_Click()
Form2.Show
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

Me.Caption = "AxLabelPlus v" & axLabelPlus1(0).Version & " - New Properties (Mod Version of Great LabelPlus from Leandro Ascierto)"
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

Private Sub OP1_Change()
axLPdc.Caption1ForeColorOpacity = CInt(OP1.Text)
End Sub

Private Sub OP2_Change()
axLPdc.Caption2ForeColorOpacity = CInt(OP2.Text)
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


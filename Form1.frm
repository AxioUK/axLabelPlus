VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{D6F84FAD-6738-419D-846A-64AC9AD4766C}#4.0#0"; "axLabelPlus.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14265
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
   ScaleHeight     =   8805
   ScaleWidth      =   14265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3045
      Left            =   10695
      TabIndex        =   78
      Top             =   1440
      Width           =   3285
      Begin VB.OptionButton opPictureEffect 
         BackColor       =   &H80000004&
         Caption         =   "eIncreaseOpacity"
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   0
         Left            =   555
         TabIndex        =   80
         Top             =   2190
         Value           =   -1  'True
         Width           =   2220
      End
      Begin VB.OptionButton opPictureEffect 
         BackColor       =   &H80000004&
         Caption         =   "eAlternateGrayColor"
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   1
         Left            =   555
         TabIndex        =   79
         Top             =   2490
         Width           =   2220
      End
      Begin AXLPCTRL.axLabelPlus axLabelPlus2 
         Height          =   1785
         Left            =   495
         TabIndex        =   81
         Top             =   315
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   3149
         BackColor       =   14737632
         Caption1        =   "Form1.frx":0000
         Caption2        =   "Form1.frx":0042
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
         ChangeOnMouseOver=   11
         GradientColorP1 =   0
         GradientColorP1Opacity=   0
         GradientColorP2 =   0
         GradientColorP2Opacity=   0
         PictureAlignmentH=   1
         PictureAlignmentV=   1
         PictureOpacity  =   50
         ShadowColorOpacity=   0
         CallOutAlign    =   0
         CallOutWidth    =   0
         CallOutLen      =   0
         PictureColor    =   12648384
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
         GlowSpeed       =   0
         GlowColor       =   0
         GlowTiks        =   0
         PicturePresent  =   -1  'True
         PictureArr      =   "Form1.frx":0084
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7770
      Left            =   10590
      ScaleHeight     =   7740
      ScaleWidth      =   15
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   195
      Width           =   45
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7770
      Left            =   5685
      ScaleHeight     =   7740
      ScaleWidth      =   15
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   135
      Width           =   45
   End
   Begin VB.TextBox txtTiks 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1125
      TabIndex        =   40
      Text            =   "10"
      Top             =   7170
      Width           =   345
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   510
      Left            =   285
      TabIndex        =   41
      Top             =   7860
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   900
      _Version        =   393216
      Min             =   1
      Max             =   100
      SelStart        =   50
      TickStyle       =   1
      Value           =   50
      TextPosition    =   1
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load Picture"
      Height          =   360
      Left            =   11175
      TabIndex        =   42
      Top             =   960
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   180
      Left            =   6075
      TabIndex        =   68
      Top             =   2100
      Width           =   4230
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Form2 Sample"
      Height          =   360
      Left            =   8760
      TabIndex        =   67
      Top             =   90
      Width           =   1485
   End
   Begin VB.TextBox OP2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8655
      TabIndex        =   64
      Text            =   "50"
      Top             =   6870
      Width           =   405
   End
   Begin VB.TextBox OP1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8655
      TabIndex        =   63
      Text            =   "50"
      Top             =   6555
      Width           =   405
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C00000&
      Height          =   330
      Left            =   7710
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   6870
      Width           =   345
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C000&
      Height          =   330
      Left            =   7710
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   6525
      Width           =   345
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   330
      Left            =   8145
      TabIndex        =   58
      Top             =   6525
      Width           =   345
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   330
      Left            =   8145
      TabIndex        =   57
      Top             =   6885
      Width           =   345
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeCaptionHotLine"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   3360
      TabIndex        =   56
      Top             =   2925
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeCaptionBorder"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   3360
      TabIndex        =   55
      Top             =   2685
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeCaptionIcon"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   3360
      TabIndex        =   54
      Top             =   2430
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeIconBorder"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   3360
      TabIndex        =   53
      Top             =   2190
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeIconOnly"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   3360
      TabIndex        =   52
      Top             =   1935
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeCaptions"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   3360
      TabIndex        =   51
      Top             =   1695
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeCaption2"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   3360
      TabIndex        =   50
      Top             =   1455
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeCaption1"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   3360
      TabIndex        =   49
      Top             =   1200
      Width           =   2040
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   6270
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccionar Fuente"
   End
   Begin VB.CommandButton cmdFont2 
      Caption         =   "..."
      Height          =   300
      Left            =   9885
      TabIndex        =   48
      Top             =   6150
      Width           =   315
   End
   Begin VB.CommandButton cmfFont1 
      Caption         =   "..."
      Height          =   300
      Left            =   9885
      TabIndex        =   47
      Top             =   5835
      Width           =   315
   End
   Begin VB.TextBox txtFont1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   7695
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "Verdana"
      Top             =   5835
      Width           =   2160
   End
   Begin VB.TextBox txtFont2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   7695
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "Tahoma"
      Top             =   6150
      Width           =   2160
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Shadow"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3120
      TabIndex        =   0
      Top             =   4905
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3120
      TabIndex        =   7
      Top             =   4605
      Width           =   1335
   End
   Begin VB.CommandButton cmdGlowing 
      Caption         =   "Glowing"
      Height          =   465
      Left            =   1665
      TabIndex        =   10
      Top             =   7260
      Width           =   765
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   7695
      TabIndex        =   26
      Text            =   "axLabelPlus2"
      Top             =   5475
      Width           =   2160
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   7695
      TabIndex        =   35
      Text            =   "axLabelPlus1"
      Top             =   5160
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1815
      TabIndex        =   31
      Top             =   4605
      Width           =   1200
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   210
      Left            =   6915
      Max             =   50
      TabIndex        =   25
      Top             =   2880
      Value           =   10
      Width           =   2625
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1215
      Left            =   6555
      Max             =   50
      TabIndex        =   24
      Top             =   3150
      Value           =   20
      Width           =   210
   End
   Begin VB.TextBox Y2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8880
      TabIndex        =   20
      Text            =   "20"
      Top             =   4800
      Width           =   405
   End
   Begin VB.TextBox Y1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8880
      TabIndex        =   18
      Text            =   "5"
      Top             =   4485
      Width           =   405
   End
   Begin VB.TextBox X2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8010
      TabIndex        =   16
      Text            =   "10"
      Top             =   4800
      Width           =   405
   End
   Begin VB.TextBox X1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8010
      TabIndex        =   14
      Text            =   "7"
      Top             =   4485
      Width           =   405
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   6315
      Max             =   50
      TabIndex        =   13
      Top             =   3150
      Value           =   5
      Width           =   210
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   210
      Left            =   6915
      Max             =   50
      TabIndex        =   12
      Top             =   2640
      Value           =   7
      Width           =   2625
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Gradient"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3360
      TabIndex        =   9
      Top             =   3615
      Width           =   1950
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeHotlineColor"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeBorderColor"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   705
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Caption         =   "eChangeNone"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   465
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "If GlowTiks value is set=0 then Glowing indefinitely"
      Height          =   195
      Left            =   240
      TabIndex        =   85
      Top             =   8445
      Width           =   3645
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GlowSpeed"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   285
      TabIndex        =   84
      Top             =   7605
      Width           =   795
   End
   Begin AXLPCTRL.axLabelPlus axLPGlow 
      Height          =   465
      Index           =   2
      Left            =   3315
      TabIndex        =   83
      Top             =   7395
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   820
      BackColor       =   8421504
      BackColorOpacity=   50
      BackColorPress  =   8421504
      BackColorPressOpacity=   50
      Border          =   -1  'True
      BorderColor     =   65280
      BorderColorOpacity=   0
      BorderCornerLeftTop=   20
      BorderCornerRightTop=   20
      BorderCornerBottomRight=   20
      BorderCornerBottomLeft=   20
      BorderWidth     =   10
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption1        =   "Form1.frx":289F
      Caption2        =   "Form1.frx":28C1
      Caption2PaddingX=   5
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
      ChangeColorOnClick=   -1  'True
      ChangeOnMouseOver=   0
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
      IconPaddingX    =   10
      IconAlignmentH  =   2
      IconAlignmentV  =   1
      GlowTiks        =   0
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLPGlow 
      Height          =   585
      Index           =   1
      Left            =   2955
      TabIndex        =   82
      Top             =   6750
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1032
      BackColor       =   255
      BackColorOpacity=   50
      BackColorPress  =   8421504
      BackColorPressOpacity=   50
      Border          =   -1  'True
      BorderColor     =   65535
      BorderColorOpacity=   0
      BorderCornerLeftTop=   20
      BorderCornerRightTop=   20
      BorderCornerBottomRight=   20
      BorderCornerBottomLeft=   20
      BorderWidth     =   10
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption1        =   "Form1.frx":28E1
      Caption2        =   "Form1.frx":2903
      Caption2PaddingX=   5
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
      Caption1ForeColor=   65535
      ChangeColorOnClick=   -1  'True
      ChangeOnMouseOver=   0
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
      IconPaddingX    =   10
      IconAlignmentH  =   2
      IconAlignmentV  =   1
      GlowTiks        =   0
      PictureArr      =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ChangeOnMouseOver PictureEffects"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   510
      Index           =   5
      Left            =   10710
      TabIndex        =   77
      Top             =   135
      Width           =   3090
      WordWrap        =   -1  'True
   End
   Begin AXLPCTRL.axLabelPlus axLPValue 
      Height          =   960
      Index           =   0
      Left            =   6525
      TabIndex        =   74
      Top             =   990
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1693
      BackColor       =   12648447
      BackColorOpacity=   90
      BackColorPress  =   128
      BackColorPressOpacity=   90
      Border          =   -1  'True
      BorderColor     =   65280
      BorderColorOpacity=   90
      ColorOnMouseOver=   12632256
      ColorOpacityOnMouseOver=   90
      BorderCornerLeftTop=   8
      BorderCornerRightTop=   8
      BorderCornerBottomRight=   8
      BorderCornerBottomLeft=   8
      BorderWidth     =   2
      Caption1        =   "Form1.frx":2923
      Caption2        =   "Form1.frx":2943
      Caption1PaddingX=   7
      Caption1PaddingY=   5
      Caption2PaddingX=   7
      Caption2PaddingY=   20
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
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      HotLine         =   -1  'True
      HotLineColor    =   255
      HotLineWidth    =   15
      HotLinePosition =   0
      OptionBehavior  =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GlowSpeed       =   0
      GlowColor       =   0
      GlowTiks        =   0
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLPValue 
      Height          =   960
      Index           =   1
      Left            =   8430
      TabIndex        =   73
      Top             =   1005
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1693
      BackColor       =   12648447
      BackColorOpacity=   90
      BackColorPress  =   128
      BackColorPressOpacity=   90
      Border          =   -1  'True
      BorderColor     =   65280
      BorderColorOpacity=   90
      ColorOnMouseOver=   12632256
      ColorOpacityOnMouseOver=   90
      BorderCornerLeftTop=   8
      BorderCornerRightTop=   8
      BorderCornerBottomRight=   8
      BorderCornerBottomLeft=   8
      BorderWidth     =   2
      Caption1        =   "Form1.frx":2965
      Caption2        =   "Form1.frx":2985
      Caption1PaddingX=   7
      Caption1PaddingY=   5
      Caption2PaddingX=   7
      Caption2PaddingY=   20
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
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      HotLine         =   -1  'True
      HotLineColor    =   49152
      HotLineWidth    =   15
      HotLinePosition =   0
      Value           =   -1  'True
      OptionBehavior  =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GlowSpeed       =   0
      GlowColor       =   0
      GlowTiks        =   0
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLPCross 
      Height          =   810
      Left            =   195
      TabIndex        =   72
      Top             =   5310
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1429
      BackColorPress  =   8421504
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":29A7
      Caption2        =   "Form1.frx":29DF
      Caption1PaddingX=   10
      Caption2PaddingX=   10
      Caption2PaddingY=   20
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ChangeColorOnClick=   -1  'True
      ChangeOnMouseOver=   0
      ShadowSize      =   10
      ShadowColor     =   8388736
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
   Begin AXLPCTRL.axLabelPlus axLPdc 
      Height          =   1140
      Left            =   6900
      TabIndex        =   71
      Top             =   3225
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   2011
      BackColor       =   14737632
      BackColorOpacity=   90
      BackColorPress  =   8421504
      BackColorPressOpacity=   90
      Border          =   -1  'True
      BorderColor     =   8421504
      BorderColorOpacity=   90
      ColorOnMouseOver=   12632256
      ColorOpacityOnMouseOver=   90
      BorderWidth     =   2
      Caption1        =   "Form1.frx":2A17
      Caption2        =   "Form1.frx":2A59
      Caption1PaddingX=   7
      Caption1PaddingY=   5
      Caption2PaddingX=   10
      Caption2PaddingY=   20
      CaptionShowPrefix=   -1  'True
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      Caption1ForeColor=   49152
      Caption1ForeColorOpacity=   50
      Caption2ForeColor=   12582912
      Caption2ForeColorOpacity=   50
      ChangeOnMouseOver=   0
      GradientColorP1 =   0
      GradientColorP1Opacity=   0
      GradientColorP2 =   0
      GradientColorP2Opacity=   0
      ShadowColorOpacity=   0
      CallOutAlign    =   3
      CallOutWidth    =   3
      CallOutLen      =   5
      CallOut         =   -1  'True
      CallOutCustomPosition=   5
      CallOutRightTriangle=   -1  'True
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61170
      IconPaddingY    =   32
      IconAlignmentH  =   1
      GlowSpeed       =   0
      GlowColor       =   0
      GlowTiks        =   0
      PictureArr      =   0
   End
   Begin AXLPCTRL.axLabelPlus axLPGlow 
      Height          =   690
      Index           =   0
      Left            =   3825
      TabIndex        =   70
      Top             =   6735
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1217
      BackColorOpacity=   50
      BackColorPress  =   8421504
      BackColorPressOpacity=   50
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderColorOpacity=   0
      BorderCornerLeftTop=   20
      BorderCornerRightTop=   20
      BorderCornerBottomRight=   20
      BorderCornerBottomLeft=   20
      BorderWidth     =   10
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption1        =   "Form1.frx":2A9F
      Caption2        =   "Form1.frx":2AC1
      Caption2PaddingX=   5
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
      Caption1ForeColor=   192
      ChangeColorOnClick=   -1  'True
      ChangeOnMouseOver=   0
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
      IconPaddingX    =   10
      IconAlignmentH  =   2
      IconAlignmentV  =   1
      GlowTiks        =   0
      PictureArr      =   0
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GlowTiks"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   285
      TabIndex        =   69
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Opacity"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   9090
      TabIndex        =   66
      Top             =   6915
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Opacity"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   9090
      TabIndex        =   65
      Top             =   6600
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color 1"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   7095
      TabIndex        =   60
      Top             =   6600
      Width           =   645
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color 2"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   7095
      TabIndex        =   59
      Top             =   6930
      Width           =   645
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font 2"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   7095
      TabIndex        =   46
      Top             =   6225
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font 1"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   7095
      TabIndex        =   45
      Top             =   5895
      Width           =   645
   End
   Begin AXLPCTRL.axLabelPlus axLPccc 
      Height          =   810
      Left            =   270
      TabIndex        =   39
      Top             =   3420
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1429
      BackColorPress  =   8421504
      Border          =   -1  'True
      BorderColor     =   16711680
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   2
      Caption1        =   "Form1.frx":2AE1
      Caption2        =   "Form1.frx":2B19
      Caption1PaddingX=   5
      Caption2PaddingX=   5
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
      Caption1        =   "Form1.frx":2B51
      Caption2        =   "Form1.frx":2B89
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
      Caption1        =   "Form1.frx":2BC1
      Caption2        =   "Form1.frx":2BF9
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
      Caption1        =   "Form1.frx":2C31
      Caption2        =   "Form1.frx":2C69
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
      BackStyle       =   0  'Transparent
      Caption         =   " Glowing "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   210
      TabIndex        =   29
      Top             =   6465
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption1"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   6960
      TabIndex        =   34
      Top             =   5220
      Width           =   645
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption2"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   6960
      TabIndex        =   33
      Top             =   5550
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CrossClose"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   4
      Left            =   210
      TabIndex        =   28
      Top             =   4485
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value [OptionBehavior=TRUE]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   3
      Left            =   6105
      TabIndex        =   27
      Top             =   555
      Width           =   4050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption2"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   6990
      TabIndex        =   23
      Top             =   4830
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption1"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   6990
      TabIndex        =   22
      Top             =   4545
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8640
      TabIndex        =   21
      Top             =   4845
      Width           =   195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8640
      TabIndex        =   19
      Top             =   4530
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7770
      TabIndex        =   17
      Top             =   4845
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7785
      TabIndex        =   15
      Top             =   4530
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DualCaption"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   2
      Left            =   6105
      TabIndex        =   11
      Top             =   2310
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ChangeColorOnClick"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   1
      Left            =   150
      TabIndex        =   8
      Top             =   3075
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ChangeOnMouseOver"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   90
      Width           =   2880
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


Private Sub axLabelPlus2_MouseEnter()
'If axLabelPlus2.IsMouseInExtender Then MsgBox "Aguaita!, 'tas sobre el labelPlus!"
End Sub

Private Sub axLPValue_ChangeValue(Index As Integer, Value As Boolean)
axLPValue(Index).Caption2 = axLPValue(Index).Value
If axLPValue(Index).Value = False Then
  axLPValue(Index).HotLineColor = &HFF&
  axLPValue(Index).BorderColor = &HFF&
Else
  axLPValue(Index).HotLineColor = &H8000&
  axLPValue(Index).BorderColor = &H8000&
End If
End Sub

Private Sub axLPValue_Click(Index As Integer)
axLPValue(Index).Caption2 = axLPValue(Index).Value
If axLPValue(Index).Value = False Then
  axLPValue(Index).HotLineColor = &HFF&
  axLPValue(Index).BorderColor = &HFF&
Else
  axLPValue(Index).HotLineColor = &H8000&
  axLPValue(Index).BorderColor = &H8000&
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
axLPCross.Shadow = Check4.Value
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
axLPGlow(0).GlowTiks = CInt(txtTiks.Text)
axLPGlow(0).GlowSpeed = Slider1.Value
axLPGlow(0).Glowing = Not axLPGlow(0).Glowing
axLPGlow(1).GlowTiks = CInt(txtTiks.Text)
axLPGlow(1).GlowSpeed = Slider1.Value
axLPGlow(1).Glowing = Not axLPGlow(1).Glowing
axLPGlow(2).GlowTiks = CInt(txtTiks.Text)
axLPGlow(2).GlowSpeed = Slider1.Value
axLPGlow(2).Glowing = Not axLPGlow(2).Glowing
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

Private Sub Command5_Click()
On Error Resume Next
With cmDlg
  .DialogTitle = "Seleccionar Imagen"
  .Filter = "Pictures|*.bmp;*.gif;*.jpg;*.jpeg;*.png;*.dib;*.rle;*.jpe;*.jfif;*.emf;*.wmf;*.tif;*.tiff;*.ico;*.cur"
  .ShowOpen
  axLabelPlus2.LoadImagefromPath .FileName
End With

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

axLPGlow(0).GlowSpeed = CInt(txtTiks.Text)
axLPGlow(1).GlowSpeed = CInt(txtTiks.Text)
axLPGlow(2).GlowSpeed = CInt(txtTiks.Text)

With axLPdc
  .Caption1 = Text1.Text
  .Caption2 = Text2.Text
End With


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If axLabelPlus2.IsMouseInExtender Then MsgBox "Aguaita!, 'tas sobre el labelPlus!"
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

Private Sub opPictureEffect_Click(Index As Integer)
With axLabelPlus2
  .ChangeOnMouseOver = eChangePictureEffects
  .PictureEffectMouseOver = Index
End With
End Sub

Private Sub Option1_Click(Index As Integer)
For I = 0 To 2
  axLabelPlus1(I).ChangeOnMouseOver = Index
Next I
End Sub

Private Sub Slider1_Click()
axLPGlow(0).GlowSpeed = Slider1.Value
axLPGlow(1).GlowSpeed = Slider1.Value
axLPGlow(2).GlowSpeed = Slider1.Value
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


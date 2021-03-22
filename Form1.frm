VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{D6F84FAD-6738-419D-846A-64AC9AD4766C}#3.1#0"; "axLabelPlusX.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   ClientHeight    =   8805
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
   ScaleHeight     =   8805
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTiks 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      TabIndex        =   74
      Text            =   "10"
      Top             =   7155
      Width           =   345
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   510
      Left            =   255
      TabIndex        =   73
      Top             =   7515
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
      Left            =   5010
      TabIndex        =   71
      Top             =   6900
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1440
      Left            =   6300
      TabIndex        =   68
      Top             =   735
      Width           =   4230
      Begin AXLPCTRL.axLabelPlus axLPValue 
         Height          =   915
         Index           =   1
         Left            =   2205
         TabIndex        =   70
         Top             =   315
         Width           =   1860
         _extentx        =   3281
         _extenty        =   1614
         border          =   -1
         bordercolor     =   16711680
         bordercornerlefttop=   5
         bordercornerrighttop=   5
         bordercornerbottomright=   5
         bordercornerbottomleft=   5
         borderwidth     =   2
         caption1        =   "Form1.frx":0000
         caption2        =   "Form1.frx":0020
         caption1paddingx=   20
         caption1paddingy=   15
         caption2paddingx=   20
         caption2paddingy=   20
         captionshadow   =   -1
         caption1font    =   "Form1.frx":0042
         caption2font    =   "Form1.frx":006A
         caption1forecolor=   16777215
         caption2forecolor=   16777215
         hotline         =   -1
         hotlinewidth    =   7
         hotlineposition =   0
         optionbehavior  =   -1
         iconfont        =   "Form1.frx":009A
         iconcharcode    =   61384
         iconforecolor   =   192
         iconpaddingx    =   5
         iconalignmenth  =   2
         iconalignmentv  =   1
      End
      Begin AXLPCTRL.axLabelPlus axLPValue 
         Height          =   915
         Index           =   0
         Left            =   210
         TabIndex        =   69
         Top             =   315
         Width           =   1860
         _extentx        =   3281
         _extenty        =   1614
         border          =   -1
         bordercolor     =   16711680
         bordercornerlefttop=   5
         bordercornerrighttop=   5
         bordercornerbottomright=   5
         bordercornerbottomleft=   5
         borderwidth     =   2
         caption1        =   "Form1.frx":00C2
         caption2        =   "Form1.frx":00E2
         caption1paddingx=   20
         caption1paddingy=   15
         caption2paddingx=   20
         caption2paddingy=   20
         captionshadow   =   -1
         caption1font    =   "Form1.frx":0104
         caption2font    =   "Form1.frx":012C
         caption1forecolor=   16777215
         caption2forecolor=   16777215
         hotline         =   -1
         hotlinewidth    =   7
         hotlineposition =   0
         optionbehavior  =   -1
         iconfont        =   "Form1.frx":015C
         iconcharcode    =   61384
         iconforecolor   =   192
         iconpaddingx    =   5
         iconalignmenth  =   2
         iconalignmentv  =   1
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Form2 Sample"
      Height          =   360
      Left            =   8940
      TabIndex        =   67
      Top             =   90
      Width           =   1485
   End
   Begin VB.TextBox OP2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8835
      TabIndex        =   64
      Text            =   "50"
      Top             =   6870
      Width           =   405
   End
   Begin VB.TextBox OP1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8835
      TabIndex        =   63
      Text            =   "50"
      Top             =   6555
      Width           =   405
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   330
      Left            =   7890
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   6870
      Width           =   345
   End
   Begin VB.PictureBox Picture1 
      Height          =   330
      Left            =   7890
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
      Left            =   8325
      TabIndex        =   58
      Top             =   6525
      Width           =   345
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   330
      Left            =   8325
      TabIndex        =   57
      Top             =   6885
      Width           =   345
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "eChangeCaptionHotLine"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   3555
      TabIndex        =   56
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
      TabIndex        =   55
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
      TabIndex        =   54
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
      TabIndex        =   53
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
      TabIndex        =   52
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
      TabIndex        =   51
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
      TabIndex        =   50
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
      TabIndex        =   49
      Top             =   1128
      Width           =   2040
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   6450
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccionar Fuente"
   End
   Begin VB.CommandButton cmdFont2 
      Caption         =   "..."
      Height          =   300
      Left            =   10065
      TabIndex        =   48
      Top             =   6150
      Width           =   315
   End
   Begin VB.CommandButton cmfFont1 
      Caption         =   "..."
      Height          =   300
      Left            =   10065
      TabIndex        =   47
      Top             =   5835
      Width           =   315
   End
   Begin VB.TextBox txtFont1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   7875
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "Font1"
      Top             =   5835
      Width           =   2160
   End
   Begin VB.TextBox txtFont2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   7875
      Locked          =   -1  'True
      TabIndex        =   43
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
      Left            =   1530
      TabIndex        =   10
      Top             =   7065
      Width           =   855
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   7875
      TabIndex        =   26
      Text            =   "axLabelPlus2"
      Top             =   5475
      Width           =   2160
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   7875
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
      Top             =   2880
      Value           =   10
      Width           =   2625
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1215
      Left            =   6735
      Max             =   50
      TabIndex        =   24
      Top             =   3150
      Value           =   20
      Width           =   210
   End
   Begin VB.TextBox Y2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9060
      TabIndex        =   20
      Text            =   "20"
      Top             =   4800
      Width           =   405
   End
   Begin VB.TextBox Y1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9060
      TabIndex        =   18
      Text            =   "00"
      Top             =   4485
      Width           =   405
   End
   Begin VB.TextBox X2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8190
      TabIndex        =   16
      Text            =   "10"
      Top             =   4800
      Width           =   405
   End
   Begin VB.TextBox X1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8190
      TabIndex        =   14
      Text            =   "10"
      Top             =   4485
      Width           =   405
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   6495
      Max             =   50
      TabIndex        =   13
      Top             =   3150
      Width           =   210
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   210
      Left            =   7095
      Max             =   50
      TabIndex        =   12
      Top             =   2640
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tiks"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   285
      TabIndex        =   75
      Top             =   7200
      Width           =   270
   End
   Begin AXLPCTRL.axLabelPlus axLPGlow 
      Height          =   495
      Left            =   2685
      TabIndex        =   42
      Top             =   7290
      Width           =   495
      _extentx        =   873
      _extenty        =   873
      backcolor       =   -2147483633
      border          =   -1
      bordercolor     =   255
      bordercoloropacity=   50
      bordercornerlefttop=   30
      bordercornerrighttop=   30
      bordercornerbottomright=   30
      bordercornerbottomleft=   30
      borderposition  =   2
      borderwidth     =   10
      captionalignmenth=   1
      captionalignmentv=   1
      caption1        =   "Form1.frx":0184
      caption2        =   "Form1.frx":01A6
      caption1font    =   "Form1.frx":01C6
      caption2font    =   "Form1.frx":01EE
      changeonmouseover=   0
      gradientcolorp1 =   0
      gradientcolorp1opacity=   0
      gradientcolorp2 =   0
      gradientcolorp2opacity=   0
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "Form1.frx":0216
      glowcolor       =   8454143
   End
   Begin AXLPCTRL.axLabelPlus axLabelPlus2 
      Height          =   1320
      Left            =   4455
      TabIndex        =   72
      Top             =   7320
      Width           =   2730
      _extentx        =   4815
      _extenty        =   2328
      backcolor       =   -2147483633
      backshadow      =   -1
      border          =   -1
      bordercolor     =   16711680
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      captionalignmenth=   2
      captionalignmentv=   2
      caption1        =   "Form1.frx":023E
      caption2        =   "Form1.frx":0272
      caption1paddingy=   15
      caption1font    =   "Form1.frx":02A4
      caption2font    =   "Form1.frx":02CC
      changeonmouseover=   0
      shadowsize      =   2
      shadowcolor     =   12582912
      shadowoffsetx   =   3
      shadowoffsety   =   3
      shadowcoloropacity=   90
      hotline         =   -1
      hotlinewidth    =   7
      hotlineposition =   0
      iconfont        =   "Form1.frx":02FC
      iconpaddingx    =   7
      iconalignmentv  =   1
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Opacity"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   9270
      TabIndex        =   66
      Top             =   6915
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
      TabIndex        =   65
      Top             =   6600
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
      TabIndex        =   60
      Top             =   6600
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
      TabIndex        =   59
      Top             =   6930
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
      TabIndex        =   46
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
      TabIndex        =   45
      Top             =   5895
      Width           =   645
   End
   Begin AXLPCTRL.axLabelPlus axLPdc 
      Height          =   1260
      Left            =   7005
      TabIndex        =   41
      Top             =   3120
      Width           =   2940
      _extentx        =   5186
      _extenty        =   2223
      border          =   -1
      bordercolor     =   16711680
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      caption1        =   "Form1.frx":0324
      caption2        =   "Form1.frx":035C
      caption1paddingx=   10
      caption2paddingx=   10
      caption2paddingy=   20
      captionshadow   =   -1
      caption1font    =   "Form1.frx":0394
      caption2font    =   "Form1.frx":03BC
      caption1forecolor=   16777215
      caption2forecolor=   65535
      caption1forecoloropacity=   50
      caption2forecoloropacity=   50
      hotline         =   -1
      hotlinewidth    =   7
      hotlineposition =   0
      iconfont        =   "Form1.frx":03EC
      iconcharcode    =   61384
      iconforecolor   =   16777215
      iconpaddingx    =   15
      iconalignmenth  =   2
      iconalignmentv  =   1
   End
   Begin AXLPCTRL.axLabelPlus axLPCross 
      Height          =   1125
      Left            =   240
      TabIndex        =   40
      Top             =   5250
      Width           =   2730
      _extentx        =   4815
      _extenty        =   1984
      backcolor       =   -2147483633
      backshadow      =   -1
      border          =   -1
      bordercolor     =   16711680
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   5
      caption1        =   "Form1.frx":0414
      caption2        =   "Form1.frx":044C
      caption1paddingx=   42
      caption1paddingy=   15
      caption2paddingx=   20
      caption2paddingy=   35
      caption1font    =   "Form1.frx":0484
      caption2font    =   "Form1.frx":04AC
      changeonmouseover=   0
      shadowsize      =   2
      shadowcolor     =   12582912
      shadowoffsetx   =   3
      shadowoffsety   =   3
      shadowcoloropacity=   90
      hotline         =   -1
      hotlinewidth    =   7
      hotlineposition =   0
      iconfont        =   "Form1.frx":04DC
      iconcharcode    =   61384
      iconpaddingx    =   7
      iconalignmentv  =   1
   End
   Begin AXLPCTRL.axLabelPlus axLPccc 
      Height          =   810
      Left            =   270
      TabIndex        =   39
      Top             =   3420
      Width           =   2835
      _extentx        =   5001
      _extenty        =   1429
      backcolor       =   -2147483633
      backcolorpress  =   8421504
      border          =   -1
      bordercolor     =   16711680
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      captionalignmenth=   2
      captionalignmentv=   2
      caption1        =   "Form1.frx":0504
      caption2        =   "Form1.frx":053C
      caption1paddingy=   5
      caption2paddingy=   3
      captionshadow   =   -1
      caption1font    =   "Form1.frx":0574
      caption2font    =   "Form1.frx":059C
      changecoloronclick=   -1
      changeonmouseover=   0
      hotline         =   -1
      hotlinewidth    =   7
      hotlineposition =   0
      iconfont        =   "Form1.frx":05CC
      iconcharcode    =   61384
      iconpaddingx    =   10
      iconalignmenth  =   2
      iconalignmentv  =   1
   End
   Begin AXLPCTRL.axLabelPlus axLabelPlus1 
      Height          =   810
      Index           =   2
      Left            =   255
      TabIndex        =   38
      Top             =   2175
      Width           =   2835
      _extentx        =   5001
      _extenty        =   1429
      backcolor       =   12648384
      border          =   -1
      bordercolor     =   16448
      coloronmouseover=   65280
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      caption1        =   "Form1.frx":05F4
      caption2        =   "Form1.frx":062C
      caption1paddingx=   10
      caption2paddingx=   10
      caption2paddingy=   20
      captionshadow   =   -1
      caption1font    =   "Form1.frx":0664
      caption2font    =   "Form1.frx":068C
      changeonmouseover=   0
      hotline         =   -1
      hotlinewidth    =   7
      hotlineposition =   0
      iconfont        =   "Form1.frx":06BC
      iconcharcode    =   61384
      iconpaddingx    =   10
      iconalignmenth  =   2
      iconalignmentv  =   1
   End
   Begin AXLPCTRL.axLabelPlus axLabelPlus1 
      Height          =   810
      Index           =   1
      Left            =   255
      TabIndex        =   37
      Top             =   1320
      Width           =   2835
      _extentx        =   5001
      _extenty        =   1429
      backcolor       =   12648384
      border          =   -1
      bordercolor     =   16448
      coloronmouseover=   33023
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      caption1        =   "Form1.frx":06E4
      caption2        =   "Form1.frx":071C
      caption1paddingx=   10
      caption2paddingx=   10
      caption2paddingy=   20
      captionshadow   =   -1
      caption1font    =   "Form1.frx":0754
      caption2font    =   "Form1.frx":077C
      changeonmouseover=   0
      hotline         =   -1
      hotlinewidth    =   7
      hotlineposition =   0
      iconfont        =   "Form1.frx":07AC
      iconcharcode    =   61384
      iconpaddingx    =   10
      iconalignmenth  =   2
      iconalignmentv  =   1
   End
   Begin AXLPCTRL.axLabelPlus axLabelPlus1 
      Height          =   810
      Index           =   0
      Left            =   255
      TabIndex        =   36
      Top             =   465
      Width           =   2835
      _extentx        =   5001
      _extenty        =   1429
      backcolor       =   12648384
      border          =   -1
      bordercolor     =   16448
      coloronmouseover=   16711935
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      caption1        =   "Form1.frx":07D4
      caption2        =   "Form1.frx":080C
      caption1paddingx=   10
      caption2paddingx=   10
      caption2paddingy=   20
      captionshadow   =   -1
      caption1font    =   "Form1.frx":0844
      caption2font    =   "Form1.frx":086C
      changeonmouseover=   0
      hotline         =   -1
      hotlinewidth    =   7
      hotlineposition =   0
      iconfont        =   "Form1.frx":089C
      iconcharcode    =   61384
      iconpaddingx    =   10
      iconalignmenth  =   2
      iconalignmentv  =   1
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
      Left            =   270
      TabIndex        =   29
      Top             =   6840
      Width           =   855
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
      Top             =   5220
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
      Top             =   5550
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
      Top             =   525
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
      Top             =   4830
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
      Top             =   4545
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
      Top             =   4845
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
      Top             =   4530
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
      Top             =   4845
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
      Top             =   4530
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
      Top             =   2370
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

Private Sub axLabelPlus3_Click()

End Sub

Private Sub axLPGlow_Click()
'Dim oShell As Long
'oShell = Shell("c:\windows\notepad.exe", vbNormalFocus)
'oShell = Shell("c:\windows\notepad.exe", vbNormalFocus)
'oShell = Shell("c:\windows\notepad.exe", vbNormalFocus)
'oShell = Shell("c:\windows\notepad.exe", vbNormalFocus)
'oShell = Shell("c:\windows\notepad.exe", vbNormalFocus)
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
axLPCross.Glowing = Not axLPCross.Glowing
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

axLPGlow.GlowTiks = CInt(txtTiks.Text)
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

Private Sub Slider1_Click()
axLPGlow.Glowspeed = Slider1.Value
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

Private Sub txtTiks_Change()
axLPGlow.GlowTiks = CInt(txtTiks.Text)
End Sub

Private Sub VScroll1_Change()
axLPdc.Caption1PaddingY = VScroll1.Value
Y1.Text = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
axLPdc.Caption2PaddingY = VScroll2.Value
Y2.Text = VScroll2.Value
End Sub


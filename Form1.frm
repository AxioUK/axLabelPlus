VERSION 5.00
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
   LockControls    =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
   Begin Proyecto1.axLabelPlus axLPCross 
      Height          =   915
      Left            =   345
      TabIndex        =   29
      Top             =   4890
      Width           =   2775
      _extentx        =   4895
      _extenty        =   1614
      backcolor       =   12648447
      backcolorpress  =   8438015
      border          =   -1
      bordercolor     =   16711680
      coloronmouseover=   255
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      captionalignmenth=   1
      captionalignmentv=   1
      caption1        =   "Form1.frx":0000
      caption2        =   "Form1.frx":003C
      caption2paddingx=   10
      caption2paddingy=   10
      crossposition   =   7
      font            =   "Form1.frx":005C
      forecoloronpress=   16777215
      changecoloronclick=   -1
      changeonmouseover=   0
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      hotline         =   -1
      hotlinewidth    =   8
      hotlineposition =   0
      iconfont        =   "Form1.frx":0084
      iconforecolor   =   0
      iconopacity     =   0
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
   Begin Proyecto1.axLabelPlus axLPValue 
      Height          =   990
      Left            =   6480
      TabIndex        =   26
      Top             =   525
      Width           =   2775
      _extentx        =   4895
      _extenty        =   1746
      backcolor       =   12648447
      backcolorpress  =   8438015
      border          =   -1
      bordercolor     =   -2147483635
      coloronmouseover=   255
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      captionalignmenth=   1
      caption1        =   "Form1.frx":00AC
      caption2        =   "Form1.frx":00DE
      caption2sizeminus=   1
      caption2paddingx=   50
      caption2paddingy=   12
      crossposition   =   7
      font            =   "Form1.frx":0108
      forecoloronpress=   16777215
      changecoloronclick=   -1
      changeonmouseover=   0
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      hotline         =   -1
      hotlinewidth    =   45
      hotlineposition =   0
      iconfont        =   "Form1.frx":0130
      iconforecolor   =   0
      iconopacity     =   0
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
   Begin Proyecto1.axLabelPlus axLPdc 
      Height          =   1245
      Left            =   7020
      TabIndex        =   10
      Top             =   2790
      Width           =   2775
      _extentx        =   4895
      _extenty        =   2196
      backcolor       =   12648447
      backcolorpress  =   8438015
      border          =   -1
      bordercolor     =   16711680
      coloronmouseover=   255
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      caption1        =   "Form1.frx":0158
      caption2        =   "Form1.frx":0188
      crossposition   =   7
      font            =   "Form1.frx":01B8
      forecoloronpress=   16777215
      changecoloronclick=   -1
      changeonmouseover=   0
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      hotline         =   -1
      hotlinewidth    =   8
      hotlineposition =   0
      iconfont        =   "Form1.frx":01E0
      iconforecolor   =   0
      iconopacity     =   0
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
   Begin Proyecto1.axLabelPlus axLPccc 
      Height          =   795
      Left            =   285
      TabIndex        =   7
      Top             =   3480
      Width           =   2775
      _extentx        =   4895
      _extenty        =   1402
      backcolor       =   12648447
      backcolorpress  =   8438015
      border          =   -1
      bordercolor     =   16711680
      coloronmouseover=   255
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      captionalignmenth=   1
      captionalignmentv=   1
      caption1        =   "Form1.frx":0208
      caption2        =   "Form1.frx":023A
      caption2paddingx=   10
      caption2paddingy=   10
      crossposition   =   7
      font            =   "Form1.frx":025A
      forecoloronpress=   16777215
      changecoloronclick=   -1
      changeonmouseover=   0
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      hotline         =   -1
      hotlinewidth    =   8
      hotlineposition =   0
      iconfont        =   "Form1.frx":0282
      iconforecolor   =   0
      iconopacity     =   0
   End
   Begin Proyecto1.axLabelPlus axLPbc 
      Height          =   795
      Index           =   2
      Left            =   300
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
      _extentx        =   4895
      _extenty        =   1402
      backcolor       =   12648447
      border          =   -1
      bordercolor     =   16711680
      coloronmouseover=   255
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      caption1        =   "Form1.frx":02AA
      caption2        =   "Form1.frx":02CA
      caption1paddingx=   10
      caption2paddingx=   10
      caption2paddingy=   10
      crossposition   =   7
      font            =   "Form1.frx":02EA
      changeonmouseover=   0
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      hotline         =   -1
      hotlinewidth    =   8
      hotlineposition =   0
      iconfont        =   "Form1.frx":0312
      iconforecolor   =   0
      iconopacity     =   0
   End
   Begin Proyecto1.axLabelPlus axLPbc 
      Height          =   795
      Index           =   1
      Left            =   300
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
      _extentx        =   4895
      _extenty        =   1402
      backcolor       =   12648447
      border          =   -1
      bordercolor     =   16711680
      coloronmouseover=   255
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      caption1        =   "Form1.frx":033A
      caption2        =   "Form1.frx":035A
      caption1paddingx=   10
      caption2paddingx=   10
      caption2paddingy=   10
      crossposition   =   7
      font            =   "Form1.frx":037A
      changeonmouseover=   0
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      hotline         =   -1
      hotlinewidth    =   8
      hotlineposition =   0
      iconfont        =   "Form1.frx":03A2
      iconforecolor   =   0
      iconopacity     =   0
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
   Begin Proyecto1.axLabelPlus axLPbc 
      Height          =   795
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   2775
      _extentx        =   4895
      _extenty        =   1402
      backcolor       =   12648447
      border          =   -1
      bordercolor     =   16711680
      coloronmouseover=   255
      bordercornerlefttop=   5
      bordercornerrighttop=   5
      bordercornerbottomright=   5
      bordercornerbottomleft=   5
      borderwidth     =   2
      caption1        =   "Form1.frx":03CA
      caption2        =   "Form1.frx":03EA
      caption1paddingx=   10
      caption2paddingx=   10
      caption2paddingy=   10
      crossposition   =   7
      font            =   "Form1.frx":040A
      changeonmouseover=   0
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      hotline         =   -1
      hotlinewidth    =   8
      hotlineposition =   0
      iconfont        =   "Form1.frx":0432
      iconforecolor   =   0
      iconopacity     =   0
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
X1.text = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
axLPdc.Caption2PaddingX = HScroll2.Value
X2.text = HScroll2.Value
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
axLPdc.Caption1 = Text1.text
End Sub

Private Sub Text2_Change()
axLPdc.Caption2 = Text2.text
End Sub

Private Sub VScroll1_Change()
axLPdc.Caption1PaddingY = VScroll1.Value
Y1.text = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
axLPdc.Caption2PaddingY = VScroll2.Value
Y2.text = VScroll2.Value
End Sub

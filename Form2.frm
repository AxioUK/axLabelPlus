VERSION 5.00
Object = "*\AaxLabelPlus.vbp"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8265
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin AXLPCTRL.axLabelPlus axlpButton 
      Height          =   3720
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   3960
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   6562
      BackColor       =   7496448
      Caption1        =   "Form2.frx":18AD76
      Caption2        =   "Form2.frx":18ADC0
      Caption2SizeMinus=   6
      Caption1PaddingX=   65
      Caption1PaddingY=   20
      Caption2PaddingX=   10
      Caption2PaddingY=   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ChangeOnMouseOver=   0
      GradientColorP1 =   0
      GradientColorP1Opacity=   0
      GradientColorP2 =   0
      GradientColorP2Opacity=   0
      PictureOpacity  =   30
      PicturePaddingX =   20
      PicturePaddingY =   20
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      PictureColor    =   16777215
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
      PicturePresent  =   -1  'True
      PictureArr      =   "Form2.frx":18ADE0
   End
   Begin AXLPCTRL.axLabelPlus axlpButton 
      Height          =   1800
      Index           =   0
      Left            =   3930
      TabIndex        =   0
      Top             =   2055
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   3175
      BackColor       =   7496448
      CaptionAlignmentH=   2
      Caption1        =   "Form2.frx":18B303
      Caption2        =   "Form2.frx":18B333
      Caption2SizeMinus=   5
      Caption1PaddingX=   80
      Caption1PaddingY=   30
      Caption2PaddingX=   10
      Caption2PaddingY=   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ChangeOnMouseOver=   0
      GradientColorP1 =   0
      GradientColorP1Opacity=   0
      GradientColorP2 =   0
      GradientColorP2Opacity=   0
      PictureAlignmentV=   1
      PictureGraysScale=   -1  'True
      PicturePaddingX =   20
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      PictureColorize =   -1  'True
      PictureColor    =   16777215
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
      PicturePresent  =   -1  'True
      PictureArr      =   "Form2.frx":18B37B
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
axlpButton(1).Caption2 = "- Productos" & vbNewLine & "- Servicios" & vbNewLine & "- Impuestos" & vbNewLine & "- Categoria de Productos" & vbNewLine & "- Categoria de Servicios"
End Sub

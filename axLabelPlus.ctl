VERSION 5.00
Begin VB.UserControl axLabelPlus 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ClipBehavior    =   0  'None
   PropertyPages   =   "axLabelPlus.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ToolboxBitmap   =   "axLabelPlus.ctx":0011
   Windowless      =   -1  'True
   Begin VB.Timer tmrMOUSEOVER 
      Left            =   1320
      Top             =   1800
   End
End
Attribute VB_Name = "axLabelPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'Original Name: LabelPlus
'Autor:  Leandro Ascierto
'Web: www.leandroascierto.com
'LastUpdate: 18/01/2020
'Version: 1.5.0
'Based on: FirenzeLabel Project :http://www.vbforums.com/showthread.php?845221-VB6-FIRENZE-LABEL-label-control-with-so-many-functions
           'Martin Vartiak, powered by Cairo Graphics and vbRichClient-Framework.
'Special thanks to: All members of the VB6 Latin group (www.leandroacierto.com/foro), vbforum.com and activevb.net
'-----------------------------------------------
'Moded Name: axLabelPlus
'Autor:  David Rojas A. [AxioUK]
'LastUpdate: 19/01/2020
'Version: 1.5.5
'Special thanks to:
'- Leandro Ascierto por la creación de este Espectacular Control, su apoyo y guía y por permitirme modificar su control.
'- YAcosta por sus ideas y por testear cada modificación.
'- Albertomi por su apoyo y guía
'-----------------------------------------------
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsSetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long, ByVal lpTlsValue As Long) As Long
'Private Declare Function TlsFree Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsAlloc Lib "kernel32.dll" () As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal flags As Long) As Long
'Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
'Private Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal Brush As Long, ByVal argb As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipFillPolygonI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByRef mPoints As Any, ByVal mCount As Long, ByVal mFillMode As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByRef mPoints As Any, ByVal mCount As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreateTexture Lib "GdiPlus.dll" (ByVal mImage As Long, ByVal mWrapMode As Long, ByRef mTexture As Long) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As RECTL, ByVal mFlags As ImageLockMode, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTF, ByVal mStringFormat As Long, ByRef mBoundingBox As RECTF, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTF, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipSetPenMode Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mPenMode As PenAlignment) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As WrapMode, ByRef mLineGradient As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTF, ByVal mFormat As Long) As Long
'Private Declare Function GdipGetGenericFontFamilySerif Lib "gdiplus" (ByRef nativeFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As StringFormatFlags) As Long
Private Declare Function GdipSetStringFormatHotkeyPrefix Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mHotkeyPrefix As HotkeyPrefix) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As StringTrimming) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As StringAlignment) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateEffect Lib "gdiplus" (ByVal dwCid1 As Long, ByVal dwCid2 As Long, ByVal dwCid3 As Long, ByVal dwCid4 As Long, ByRef Effect As Long) As Long
Private Declare Function GdipSetEffectParameters Lib "gdiplus" (ByVal Effect As Long, ByRef params As Any, ByVal Size As Long) As Long
'Private Declare Function GdipDeleteEffect Lib "gdiplus" (ByVal Effect As Long) As Long
Private Declare Function GdipDrawImageFX Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByRef Source As RECTF, ByVal xForm As Long, ByVal Effect As Long, ByVal imageAttributes As Long, ByVal srcUnit As Long) As Long
'Private Declare Function GdipDrawPie Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipDrawArc Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipResetPath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long

Private Type BlurParams
    Radius As Single
    ExpandEdge As Long
End Type

Private Type RECTF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type COLORMATRIX
    m(0 To 4, 0 To 4)           As Single
End Type

Private Type PicBmp
    Size As Long
    type As Long
    hBmp As Long
    hpal As Long
    Reserved As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type POINTL
    X As Long
    Y As Long
End Type

Private Type BitmapData
    Width                       As Long
    Height                      As Long
    stride                      As Long
    PixelFormat                 As Long
    Scan0Ptr                    As Long
    ReservedPtr                 As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum eCallOutPosition
    coLeft
    coTop
    coRight
    coBottom
End Enum

Public Enum eCallOutAlign
    coFirstCorner
    coMidle
    coSecondCorner
    coCustomPosition
End Enum

Public Enum eBorderPosition
    bpInside
    bpCenter
    bpOutside
End Enum

Public Enum CaptionAlignmentH
    cLeft
    cCenter
    cRight
End Enum

Public Enum CaptionAlignmentV
    cTop
    cMiddle
    cBottom
End Enum

Public Enum PictureAlignmentH
    pLeft
    pCenter
    pRight
End Enum

Public Enum PictureAlignmentV
    pTop
    pMiddle
    pBottom
End Enum

Public Enum HotLinePosition
    hlLeft
    hlTop
    hlRight
    hlBottom
End Enum

Public Enum CrossPos
    cTopRight
    cMiddleRight
    cBottomRight
    cTopLeft
    cMiddleLeft
    cBottomLeft
    cMiddleTop
    cMiddleBottom
End Enum

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum
  
Public Enum StringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum
  
Public Enum StringTrimming
    StringTrimmingNone = &H0
    StringTrimmingCharacter = &H1
    StringTrimmingWord = &H2
    StringTrimmingEllipsisCharacter = &H3
    StringTrimmingEllipsisWord = &H4
    StringTrimmingEllipsisPath = &H5
End Enum

Public Enum StringFormatFlags
    StringFormatFlagsNone = &H0
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum

Private Enum HotkeyPrefix
    HotkeyPrefixNone = &H0
    HotkeyPrefixShow = &H1
    HotkeyPrefixHide = &H2
End Enum

Private Enum WrapMode
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipy = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum

Private Enum ImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum
 
Private Enum PenAlignment
    PenAlignmentCenter = &H0
    PenAlignmentInset = &H1
End Enum

Public Enum eChangeOnMouse
    eChangeNone
    eChangeBorderColor
    eChangeHotlineColor
End Enum

Private Const TLS_MINIMUM_AVAILABLE     As Long = 64
Private Const IDC_HAND                  As Long = 32649
Private Const GWL_WNDPROC               As Long = -4
Private Const GW_OWNER                  As Long = 4
Private Const WS_CHILD                  As Long = &H40000000
Private Const UnitPixel                 As Long = &H2&
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const PixelFormat32bppARGB      As Long = &H26200A
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const CombineModeExclude        As Long = &H4

Public Event Click()
Public Event DblClick()
Public Event Change()
Public Event ChangeValue(Value As Boolean)
Public Event CrossClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event PrePaint(hdc As Long, X As Long, Y As Long)
Public Event PostPaint(ByVal hdc As Long)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event PictureDownloadProgress(BytesMax As Long, BytesLeidos As Long)
Public Event PictureDownloadComplete()
Public Event PictureDownloadError()

'Default Property Values:
Const m_def_BackColorOpacity = 100
Const m_def_BackColorP = &HE0E0E0
Const m_def_BackColorPOpacity = 100
Const m_def_Border = False
Const m_def_BorderColor = vbActiveBorder
Const m_def_BorderColorOpacity = 100
Const m_def_ColorOnMouseOver = vbWhite
Const m_def_ColorOnMouseOverOpacity = 100
Const m_def_BorderPosition = 1
Const m_def_BorderWidth = 0
Const m_def_CaptionAlignmentH = 0
Const m_def_CaptionAlignmentV = 0
Const m_def_ForeColor = &H80000012
Const m_def_ForeColorOpacity = 100
Const m_def_ForeColorP = &HC0C0C0
Const m_def_ForeColorPOpacity = 100
Const m_def_ChangeColorOnClick = False
Const m_def_Gradient = False
Const m_def_GradientAngle = 0
Const m_def_GradientColor1 = &HD3A042
Const m_def_GradientColor1Opacity = 100
Const m_def_GradientColor2 = &HE96E9B
Const m_def_GradientColor2Opacity = 100
Const m_def_GradientColorP1 = &HD0BB97
Const m_def_GradientColorP1Opacity = 100
Const m_def_GradientColorP2 = &HC1A06F
Const m_def_GradientColorP2Opacity = 100
Const m_def_PictureOpacity = 100
Const m_def_WordWrap = True
Const m_def_Value = False
Const m_def_OptionBehavior = False

'Property Variables:
Dim m_CallOut As Boolean
Dim m_CallOutPosicion As eCallOutPosition
Dim m_CallOutAlign As eCallOutAlign
Dim m_coLen As Long
Dim m_coWidth As Long
Dim m_coCustomPos As Long
Dim m_coRightTriangle As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_BackColorOpacity As Integer
Dim m_BackColorP As OLE_COLOR
Dim m_BackColorPOpacity As Integer
Dim m_BackAcrylicBlur As Boolean
Dim m_BackShadow As Boolean
Dim m_Border As Boolean
Dim m_BorderColor As OLE_COLOR
Dim m_BorderColorOpacity As Integer
Dim m_ColorOnMouseOver As OLE_COLOR
Dim m_ColorOnMouseOverOpacity As Integer
Dim m_BorderPosition As eBorderPosition
Dim m_BorderCornerLeftTop As Integer
Dim m_BorderCornerRightTop As Integer
Dim m_BorderCornerBottomLeft As Integer
Dim m_BorderCornerBottomRight As Integer
Dim m_BorderWidth As Integer
Dim hImgShadow As Long
Dim m_ShadowSize As Integer
Dim m_ShadowColor As OLE_COLOR
Dim m_ShadowOffsetX As Integer
Dim m_ShadowOffsetY As Integer
Dim m_ShadowColorOpacity As Integer
Dim hImgCaptionShadow As Long
Dim m_Caption1() As Byte
Dim m_Caption2() As Byte
Dim m_CaptionAlignmentH As CaptionAlignmentH
Dim m_CaptionAlignmentV As CaptionAlignmentV
Dim m_Caption1PaddingX As Integer
Dim m_Caption1PaddingY As Integer
Dim m_Caption2PaddingX As Integer
Dim m_Caption2PaddingY As Integer
Dim m_SizeMinus As Integer
Dim m_CaptionTriming As StringTrimming
Dim m_CaptionBorderWidth As Integer
Dim m_CaptionBorderColor As OLE_COLOR
Dim m_CaptionShadow As Boolean
Dim m_CaptionAngle As Integer
Dim m_CaptionShowPrefix As Boolean
Dim m_AutoSize As Boolean
Dim m_MousePointerHands As Boolean
Dim WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
Dim m_ForeColor As OLE_COLOR
Dim m_ForeColorOpacity As Integer
Dim m_ForeColorP As OLE_COLOR
Dim m_ForeColorPOpacity As Integer
Dim m_ChangeColorOnClick As Boolean
Dim m_ChangeOnMouseOver As eChangeOnMouse
Dim m_Gradient As Boolean
Dim m_GradientAngle As Integer
Dim m_GradientColor1 As OLE_COLOR
Dim m_GradientColor1Opacity As Integer
Dim m_GradientColor2 As OLE_COLOR
Dim m_GradientColor2Opacity As Integer
Dim m_GradientColorP1 As OLE_COLOR
Dim m_GradientColorP1Opacity As Integer
Dim m_GradientColorP2 As OLE_COLOR
Dim m_GradientColorP2Opacity As Integer
Dim m_PictureAngle As Integer
Dim m_PictureAlignmentH As PictureAlignmentH
Dim m_PictureAlignmentV As PictureAlignmentV
Dim m_PicturePaddingX As Integer
Dim m_PicturePaddingY As Integer
Dim m_PictureRealWidth As Long
Dim m_PictureRealHeight As Long
Dim m_PictureSetWidth As Long
Dim m_PictureSetHeight As Long
Dim m_PictureArr()  As Byte
Dim m_PicturePresent As Boolean
Dim m_PictureGraysScale As Boolean
Dim m_PictureContrast As Integer
Dim m_PictureBrightness As Integer
Dim m_PictureOpacity As Integer
Dim m_PictureColor As OLE_COLOR
Dim m_PictureColorize As Boolean
Dim m_PictureShadow As Boolean
Dim m_MouseToParent As Boolean
Dim m_HotLine As Boolean
Dim m_HotLineColor As OLE_COLOR
Dim m_HotLineColorOpacity As Integer
Dim m_HotLineWidth As Long
Dim m_HotLinePosition As HotLinePosition
Dim m_Value As Boolean
Dim m_OptionBehavior As Boolean
Dim m_Clicked As Boolean
Dim m_MouseOver As Boolean
Dim YCrossPos As Long
Dim XCrossPos As Long
Dim m_CrossPosition As CrossPos
Dim m_CrossVisible As Boolean

Dim m_WordWrap As Boolean
Dim m_IconFont As StdFont
Dim m_IconCharCode As Long
Dim m_IconForeColor As Long
Dim m_IconPaddingX As Integer
Dim m_IconPaddingY As Integer
Dim m_IconAlignmentH As CaptionAlignmentH
Dim m_IconAlignmentV As CaptionAlignmentV
Dim m_IconOpacity As Integer

Dim hCur As Long
Dim c_lhWnd As Long
Dim nScale As Single
Dim hDCMemory As Long
Dim hBmp As Long
Dim OldhBmp As Long
Dim bRecreateShadowCaption As Boolean
Dim m_DrawProgress As Boolean
Dim c_AsyncProp As AsyncProperty
Dim m_PictureBrush As Long
Dim hFontCollection As Long

Private Sub DrawCross(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Heigth As Long, ByVal PenWidth As Long, ByVal ForeColor As OLE_COLOR)
    UserControl.DrawWidth = PenWidth
    UserControl.ForeColor = ForeColor
    UserControl.Line (Left, Top)-(Left + Width, Top + Heigth)
    UserControl.Line (Left, Top + Heigth)-(Left + Width, Top)
End Sub

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Private Sub OptBehavior()
Dim Frm As Form
    Set Frm = Extender.Parent
    
Dim lHWnd As Long
    lHWnd = Extender.Container.hwnd

    Dim Ctrl As Control
    For Each Ctrl In Frm.Controls
        With Ctrl
           If TypeOf Ctrl Is axLabelPlus Then
              If .OptionBehavior = True Then
                 'If (.Container.hWnd = lHWnd) And (Ctrl.hWnd <> UserControl.hWnd) Then
                 If (.Container.hwnd = lHWnd) And ObjPtr(Ctrl) <> ObjPtr(Extender) Then
                  If .Value Then .Value = False
                 End If
              End If
           End If
        End With
    Next
End Sub

Public Sub Draw(ByVal hdc As Long, ByVal hGraphics As Long, ByVal PosX As Long, PosY As Long)
    Dim hPath As Long
    Dim hBrush As Long
    Dim hPen As Long
    Dim X As Long, Y As Long
    Dim Xx As Long, Yy As Long
    Dim lWidth As Long, lHeight As Long
    Dim WW As Long, HH As Long
    Dim ShadowSize As Integer
    Dim ShadowOffsetX As Integer
    Dim ShadowOffsetY As Integer
    Dim BorderWidth As Integer

    ShadowSize = m_ShadowSize * nScale
    ShadowOffsetX = m_ShadowOffsetX * nScale
    ShadowOffsetY = m_ShadowOffsetY * nScale
    BorderWidth = m_BorderWidth * nScale
    
    If m_BackAcrylicBlur Then
        BitBlt hDCMemory, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hdc, 0, 0, vbSrcCopy
    End If
     
    If hGraphics = 0 Then GdipCreateFromHDC hdc, hGraphics
    
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipTranslateWorldTransform hGraphics, PosX, PosY, &H1
    
    If m_BorderPosition = bpInside Then
        lWidth = UserControl.ScaleWidth
        lHeight = UserControl.ScaleHeight
    ElseIf m_BorderPosition = bpCenter Then
        X = (BorderWidth \ 2)
        Y = (BorderWidth \ 2)
        lWidth = UserControl.ScaleWidth - BorderWidth
        lHeight = UserControl.ScaleHeight - BorderWidth
    Else
        X = BorderWidth
        Y = BorderWidth
        lWidth = UserControl.ScaleWidth - (BorderWidth * 2)
        lHeight = UserControl.ScaleHeight - (BorderWidth * 2)
    End If
    
    If hImgShadow Then
        Xx = IIf(ShadowOffsetX > 0, ShadowOffsetX, 0) '+ PosX
        Yy = IIf(ShadowOffsetY > 0, ShadowOffsetY, 0) '+ PosY
        GdipDrawImageRectI hGraphics, hImgShadow, Xx, Yy, UserControl.ScaleWidth - Abs(ShadowOffsetX), UserControl.ScaleHeight - Abs(ShadowOffsetY)
    End If
    
    If m_BackShadow = True And m_ShadowSize > 0 Then
        X = X + ShadowSize + IIf(ShadowOffsetX < 0, Abs(ShadowOffsetX), 0) '+ PosX
        Y = Y + ShadowSize + IIf(ShadowOffsetY < 0, Abs(ShadowOffsetY), 0) '+ PosY
        lWidth = lWidth - (ShadowSize * 2) - Abs(ShadowOffsetX)
        lHeight = lHeight - (ShadowSize * 2) - Abs(ShadowOffsetY)
    End If
  
    Xx = X:         Yy = Y
    WW = lWidth:    HH = lHeight

    hPath = RoundRectangle(Xx, Yy, WW, HH)
    
    If m_BackAcrylicBlur Then
        DrawAcrylicBlur hGraphics, hPath
    End If

    If m_Gradient Then
        Dim RECTL As RECTL
        SetRect RECTL, X, Y, lWidth, lHeight
        If m_ChangeColorOnClick And m_Clicked Then
          GdipCreateLineBrushFromRectWithAngleI RECTL, ConvertColor(m_GradientColorP1, m_GradientColorP1Opacity), _
                                                      ConvertColor(m_GradientColorP2, m_GradientColorP2Opacity), _
                                                      m_GradientAngle + 90, 0, WrapModeTileFlipXY, hBrush
        Else
          GdipCreateLineBrushFromRectWithAngleI RECTL, ConvertColor(m_GradientColor1, m_GradientColor1Opacity), _
                                                      ConvertColor(m_GradientColor2, m_GradientColor2Opacity), _
                                                      m_GradientAngle + 90, 0, WrapModeTileFlipXY, hBrush
        End If
    Else
        If m_ChangeColorOnClick And m_Clicked Then
            GdipCreateSolidFill ConvertColor(m_BackColorP, m_BackColorPOpacity), hBrush
        Else
            GdipCreateSolidFill ConvertColor(m_BackColor, m_BackColorOpacity), hBrush
        End If
    End If
        
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush

    If m_PicturePresent Then
        If m_PictureBrush = 0 Then m_PictureBrush = CreateBrushTexture(hGraphics, hPath, Xx, Yy, WW, HH)
        GdipFillPath hGraphics, m_PictureBrush, hPath
    End If

    If Not c_AsyncProp Is Nothing Then
        If c_AsyncProp.BytesMax = c_AsyncProp.BytesRead Then
            Set c_AsyncProp = Nothing
        Else
            DrawProgress hGraphics, Xx, Yy, WW, HH, c_AsyncProp.BytesRead, c_AsyncProp.BytesMax
        End If
    End If
    
    If m_CaptionShadow Then
        CreateCaptionShadow Xx, Yy, WW, HH
        
        If hImgCaptionShadow <> 0 Then
            X = Xx - ShadowSize + IIf(ShadowOffsetX > 0, ShadowOffsetX, ShadowOffsetX * 2) + PosX
            Y = Yy - ShadowSize + IIf(ShadowOffsetY > 0, ShadowOffsetY, ShadowOffsetY * 2) + PosY
            GdipDrawImageRectI hGraphics, hImgCaptionShadow, X, Y, WW + ShadowSize * 2, HH + ShadowSize * 2
        End If
    End If
    
    If m_HotLine Then
          DrawHotLine hGraphics, hPath, PosX, PosY
    End If
    
    GDIP_AddPathString hGraphics, Xx, Yy, WW, HH

    If m_Border And BorderWidth > 0 Then
        If m_ChangeOnMouseOver = eChangeBorderColor And m_MouseOver Then
          GdipCreatePen1 ConvertColor(m_ColorOnMouseOver, m_ColorOnMouseOverOpacity), BorderWidth, UnitPixel, hPen
        Else
          GdipCreatePen1 ConvertColor(m_BorderColor, m_BorderColorOpacity), BorderWidth, UnitPixel, hPen
        End If
        If m_BorderPosition = bpInside Then
            GdipSetPenMode hPen, PenAlignmentInset
        ElseIf m_BorderPosition = bpOutside Then
    
            GdipDeletePath hPath
            X = (BorderWidth / 2) + PosX   '+ ShadowSize + ShadowOffsetX
            Y = (BorderWidth / 2) + PosY  '+ ShadowSize + ShadowOffsetX
            lWidth = UserControl.ScaleWidth - BorderWidth '- (ShadowSize * 2)
            lHeight = UserControl.ScaleHeight - BorderWidth  '- (ShadowSize * 2)
            hPath = RoundRectangle(X, Y, lWidth, lHeight, True)
        End If
        
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
    End If
    
    GdipDeletePath hPath
    If hdc <> 0 Then GdipDeleteGraphics hGraphics
      
  'Cross
  If m_CrossVisible Then
    Select Case m_CrossPosition
      Case Is = cTopRight
          If m_BackShadow Then
            YCrossPos = (BorderWidth + m_ShadowSize + 5)
            XCrossPos = UserControl.ScaleWidth - (BorderWidth + m_ShadowSize + 12)
          Else
            YCrossPos = BorderWidth + 10
            XCrossPos = UserControl.ScaleWidth - (BorderWidth + 15)
          End If
      Case Is = cMiddleRight
          If m_BackShadow Then
            YCrossPos = (UserControl.ScaleHeight / 2) - 5
            XCrossPos = UserControl.ScaleWidth - (BorderWidth + m_ShadowSize + 12)
          Else
            YCrossPos = UserControl.ScaleHeight / 2
            XCrossPos = UserControl.ScaleWidth - (BorderWidth + 15)
          End If
      Case Is = cBottomRight
          If m_BackShadow Then
            YCrossPos = UserControl.ScaleHeight - (BorderWidth + m_ShadowSize + 12)
            XCrossPos = UserControl.ScaleWidth - (BorderWidth + m_ShadowSize + 12)
          Else
            YCrossPos = UserControl.ScaleHeight - (BorderWidth + 14)
            XCrossPos = UserControl.ScaleWidth - (BorderWidth + 15)
          End If
      Case Is = cTopLeft
          If m_BackShadow Then
            YCrossPos = (BorderWidth + m_ShadowSize + 5)
            XCrossPos = (BorderWidth + m_ShadowSize + 5)
          Else
            YCrossPos = BorderWidth + 10
            XCrossPos = (BorderWidth + 10)
          End If
      Case Is = cMiddleLeft
          If m_BackShadow Then
            YCrossPos = (UserControl.ScaleHeight / 2) - 5
            XCrossPos = (BorderWidth + m_ShadowSize + 5)
          Else
            YCrossPos = UserControl.ScaleHeight / 2
            XCrossPos = (BorderWidth + 10)
          End If
      Case Is = cBottomLeft
          If m_BackShadow Then
            YCrossPos = UserControl.ScaleHeight - (BorderWidth + m_ShadowSize + 12)
            XCrossPos = (BorderWidth + m_ShadowSize + 5)
          Else
            YCrossPos = UserControl.ScaleHeight - (BorderWidth + 14)
            XCrossPos = (BorderWidth + 10)
          End If
      Case Is = cMiddleTop
          If m_BackShadow Then
            YCrossPos = (BorderWidth + m_ShadowSize + 5)
            XCrossPos = (UserControl.ScaleWidth / 2)
          Else
            YCrossPos = BorderWidth + 10
            XCrossPos = (UserControl.ScaleWidth / 2)
          End If
      Case Is = cMiddleBottom
          If m_BackShadow Then
            YCrossPos = UserControl.ScaleHeight - (BorderWidth + m_ShadowSize + 12)
            XCrossPos = (UserControl.ScaleWidth / 2)
          Else
            YCrossPos = UserControl.ScaleHeight - (BorderWidth + 14)
            XCrossPos = (UserControl.ScaleWidth / 2)
          End If
    End Select

    DrawCross XCrossPos, YCrossPos, 6, 6, 2, m_BorderColor
  End If
End Sub

Public Function DrawLine(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional ByVal oColor As OLE_COLOR = vbBlack, Optional ByVal Opacity As Integer = 100, Optional ByVal PenWidth As Integer = 1) As Boolean
    Dim hGraphics As Long, hPen As Long
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipCreatePen1 ConvertColor(oColor, Opacity), PenWidth * nScale, UnitPixel, hPen
    DrawLine = GdipDrawLineI(hGraphics, hPen, X1 * nScale, Y1 * nScale, X2 * nScale, Y2 * nScale) = 0
    GdipDeletePen hPen
    GdipDeleteGraphics hGraphics
End Function

Function DrawText(ByVal hdc As Long, ByVal text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal ForeColor As OLE_COLOR, Optional ByVal ColorOpacity As Integer = 100, Optional HAlign As CaptionAlignmentH, Optional VAlign As CaptionAlignmentV, Optional bWordWrap As Boolean) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RECTF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim hGraphics As Long
    
    SafeRange ColorOpacity, 0, 100
    
    GdipCreateFromHDC hdc, hGraphics
  
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
        'If GdipGetGenericFontFamilySerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        'GdipSetStringFormatFlags hFormat, HotkeyPrefixShow
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
    End If
        
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        

    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)

    layoutRect.Left = X * nScale: layoutRect.Top = Y * nScale
    layoutRect.Width = Width * nScale: layoutRect.Height = Height * nScale

    GdipCreateSolidFill ConvertColor(ForeColor, ColorOpacity), hBrush
            
    'GdipSetTextRenderingHint hGraphics, TextRenderingHintClearTypeGridFit
    'GdipSetTextContrast hGraphics, 12
    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    GdipDrawString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, hBrush
    
    Dim BB As RECTF, CF As Long, LF As Long

    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    GdipMeasureString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, BB, CF, LF

    
    If bWordWrap Then
        DrawText = BB.Height / nScale
    Else
        DrawText = BB.Width / nScale
    End If
    
    GdipDeleteFont hFont
    GdipDeleteBrush hBrush
    GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily
    GdipDeleteGraphics hGraphics

End Function

Public Function GetSystemHandCursor() As Picture
    Dim Pic As PicBmp, IPic As IPicture, GUID(0 To 3) As Long
    
    If hCur Then DestroyCursor hCur: hCur = 0
    
    hCur = LoadCursor(ByVal 0&, IDC_HAND)
     
    GUID(0) = &H7BF80980
    GUID(1) = &H101ABF32
    GUID(2) = &HAA00BB8B
    GUID(3) = &HAB0C3000
 
    With Pic
        .Size = Len(Pic)
        .type = vbPicTypeIcon
        .hBmp = hCur
        .hpal = 0
    End With
 
    Call OleCreatePictureIndirect(Pic, GUID(0), 1, IPic)
 
    Set GetSystemHandCursor = IPic
    
End Function

Public Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Public Function IsMouseInExtender() As Boolean
    Dim PT As POINTAPI
    Dim CPT As POINTAPI
    Dim TR As RECT
    Dim bArea As Boolean
    
    Call GetCursorPos(PT)
    Call ClientToScreen(c_lhWnd, CPT)
    
    CPT.X = PT.X - CPT.X
    CPT.Y = PT.Y - CPT.Y

    With TR
        .Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode) ' / nScale
        .Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode) ' / nScale
        .Right = .Left + UserControl.ScaleWidth
        .Bottom = .Top + UserControl.ScaleHeight
    End With
    
    bArea = PtInRect(TR, CPT.X, CPT.Y)
    
    If bArea And WindowFromPoint(PT.X, PT.Y) = c_lhWnd Then
        IsMouseInExtender = True
    End If

End Function

Public Function PictureDelete()
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Erase m_PictureArr
    m_PicturePresent = False
    Call PropertyChanged("PicturePresent")
    Call PropertyChanged("PictureArr")
    UserControl.Refresh
End Function


Public Function PictureFromStream(ByRef bvStream() As Byte) As Boolean
    Dim hImage As Long

    If LoadImageFromStream(bvStream, hImage) Then
        GdipGetImageWidth hImage, m_PictureRealWidth
        GdipGetImageHeight hImage, m_PictureRealHeight
        GdipDisposeImage hImage
        m_PictureArr() = bvStream
        PictureFromStream = True
        m_PicturePresent = True
    Else
        Erase m_PictureArr
        m_PicturePresent = False
    End If
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Call PropertyChanged("PicturePresent")
    Call PropertyChanged("PictureArr")
    UserControl.Refresh
End Function

Public Function PictureFromURL(ByVal sUrl As String, Optional ByVal UseCache As Boolean, Optional ByVal DrawProgress As Boolean = True) As Boolean
    On Error Resume Next
    m_DrawProgress = DrawProgress
    UserControl.CancelAsyncRead "URL"
    Err.Clear
    Call AsyncRead(sUrl, vbAsyncTypeByteArray, "URL", IIf(UseCache, 0, vbAsyncReadForceUpdate))
    PictureFromURL = Err.Number = 0
End Function

Public Function PictureGetStream() As Byte()
    PictureGetStream = m_PictureArr
End Function


Public Function Polygon(ByVal hdc As Long, ByVal PenWidth As Long, ByVal oColor As OLE_COLOR, ByVal Opacity As Integer, ParamArray vPoints() As Variant) As Boolean
    Dim hGraphics As Long, hBrush As Long, hPen As Long
    Dim lPoints() As Long
    Dim lCount As Long
    Dim I As Long
    
    If UBound(vPoints) = 1 Then
        lCount = vPoints(1)
        ReDim lPoints(lCount - 1)
        CopyMemory lPoints(0), ByVal CLng(vPoints(0)), lCount * 4
    Else
        lCount = UBound(vPoints) + 1
        ReDim lPoints(lCount - 1)
        For I = 0 To lCount - 1
            lPoints(I) = vPoints(I) * nScale
        Next
    End If
    GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    If PenWidth > 0 Then
        GdipCreatePen1 ConvertColor(oColor, Opacity), PenWidth, UnitPixel, hPen
        Call GdipDrawPolygonI(hGraphics, hPen, lPoints(0), lCount / 2)
        GdipDeletePen hPen
    Else
        GdipCreateSolidFill ConvertColor(oColor, Opacity), hBrush
        Call GdipFillPolygonI(hGraphics, hBrush, lPoints(0), lCount / 2, &H1)
        GdipDeleteBrush hBrush
    End If
    
    GdipDeleteGraphics hGraphics
End Function

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Function ConvertColor(ByVal Color As Long, ByVal Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function


Private Function CreateBlurShadowImage(ByVal hImage As Long, ByVal Color As Long, blurDepth As Integer, _
                                        Optional ByVal Left As Long, Optional ByVal Top As Long, _
                                        Optional ByVal Width As Long, Optional ByVal Height As Long) As Long
                                        
    Dim REC As RECTL
    Dim X As Long, Y As Long
    Dim hImgShadow As Long
    Dim bmpData1 As BitmapData
    Dim bmpData2 As BitmapData
    Dim t2xBlur As Long
    Dim R As Long, G As Long, B As Long
    Dim Alpha As Byte
    Dim lSrcAlpha As Long, lDestAlpha As Long
    Dim dBytes() As Byte
    Dim srcBytes() As Byte
    Dim vTally() As Long
    Dim tAlpha As Long, tColumn As Long, tAvg As Long
    Dim initY As Long, initYstop As Long, initYstart As Long
    Dim initX As Long, initXstop As Long
    
    If hImage = 0& Then Exit Function
 
    If Width = 0& Then Call GdipGetImageWidth(hImage, Width)
    If Height = 0& Then Call GdipGetImageHeight(hImage, Height)
 
    t2xBlur = blurDepth * 2
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
 
    SetRect REC, Left, Top, Width, Height
 
    ReDim srcBytes(REC.Width * 4 - 1&, REC.Height - 1&)
  
    With bmpData1
        .Scan0Ptr = VarPtr(srcBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
   
    Call GdipBitmapLockBits(hImage, REC, ImageLockModeUserInputBuf Or ImageLockModeRead, PixelFormat32bppPARGB, bmpData1)
 
    SetRect REC, Left, Top, Width + t2xBlur, Height + t2xBlur
    
    Call GdipCreateBitmapFromScan0(REC.Width, REC.Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImgShadow)

    ReDim dBytes(REC.Width * 4 - 1&, REC.Height - 1&)
    
    With bmpData2
        .Scan0Ptr = VarPtr(dBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
    
    Call GdipBitmapLockBits(hImgShadow, REC, ImageLockModeUserInputBuf Or ImageLockModeRead Or ImageLockModeWrite, PixelFormat32bppPARGB, bmpData2)
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
    
    tAvg = (t2xBlur + 1) * (t2xBlur + 1)    ' how many pixels are being blurred
    
    ReDim vTally(0 To t2xBlur)              ' number of blur columns per pixel
    
    For Y = 0 To Height + t2xBlur - 1     ' loop thru shadow dib
    
        FillMemory vTally(0), (t2xBlur + 1) * 4, 0  ' reset column totals
        
        If Y < t2xBlur Then         ' y does not exist in source
            initYstart = 0          ' use 1st row
        Else
            initYstart = Y - t2xBlur ' start n blur rows above y
        End If
        ' how may source rows can we use for blurring?
        If Y < Height Then initYstop = Y Else initYstop = Height - 1
        
        tAlpha = 0  ' reset alpha sum
        tColumn = 0    ' reset column counter
        
        ' the first n columns will all be zero
        ' only the far right blur column has values; tally them
        For initY = initYstart To initYstop
            tAlpha = tAlpha + srcBytes(3, initY)
        Next
        ' assign the right column value
        vTally(t2xBlur) = tAlpha
        
        For X = 3 To (Width - 2) * 4 - 1 Step 4
            ' loop thru each source pixel's alpha
            
            ' set shadow alpha using blur average
            dBytes(X, Y) = tAlpha \ tAvg
            ' and set shadow color
            Select Case dBytes(X, Y)
            Case 255
                dBytes(X - 1, Y) = R
                dBytes(X - 2, Y) = G
                dBytes(X - 3, Y) = B
            Case 0
            Case Else
                dBytes(X - 1, Y) = R * dBytes(X, Y) \ 255
                dBytes(X - 2, Y) = G * dBytes(X, Y) \ 255
                dBytes(X - 3, Y) = B * dBytes(X, Y) \ 255
            End Select
            ' remove the furthest left column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' count the next column of alphas
            vTally(tColumn) = 0&
            For initY = initYstart To initYstop
                vTally(tColumn) = vTally(tColumn) + srcBytes(X + 4, initY)
            Next
            ' add the new column's sum to the overall sum
            tAlpha = tAlpha + vTally(tColumn)
            ' set the next column to be recalculated
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
        
        ' now to finish blurring from right edge of source
        For X = X To (Width + t2xBlur - 1) * 4 - 1 Step 4
            dBytes(X, Y) = tAlpha \ tAvg
            Select Case dBytes(X, Y)
            Case 255
                dBytes(X - 1, Y) = R
                dBytes(X - 2, Y) = G
                dBytes(X - 3, Y) = B
            Case 0
            Case Else
                dBytes(X - 1, Y) = R * dBytes(X, Y) \ 255
                dBytes(X - 2, Y) = G * dBytes(X, Y) \ 255
                dBytes(X - 3, Y) = B * dBytes(X, Y) \ 255
            End Select
            ' remove this column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' set next column to be removed
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
    Next
 
    Call GdipBitmapUnlockBits(hImage, bmpData1)
    Call GdipBitmapUnlockBits(hImgShadow, bmpData2)
    
    CreateBlurShadowImage = hImgShadow
End Function

Private Function CreateBrushTexture(hGraphics As Long, hPath As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
    Dim hBrush As Long
    Dim hImage As Long
    Dim hGraphics2 As Long, hImage2 As Long
    Dim tMatrixColor    As COLORMATRIX, tMatrixGray    As COLORMATRIX
    Dim hAttributes As Long
    Dim ReqWidth As Long, ReqHeight As Long
    Dim HScale As Double, VScale As Double
    Dim MyScale As Double
    Dim imgWidth As Long
    Dim imgHeight As Long

    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0

    If LoadImageFromStream(m_PictureArr, hImage) Then
 
        Call GdipCreateImageAttributes(hAttributes)
        
        
        With tMatrixColor
            If m_PictureColorize Then
                Dim R As Byte, G As Byte, B As Byte

                B = ((m_PictureColor \ &H10000) And &HFF)
                G = ((m_PictureColor \ &H100) And &HFF)
                R = (m_PictureColor And &HFF)
                
                .m(0, 0) = R / 255
                .m(1, 0) = G / 255
                .m(2, 0) = B / 255
                .m(0, 4) = R / 255
                .m(1, 4) = G / 255
                .m(2, 4) = B / 255
            Else
                .m(0, 0) = 1
                .m(1, 1) = 1
                .m(2, 2) = 1

            End If
            .m(3, 3) = m_PictureOpacity / 100
            .m(4, 4) = 1
 
            If Not m_PictureContrast = 0 Then
                .m(0, 0) = 1 + m_PictureContrast
                .m(1, 1) = .m(0, 0)
                .m(2, 2) = .m(0, 0)
                .m(0, 4) = 0.5 * -m_PictureContrast
                .m(1, 4) = .m(0, 4)
                .m(2, 4) = .m(0, 4)
            End If
            
            If m_PictureBrightness <> 0 Then
                .m(0, 4) = .m(0, 4) + m_PictureBrightness / 100
                .m(1, 4) = .m(1, 4) + m_PictureBrightness / 100
                .m(2, 4) = .m(2, 4) + m_PictureBrightness / 100
            End If
            
            If m_PictureGraysScale Then
                .m(0, 0) = 0.299
                .m(1, 0) = 0.299
                .m(2, 0) = 0.299
                .m(0, 1) = 0.587
                .m(1, 1) = 0.587
                .m(2, 1) = 0.587
                .m(0, 2) = 0.114
                .m(1, 2) = 0.114
                .m(2, 2) = 0.114
            End If
        End With

        Call GdipCreateBitmapFromScan0(UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, PixelFormat32bppARGB, ByVal 0&, hImage2)
        Call GdipGetImageGraphicsContext(hImage2, hGraphics2)
        GdipSetSmoothingMode hGraphics2, SmoothingModeAntiAlias
        
        If m_PictureSetWidth = 0 Then ReqWidth = m_PictureRealWidth Else ReqWidth = m_PictureSetWidth
        If m_PictureSetHeight = 0 Then ReqHeight = m_PictureRealHeight Else ReqHeight = m_PictureSetHeight

        HScale = ReqWidth / m_PictureRealWidth
        VScale = ReqHeight / m_PictureRealHeight
        
        MyScale = IIf(VScale >= HScale, HScale, VScale)

        ReqWidth = m_PictureRealWidth * MyScale * nScale
        ReqHeight = m_PictureRealHeight * MyScale * nScale

        If m_PictureAlignmentH = pLeft Then X = X + (m_PicturePaddingX * nScale)
        If m_PictureAlignmentH = pCenter Then X = X + (Width / 2) - (ReqWidth / 2) + (m_PicturePaddingX * nScale)
        If m_PictureAlignmentH = pRight Then X = X + Width - ReqWidth - (m_PicturePaddingX * nScale)
        If m_PictureAlignmentV = pTop Then Y = Y + (m_PicturePaddingY * nScale)
        If m_PictureAlignmentV = pMiddle Then Y = Y + (Height / 2) - (ReqHeight / 2) + (m_PicturePaddingY * nScale)
        If m_PictureAlignmentV = pBottom Then Y = Y + Height - ReqHeight - (m_PicturePaddingY * nScale)

        If m_PictureShadow = True And m_ShadowSize > 0 And m_ShadowColorOpacity > 0 Then
            Dim hPictureShadow As Long
            Dim ShadowSize As Integer
            Dim W As Long, H As Long
            
            ShadowSize = m_ShadowSize * nScale
            hPictureShadow = CreateBlurShadowImage(hImage, m_ShadowColor, ShadowSize, 0, 0, m_PictureRealWidth, m_PictureRealHeight)
            tMatrixColor.m(3, 3) = m_ShadowColorOpacity / 100
            GdipSetImageAttributesColorMatrix hAttributes, &H0, True, tMatrixColor, tMatrixGray, &H0
            If m_PictureAngle <> 0 Then
                W = ReqWidth + ShadowSize * 2
                H = ReqHeight + ShadowSize * 2
                Call GdipRotateWorldTransform(hGraphics2, m_PictureAngle + 180, 0)
                Call GdipTranslateWorldTransform(hGraphics2, X + (W \ 2) - ShadowSize + m_ShadowOffsetX, Y + (H \ 2) - ShadowSize + m_ShadowOffsetY, 1)
                GdipDrawImageRectRectI hGraphics2, hPictureShadow, W \ 2, H \ 2, -W, -H, 0, 0, m_PictureRealWidth + ShadowSize * 2, m_PictureRealHeight + ShadowSize * 2, UnitPixel, hAttributes
                GdipResetWorldTransform hGraphics2
            Else
                GdipDrawImageRectRectI hGraphics2, hPictureShadow, X - ShadowSize + m_ShadowOffsetX, Y - ShadowSize + m_ShadowOffsetY, ReqWidth + ShadowSize * 2, ReqHeight + ShadowSize * 2, 0, 0, m_PictureRealWidth + ShadowSize * 2, m_PictureRealHeight + ShadowSize * 2, UnitPixel, hAttributes
            End If
            GdipDisposeImage hPictureShadow
            
            tMatrixColor.m(3, 3) = m_PictureOpacity / 100
        End If
                
        GdipSetImageAttributesColorMatrix hAttributes, &H0, True, tMatrixColor, tMatrixGray, &H0

        If m_PictureAngle <> 0 Then
            Call GdipRotateWorldTransform(hGraphics2, m_PictureAngle + 180, 0)
            Call GdipTranslateWorldTransform(hGraphics2, X + (ReqWidth \ 2), Y + (ReqHeight \ 2), 1)
            GdipDrawImageRectRectI hGraphics2, hImage, ReqWidth \ 2, ReqHeight \ 2, -ReqWidth, -ReqHeight, 0, 0, m_PictureRealWidth, m_PictureRealHeight, UnitPixel, hAttributes
        Else
            GdipDrawImageRectRectI hGraphics2, hImage, X, Y, ReqWidth, ReqHeight, 0, 0, m_PictureRealWidth, m_PictureRealHeight, UnitPixel, hAttributes
        End If
                
        GdipDisposeImage hImage
        Call GdipDisposeImageAttributes(hAttributes)
        
        GdipCreateTexture hImage2, &H0, hBrush
        
        CreateBrushTexture = hBrush

        Call GdipDeleteGraphics(hGraphics2)
        Call GdipDisposeImage(hImage2)
        
    End If
End Function

Private Sub CreateBuffer() 'Acrylic buffer
    Dim DC As Long
    If OldhBmp Then DeleteObject SelectObject(hDCMemory, OldhBmp): OldhBmp = 0
    If hDCMemory Then DeleteDC hDCMemory: hDCMemory = 0
 
    DC = GetDC(0)
    hDCMemory = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, UserControl.ScaleWidth, UserControl.ScaleHeight)
    ReleaseDC 0&, DC
    OldhBmp = SelectObject(hDCMemory, hBmp)
End Sub

Private Sub CreateCaptionShadow(ByVal X As Long, ByVal Y As Long, ByVal lWidth As Long, ByVal lHeight As Long)
    Dim hGraphics As Long
    Dim hPath As Long
    Dim hBrush As Long, hPen As Long
    Dim hImage As Long
    Dim RecL As RECTL
    
    If bRecreateShadowCaption = False Then Exit Sub
    If hImgCaptionShadow Then GdipDisposeImage hImgCaptionShadow: hImgCaptionShadow = 0
    If m_ShadowSize = 0 Then Exit Sub
    If UBound(m_Caption1) <= 0 Then Exit Sub
    
    GdipCreateBitmapFromScan0 lWidth, lHeight, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    GdipGetImageGraphicsContext hImage, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

    GDIP_AddPathString hGraphics, 0, 0, lWidth, lHeight, True

    hImgCaptionShadow = CreateBlurShadowImage(hImage, m_ShadowColor, m_ShadowSize * nScale, 0, 0, lWidth, lHeight)
    bRecreateShadowCaption = False
    
    GdipDeletePath hPath
    GdipDeleteGraphics hGraphics
    GdipDisposeImage hImage
    
End Sub

Private Sub CreateShadow()
    Dim hImage As Long
    Dim hGraphics As Long
    Dim hPath As Long
    Dim hBrush As Long, hPen As Long
    Dim lWidth As Long, lHeight As Long
    Dim ShadowSize As Integer
    
    If hImgShadow Then GdipDisposeImage hImgShadow: hImgShadow = 0
    If m_BackShadow = False Then Exit Sub
    
    bRecreateShadowCaption = True
        
    If m_ShadowSize = 0 Then Exit Sub
    If m_BackColorOpacity = 0 And m_Border = False Then Exit Sub
    If m_ShadowColorOpacity = 0 Then Exit Sub
   
    ShadowSize = m_ShadowSize * nScale
    lWidth = UserControl.ScaleWidth - (ShadowSize * 2)
    lHeight = UserControl.ScaleHeight - (ShadowSize * 2)
    
    GdipCreateBitmapFromScan0 lWidth, lHeight, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    GdipGetImageGraphicsContext hImage, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    hPath = RoundRectangle(0, 0, lWidth - 0, lHeight - 0, True, True)
    
    If m_BackColorOpacity > 0 Then
        GdipCreateSolidFill ConvertColor(m_ShadowColor, m_ShadowColorOpacity), hBrush
        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush
    Else
        GdipCreatePen1 ConvertColor(m_ShadowColor, m_ShadowColorOpacity), (m_BorderWidth * nScale) * 2, UnitPixel, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
    End If
        
    hImgShadow = CreateBlurShadowImage(hImage, m_ShadowColor, ShadowSize, 0, 0, lWidth, lHeight)
    
    GdipDeletePath hPath
    GdipDeleteGraphics hGraphics
    GdipDisposeImage hImage
    
End Sub


Private Sub DrawAcrylicBlur(hGraphics As Long, hPath As Long)
    Dim hBrush As Long
    Dim hImage As Long
    Dim hGraphics2 As Long, hImage2 As Long
    Dim lEffect As Long
    Dim bp As BlurParams
    Dim rcSource As RECTF

    If GdipCreateBitmapFromHBITMAP(hBmp, 0, hImage) = 0 Then
        Call GdipCreateBitmapFromScan0(UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, PixelFormat32bppARGB, ByVal 0&, hImage2)
        Call GdipGetImageGraphicsContext(hImage2, hGraphics2)

        bp.ExpandEdge = 0
        bp.Radius = 25
              
        Call GdipCreateEffect(&H633C80A4, &H482B1843, &H28BEF29E, &HD4FDC534, lEffect)
        Call GdipSetEffectParameters(lEffect, bp, Len(bp))
              
        rcSource.Width = UserControl.ScaleWidth
        rcSource.Height = UserControl.ScaleHeight
    
        Call GdipDrawImageFX(hGraphics2, hImage, rcSource, 0, lEffect, 0, UnitPixel)
        
        GdipCreateTexture hImage2, &H0, hBrush
        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush
        Call GdipDeleteGraphics(hGraphics2)
        Call GdipDisposeImage(hImage2)
    End If
End Sub

Private Function DrawHotLine(hGraphics As Long, hPath As Long, ByVal PosX As Long, ByVal PosY As Long)
    Dim hBrush As Long
    Dim X As Long, Y As Long
    Dim WW As Long, HH As Long
    Dim BW As Long
    Dim LW As Long
    Dim CL As Long
    Dim SS As Long
    
    Select Case m_BorderPosition
        Case bpOutside: BW = BorderWidth * nScale
        Case bpCenter: BW = BorderWidth / 2 * nScale
        Case bpInside
    End Select
    
    If m_BorderPosition = bpOutside Then
        If m_HotLinePosition = hlLeft Or m_HotLinePosition = hlTop Then
            X = PosX + BW
            Y = PosY + BW
            WW = UserControl.ScaleWidth - BW
            HH = UserControl.ScaleHeight - BW
        Else
            X = PosX
            Y = PosY
            WW = UserControl.ScaleWidth - BW
            HH = UserControl.ScaleHeight - BW
        End If
    Else
        X = PosX + BW
        Y = PosY + BW
        WW = UserControl.ScaleWidth - BW * 2
        HH = UserControl.ScaleHeight - BW * 2
    End If
    
    If m_BackShadow Then
        SS = ShadowSize * nScale
        If m_HotLinePosition = hlRight Or m_HotLinePosition = hlBottom Then
            X = X - SS
            Y = Y - SS
        Else
            X = X + SS
            Y = Y + SS
        End If
        If ShadowOffsetX < 0 Then X = X + Abs(ShadowOffsetX * nScale)
        If ShadowOffsetY < 0 Then Y = Y + Abs(ShadowOffsetY * nScale)
    End If
    
    LW = m_HotLineWidth * nScale
    Select Case m_HotLinePosition
        Case hlLeft: X = X + LW
        Case hlTop: Y = Y + LW
        Case hlRight: WW = WW - LW
        Case hlBottom: HH = HH - LW
    End Select
    
    If m_CallOut Then
        CL = m_coLen * nScale
        Select Case m_CallOutPosicion
            Case coLeft: If m_HotLinePosition = hlLeft Then X = X + CL
            Case coTop: If m_HotLinePosition = hlTop Then Y = Y + CL
            Case coRight: If m_HotLinePosition = hlRight Then WW = WW - CL
            Case coBottom: If m_HotLinePosition = hlBottom Then HH = HH - CL
        End Select
    End If
        
    GdipSetClipRectI hGraphics, X, Y, WW, HH, CombineModeExclude
    If m_ChangeOnMouseOver = eChangeHotlineColor And m_MouseOver Then
      GdipCreateSolidFill ConvertColor(m_ColorOnMouseOver, m_ColorOnMouseOverOpacity), hBrush
    Else
      GdipCreateSolidFill ConvertColor(m_HotLineColor, m_HotLineColorOpacity), hBrush
    End If
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipResetClip hGraphics
        
End Function

Private Sub DrawProgress(hGraphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Value As Long, ByVal Max As Long)
    Dim hPen As Long
    Dim ReqWidth As Long, ReqHeight As Long
    Dim HScale As Double, VScale As Double
    Dim MyScale As Double
    Dim imgWidth As Long
    Dim imgHeight As Long
    Dim nSize As Long
    Static Angle As Long

    If m_PictureSetWidth = 0 Then ReqWidth = UserControl.ScaleWidth \ 2 Else ReqWidth = m_PictureSetWidth \ 2
    If m_PictureSetHeight = 0 Then ReqHeight = UserControl.ScaleHeight \ 2 Else ReqHeight = m_PictureSetHeight \ 2

    MyScale = IIf(ReqHeight >= ReqWidth, ReqWidth, ReqHeight)

    ReqWidth = MyScale * nScale
    ReqHeight = MyScale * nScale

        '----------------
    If m_PictureAlignmentH = pLeft Then X = X + (m_PicturePaddingX * nScale)
    If m_PictureAlignmentH = pCenter Then X = X + (Width \ 2) - (ReqWidth \ 2) + (m_PicturePaddingX * nScale)
    If m_PictureAlignmentH = pRight Then X = X + Width - ReqWidth - (m_PicturePaddingX * nScale)
    If m_PictureAlignmentV = pTop Then Y = Y + (m_PicturePaddingY * nScale)
    If m_PictureAlignmentV = pMiddle Then Y = Y + (Height \ 2) - (ReqHeight \ 2) + (m_PicturePaddingY * nScale)
    If m_PictureAlignmentV = pBottom Then Y = Y + Height - ReqHeight - (m_PicturePaddingY * nScale)
    
    GdipCreatePen1 ConvertColor(vbBlack, 50), 3 * nScale, UnitPixel, hPen
    GdipDrawArc hGraphics, hPen, X, Y, ReqWidth, ReqHeight, 0, 360
    GdipDeletePen hPen
    GdipCreatePen1 ConvertColor(&HFFCC00, 50), 3 * nScale, UnitPixel, hPen
    '
    If Max = 0 Then
        Angle = Angle + 36
        GdipDrawArc hGraphics, hPen, X, Y, ReqWidth, ReqHeight, -90 + Angle, 60
    Else
        GdipDrawArc hGraphics, hPen, X, Y, ReqWidth, ReqHeight, -90, 360 * Value / Max
    End If
    GdipDeletePen hPen

End Sub

Private Function GDIP_AddPathString(ByVal hGraphics As Long, X As Long, Y As Long, Width As Long, Height As Long, Optional ForShadow As Boolean, Optional GetMeasureString As Boolean) As Boolean
    Dim hPath As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RECTF
    Dim layoutRect2 As RECTF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim hdc As Long

    If GdipCreatePath(&H0, hPath) = 0 Then
    
        If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
            If Not m_WordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
            If m_CaptionShowPrefix Then GdipSetStringFormatHotkeyPrefix hFormat, HotkeyPrefixShow
            GdipSetStringFormatTrimming hFormat, m_CaptionTriming
            GdipSetStringFormatAlign hFormat, m_CaptionAlignmentH
            GdipSetStringFormatLineAlign hFormat, m_CaptionAlignmentV
        End If

        GetFontStyleAndSize m_Font, lFontStyle, lFontSize
        
        If GdipCreateFontFamilyFromName(StrPtr(m_Font.Name), 0, hFontFamily) Then
            If hFontCollection Then
                If GdipCreateFontFamilyFromName(StrPtr(m_Font.Name), hFontCollection, hFontFamily) Then
                    If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
                End If
            Else
                If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
            End If
        End If
        
        If GetMeasureString Then
            Dim BB As RECTF, CF As Long, LF As Long
            
            With layoutRect
                .Left = X: .Top = Y
                .Width = Width: .Height = Height
            End With
                        
            With layoutRect2
                .Left = X: .Top = Y + lFontSize
                .Width = Width: .Height = Height
            End With
                        
            Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
            GdipMeasureString hGraphics, StrPtr(m_Caption1), -1, hFont, layoutRect, hFormat, BB, CF, LF
            GdipDeleteFont hFont
            Call GdipCreateFont(hFontFamily, lFontSize - m_SizeMinus, lFontStyle, UnitPixel, hFont)
            GdipMeasureString hGraphics, StrPtr(m_Caption2), -1, hFont, layoutRect2, hFormat, BB, CF, LF
            GdipDeleteFont hFont
                      
            X = BB.Left
            Y = BB.Top
            Width = BB.Width
            Height = BB.Height
            GdipDeleteFontFamily hFontFamily
        Else
            With layoutRect
                .Left = X + m_Caption1PaddingX * nScale: .Width = Width - (m_Caption1PaddingX * nScale) * 2
                .Top = Y + m_Caption1PaddingY * nScale: .Height = Height - (m_Caption1PaddingY * nScale) * 2
            End With
            
            With layoutRect2
                .Left = X + m_Caption2PaddingX * nScale: .Width = Width - (m_Caption2PaddingX * nScale) * 2
                .Top = Y + lFontSize + m_Caption2PaddingY * nScale: .Height = Height - (m_Caption2PaddingY * nScale) * 2
            End With
            
            If m_CaptionAngle <> 0 Then
                If ForShadow Then
                    layoutRect.Left = layoutRect.Left - (Width / 2)
                    layoutRect.Top = layoutRect.Top - (Height / 2)
                    Call GdipTranslateWorldTransform(hGraphics, (Width / 2), (Height / 2), 0)
                Else
                    layoutRect.Left = layoutRect.Left - (UserControl.ScaleWidth / 2)
                    layoutRect.Top = layoutRect.Top - (UserControl.ScaleHeight / 2)
                    Call GdipTranslateWorldTransform(hGraphics, (UserControl.ScaleWidth / 2), (UserControl.ScaleHeight / 2), 0)
                End If
                Call GdipRotateWorldTransform(hGraphics, m_CaptionAngle, 0)
            End If
            
            GdipAddPathString hPath, StrPtr(m_Caption1), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
            GdipAddPathString hPath, StrPtr(m_Caption2), -1, hFontFamily, lFontStyle, lFontSize - m_SizeMinus, layoutRect2, hFormat
            GdipDeleteStringFormat hFormat
            
            If m_ChangeColorOnClick And m_Clicked Then
                GdipCreateSolidFill ConvertColor(m_ForeColorP, IIf(ForShadow, m_ShadowColorOpacity, m_ForeColorPOpacity)), hBrush
            Else
                GdipCreateSolidFill ConvertColor(m_ForeColor, IIf(ForShadow, m_ShadowColorOpacity, m_ForeColorOpacity)), hBrush
            End If
            
            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            
            If m_CaptionBorderWidth > 0 Then
               GdipCreatePen1 ConvertColor(m_CaptionBorderColor, IIf(ForShadow, m_ShadowColorOpacity, 100)), m_CaptionBorderWidth, UnitPixel, hPen
               GdipDrawPath hGraphics, hPen, hPath
               GdipDeletePen hPen
            End If
            
            If m_CaptionAngle <> 0 Then GdipResetWorldTransform hGraphics
            
            GdipDeleteFontFamily hFontFamily
            
            If m_IconCharCode Then
                
                If GdipCreateFontFamilyFromName(StrPtr(m_IconFont.Name), 0, hFontFamily) Then
                    If GdipCreateFontFamilyFromName(StrPtr(m_IconFont.Name), hFontCollection, hFontFamily) Then
                        GdipDeletePath hPath
                        Exit Function
                    End If
                End If
                
                With layoutRect
                    .Left = X + m_IconPaddingX * nScale: .Width = Width - (m_IconPaddingX * nScale) * 2
                    .Top = Y + m_IconPaddingY * nScale: .Height = Height - (m_IconPaddingY * nScale) * 2
                End With
                
                If m_CaptionAngle <> 0 Then
                    If ForShadow Then
                        layoutRect.Left = layoutRect.Left - (Width / 2)
                        layoutRect.Top = layoutRect.Top - (Height / 2)
                        Call GdipTranslateWorldTransform(hGraphics, (Width / 2), (Height / 2), 0)
                    Else
                        layoutRect.Left = layoutRect.Left - (UserControl.ScaleWidth / 2)
                        layoutRect.Top = layoutRect.Top - (UserControl.ScaleHeight / 2)
                        Call GdipTranslateWorldTransform(hGraphics, (UserControl.ScaleWidth / 2), (UserControl.ScaleHeight / 2), 0)
                    End If
                    Call GdipRotateWorldTransform(hGraphics, m_CaptionAngle, 0)
                End If
                GetFontStyleAndSize m_IconFont, lFontStyle, lFontSize
                
                If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
                    GdipSetStringFormatAlign hFormat, m_IconAlignmentH
                    GdipSetStringFormatLineAlign hFormat, m_IconAlignmentV
                End If
                                
                GdipResetPath hPath
                GdipAddPathString hPath, StrPtr(ChrW2(m_IconCharCode)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
                GdipDeleteStringFormat hFormat
            
                GdipCreateSolidFill ConvertColor(m_IconForeColor, IIf(ForShadow, m_ShadowColorOpacity, m_IconOpacity)), hBrush
                GdipFillPath hGraphics, hBrush, hPath
                GdipDeleteBrush hBrush
                
                If m_CaptionBorderWidth > 0 Then
                   GdipCreatePen1 ConvertColor(m_CaptionBorderColor, IIf(ForShadow, m_ShadowColorOpacity, 100)), m_CaptionBorderWidth, UnitPixel, hPen
                   GdipDrawPath hGraphics, hPen, hPath
                   GdipDeletePen hPen
                End If
                
                GdipDeleteFontFamily hFontFamily
                
                If m_CaptionAngle <> 0 Then GdipResetWorldTransform hGraphics
            End If
        End If
        
        GdipDeletePath hPath
    End If

End Function

Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long)
        Dim hdc As Long
        lFontStyle = 0
        If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
        If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
        If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
        If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        
        hdc = GetDC(0&)
        lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)
        ReleaseDC 0&, hdc
End Function

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Function LoadImageFromStream(ByRef bvData() As Byte, ByRef hImage As Long) As Boolean
    On Local Error GoTo LoadImageFromStream_Error
    
    Dim IStream     As IUnknown
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, hImage) = 0 Then
            LoadImageFromStream = True
        End If
    End If

    Set IStream = Nothing
    
LoadImageFromStream_Error:
End Function

Private Function ManageGDIToken(ByVal projectHwnd As Long) As Long ' by LaVolpe
    If projectHwnd = 0& Then Exit Function
    
    Dim hwndGDIsafe     As Long                 'API window to monitor IDE shutdown
    
    Do
        hwndGDIsafe = GetParent(projectHwnd)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    ' ok, got the highest level parent, now find highest level owner
    Do
        hwndGDIsafe = GetWindow(projectHwnd, GW_OWNER)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    
    hwndGDIsafe = FindWindowEx(projectHwnd, 0&, "Static", "GDI+Safe Patch")
    If hwndGDIsafe Then
        ManageGDIToken = hwndGDIsafe    ' we already have a manager running for this VB instance
        Exit Function                   ' can abort
    End If
    
    Dim gdiSI           As GdiplusStartupInput  'GDI+ startup info
    Dim gToken          As Long                 'GDI+ instance token
    
    On Error Resume Next
    gdiSI.GdiplusVersion = 1                    ' attempt to start GDI+
    GdiplusStartup gToken, gdiSI
    If gToken = 0& Then                         ' failed to start
        If Err Then Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    Dim z_ScMem         As Long                 'Thunk base address
    Dim z_Code()        As Long                 'Thunk machine-code initialised here
    Dim nAddr           As Long                 'hwndGDIsafe prev window procedure

    Const WNDPROC_OFF   As Long = &H30          'Offset where window proc starts from z_ScMem
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const MEM_LEN       As Long = &HD4          'Byte length of thunk
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    If z_ScMem <> 0 Then                                     'Ensure the allocation succeeded
        ' we make the api window a child so we can use FindWindowEx to locate it easily
        hwndGDIsafe = CreateWindowExA(0&, "Static", "GDI+Safe Patch", WS_CHILD, 0&, 0&, 0&, 0&, projectHwnd, 0&, App.hInstance, ByVal 0&)
        If hwndGDIsafe <> 0 Then
        
            ReDim z_Code(0 To MEM_LEN \ 4 - 1)
        
            z_Code(12) = &HD231C031: z_Code(13) = &HBBE58960: z_Code(14) = &H12345678: z_Code(15) = &H3FFF631: z_Code(16) = &H74247539: z_Code(17) = &H3075FF5B: z_Code(18) = &HFF2C75FF: z_Code(19) = &H75FF2875
            z_Code(20) = &H2C73FF24: z_Code(21) = &H890853FF: z_Code(22) = &HBFF1C45: z_Code(23) = &H2287D81: z_Code(24) = &H75000000: z_Code(25) = &H443C707: z_Code(26) = &H2&: z_Code(27) = &H2C753339: z_Code(28) = &H2047B81: z_Code(29) = &H75000000
            z_Code(30) = &H2C73FF23: z_Code(31) = &HFFFFFC68: z_Code(32) = &H2475FFFF: z_Code(33) = &H681C53FF: z_Code(34) = &H12345678: z_Code(35) = &H3268&: z_Code(36) = &HFF565600: z_Code(37) = &H43892053: z_Code(38) = &H90909020: z_Code(39) = &H10C261
            z_Code(40) = &H562073FF: z_Code(41) = &HFF2453FF: z_Code(42) = &H53FF1473: z_Code(43) = &H2873FF18: z_Code(44) = &H581053FF: z_Code(45) = &H89285D89: z_Code(46) = &H45C72C75: z_Code(47) = &H800030: z_Code(48) = &H20458B00: z_Code(49) = &H89145D89
            z_Code(50) = &H81612445: z_Code(51) = &H4C4&: z_Code(52) = &HC63FF00

            z_Code(1) = 0                                                   ' shutDown mode; used internally by ASM
            z_Code(2) = zFnAddr("user32", "CallWindowProcA")                ' function pointer CallWindowProc
            z_Code(3) = zFnAddr("kernel32", "VirtualFree")                  ' function pointer VirtualFree
            z_Code(4) = zFnAddr("kernel32", "FreeLibrary")                  ' function pointer FreeLibrary
            z_Code(5) = gToken                                              ' Gdi+ token
            z_Code(10) = LoadLibrary("gdiplus")                             ' library pointer (add reference)
            z_Code(6) = GetProcAddress(z_Code(10), "GdiplusShutdown")       ' function pointer GdiplusShutdown
            z_Code(7) = zFnAddr("user32", "SetWindowLongA")                 ' function pointer SetWindowLong
            z_Code(8) = zFnAddr("user32", "SetTimer")                       ' function pointer SetTimer
            z_Code(9) = zFnAddr("user32", "KillTimer")                      ' function pointer KillTimer
        
            z_Code(14) = z_ScMem                                            ' ASM ebx start point
            z_Code(34) = z_ScMem + WNDPROC_OFF                              ' subclass window procedure location
        
            RtlMoveMemory z_ScMem, VarPtr(z_Code(0)), MEM_LEN               'Copy the thunk code/data to the allocated memory
        
            nAddr = SetWindowLong(hwndGDIsafe, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Subclass our API window
            RtlMoveMemory z_ScMem + 44, VarPtr(nAddr), 4& ' Add prev window procedure to the thunk
            gToken = 0& ' zeroize so final check below does not release it
            
            ManageGDIToken = hwndGDIsafe    ' return handle of our GDI+ manager
        Else
            VirtualFree z_ScMem, 0, MEM_RELEASE     ' failure - release memory
            z_ScMem = 0&
        End If
    Else
        VirtualFree z_ScMem, 0, MEM_RELEASE           ' failure - release memory
        z_ScMem = 0&
    End If
    
    If gToken Then GdiplusShutdown gToken       ' release token if error occurred
    
End Function

'Autor: Cobein
Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
    Dim I       As Long
    For I = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(I) = lProp Then
            ReadValue = TlsGetValue(I + 1)
            Exit Function
        End If
    Next
    ReadValue = Default
End Function

Private Function RoundRectangle(X As Long, Y As Long, Width As Long, Height As Long, Optional Inflate As Boolean, Optional nn As Boolean) As Long
    Dim mPath As Long
    Dim BCLT As Integer
    Dim BCRT As Integer
    Dim BCBR As Integer
    Dim BCBL As Integer
    Dim Xx As Long, Yy As Long
    Dim MidBorder As Long
    Dim coLen As Long
    Dim coWidth As Long
    Dim lMax As Long
    Dim coAngle  As Long

    Width = Width - 1 'Antialias pixel
    Height = Height - 1 'Antialias pixel
        
    coWidth = m_coWidth * nScale
    coLen = m_coLen * nScale
    coAngle = IIf(m_coRightTriangle, 0, coWidth / 2)

    If nn Then
        If m_BorderPosition = bpCenter Then
            coWidth = coWidth + m_BorderWidth * nScale / 2
        ElseIf m_BorderPosition = bpOutside Then
            coWidth = coWidth + m_BorderWidth * nScale
        ElseIf m_BorderPosition = bpInside Then
            coWidth = coWidth - m_BorderWidth * nScale / 2
        End If
    End If
    

    If Inflate Then MidBorder = m_BorderWidth / 2
    BCLT = GetSafeRound((m_BorderCornerLeftTop + MidBorder) * nScale, Width, Height)
    BCRT = GetSafeRound((m_BorderCornerRightTop + MidBorder) * nScale, Width, Height)
    BCBR = GetSafeRound((m_BorderCornerBottomRight + MidBorder) * nScale, Width, Height)
    BCBL = GetSafeRound((m_BorderCornerBottomLeft + MidBorder) * nScale, Width, Height)
    
    If m_CallOut Then
        Select Case m_CallOutPosicion
            Case coLeft
                X = X + coLen
                Width = Width - coLen
                lMax = Height - BCLT - BCBL
                If coWidth > lMax Then coWidth = lMax
            Case coTop
                Y = Y + coLen
                Height = Height - coLen
                lMax = Width - BCLT - BCBL
                If coWidth > lMax Then coWidth = lMax
            Case coRight
                Width = Width - coLen
                lMax = Height - BCRT - BCBR
                If coWidth > lMax Then coWidth = lMax
            Case coBottom
                Height = Height - coLen
                lMax = Width - BCBL - BCBR
                If coWidth > lMax Then coWidth = lMax
        End Select
    End If

    Call GdipCreatePath(&H0, mPath)
                    
                    
    If BCLT Then GdipAddPathArcI mPath, X, Y, BCLT * 2, BCLT * 2, 180, 90

    If m_CallOutPosicion = coTop And m_CallOut Then
        Select Case m_CallOutAlign
            Case coFirstCorner: Xx = X + BCLT
            Case coMidle: Xx = X + BCLT + ((Width - BCLT - BCRT) \ 2) - (coWidth \ 2)
            Case coSecondCorner: Xx = X + Width - coWidth - BCRT
            Case coCustomPosition: Xx = X + (m_coCustomPos * nScale)
        End Select
        
        If (Xx > Width / 2) And coAngle = 0 Then
            GdipAddPathLineI mPath, Xx, Y, Xx + coWidth, Y - coLen
            GdipAddPathLineI mPath, Xx + coWidth, Y - coLen, Xx + coWidth, Y
        Else
            If BCLT = 0 Then GdipAddPathLineI mPath, X, Y, X, Y
            GdipAddPathLineI mPath, Xx, Y, Xx + coAngle, Y - coLen
            GdipAddPathLineI mPath, Xx + coAngle, Y - coLen, Xx + coWidth, Y
        End If
    Else
        If BCLT = 0 Then GdipAddPathLineI mPath, X, Y, X + Width - BCRT, Y
    End If


    If BCRT Then GdipAddPathArcI mPath, X + Width - BCRT * 2, Y, BCRT * 2, BCRT * 2, 270, 90

    If m_CallOutPosicion = coRight And m_CallOut Then
        Select Case m_CallOutAlign
            Case coFirstCorner: Yy = Y + BCRT
            Case coMidle: Yy = Y + BCRT + ((Height - BCRT - BCBR) \ 2) - (coWidth \ 2)
            Case coSecondCorner: Yy = Y + Height - coWidth - BCBR
            Case coCustomPosition: Yy = Y + (m_coCustomPos * nScale)
        End Select
        Xx = X + Width
        If (Yy > Height / 2) And coAngle = 0 Then
            GdipAddPathLineI mPath, Xx, Yy, Xx + coLen, Yy + coWidth
            GdipAddPathLineI mPath, Xx + coLen, Yy + coWidth, Xx, Yy + coWidth
            
        Else
            If BCRT = 0 Then GdipAddPathLineI mPath, X + Width, Y, X + Width, Y
            GdipAddPathLineI mPath, Xx, Yy, Xx + coLen, Yy + coAngle
            GdipAddPathLineI mPath, Xx + coLen, Yy + coAngle, Xx, Yy + coWidth
        End If
    Else
        If BCRT = 0 Then GdipAddPathLineI mPath, X + Width, Y, X + Width, Y + Height - BCBR
    End If

    If BCBR Then GdipAddPathArcI mPath, X + Width - BCBR * 2, Y + Height - BCBR * 2, BCBR * 2, BCBR * 2, 0, 90


    If m_CallOutPosicion = coBottom And m_CallOut Then
        Select Case m_CallOutAlign
            Case coFirstCorner: Xx = X + BCBL
            Case coMidle: Xx = X + BCBL + ((Width - BCBR - BCBL) \ 2) - (coWidth \ 2)
            Case coSecondCorner: Xx = X + Width - coWidth - BCBR
            Case coCustomPosition: Xx = X + (m_coCustomPos * nScale)
        End Select
        
        Yy = Y + Height
        If (Xx > Width / 2) And coAngle = 0 Then
            GdipAddPathLineI mPath, Xx + coWidth, Yy, Xx + coWidth, Yy + coLen
            GdipAddPathLineI mPath, Xx + coWidth, Yy + coLen, Xx, Yy
        Else
            If BCBR = 0 Then GdipAddPathLineI mPath, X + Width, Y + Height, X + Width, Y + Height
            GdipAddPathLineI mPath, Xx + coWidth, Yy, Xx + coAngle, Yy + coLen
            GdipAddPathLineI mPath, Xx + coAngle, Yy + coLen, Xx, Yy
        End If
    Else
        If BCBR = 0 Then GdipAddPathLineI mPath, X + Width, Y + Height, X + BCBL, Y + Height
    End If

    If BCBL Then GdipAddPathArcI mPath, X, Y + Height - BCBL * 2, BCBL * 2, BCBL * 2, 90, 90
    
    If m_CallOutPosicion = coLeft And m_CallOut Then
        Select Case m_CallOutAlign
            Case coFirstCorner: Yy = Y + BCLT
            Case coMidle: Yy = Y + BCLT + ((Height - BCBL - BCLT) \ 2) - (coWidth \ 2)
            Case coSecondCorner: Yy = Y + Height - coWidth - BCBL
            Case coCustomPosition: Yy = Y + (m_coCustomPos * nScale)
        End Select
        
        If (Yy > Height / 2) And coAngle = 0 Then
            GdipAddPathLineI mPath, X, Yy + coWidth, X - coLen, Yy + coWidth
            GdipAddPathLineI mPath, X - coLen, Yy + coWidth, X, Yy
        Else
            If BCBL = 0 Then GdipAddPathLineI mPath, X, Y + Height, X, Y + Height
            GdipAddPathLineI mPath, X, Yy + coWidth, X - coLen, Yy + coAngle
            GdipAddPathLineI mPath, X - coLen, Yy + coAngle, X, Yy
        End If
    Else
        If BCBL = 0 Then GdipAddPathLineI mPath, X, Y + Height, X, Y + BCLT
    End If
   
    GdipClosePathFigures mPath
  
    RoundRectangle = mPath

End Function

Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

'las dos funciones a continuacion son de cobein y con algunas modificaciones mias,
'las he utilizado para crear una bandera publica sin tener que agregar un modulo publico.
Private Function WriteValue(ByVal lProp As Long, ByVal lValue As Long) As Boolean
    Dim lFlagIndex As Long
    Dim I       As Long
    Dim lIndex  As Long: lIndex = -1
    
    For I = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(I) = lProp Then
            lIndex = I + 1
            Exit For
        End If
    Next

    If lIndex = -1 Then
        Do
            lFlagIndex = TlsAlloc '// Find two consecutive slots
            lIndex = TlsAlloc
            If lIndex >= TLS_MINIMUM_AVAILABLE Then Exit Function
        Loop While Not lFlagIndex + 1 = lIndex
        Call TlsSetValue(lFlagIndex, lProp)
        Call TlsSetValue(lIndex, lValue)
        WriteValue = True
    End If
End Function

Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)  'Get the specified procedure address
End Function

Private Sub tmrMOUSEOVER_Timer()
    If Not IsMouseInExtender Then
        tmrMOUSEOVER.Interval = 0
        RaiseEvent MouseLeave
        'OnMouseOver
        m_MouseOver = False
        Refresh
    End If
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    On Error GoTo PropErr
    
    If PictureFromStream(AsyncProp.Value) Then
        RaiseEvent PictureDownloadComplete
    Else
        RaiseEvent PictureDownloadError
    End If
    
    Set c_AsyncProp = Nothing
    Exit Sub
PropErr:
    RaiseEvent PictureDownloadError
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    On Error Resume Next
    If m_DrawProgress Then
        If c_AsyncProp Is Nothing Then Set c_AsyncProp = AsyncProp
        UserControl.Refresh
    End If
    RaiseEvent PictureDownloadProgress(AsyncProp.BytesMax, AsyncProp.BytesRead)
End Sub

'===== CUSTOM EVENTS ===== ======================================= CUSTOM EVENTS ===================================== CUSTOM EVENTS ======================================

Private Sub UserControl_Click()
    RaiseEvent Click
    If m_OptionBehavior = True Then
      m_Value = True
      OptBehavior
    Else
      m_Value = Not m_Value
    End If
    
    PropertyChanged "Value"
    RaiseEvent ChangeValue(m_Value)

End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If UserControl.Enabled Then
        If Not MouseToParent Then
            HitResult = vbHitResultHit
        ElseIf Not Ambient.UserMode Then
            HitResult = vbHitResultHit
        End If
        If Ambient.UserMode Then
            If tmrMOUSEOVER.Interval = 0 Then
                tmrMOUSEOVER.Interval = 1
                RaiseEvent MouseEnter
                '----------------->
                'OnMouseOver
                m_MouseOver = True
                Refresh
            End If
        End If
    ElseIf Not Ambient.UserMode Then
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_Initialize()
    nScale = GetWindowsDPI
End Sub

Private Sub UserControl_InitProperties()
    hFontCollection = ReadValue(&HFC)
    m_BackColor = Ambient.BackColor
    m_BackColorOpacity = m_def_BackColorOpacity
    m_BackColorP = m_def_BackColorP
    m_BackColorPOpacity = m_def_BackColorPOpacity
    m_Border = m_def_Border
    m_BorderColor = m_def_BorderColor
    m_BorderColorOpacity = m_def_BorderColorOpacity
    m_ColorOnMouseOver = m_def_ColorOnMouseOver
    m_ColorOnMouseOverOpacity = m_def_ColorOnMouseOverOpacity
    m_BorderPosition = m_def_BorderPosition
    m_BorderWidth = m_def_BorderWidth
    m_Caption1 = "Cap1_" & Ambient.DisplayName
    m_Caption2 = "Cap2_" & Ambient.DisplayName
    m_SizeMinus = 3
    m_CaptionAlignmentH = m_def_CaptionAlignmentH
    m_CaptionAlignmentV = m_def_CaptionAlignmentV
    m_CaptionShadow = False
    m_CrossPosition = cTopRight
    m_CrossVisible = False
    Set m_Font = UserControl.Ambient.Font
    m_ForeColor = m_def_ForeColor
    m_ForeColorOpacity = m_def_ForeColorOpacity
    m_ForeColorP = m_def_ForeColorP
    m_ForeColorPOpacity = m_def_ForeColorPOpacity
    m_ChangeColorOnClick = m_def_ChangeColorOnClick
    m_ChangeOnMouseOver = eChangeNone
    m_Gradient = m_def_Gradient
    m_GradientAngle = m_def_GradientAngle
    m_GradientColor1 = m_def_GradientColor1
    m_GradientColor1Opacity = m_def_GradientColor1Opacity
    m_GradientColor2 = m_def_GradientColor2
    m_GradientColor2Opacity = m_def_GradientColor2Opacity
    m_PictureAlignmentH = 0
    m_PictureAlignmentV = 0
    m_PictureOpacity = m_def_PictureOpacity
    m_PicturePaddingX = 0
    m_PicturePaddingY = 0
    m_WordWrap = m_def_WordWrap
    Set m_IconFont = UserControl.Ambient.Font
    m_IconForeColor = UserControl.Ambient.ForeColor
    m_IconPaddingX = 0
    m_IconPaddingY = 0
    m_IconAlignmentH = 0
    m_IconAlignmentV = 0
    m_IconOpacity = m_def_PictureOpacity
    m_Value = m_def_Value
    m_OptionBehavior = m_def_OptionBehavior
    
    c_lhWnd = UserControl.ContainerHwnd
    Call ManageGDIToken(c_lhWnd)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
     RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hCur Then SetCursor hCur
    RaiseEvent MouseDown(Button, Shift, X, Y)
    m_Clicked = True
    Refresh
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hCur Then SetCursor hCur
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    m_Clicked = False
    Refresh
    '-------------
    If X > XCrossPos And X < XCrossPos + 6 And Y > (YCrossPos) And Y < (YCrossPos) + 6 Then
        If m_CrossVisible Then
          RaiseEvent CrossClick
          Extender.Visible = False
        End If
    End If
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_Paint()
    Dim lHdc As Long
    Dim X As Long, Y As Long
    lHdc = UserControl.hdc
    RaiseEvent PrePaint(lHdc, X, Y)
    Call Draw(lHdc, 0, X, Y)
    RaiseEvent PostPaint(UserControl.hdc)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    hFontCollection = ReadValue(&HFC)
    c_lhWnd = UserControl.ContainerHwnd
    Call ManageGDIToken(c_lhWnd)
    
    With PropBag
        m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_BackColorOpacity = .ReadProperty("BackColorOpacity", m_def_BackColorOpacity)
        m_BackColorP = .ReadProperty("BackColorPress", m_def_BackColorP)
        m_BackColorPOpacity = .ReadProperty("BackColorPressOpacity", m_def_BackColorPOpacity)
        m_BackAcrylicBlur = .ReadProperty("BackAcrylicBlur", False)
        m_BackShadow = .ReadProperty("BackShadow", False)
        m_Border = .ReadProperty("Border", m_def_Border)
        m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
        m_BorderColorOpacity = .ReadProperty("BorderColorOpacity", m_def_BorderColorOpacity)
        m_ColorOnMouseOver = .ReadProperty("ColorOnMouseOver", m_def_ColorOnMouseOver)
        m_ColorOnMouseOverOpacity = .ReadProperty("ColorOpacityOnMouseOver", m_def_ColorOnMouseOverOpacity)
        m_BorderCornerLeftTop = .ReadProperty("BorderCornerLeftTop", 0)
        m_BorderCornerRightTop = .ReadProperty("BorderCornerRightTop", 0)
        m_BorderCornerBottomRight = .ReadProperty("BorderCornerBottomRight", 0)
        m_BorderCornerBottomLeft = .ReadProperty("BorderCornerBottomLeft", 0)
        m_BorderPosition = .ReadProperty("BorderPosition", m_def_BorderPosition)
        m_BorderWidth = .ReadProperty("BorderWidth", m_def_BorderWidth)
        m_CaptionAlignmentH = .ReadProperty("CaptionAlignmentH", m_def_CaptionAlignmentH)
        m_CaptionAlignmentV = .ReadProperty("CaptionAlignmentV", m_def_CaptionAlignmentV)
        m_Caption1 = .ReadProperty("Caption1", Ambient.DisplayName)
        m_Caption2 = .ReadProperty("Caption2", Ambient.DisplayName)
        m_SizeMinus = .ReadProperty("Caption2SizeMinus", 3)
        m_Caption1PaddingX = .ReadProperty("Caption1PaddingX", 0)
        m_Caption1PaddingY = .ReadProperty("Caption1PaddingY", 0)
        m_Caption2PaddingX = .ReadProperty("Caption2PaddingX", 0)
        m_Caption2PaddingY = .ReadProperty("Caption2PaddingY", 0)
        m_CaptionTriming = .ReadProperty("CaptionTriming", StringTrimmingNone)
        m_CaptionBorderWidth = .ReadProperty("CaptionBorderWidth", 0)
        m_CaptionBorderColor = .ReadProperty("CaptionBorderColor", vbHighlightText)
        m_CaptionShadow = .ReadProperty("CaptionShadow", False)
        m_CaptionAngle = .ReadProperty("CaptionAngle", 0)
        m_CaptionShowPrefix = .ReadProperty("CaptionShowPrefix", False)
        m_CrossPosition = .ReadProperty("CrossPosition", cTop)
        m_CrossVisible = .ReadProperty("CrossVisible", False)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set m_Font = .ReadProperty("Font", UserControl.Ambient.Font)
        m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
        m_ForeColorOpacity = .ReadProperty("ForeColorOpacity", m_def_ForeColorOpacity)
        m_ForeColorP = .ReadProperty("ForeColorOnPress", m_def_ForeColorP)
        m_ForeColorPOpacity = .ReadProperty("ForeColorOnPressOpacity", m_def_ForeColorPOpacity)
        m_ChangeColorOnClick = .ReadProperty("ChangeColorOnClick", m_def_ChangeColorOnClick)
        m_ChangeOnMouseOver = .ReadProperty("ChangeOnMouseOver", eChangeBorderColor)
        m_Gradient = .ReadProperty("Gradient", m_def_Gradient)
        m_GradientAngle = .ReadProperty("GradientAngle", m_def_GradientAngle)
        m_GradientColor1 = .ReadProperty("GradientColor1", m_def_GradientColor1)
        m_GradientColor1Opacity = .ReadProperty("GradientColor1Opacity", m_def_GradientColor1Opacity)
        m_GradientColor2 = .ReadProperty("GradientColor2", m_def_GradientColor2)
        m_GradientColor2Opacity = .ReadProperty("GradientColor2Opacity", m_def_GradientColor2Opacity)
        m_GradientColorP1 = .ReadProperty("GradientColorP1", m_def_GradientColorP1)
        m_GradientColorP1Opacity = .ReadProperty("GradientColorP1Opacity", m_def_GradientColorP1Opacity)
        m_GradientColorP2 = .ReadProperty("GradientColorP2", m_def_GradientColorP2)
        m_GradientColorP2Opacity = .ReadProperty("GradientColorP2Opacity", m_def_GradientColorP2Opacity)
        m_PictureAngle = .ReadProperty("PictureAngle", 0)
        m_PictureAlignmentH = .ReadProperty("PictureAlignmentH", 0)
        m_PictureAlignmentV = .ReadProperty("PictureAlignmentV", 0)
        m_PictureOpacity = .ReadProperty("PictureOpacity", m_def_PictureOpacity)
        m_PictureBrightness = .ReadProperty("PictureBrightness", 0)
        m_PictureContrast = .ReadProperty("PictureContrast", 0)
        m_PictureGraysScale = .ReadProperty("PictureGraysScale", False)
        m_PicturePaddingX = .ReadProperty("PicturePaddingX", 0)
        m_PicturePaddingY = .ReadProperty("PicturePaddingY", 0)
        m_PictureSetWidth = .ReadProperty("PictureSetWidth", 0)
        m_PictureSetHeight = .ReadProperty("PictureSetHeight", 0)
        m_WordWrap = .ReadProperty("WordWrap", m_def_WordWrap)
        m_ShadowSize = .ReadProperty("ShadowSize", 0)
        m_ShadowColor = .ReadProperty("ShadowColor", vbBlack)
        m_ShadowOffsetX = .ReadProperty("ShadowOffsetX", 0)
        m_ShadowOffsetY = .ReadProperty("ShadowOffsetY", 0)
        m_ShadowColorOpacity = .ReadProperty("ShadowColorOpacity", 50)
        m_CallOutAlign = .ReadProperty("CallOutAlign", coMidle)
        m_CallOutPosicion = .ReadProperty("CallOutPosicion", coLeft)
        m_coWidth = .ReadProperty("CallOutWidth", 10)
        m_coLen = .ReadProperty("CallOutLen", 10)
        m_CallOut = .ReadProperty("CallOut", False)
        m_coCustomPos = .ReadProperty("CallOutCustomPosition", 0)
        m_coRightTriangle = .ReadProperty("CallOutRightTriangle", False)
        m_PictureColorize = .ReadProperty("PictureColorize", False)
        m_PictureShadow = .ReadProperty("PictureShadow", False)
        m_PictureColor = .ReadProperty("PictureColor", vbBlack)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbArrow)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        m_MousePointerHands = .ReadProperty("MousePointerHands", False)
        m_MouseToParent = .ReadProperty("MouseToParent", False)
        UserControl.OLEDropMode = .ReadProperty("OLEDropMode", 0&)
        m_HotLine = .ReadProperty("HotLine", False)
        m_HotLineColor = .ReadProperty("HotLineColor", vbHighlight)
        m_HotLineColorOpacity = .ReadProperty("HotLineColorOpacity", 100)
        m_HotLineWidth = .ReadProperty("HotLineWidth", 5&)
        m_HotLinePosition = .ReadProperty("HotLinePosition", hlBottom)
        m_Value = .ReadProperty("Value", m_def_Value)
        m_OptionBehavior = .ReadProperty("OptionBehavior", m_def_OptionBehavior)
        Set m_IconFont = .ReadProperty("IconFont", UserControl.Ambient.Font)
        m_IconCharCode = .ReadProperty("IconCharCode", 0)
        m_IconForeColor = .ReadProperty("IconForeColor", vbButtonText)
        m_IconPaddingX = .ReadProperty("IconPaddingX", 0)
        m_IconPaddingY = .ReadProperty("IconPaddingY", 0)
        m_IconAlignmentH = .ReadProperty("IconAlignmentH", 0)
        m_IconAlignmentV = .ReadProperty("IconAlignmentV", 0)
        m_IconOpacity = .ReadProperty("IconOpacity", 100)
        
        If m_MousePointerHands Then
            If Ambient.UserMode Then
                UserControl.MousePointer = vbCustom
                UserControl.MouseIcon = GetSystemHandCursor
            End If
        End If
    
        If CBool(.ReadProperty("PicturePresent", False)) Then
            m_PictureArr() = .ReadProperty("PictureArr")
            Call PictureFromStream(m_PictureArr)
        End If
        bRecreateShadowCaption = True
        CreateShadow
        If m_BackAcrylicBlur Then CreateBuffer
    End With
End Sub

Private Sub UserControl_Resize()
    If m_BackAcrylicBlur Then CreateBuffer
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    CreateShadow
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub

Private Sub UserControl_Terminate()
    If m_PictureBrush Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    If hImgShadow Then GdipDisposeImage hImgShadow
    If hImgCaptionShadow Then GdipDisposeImage hImgCaptionShadow
    If hCur Then DestroyCursor hCur
    If OldhBmp Then DeleteObject SelectObject(hDCMemory, OldhBmp): OldhBmp = 0
    If hDCMemory Then DeleteDC hDCMemory: hDCMemory = 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", m_BackColor, Ambient.BackColor)
        Call .WriteProperty("BackColorOpacity", m_BackColorOpacity, m_def_BackColorOpacity)
        Call .WriteProperty("BackColorPress", m_BackColorP, m_def_BackColorP)
        Call .WriteProperty("BackColorPressOpacity", m_BackColorPOpacity, m_def_BackColorPOpacity)
        Call .WriteProperty("BackAcrylicBlur", m_BackAcrylicBlur, False)
        Call .WriteProperty("BackShadow", m_BackShadow, False)
        Call .WriteProperty("Border", m_Border, m_def_Border)
        Call .WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
        Call .WriteProperty("BorderColorOpacity", m_BorderColorOpacity, m_def_BorderColorOpacity)
        Call .WriteProperty("ColorOnMouseOver", m_ColorOnMouseOver, m_def_ColorOnMouseOver)
        Call .WriteProperty("ColorOpacityOnMouseOver", m_ColorOnMouseOverOpacity, m_def_ColorOnMouseOverOpacity)
        Call .WriteProperty("BorderCornerLeftTop", m_BorderCornerLeftTop, 0)
        Call .WriteProperty("BorderCornerRightTop", m_BorderCornerRightTop, 0)
        Call .WriteProperty("BorderCornerBottomRight", m_BorderCornerBottomRight, 0)
        Call .WriteProperty("BorderCornerBottomLeft", m_BorderCornerBottomLeft, 0)
        Call .WriteProperty("BorderPosition", m_BorderPosition, m_def_BorderPosition)
        Call .WriteProperty("BorderWidth", m_BorderWidth, m_def_BorderWidth)
        Call .WriteProperty("CaptionAlignmentH", m_CaptionAlignmentH, m_def_CaptionAlignmentH)
        Call .WriteProperty("CaptionAlignmentV", m_CaptionAlignmentV, m_def_CaptionAlignmentV)
        Call .WriteProperty("Caption1", m_Caption1, Ambient.DisplayName)
        Call .WriteProperty("Caption2", m_Caption2, Ambient.DisplayName)
        Call .WriteProperty("Caption2SizeMinus", m_SizeMinus, 3)
        Call .WriteProperty("Caption1PaddingX", m_Caption1PaddingX, 0)
        Call .WriteProperty("Caption1PaddingY", m_Caption1PaddingY, 0)
        Call .WriteProperty("Caption2PaddingX", m_Caption2PaddingX, 0)
        Call .WriteProperty("Caption2PaddingY", m_Caption2PaddingY, 0)
        Call .WriteProperty("CaptionBorderWidth", m_CaptionBorderWidth, 0)
        Call .WriteProperty("CaptionBorderColor", m_CaptionBorderColor, vbHighlightText)
        Call .WriteProperty("CaptionShadow", m_CaptionShadow, False)
        Call .WriteProperty("CaptionAngle", m_CaptionAngle, 0)
        Call .WriteProperty("CaptionTriming", m_CaptionTriming, StringTrimmingNone)
        Call .WriteProperty("CaptionShowPrefix", m_CaptionShowPrefix, False)
        Call .WriteProperty("CrossPosition", m_CrossPosition, cTop)
        Call .WriteProperty("CrossVisible", m_CrossVisible, False)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Font", m_Font, UserControl.Ambient.Font)
        Call .WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
        Call .WriteProperty("ForeColorOpacity", m_ForeColorOpacity, m_def_ForeColorOpacity)
        Call .WriteProperty("ForeColorOnPress", m_ForeColorP, m_def_ForeColorP)
        Call .WriteProperty("ForeColorOnPressOpacity", m_ForeColorPOpacity, m_def_ForeColorPOpacity)
        Call .WriteProperty("ChangeColorOnClick", m_ChangeColorOnClick, m_def_ChangeColorOnClick)
        Call .WriteProperty("ChangeOnMouseOver", m_ChangeOnMouseOver, eChangeBorderColor)
        Call .WriteProperty("Gradient", m_Gradient, m_def_Gradient)
        Call .WriteProperty("GradientAngle", m_GradientAngle, m_def_GradientAngle)
        Call .WriteProperty("GradientColor1", m_GradientColor1, m_def_GradientColor1)
        Call .WriteProperty("GradientColor1Opacity", m_GradientColor1Opacity, m_def_GradientColor1Opacity)
        Call .WriteProperty("GradientColor2", m_GradientColor2, m_def_GradientColor2)
        Call .WriteProperty("GradientColor2Opacity", m_GradientColor2Opacity, m_def_GradientColor2Opacity)
        Call .WriteProperty("GradientColorP1", m_GradientColorP1, m_def_GradientColorP1)
        Call .WriteProperty("GradientColorP1Opacity", m_GradientColorP1Opacity, m_def_GradientColorP1Opacity)
        Call .WriteProperty("GradientColorP2", m_GradientColorP2, m_def_GradientColorP2)
        Call .WriteProperty("GradientColorP2Opacity", m_GradientColorP2Opacity, m_def_GradientColorP2Opacity)
        Call .WriteProperty("PictureAngle", m_PictureAngle, 0)
        Call .WriteProperty("PictureAlignmentH", m_PictureAlignmentH, pLeft)
        Call .WriteProperty("PictureAlignmentV", m_PictureAlignmentV, pTop)
        Call .WriteProperty("PictureOpacity", m_PictureOpacity, m_def_PictureOpacity)
        Call .WriteProperty("PictureBrightness", m_PictureBrightness, 0)
        Call .WriteProperty("PictureContrast", m_PictureContrast, 0)
        Call .WriteProperty("PictureGraysScale", m_PictureGraysScale, False)
        Call .WriteProperty("PicturePaddingX", m_PicturePaddingX, 0)
        Call .WriteProperty("PicturePaddingY", m_PicturePaddingY, 0)
        Call .WriteProperty("PictureSetWidth", m_PictureSetWidth, 0)
        Call .WriteProperty("PictureSetHeight", m_PictureSetHeight, 0)
        Call .WriteProperty("WordWrap", m_WordWrap, True)
        Call .WriteProperty("ShadowSize", m_ShadowSize, 0)
        Call .WriteProperty("ShadowColor", m_ShadowColor, vbBlack)
        Call .WriteProperty("ShadowOffsetX", m_ShadowOffsetX, 0)
        Call .WriteProperty("ShadowOffsetY", m_ShadowOffsetY, 0)
        Call .WriteProperty("ShadowColorOpacity", m_ShadowColorOpacity, 50)
        Call .WriteProperty("CallOutAlign", m_CallOutAlign, coMidle)
        Call .WriteProperty("CallOutPosicion", m_CallOutPosicion, coLeft)
        Call .WriteProperty("CallOutWidth", m_coWidth, 10)
        Call .WriteProperty("CallOutLen", m_coLen, 10)
        Call .WriteProperty("CallOut", m_CallOut, False)
        Call .WriteProperty("CallOutCustomPosition", m_coCustomPos, 0)
        Call .WriteProperty("CallOutRightTriangle", m_coRightTriangle, 0)
        Call .WriteProperty("PictureColorize", m_PictureColorize, False)
        Call .WriteProperty("PictureShadow", m_PictureShadow, False)
        Call .WriteProperty("PictureColor", m_PictureColor, False)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, vbArrow)
        Call .WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)
        Call .WriteProperty("MousePointerHands", m_MousePointerHands, False)
        Call .WriteProperty("MouseToParent", m_MouseToParent, False)
        Call .WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0&)
        Call .WriteProperty("HotLine", m_HotLine, False)
        Call .WriteProperty("HotLineColor", m_HotLineColor, vbHighlight)
        Call .WriteProperty("HotLineColorOpacity", m_HotLineColorOpacity, 100)
        Call .WriteProperty("HotLineWidth", m_HotLineWidth, 5)
        Call .WriteProperty("HotLinePosition", m_HotLinePosition, hlBottom)
        Call .WriteProperty("Value", m_Value, m_def_Value)
        Call .WriteProperty("OptionBehavior", m_OptionBehavior, m_def_OptionBehavior)
        Call .WriteProperty("IconFont", m_IconFont, UserControl.Ambient.Font)
        Call .WriteProperty("IconCharCode", m_IconCharCode, 0)
        Call .WriteProperty("IconForeColor", m_IconForeColor, vbButtonText)
        Call .WriteProperty("IconPaddingX", m_IconPaddingX, 0)
        Call .WriteProperty("IconPaddingY", m_IconPaddingY, 0)
        Call .WriteProperty("IconAlignmentH", m_IconAlignmentH, 0)
        Call .WriteProperty("IconAlignmentV", m_IconAlignmentV, 0)
        Call .WriteProperty("IconOpacity", m_IconOpacity, 100)
        
        Call .WriteProperty("PicturePresent", m_PicturePresent, False)
        If m_PicturePresent Then
            Call .WriteProperty("PictureArr", m_PictureArr, 0)
        Else
            Call .WriteProperty("PictureArr", 0)
        End If
        
    End With

End Sub

Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal NewValue As Boolean)
    Dim hGraphics As Long, hImage As Long
    Dim lWidth As Long, lHeight As Long
    Dim lDif As Long
    
    m_AutoSize = NewValue
    If m_AutoSize = False Then Exit Property
    
    GdipCreateBitmapFromScan0 UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    GdipGetImageGraphicsContext hImage, hGraphics
    
    lDif = (m_BorderWidth * 2) + IIf((m_Caption2PaddingX * 2) > (m_Caption1PaddingX * 2), (m_Caption2PaddingX * 2), (m_Caption1PaddingX * 2))
    If m_BackShadow Then lDif = lDif + (m_ShadowSize * 2)
    If m_CallOut = True And m_CallOutPosicion = coLeft Or m_CallOutPosicion = coRight Then
        lDif = lDif + m_coLen
    End If
    lDif = lDif * nScale
    
    If m_WordWrap Then
        lWidth = UserControl.ScaleWidth - lDif
    Else
        lWidth = Screen.Width
    End If
    
    GDIP_AddPathString hGraphics, 0, 0, lWidth, lHeight, False, True
    lWidth = lWidth + lDif + 1 'NO SE QUE FALLA QUE DEVO SUMAR 1
    lDif = ((m_BorderWidth * 2) + (m_Caption1PaddingY * 2))
    If m_BackShadow Then lDif = lDif + (m_ShadowSize * 2)
    If m_CallOut = True And m_CallOutPosicion = coTop Or m_CallOutPosicion = coBottom Then
        lDif = lDif + m_coLen
    End If
    lDif = lDif * nScale
    lHeight = lHeight + lDif
    
    UserControl.Size (lWidth + 1) * Screen.TwipsPerPixelX, (lHeight + 1) * Screen.TwipsPerPixelY
    
    GdipDeleteGraphics hGraphics
    GdipDisposeImage hImage
End Property

Public Property Get BackAcrylicBlur() As Boolean
    BackAcrylicBlur = m_BackAcrylicBlur
End Property

Public Property Let BackAcrylicBlur(ByVal New_Value As Boolean)
    m_BackAcrylicBlur = New_Value
    PropertyChanged "BackAcrylicBlur"
    
    If New_Value Then
        CreateBuffer
    Else
        If OldhBmp Then DeleteObject SelectObject(hDCMemory, OldhBmp): OldhBmp = 0
        If hDCMemory Then DeleteDC hDCMemory: hDCMemory = 0
    End If
    CreateShadow
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get BackColorOpacity() As Integer
    BackColorOpacity = m_BackColorOpacity
End Property

Public Property Let BackColorOpacity(ByVal New_BackColorOpacity As Integer)
    m_BackColorOpacity = New_BackColorOpacity
    SafeRange m_BackColorOpacity, 0, 100
    PropertyChanged "BackColorOpacity"
    CreateShadow
    Refresh
End Property
'<<------------------------------------------------------->>
Public Property Get BackColorPress() As OLE_COLOR
    BackColorPress = m_BackColorP
End Property

Public Property Let BackColorPress(ByVal New_BackColorP As OLE_COLOR)
    m_BackColorP = New_BackColorP
    PropertyChanged "BackColorPress"
    Refresh
End Property

Public Property Get BackColorPressOpacity() As Integer
    BackColorPressOpacity = m_BackColorPOpacity
End Property

Public Property Let BackColorPressOpacity(ByVal New_BackColorPOpacity As Integer)
    m_BackColorPOpacity = New_BackColorPOpacity
    SafeRange m_BackColorPOpacity, 0, 100
    PropertyChanged "BackColorPressOpacity"
    CreateShadow
    Refresh
End Property

Public Property Get BackShadow() As Boolean
    BackShadow = m_BackShadow
End Property

Public Property Let BackShadow(ByVal New_Value As Boolean)
    m_BackShadow = New_Value
    PropertyChanged "BackShadow"
    CreateShadow
    Refresh
End Property
'm_CrossPosition
Public Property Get CrossPosition() As CrossPos
    CrossPosition = m_CrossPosition
End Property

Public Property Let CrossPosition(ByVal New_Value As CrossPos)
    m_CrossPosition = New_Value
    PropertyChanged "CrossPosition"
    Refresh
End Property
'm_CrossVisible
Public Property Get CrossVisible() As Boolean
    CrossVisible = m_CrossVisible
End Property

Public Property Let CrossVisible(ByVal New_Value As Boolean)
    m_CrossVisible = New_Value
    PropertyChanged "CrossVisible"
    Refresh
End Property
'------------------------------------------
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    Refresh
End Property
'<<--------------------------------------------->>
Public Property Get ColorOnMouseOver() As OLE_COLOR
    ColorOnMouseOver = m_ColorOnMouseOver
End Property

Public Property Let ColorOnMouseOver(ByVal New_ColorOnMouseOver As OLE_COLOR)
    m_ColorOnMouseOver = New_ColorOnMouseOver
    PropertyChanged "ColorOnMouseOver"
    Refresh
End Property

Public Property Get BorderColorOpacity() As Integer
    BorderColorOpacity = m_BorderColorOpacity
End Property

Public Property Let BorderColorOpacity(ByVal New_BorderColorOpacity As Integer)
    m_BorderColorOpacity = New_BorderColorOpacity
    SafeRange m_BorderColorOpacity, 0, 100
    PropertyChanged "BorderColorOpacity"
    Refresh
End Property

Public Property Get ColorOpacityOnMouseOver() As Integer
    ColorOpacityOnMouseOver = m_ColorOnMouseOverOpacity
End Property

Public Property Let ColorOpacityOnMouseOver(ByVal New_ColorOnMouseOverOpacity As Integer)
    m_ColorOnMouseOverOpacity = New_ColorOnMouseOverOpacity
    SafeRange m_ColorOnMouseOverOpacity, 0, 100
    PropertyChanged "BorderColorOpacityOnMouseOver"
    Refresh
End Property

'<<--------------------------------------------->>
Public Property Get BorderCornerBottomLeft() As Integer
    BorderCornerBottomLeft = m_BorderCornerBottomLeft
End Property

Public Property Let BorderCornerBottomLeft(ByVal New_Value As Integer)
    m_BorderCornerBottomLeft = New_Value
    PropertyChanged "BorderCornerBottomLeft"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get BorderCornerBottomRight() As Integer
    BorderCornerBottomRight = m_BorderCornerBottomRight
End Property

Public Property Let BorderCornerBottomRight(ByVal New_Value As Integer)
    m_BorderCornerBottomRight = New_Value
    PropertyChanged "BorderCornerBottomRight"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get BorderCornerLeftTop() As Integer
    BorderCornerLeftTop = m_BorderCornerLeftTop
End Property

Public Property Let BorderCornerLeftTop(ByVal New_Value As Integer)
    m_BorderCornerLeftTop = New_Value
    PropertyChanged "BorderCornerLeftTop"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get BorderCornerRightTop() As Integer
    BorderCornerRightTop = m_BorderCornerRightTop
End Property

Public Property Let BorderCornerRightTop(ByVal New_Value As Integer)
    m_BorderCornerRightTop = New_Value
    PropertyChanged "BorderCornerRightTop"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    PropertyChanged "Border"
    CreateShadow
    Refresh
End Property

Public Property Get BorderPosition() As eBorderPosition
    BorderPosition = m_BorderPosition
End Property

Public Property Let BorderPosition(ByVal New_BorderPosition As eBorderPosition)
    m_BorderPosition = New_BorderPosition
    PropertyChanged "BorderPosition"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    CreateShadow
    Refresh
End Property

Public Property Get BorderWidth() As Integer
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    m_BorderWidth = New_BorderWidth
    PropertyChanged "BorderWidth"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    CreateShadow
    Refresh
End Property

Public Property Get CallOutAlign() As eCallOutAlign
    CallOutAlign = m_CallOutAlign
End Property

Public Property Let CallOutAlign(ByVal New_Value As eCallOutAlign)
    m_CallOutAlign = New_Value
    PropertyChanged "CallOutAlign"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get CallOutCustomPosition() As Long
    CallOutCustomPosition = m_coCustomPos
End Property

Public Property Let CallOutCustomPosition(ByVal New_Value As Long)
    m_coCustomPos = New_Value
    PropertyChanged "CallOutCustomPosition"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get CallOut() As Boolean
    CallOut = m_CallOut
End Property
    
Public Property Get CallOutLen() As Integer
    CallOutLen = m_coLen
End Property

Public Property Let CallOutLen(ByVal New_Value As Integer)
    m_coLen = New_Value
    PropertyChanged "CallOutLen"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Let CallOut(ByVal New_Value As Boolean)
    m_CallOut = New_Value
    PropertyChanged "CallOut"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get CallOutPosicion() As eCallOutPosition
    CallOutPosicion = m_CallOutPosicion
End Property

Public Property Let CallOutPosicion(ByVal New_Value As eCallOutPosition)
    m_CallOutPosicion = New_Value
    PropertyChanged "CallOutPosicion"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get CallOutRightTriangle() As Boolean
    CallOutRightTriangle = m_coRightTriangle
End Property

Public Property Let CallOutRightTriangle(ByVal New_Value As Boolean)
    m_coRightTriangle = New_Value
    PropertyChanged "CallOutRightTriangle"
    CreateShadow
    Refresh
End Property
    
Public Property Get CallOutWidth() As Integer
    CallOutWidth = m_coWidth
End Property

Public Property Let CallOutWidth(ByVal New_Value As Integer)
    m_coWidth = New_Value
    PropertyChanged "CallOutWidth"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get CaptionAlignmentH() As CaptionAlignmentH
    CaptionAlignmentH = m_CaptionAlignmentH
End Property

Public Property Let CaptionAlignmentH(ByVal New_CaptionAlignmentH As CaptionAlignmentH)
    m_CaptionAlignmentH = New_CaptionAlignmentH
    PropertyChanged "CaptionAlignmentH"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionAlignmentV() As CaptionAlignmentV
    CaptionAlignmentV = m_CaptionAlignmentV
End Property

Public Property Let CaptionAlignmentV(ByVal New_CaptionAlignmentV As CaptionAlignmentV)
    m_CaptionAlignmentV = New_CaptionAlignmentV
    PropertyChanged "CaptionAlignmentV"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionAngle() As Integer
    CaptionAngle = m_CaptionAngle
End Property

Public Property Let CaptionAngle(ByVal New_Value As Integer)
    m_CaptionAngle = New_Value
    PropertyChanged "CaptionAngle"
    'If hImgCaptionShadow Then GdipDisposeImage hImgCaptionShadow: hImgCaptionShadow = 0
    bRecreateShadowCaption = True
    Refresh
End Property
'Caption2SizeMinus
Public Property Get Caption2SizeMinus() As Integer
    Caption2SizeMinus = m_SizeMinus
End Property

Public Property Let Caption2SizeMinus(ByVal New_Value As Integer)
    m_SizeMinus = New_Value
    PropertyChanged "Caption2SizeMinus"
    Refresh
End Property

Public Property Get CaptionBorderColor() As OLE_COLOR
    CaptionBorderColor = m_CaptionBorderColor
End Property

Public Property Let CaptionBorderColor(ByVal New_Value As OLE_COLOR)
    m_CaptionBorderColor = New_Value
    PropertyChanged "CaptionBorderColor"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionBorderWidth() As Integer
    CaptionBorderWidth = m_CaptionBorderWidth
End Property

Public Property Let CaptionBorderWidth(ByVal New_Value As Integer)
    m_CaptionBorderWidth = New_Value
    If m_CaptionBorderWidth < 0 Then m_CaptionBorderWidth = 0
    PropertyChanged "CaptionBorderWidth"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get Caption1() As String
    Caption1 = m_Caption1
End Property

Public Property Let Caption1(ByRef New_Caption As String)
    m_Caption1 = New_Caption
    PropertyChanged "Caption1"
    If m_AutoSize Then Me.AutoSize = True
    RaiseEvent Change
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get Caption2() As String
    Caption2 = m_Caption2
End Property

Public Property Let Caption2(ByRef New_Caption As String)
    m_Caption2 = New_Caption
    PropertyChanged "Caption2"
    If m_AutoSize Then Me.AutoSize = True
    RaiseEvent Change
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get Caption1PaddingX() As Integer
    Caption1PaddingX = m_Caption1PaddingX
End Property

Public Property Let Caption1PaddingX(ByVal New_Caption1PaddingX As Integer)
    m_Caption1PaddingX = New_Caption1PaddingX
    PropertyChanged "Caption1PaddingX"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get Caption1PaddingY() As Integer
    Caption1PaddingY = m_Caption1PaddingY
End Property

Public Property Let Caption1PaddingY(ByVal New_Caption1PaddingY As Integer)
    m_Caption1PaddingY = New_Caption1PaddingY
    PropertyChanged "Caption1PaddingY"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get Caption2PaddingX() As Integer
    Caption2PaddingX = m_Caption2PaddingX
End Property

Public Property Let Caption2PaddingX(ByVal New_Caption2PaddingX As Integer)
    m_Caption2PaddingX = New_Caption2PaddingX
    PropertyChanged "Caption2PaddingX"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get Caption2PaddingY() As Integer
    Caption2PaddingY = m_Caption2PaddingY
End Property

Public Property Let Caption2PaddingY(ByVal New_Caption2PaddingY As Integer)
    m_Caption2PaddingY = New_Caption2PaddingY
    PropertyChanged "Caption2PaddingY"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionShadow() As Boolean
    CaptionShadow = m_CaptionShadow
End Property

Public Property Let CaptionShadow(ByVal New_Value As Boolean)
    m_CaptionShadow = New_Value
    PropertyChanged "CaptionShadow"
    If hImgCaptionShadow Then GdipDisposeImage hImgCaptionShadow: hImgCaptionShadow = 0
    bRecreateShadowCaption = True
    Refresh
End Property


Public Property Get CaptionShowPrefix() As Boolean
    CaptionShowPrefix = m_CaptionShowPrefix
End Property

Public Property Let CaptionShowPrefix(ByVal New_Value As Boolean)
    m_CaptionShowPrefix = New_Value
    PropertyChanged "CaptionShowPrefix"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionTriming() As StringTrimming
    CaptionTriming = m_CaptionTriming
End Property

Public Property Let CaptionTriming(ByVal New_Value As StringTrimming)
    m_CaptionTriming = New_Value
    PropertyChanged "CaptionTriming"
    Refresh
End Property
'ChangeOnMouseOver
Public Property Get ChangeOnMouseOver() As eChangeOnMouse
    ChangeOnMouseOver = m_ChangeOnMouseOver
End Property

Public Property Let ChangeOnMouseOver(ByVal New_Change As eChangeOnMouse)
    m_ChangeOnMouseOver = New_Change
    PropertyChanged "ChangeOnMouseOver"
    Refresh
End Property

Public Property Get ChangeColorOnClick() As Boolean
    ChangeColorOnClick = m_ChangeColorOnClick
End Property

Public Property Let ChangeColorOnClick(ByVal New_Change As Boolean)
    m_ChangeColorOnClick = New_Change
    PropertyChanged "ChangeColorOnClick"
    Refresh
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(New_Font As StdFont)
    With m_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With
    PropertyChanged "Font"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Refresh
End Property
''>>----------------------------------------
Public Property Get ForeColorOnPress() As OLE_COLOR
    ForeColorOnPress = m_ForeColorP
End Property

Public Property Let ForeColorOnPress(ByVal New_ForeColorP As OLE_COLOR)
    m_ForeColorP = New_ForeColorP
    PropertyChanged "ForeColorOnPress"
    Refresh
End Property

Public Property Get ForeColorOnPressOpacity() As Integer
    ForeColorOnPressOpacity = m_ForeColorOpacity
End Property

Public Property Let ForeColorOnPressOpacity(ByVal New_ForeColorPOpacity As Integer)
    m_ForeColorPOpacity = New_ForeColorPOpacity
    SafeRange m_ForeColorPOpacity, 0, 100
    PropertyChanged "ForeColorOnPressOpacity"
    Refresh
End Property

Public Property Get ForeColorOpacity() As Integer
    ForeColorOpacity = m_ForeColorOpacity
End Property

Public Property Let ForeColorOpacity(ByVal New_ForeColorOpacity As Integer)
    m_ForeColorOpacity = New_ForeColorOpacity
    SafeRange m_ForeColorOpacity, 0, 100
    PropertyChanged "ForeColorOpacity"
    Refresh
End Property

Public Property Get GradientAngle() As Integer
    GradientAngle = m_GradientAngle
End Property

Public Property Let GradientAngle(ByVal New_GradientAngle As Integer)
    m_GradientAngle = New_GradientAngle
    SafeRange m_GradientAngle, 0, 359
    PropertyChanged "GradientAngle"
    Refresh
End Property

Public Property Get GradientColor1() As OLE_COLOR
    GradientColor1 = m_GradientColor1
End Property

Public Property Let GradientColor1(ByVal New_GradientColor1 As OLE_COLOR)
    m_GradientColor1 = New_GradientColor1
    PropertyChanged "GradientColor1"
    Refresh
End Property

Public Property Get GradientColor1Opacity() As Integer
    GradientColor1Opacity = m_GradientColor1Opacity
End Property

Public Property Let GradientColor1Opacity(ByVal New_GradientColor1Opacity As Integer)
    m_GradientColor1Opacity = New_GradientColor1Opacity
    SafeRange m_GradientColor1Opacity, 0, 100
    PropertyChanged "GradientColor1Opacity"
    Refresh
End Property

Public Property Get GradientColor2() As OLE_COLOR
    GradientColor2 = m_GradientColor2
End Property

Public Property Let GradientColor2(ByVal New_GradientColor2 As OLE_COLOR)
    m_GradientColor2 = New_GradientColor2
    PropertyChanged "GradientColor2"
    Refresh
End Property

Public Property Get GradientColor2Opacity() As Integer
    GradientColor2Opacity = m_GradientColor2Opacity
End Property

Public Property Let GradientColor2Opacity(ByVal New_GradientColor2Opacity As Integer)
    m_GradientColor2Opacity = New_GradientColor2Opacity
    SafeRange m_GradientColor2Opacity, 0, 100
    PropertyChanged "GradientColor2Opacity"
    Refresh
End Property

Public Property Get GradientColorP1() As OLE_COLOR
    GradientColorP1 = m_GradientColorP1
End Property

Public Property Let GradientColorP1(ByVal New_GradientColorP1 As OLE_COLOR)
    m_GradientColorP1 = New_GradientColorP1
    PropertyChanged "GradientColorP1"
    Refresh
End Property

Public Property Get GradientColorP1Opacity() As Integer
    GradientColorP1Opacity = m_GradientColorP1Opacity
End Property

Public Property Let GradientColorP1Opacity(ByVal New_GradientColorP1Opacity As Integer)
    m_GradientColorP1Opacity = New_GradientColorP1Opacity
    SafeRange m_GradientColorP1Opacity, 0, 100
    PropertyChanged "GradientColorP1Opacity"
    Refresh
End Property

Public Property Get GradientColorP2() As OLE_COLOR
    GradientColorP2 = m_GradientColorP2
End Property

Public Property Let GradientColorP2(ByVal New_GradientColorP2 As OLE_COLOR)
    m_GradientColorP2 = New_GradientColorP2
    PropertyChanged "GradientColorP2"
    Refresh
End Property

Public Property Get GradientColorP2Opacity() As Integer
    GradientColorP2Opacity = m_GradientColorP2Opacity
End Property

Public Property Let GradientColorP2Opacity(ByVal New_GradientColorP2Opacity As Integer)
    m_GradientColorP2Opacity = New_GradientColorP2Opacity
    SafeRange m_GradientColorP2Opacity, 0, 100
    PropertyChanged "GradientColorP2Opacity"
    Refresh
End Property

Public Property Get Gradient() As Boolean
    Gradient = m_Gradient
End Property

Public Property Let Gradient(ByVal New_Gradient As Boolean)
    m_Gradient = New_Gradient
    PropertyChanged "Gradient"
    Refresh
End Property

Public Property Get HotLineColor() As OLE_COLOR
    HotLineColor = m_HotLineColor
End Property

Public Property Let HotLineColor(ByVal New_Value As OLE_COLOR)
    m_HotLineColor = New_Value
    PropertyChanged "HotLineColor "
    Refresh
End Property

Public Property Get HotLineColorOpacity() As Integer
    HotLineColorOpacity = m_HotLineColorOpacity
End Property

Public Property Let HotLineColorOpacity(ByVal New_Value As Integer)
    m_HotLineColorOpacity = New_Value
    SafeRange m_HotLineColorOpacity, 0, 100
    PropertyChanged "HotLineColorOpacity "
    Refresh
End Property

Public Property Get HotLine() As Boolean
    HotLine = m_HotLine
End Property

Public Property Let HotLine(ByVal New_Value As Boolean)
    m_HotLine = New_Value
    PropertyChanged "HotLine"
    Refresh
End Property

Public Property Get HotLinePosition() As HotLinePosition
    HotLinePosition = m_HotLinePosition
End Property

Public Property Let HotLinePosition(ByVal New_Value As HotLinePosition)
    m_HotLinePosition = New_Value
    PropertyChanged "HotLinePosition"
    Refresh
End Property

Public Property Get HotLineWidth() As Integer
    HotLineWidth = m_HotLineWidth
End Property

Public Property Let HotLineWidth(ByVal New_Value As Integer)
    m_HotLineWidth = New_Value
    PropertyChanged "HotLineWidth"
    Refresh
End Property

Public Property Get IconAlignmentH() As CaptionAlignmentH
    IconAlignmentH = m_IconAlignmentH
End Property

Public Property Let IconAlignmentH(ByVal New_IconAlignmentH As CaptionAlignmentH)
    m_IconAlignmentH = New_IconAlignmentH
    PropertyChanged "IconAlignmentH"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get IconAlignmentV() As CaptionAlignmentV
    IconAlignmentV = m_IconAlignmentV
End Property

Public Property Let IconAlignmentV(ByVal New_IconAlignmentV As CaptionAlignmentV)
    m_IconAlignmentV = New_IconAlignmentV
    PropertyChanged "IconAlignmentV"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get IconCharCode() As String
    IconCharCode = "&H" & Hex(m_IconCharCode)
End Property

Public Property Let IconCharCode(ByVal New_IconCharCode As String)
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not Left(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        m_IconCharCode = "&H" & New_IconCharCode
    Else
        m_IconCharCode = New_IconCharCode
    End If
    PropertyChanged "IconCharCode"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property

Public Property Set IconFont(New_Font As StdFont)
    With m_IconFont
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With
    PropertyChanged "IconFont"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property

Public Property Let IconForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "IconForeColor"
    Refresh
End Property

Public Property Get IconOpacity() As Integer
    IconOpacity = m_IconOpacity
End Property

Public Property Let IconOpacity(ByVal New_IconOpacity As Integer)
    m_IconOpacity = New_IconOpacity
    SafeRange m_IconOpacity, 0, 100
    PropertyChanged "IconOpacity"
    Refresh
End Property

Public Property Get IconPaddingX() As Integer
    IconPaddingX = m_IconPaddingX
End Property

Public Property Let IconPaddingX(ByVal New_IconPaddingX As Integer)
    m_IconPaddingX = New_IconPaddingX
    PropertyChanged "IconPaddingX"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get IconPaddingY() As Integer
    IconPaddingY = m_IconPaddingY
End Property

Public Property Let IconPaddingY(ByVal New_IconPaddingY As Integer)
    m_IconPaddingY = New_IconPaddingY
    PropertyChanged "IconPaddingY"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get Image() As IPicture
    Dim DC As Long, TempHdc As Long
    Dim hBmp As Long, OldhBmp As Long
    Dim Pic As PicBmp, IID_IDispatch As GUID
    
    DC = GetDC(0)
    TempHdc = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, UserControl.ScaleWidth, UserControl.ScaleHeight)
    ReleaseDC 0&, DC
    OldhBmp = SelectObject(TempHdc, hBmp)
    
    BitBlt TempHdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hdc, 0, 0, vbSrcCopy
    
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
  
    With Pic
        .Size = Len(Pic)
        .type = vbPicTypeBitmap
        .hBmp = hBmp
        .hpal = 0
    End With

    If OldhBmp Then Call SelectObject(hDCMemory, OldhBmp): OldhBmp = 0
    If TempHdc Then DeleteDC TempHdc: TempHdc = 0

    Call OleCreatePictureIndirect(Pic, IID_IDispatch, 1, Image)

End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal NewValue As IPictureDisp)
    UserControl.MouseIcon = NewValue
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Get MousePointerHands() As Boolean
    MousePointerHands = m_MousePointerHands
End Property

Public Property Let MousePointerHands(ByVal NewValue As Boolean)
    m_MousePointerHands = NewValue
    If NewValue Then
        If Ambient.UserMode Then
            UserControl.MousePointer = vbCustom
            UserControl.MouseIcon = GetSystemHandCursor
        End If
    Else
        If hCur Then DestroyCursor hCur: hCur = 0
        UserControl.MousePointer = vbDefault
        UserControl.MouseIcon = Nothing
    End If
    PropertyChanged "MousePointerHands"
End Property

Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
    UserControl.MousePointer = NewValue
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseToParent() As Boolean
    MouseToParent = m_MouseToParent
End Property

Public Property Let MouseToParent(ByVal New_Value As Boolean)
    m_MouseToParent = New_Value
    PropertyChanged "MouseToParent"
End Property
Public Property Get OLEDropMode() As OLEDropConstants
    OLEDropMode = UserControl.OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Public Property Let OLEDropMode(ByVal New_Value As OLEDropConstants)
    UserControl.OLEDropMode = New_Value
End Property
'-------------------------------------->
Public Property Get OptionBehavior() As Boolean
   OptionBehavior = m_OptionBehavior
End Property

Public Property Let OptionBehavior(ByVal bOptionBehavior As Boolean)
   m_OptionBehavior = bOptionBehavior
   PropertyChanged "OptionBehavior"
End Property
'-------------------------------------->
Public Property Get PictureAlignmentH() As PictureAlignmentH
    PictureAlignmentH = m_PictureAlignmentH
End Property

Public Property Let PictureAlignmentH(ByVal New_PictureAlignmentH As PictureAlignmentH)
    m_PictureAlignmentH = New_PictureAlignmentH
    PropertyChanged "PictureAlignmentH"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureAlignmentV() As PictureAlignmentV
    PictureAlignmentV = m_PictureAlignmentV
End Property

Public Property Let PictureAlignmentV(ByVal New_PictureAlignmentV As PictureAlignmentV)
    m_PictureAlignmentV = New_PictureAlignmentV
    PropertyChanged "PictureAlignmentV"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureAngle() As Integer
    PictureAngle = m_PictureAngle
End Property

Public Property Let PictureAngle(ByVal New_Value As Integer)
    m_PictureAngle = New_Value
    PropertyChanged "PictureAngle"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureBrightness() As Long
    PictureBrightness = m_PictureBrightness
End Property

Public Property Let PictureBrightness(ByVal New_Value As Long)
    m_PictureBrightness = New_Value
    SafeRange m_PictureBrightness, -100, 100
    PropertyChanged "PictureBrightness"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureColor() As OLE_COLOR
    PictureColor = m_PictureColor
End Property

Public Property Get PictureColorize() As Boolean
    PictureColorize = m_PictureColorize
End Property

Public Property Let PictureColorize(ByVal New_Value As Boolean)
    m_PictureColorize = New_Value
    PropertyChanged "PictureColorize"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Let PictureColor(ByVal New_Value As OLE_COLOR)
    m_PictureColor = New_Value
    PropertyChanged "PictureColor"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureContrast() As Long
    PictureContrast = m_PictureContrast
End Property

Public Property Let PictureContrast(ByVal New_Value As Long)
    m_PictureContrast = New_Value
    SafeRange m_PictureContrast, -100, 100
    PropertyChanged "PictureContrast"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureExist() As Boolean
    PictureExist = m_PicturePresent
End Property

Public Property Get PictureGetHeight() As Long
    PictureGetHeight = m_PictureRealHeight
End Property

Public Property Get PictureGetWidth() As Long
    PictureGetWidth = m_PictureRealWidth
End Property

Public Property Get PictureGrayScale() As Boolean
    PictureGrayScale = m_PictureGraysScale
End Property

Public Property Let PictureGrayScale(ByVal New_Value As Boolean)
    m_PictureGraysScale = New_Value
    PropertyChanged "PictureGrayScale"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureOpacity() As Integer
    PictureOpacity = m_PictureOpacity
End Property

Public Property Let PictureOpacity(ByVal New_PictureOpacity As Integer)
    m_PictureOpacity = New_PictureOpacity
    SafeRange m_PictureOpacity, 0, 100
    PropertyChanged "PictureOpacity"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PicturePaddingX() As Integer
    PicturePaddingX = m_PicturePaddingX
End Property

Public Property Let PicturePaddingX(ByVal New_PicturePaddingX As Integer)
    m_PicturePaddingX = New_PicturePaddingX
    PropertyChanged "PicturePaddingX"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PicturePaddingY() As Integer
    PicturePaddingY = m_PicturePaddingY
End Property

Public Property Let PicturePaddingY(ByVal New_PicturePaddingY As Integer)
    m_PicturePaddingY = New_PicturePaddingY
    PropertyChanged "PicturePaddingY"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureSetHeight() As Long
    PictureSetHeight = m_PictureSetHeight
End Property

Public Property Let PictureSetHeight(ByVal New_Value As Long)
    m_PictureSetHeight = New_Value
    PropertyChanged "PictureSetHeight"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureSetWidth() As Long
    PictureSetWidth = m_PictureSetWidth
End Property

Public Property Let PictureSetWidth(ByVal New_Value As Long)
    m_PictureSetWidth = New_Value
    PropertyChanged "PictureSetWidth"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureShadow() As Boolean
    PictureShadow = m_PictureShadow
End Property

Public Property Let PictureShadow(ByVal New_Value As Boolean)
    m_PictureShadow = New_Value
    PropertyChanged "PictureShadow"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(ByVal New_Value As OLE_COLOR)
    m_ShadowColor = New_Value
    PropertyChanged "ShadowColor"
    If (m_PictureBrush <> 0) And (m_PictureShadow = True) Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    CreateShadow
    Refresh
End Property

Public Property Get ShadowColorOpacity() As Integer
    ShadowColorOpacity = m_ShadowColorOpacity
End Property

Public Property Let ShadowColorOpacity(ByVal New_Value As Integer)
    m_ShadowColorOpacity = New_Value
    SafeRange m_ShadowColorOpacity, 0, 100
    PropertyChanged "ShadowColorOpacity"
    If (m_PictureBrush <> 0) And (m_PictureShadow = True) Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    CreateShadow
    Refresh
End Property

Public Property Get ShadowOffsetX() As Integer
    ShadowOffsetX = m_ShadowOffsetX
End Property

Public Property Let ShadowOffsetX(ByVal New_Value As Integer)
    m_ShadowOffsetX = New_Value
    PropertyChanged "ShadowOffsetX"
    If (m_PictureBrush <> 0) And (m_PictureShadow = True) Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get ShadowOffsetY() As Integer
    ShadowOffsetY = m_ShadowOffsetY
End Property

Public Property Let ShadowOffsetY(ByVal New_Value As Integer)
    m_ShadowOffsetY = New_Value
    PropertyChanged "ShadowOffsetY"
    If (m_PictureBrush <> 0) And (m_PictureShadow = True) Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get ShadowSize() As Integer
    ShadowSize = m_ShadowSize
End Property

Public Property Let ShadowSize(ByVal New_Value As Integer)
    m_ShadowSize = New_Value
    SafeRange m_ShadowSize, 0, 100
    PropertyChanged "ShadowSize"
    CreateShadow
    If (m_PictureBrush <> 0) And (m_PictureShadow = True) Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Boolean)
  If m_OptionBehavior And NewValue Then OptBehavior
    m_Value = NewValue
    PropertyChanged "Value"
    RaiseEvent ChangeValue(m_Value)
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    m_WordWrap = New_WordWrap
    PropertyChanged "WordWrap"
    Refresh
End Property



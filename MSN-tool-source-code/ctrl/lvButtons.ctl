VERSION 5.00
Begin VB.UserControl lvButtons_H 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ToolboxBitmap   =   "lvButtons.ctx":0000
End
Attribute VB_Name = "lvButtons_H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
'Please read before making use of this code!
'Disclaimer: This is illegal if executed on real victims and could land you in prison for sure.
'This is intended for educational purposes only. We take no responsibility at all for your actions.
'This code is provided by EEEDS Eagle Eye Digital Security (Oman) for education purpose only.
'For more educational source codes please visit us http://www.digi77.com
'Author of this code W. Al Maawali Founder of  Eagle Eye Digital Solutions and Oman0.net can be reached via warith@digi77.com .

'Sharing knowledge is not about giving people something, or getting something from them.
'That is only valid for information sharing.
'Sharing knowledge occurs when people are genuinely interested in helping one another develop new capacities for action;
'it is about creating learning processes.
'Peter Senge
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!


Option Explicit
Option Compare Text
' See the Readme.html file provided
' This button control was inspired by Gonchuki's Chameleon Button v1.x. The
' 1st three versions of this control were based off of his control but
' eventually scrapped because of memory leaks, faulty logic and some buggy
' code. Any code in this version with the exception of the following routines
' that are similar to Gonchuki's control are coincidence. The ShadeColor and
' the Step calculations in DrawButtonBackground routines for XP colors are
' formulas found in Gonchuki's v1.x and credit goes to
' him & Ghuran Kartal for the formula.
'/////// Public Events sent back to the parent container
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOnButton(OnButton As Boolean)
Public Event Click()
Attribute Click.VB_MemberFlags = "200"
Public Event DoubleClick(Button As Integer)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
' GDI32 Function Calls
' =====================================================================
' DC manipulation
' Shape Forming functions
Private Const RGN_DIFF                    As Integer = 4
' Other drawing functions
' KERNEL32 Function Calls
' =====================================================================
' USER32 Function Calls
' =====================================================================
' General Windows related functions
' Standard TYPE Declarations used
' =====================================================================
Private Type POINTAPI                ' general use. Typically used for cursor location
    X                                       As Long
    Y                                       As Long
End Type
Private Type RECT                    ' used to set/ref boundaries of a rectangle
    Left                                    As Long
    Top                                     As Long
    Right                                   As Long
    Bottom                                  As Long
End Type
Private Type BITMAP                  ' used to determine if an image is a bitmap
    bmType                                  As Long
    bmWidth                                 As Long
    bmHeight                                As Long
    bmWidthBytes                            As Long
    bmPlanes                                As Integer
    bmBitsPixel                             As Integer
    bmBits                                  As Long
End Type
Private Type ICONINFO                ' used to determine if image is an icon
    fIcon                                   As Long
    xHotSpot                                As Long
    yHotSpot                                As Long
    hbmMask                                 As Long
    hbmColor                                As Long
End Type
Private Type LOGFONT               ' used to create fonts
    lfHeight                                As Long
    lfWidth                                 As Long
    lfEscapement                            As Long
    lfOrientation                           As Long
    lfWeight                                As Long
    lfItalic                                As Byte
    lfUnderline                             As Byte
    lfStrikeOut                             As Byte
    lfCharSet                               As Byte
    lfOutPrecision                          As Byte
    lfClipPrecision                         As Byte
    lfQuality                               As Byte
    lfPitchAndFamily                        As Byte
    lfFaceName                              As String * 32
End Type
' Custom TYPE Declarations used
' =====================================================================
Private Type ButtonDCInfo    ' used to manage the drawing DC
    hDC                                     As Long
    OldBitmap                               As Long
    OldPen                                  As Long
    OldBrush                                As Long
    ClipRgn                                 As Long
    OldFont                                 As Long
End Type
Private Type ButtonProperties
    bCaption                                As String
    bCaptionAlign                           As AlignmentConstants
    bCaptionStyle                           As CaptionEffectConstants
    bBackStyle                              As BackStyleConstants
    bStatus                                 As Integer
    bShape                                  As ButtonStyleConstants
    bSegPts                                 As POINTAPI
    bRect                                   As RECT
    bShowFocus                              As Boolean
    bBackHover                              As Long
    bForeHover                              As Long
    bLockHover                              As HoverLockConstants
    bGradient                               As GradientConstants
    bGradientColor                          As Long
    bMode                                   As ButtonModeConstants
    bValue                                  As Boolean
End Type
Private Type ImageProperties
    Image                                   As StdPicture
    Align                                   As ImagePlacementConstants
    Size                                    As Integer
    iRect                                   As RECT
    SourceSize                              As POINTAPI
Type                                    As Long
End Type
' Standard CONSTANTS as Constants or Enumerators
' =====================================================================
Private Const WHITENESS                   As Long = &HFF0062
Private Const CI_BITMAP                   As Long = &H0
Private Const CI_ICON                     As Long = &H1
Private Const WM_KEYDOWN                  As Long = &H100
' //////////// Custom Colors \\\\\\\\\\\\\\\\\
Private Const vbGray                      As Long = 8421504
' //////////// DrawText API Constants \\\\\\\\\\\\\\
Private Const DT_CALCRECT                 As Long = &H400
Private Const DT_CENTER                   As Long = &H1
Private Const DT_LEFT                     As Long = &H0
Private Const DT_RIGHT                    As Long = &H2
Private Const DT_WORDBREAK                As Long = &H10
' ///////////////// PROJECT-WIDE VARIABLES \\\\\\\\\\\\\\
Private ButtonDC                          As ButtonDCInfo
Private myProps                           As ButtonProperties
Private myImage                           As ImageProperties
Private bNoRefresh                        As Boolean
Private curBackColor                      As Long
Private adjBackColorUp                    As Long
Private adjBackColorDn                    As Long
Private adjHoverColor                     As Long
Private mButton                           As Integer
Private bTimerActive                      As Boolean
Private cParentBC                         As Long
Private cCheckBox                         As Long
Private bKeyDown                          As Boolean
' Custom CONSTANTS as Constants or Enumerators
' =====================================================================
' ////////////// Used to set/reset HDC objects \\\\\\\\\\\\\\
Private Enum ColorObjects
    cObj_Brush = 0
    cObj_Pen = 1
    cObj_Text = 2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private cObj_Brush, cObj_Pen, cObj_Text
#End If
' ////////////// Button Properties \\\\\\\\\\\\\\\
Public Enum ImagePlacementConstants ' image alignment
    lv_LeftEdge = 0
    lv_LeftOfCaption = 1
    lv_RightEdge = 2
    lv_RightOfCaption = 3
    lv_TopCenter = 4
    lv_BottomCenter = 5
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private lv_LeftEdge, lv_LeftOfCaption, lv_RightEdge, lv_RightOfCaption
Private lv_TopCenter, lv_BottomCenter
#End If

Public Enum ImageSizeConstants      ' image sizes
    lv_16x16 = 0
    lv_24x24 = 1
    lv_32x32 = 2
    lv_Fill_Stretch = 3
    lv_Fill_ScaleUpDown = 4
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private lv_16x16, lv_24x24, lv_32x32, lv_Fill_Stretch, lv_Fill_ScaleUpDown
#End If
Public Enum ButtonModeConstants
    lv_CommandButton = 0
    lv_CheckBox = 1
    lv_OptionButton = 2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private lv_CommandButton
Private lv_CheckBox, lv_OptionButton
#End If
Public Enum ButtonStyleConstants    ' button shapes
    lv_Rectangular = 0
    lv_LeftDiagonal = 1
    lv_RightDiagonal = 2
    lv_FullDiagonal = 3
    lv_Round3D = 4                  ' border changes gradients when clicked
    lv_Round3DFixed = 5             ' border does not change gradients
    lv_RoundFlat = 6                ' 1-pixel black border
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private lv_Rectangular, lv_LeftDiagonal, lv_RightDiagonal, lv_FullDiagonal
Private lv_Round3D, lv_Round3DFixed, lv_RoundFlat
#End If
Public Enum HoverLockConstants      ' hover lock options
    lv_LockTextandBackColor = 0
    lv_LockTextColorOnly = 1
    lv_LockBackColorOnly = 2
    lv_NoLocks = 3
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private lv_LockTextandBackColor, lv_LockTextColorOnly
Private lv_LockBackColorOnly, lv_NoLocks
#End If
Public Enum GradientConstants       ' gradient directions
    lv_NoGradient = 0
    lv_Left2Right = 1
    lv_Right2Left = 2
    lv_Top2Bottom = 3
    lv_Bottom2Top = 4
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private lv_NoGradient, lv_Left2Right, lv_Right2Left, lv_Top2Bottom
Private lv_Bottom2Top
#End If
Public Enum CaptionEffectConstants  ' caption styles
    lv_default = 0
    lv_Sunken = 1
    lv_Raised = 2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private lv_default, lv_Sunken, lv_Raised
#End If
Public Enum FontStyles
    lv_PlainStyle = 0
    lv_Bold = 2
    lv_Italic = 4
    lv_Underline = 8
    lv_BoldItalic = 2 Or 4
    lv_BoldUnderline = 2 Or 8
    lv_ItalicUnderline = 4 Or 8
    lv_BoldItalicUnderline = 2 Or 4 Or 8
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private lv_PlainStyle, lv_Bold, lv_Italic, lv_Underline
Private lv_BoldItalic, lv_BoldUnderline, lv_ItalicUnderline, lv_BoldItalicUnderline
#End If
Public Enum BackStyleConstants      ' button styles
    lv_w95 = 0
    lv_w31 = 1
    lv_XP = 2
    lv_Java = 3
    lv_Flat = 4
    lv_Hover = 5
    lv_Netscape = 6
    lv_Macintosh = 7
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private lv_w95, lv_w31
Private lv_XP, lv_Java, lv_Flat, lv_Hover, lv_Netscape, lv_Macintosh
#End If
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Integer
Private Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, _
                                                 ByVal nMapMode As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
                                                 ByVal hSrcRgn1 As Long, _
                                                 ByVal hSrcRgn2 As Long, _
                                                 ByVal nCombineMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, _
                                                        ByVal Y1 As Long, _
                                                        ByVal X2 As Long, _
                                                        ByVal Y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, _
                                                       ByVal nCount As Long, _
                                                       ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
                                                    ByVal Y1 As Long, _
                                                    ByVal X2 As Long, _
                                                    ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, _
                                                    ByVal hRgn As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hDC As Long, _
                                          ByVal nLeftRect As Long, _
                                          ByVal nTopRect As Long, _
                                          ByVal nRightRect As Long, _
                                          ByVal nBottomRect As Long, _
                                          ByVal nXStartArc As Long, _
                                          ByVal nYStartArc As Long, _
                                          ByVal nXEndArc As Long, _
                                          ByVal nYEndArc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, _
                                                   ByVal nHeight As Long, _
                                                   ByVal nPlanes As Long, _
                                                   ByVal nBitCount As Long, _
                                                   lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal dwRop As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, _
                                               lpPoint As POINTAPI, _
                                               ByVal nCount As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, _
                                                ByVal X1 As Long, _
                                                ByVal Y1 As Long, _
                                                ByVal X2 As Long, _
                                                ByVal Y2 As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, _
                                                    ByVal hPalette As Long, _
                                                    ByVal bForceBackground As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, _
                                                 ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, _
                                                ByVal nBkMode As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, _
                                                   ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, _
                                                 ByVal X As Long, _
                                                 ByVal Y As Long, _
                                                 ByVal nWidth As Long, _
                                                 ByVal nHeight As Long, _
                                                 ByVal hSrcDC As Long, _
                                                 ByVal xSrc As Long, _
                                                 ByVal ySrc As Long, _
                                                 ByVal nSrcWidth As Long, _
                                                 ByVal nSrcHeight As Long, _
                                                 ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal ByteLen As Long)
Private Declare Function CopyImage Lib "user32" (ByVal HANDLE As Long, _
                                                 ByVal imageType As Long, _
                                                 ByVal newWidth As Long, _
                                                 ByVal newHeight As Long, _
                                                 ByVal lFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
                                                  ByVal xLeft As Long, _
                                                  ByVal yTop As Long, _
                                                  ByVal hIcon As Long, _
                                                  ByVal cxWidth As Long, _
                                                  ByVal cyWidth As Long, _
                                                  ByVal istepIfAniCur As Long, _
                                                  ByVal hbrFlickerFreeDraw As Long, _
                                                  ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, _
                                                   piconinfo As ICONINFO) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, _
                                                                ByVal lpString As String) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
                                                 ByVal nIDEvent As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
                                                 ByVal hDC As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, _
                                                                      ByVal lpString As String) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, _
                                                      lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, _
                                                                ByVal lpString As String, _
                                                                ByVal hData As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, _
                                                    ByVal hRgn As Long, _
                                                    ByVal bRedraw As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, _
                                                       ByVal yPoint As Long) As Long

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Button back color. See also ResetDefaultColors"

    BackColor = curBackColor

End Property

Public Property Let BackColor(nColor As OLE_COLOR)

  ' Sets the backcolor of the button

    curBackColor = ConvertColor(nColor)
    '<:-)Auto-inserted With End...With Structure
    With myProps
        If .bLockHover = lv_LockBackColorOnly Or .bLockHover = lv_LockTextandBackColor Then
            If .bGradient Then
                Me.HoverBackColor = .bGradientColor
             Else
                Me.HoverBackColor = nColor
            End If
        End If
    End With 'myProps
    GetGDIMetrics "BackColor"
    Refresh
    PropertyChanged "cBack"

End Property

Public Property Get ButtonShape() As ButtonStyleConstants
Attribute ButtonShape.VB_Description = "Rectangular or various diagonal shapes"

    ButtonShape = myProps.bShape

End Property

Public Property Let ButtonShape(nShape As ButtonStyleConstants)

  ' Sets the button's shape (rectangular, diagonal, or circular)

    If nShape < lv_Rectangular Or nShape > lv_RoundFlat Then
        Exit Property
    End If
    '<:-)Auto-inserted With End...With Structure
    With myProps
        .bShape = nShape
        If .bCaptionAlign <> vbCenter Then
            .bCaptionAlign = vbCenter
        End If
    End With 'myProps
    Call UserControl_Resize
    myProps.bCaptionAlign = Me.CaptionAlign
    DelayDrawing False
    PropertyChanged "Shape"

End Property

Public Property Get ButtonStyle() As BackStyleConstants
Attribute ButtonStyle.VB_Description = "Various operating system button styles"

    ButtonStyle = myProps.bBackStyle

End Property

Public Property Let ButtonStyle(Style As BackStyleConstants)

  ' Sets the style of button to be displayed

    If Style < 0 Or Style > 7 Then
        Exit Property
    End If
    myProps.bBackStyle = Style
    CreateButtonRegion                 ' re-create the button shape
    CalculateBoundingRects             ' recalculate the text/image bounding rectangles
    GetGDIMetrics "BackColor"          ' cache base colors
    Refresh
    PropertyChanged "BackStyle"

End Property

Private Sub CalculateBoundingRects()

  ' Routine measures and places the rectangles to draw
  ' the caption and image on the control. The results
  ' are cached so this routine doesn't need to run
  ' every time the button is redrawn/painted
  
  Dim cRect         As RECT
  Dim tRect         As RECT
  Dim iRect         As RECT

    '<:-):UPDATED: Multiple Dim line separated
  Dim imgOffset     As RECT
  Dim bImgWidthAdj  As Boolean
  Dim bImgHeightAdj As Boolean
    '<:-):UPDATED: Multiple Dim line separated
  Dim rEdge         As Long
  Dim lEdge         As Long
  Dim adjWidth      As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim sCaption      As String
  Dim ratio(0 To 1) As Single
    ' calculations needed for diagonal buttons
    Select Case myProps.bShape
     Case lv_RightDiagonal
        rEdge = myProps.bSegPts.Y + ((ScaleWidth - myProps.bSegPts.Y) \ 3)
        adjWidth = rEdge
     Case lv_LeftDiagonal
        lEdge = myProps.bSegPts.X - (myProps.bSegPts.X \ 3) + 3
        rEdge = UserControl.ScaleWidth
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        adjWidth = UserControl.ScaleWidth - lEdge
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
     Case lv_FullDiagonal
        lEdge = myProps.bSegPts.X - (myProps.bSegPts.X \ 3) + 3
        rEdge = myProps.bSegPts.Y + ((ScaleWidth - myProps.bSegPts.Y) \ 3)
        adjWidth = rEdge - lEdge
     Case Else
        adjWidth = myProps.bSegPts.Y
        rEdge = UserControl.ScaleWidth
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    End Select
    '<:-)Auto-inserted With End...With Structure
    With myImage
        If (.SourceSize.X + .SourceSize.Y) > 0 Then
            ' image in use, calculations for image rectangle
            If .Size < 33 Then
                Select Case .Align
                 Case lv_LeftEdge, lv_LeftOfCaption
                    imgOffset.Left = .Size
                    bImgWidthAdj = True
                 Case lv_RightEdge, lv_RightOfCaption
                    imgOffset.Right = .Size
                    bImgWidthAdj = True
                 Case lv_TopCenter
                    imgOffset.Top = .Size
                    bImgHeightAdj = True
                 Case lv_BottomCenter
                    imgOffset.Bottom = .Size
                    bImgHeightAdj = True
                End Select
            End If
        End If
    End With 'myImage
    If Len(myProps.bCaption) Then
        sCaption = Replace$(myProps.bCaption, "||", vbNewLine)
        ' calculate total available button width available for text
        cRect.Right = adjWidth - 8 - (myImage.Size * Abs(CInt(bImgWidthAdj)))
        cRect.Bottom = UserControl.ScaleHeight - 8 - (myImage.Size * Abs(CInt(bImgHeightAdj And myImage.Align > lv_RightOfCaption)))
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        '<:-)Pleonasm Removed
        ' calculate size of rectangle to hold that text, using multiline flag
        DrawText ButtonDC.hDC, sCaption, Len(sCaption), cRect, DT_CALCRECT Or DT_WORDBREAK
        If myProps.bCaptionStyle Then
            cRect.Right = cRect.Right + 2
            cRect.Bottom = cRect.Bottom + 2
        End If
    End If
    ' now calculate the position of the text rectangle
    If Len(myProps.bCaption) Then
        tRect = cRect
        Select Case myProps.bCaptionAlign
         Case vbLeftJustify
            OffsetRect tRect, imgOffset.Left + lEdge + 4 + (Abs(CInt(imgOffset.Left > 0) * 4)), 0
         Case vbRightJustify
            OffsetRect tRect, rEdge - imgOffset.Right - 4 - cRect.Right - (Abs(CInt(imgOffset.Right > 0) * 4)), 0
         Case vbCenter
            If imgOffset.Left > 0 And myImage.Align = lv_LeftOfCaption Then
                OffsetRect tRect, (adjWidth - (imgOffset.Left + cRect.Right + 4)) \ 2 + lEdge + 4 + imgOffset.Left, 0
             Else
                If imgOffset.Right > 0 And myImage.Align = lv_RightOfCaption Then
                    OffsetRect tRect, (adjWidth - (imgOffset.Right + cRect.Right + 4)) \ 2 + lEdge, 0
                 Else
                    OffsetRect tRect, ((adjWidth - (imgOffset.Left + imgOffset.Right)) - cRect.Right) \ 2 + lEdge + imgOffset.Left, 0
                End If
            End If
        End Select
     Else
        cRect.Bottom = -4
    End If
    If (myImage.SourceSize.X + myImage.SourceSize.Y) > 0 Then
        ' finalize image rectangle position
        Select Case myImage.Align
         Case lv_LeftEdge
            iRect.Left = lEdge + 4
         Case lv_LeftOfCaption
            If Len(myProps.bCaption) Then
                iRect.Left = tRect.Left - 4 - imgOffset.Left
             Else
                iRect.Left = lEdge + 4
            End If
         Case lv_RightOfCaption
            If Len(myProps.bCaption) Then
                iRect.Left = tRect.Right + 4
             Else
                iRect.Left = rEdge - 4 - imgOffset.Right
            End If
         Case lv_RightEdge
            iRect.Left = rEdge - 4 - imgOffset.Right
         Case lv_TopCenter
            iRect.Top = (ScaleHeight - (cRect.Bottom + imgOffset.Top)) \ 2
            OffsetRect tRect, 0, iRect.Top + 2 + imgOffset.Top
         Case lv_BottomCenter
            iRect.Top = (ScaleHeight - (cRect.Bottom + imgOffset.Bottom)) \ 2 + cRect.Bottom + 4
            OffsetRect tRect, 0, iRect.Top - 2 - cRect.Bottom
        End Select
        If myImage.Align < lv_TopCenter Then
            OffsetRect tRect, 0, (ScaleHeight - cRect.Bottom) \ 2
            iRect.Top = (ScaleHeight - myImage.Size) \ 2
         Else
            iRect.Left = (adjWidth - myImage.Size) \ 2 + lEdge
        End If
        iRect.Right = iRect.Left + myImage.Size
        iRect.Bottom = iRect.Top + myImage.Size
     Else
        OffsetRect tRect, 0, (ScaleHeight - cRect.Bottom) \ 2
    End If
    ' sanity checks
    If tRect.Top < 4 Then
        tRect.Top = 4
    End If
    If tRect.Left < 4 + lEdge Then
        tRect.Left = 4 + lEdge
    End If
    If tRect.Right > rEdge - 4 Then
        tRect.Right = rEdge - 4
    End If
    If tRect.Bottom > UserControl.ScaleHeight - 5 Then
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        tRect.Bottom = UserControl.ScaleHeight - 5
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    End If
    myProps.bRect = tRect
    Select Case myImage.Size
     Case Is < 33
        If iRect.Top < 4 Then
            iRect.Top = 4
        End If
        If iRect.Left < 4 + lEdge Then
            iRect.Left = 4 + lEdge
        End If
        If iRect.Right > rEdge - 4 Then
            iRect.Right = rEdge - 4
        End If
        If iRect.Bottom > UserControl.ScaleHeight - 5 Then
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            iRect.Bottom = UserControl.ScaleHeight - 5
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        End If
     Case 40 ' stretch
        '<:-)Auto-inserted With End...With Structure
        With iRect
            .Left = 0
            .Top = 0
            .Right = UserControl.ScaleWidth
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            .Bottom = UserControl.ScaleHeight
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        End With 'iRect
     Case Else   ' scale
        ratio(0) = (adjWidth - 12) / myImage.SourceSize.X
        ratio(1) = (ScaleHeight - 12) / myImage.SourceSize.Y
        If ratio(1) < ratio(0) Then
            ratio(0) = ratio(1)
        End If
        ratio(1) = myImage.SourceSize.Y * ratio(0)
        ratio(0) = myImage.SourceSize.X * ratio(0)
        '<:-)Auto-inserted With End...With Structure
        With iRect
            .Left = (adjWidth - CLng(ratio(0))) \ 2 + lEdge
            .Top = (ScaleHeight - CLng(ratio(1))) \ 2
            .Right = .Left + CLng(ratio(0))
            .Bottom = .Top + CLng(ratio(1))
        End With 'iRect
        Erase ratio
    End Select
    myImage.iRect = iRect

End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "The caption of the button. Double pipe (||) is a line break."
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"

    Caption = myProps.bCaption

End Property

Public Property Let Caption(sCaption As String)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'sCaption'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' Sets the button caption & hot key for the control
  
  Dim i As Integer
  Dim j As Integer

    '<:-):UPDATED: Multiple Dim line separated
    ' We look from right to left. VB uses this logic & so do I
    i = InStrRev(sCaption, "&")
    Do While i
        If Mid$(sCaption, i, 2) = "&&" Then
            i = InStrRev(i - 1, sCaption, "&")
         Else
            j = i + 1
            i = 0
        End If
    Loop
    ' if found, we use the next character as a hot key
    If j Then
        UserControl.AccessKeys = Mid$(sCaption, j, 1)
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    End If
    myProps.bCaption = sCaption                     ' cache the caption
    CalculateBoundingRects                          ' recalculate button text/image bounding rects
    Refresh
    PropertyChanged "Caption"

End Property

Public Property Get CaptionAlign() As AlignmentConstants
Attribute CaptionAlign.VB_Description = "Horizontal alignment of caption on the button."

    CaptionAlign = myProps.bCaptionAlign

End Property

Public Property Let CaptionAlign(nAlign As AlignmentConstants)

  ' Caption options: Left, Right or Center Justified

    If nAlign < vbLeftJustify Or nAlign > vbCenter Then
        Exit Property
    End If
    If myImage.Align > lv_RightOfCaption Then
        If nAlign < vbCenter Then
            If (myImage.SourceSize.X + myImage.SourceSize.Y) > 0 Then
                '<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
                ' also prevent left/right justifying captions when image is centered in caption
                If UserControl.Ambient.UserMode = False Then
                    ' if not in user mode, then explain whey it is prevented
                    MsgBox "When button images are aligned top/bottom center, " & vbNewLine & "button captions can only be center aligned", vbOKOnly + vbInformation
                End If
                Exit Property
            End If
        End If '<:-)Short Circuit inserted this line
    End If '<:-)Short Circuit inserted this line
    myProps.bCaptionAlign = nAlign
    CalculateBoundingRects              ' recalculate text/image bounding rects
    Refresh
    PropertyChanged "CapAlign"

End Property

Public Property Get CaptionStyle() As CaptionEffectConstants
Attribute CaptionStyle.VB_Description = "Flat, Embossed or Engraved effects"

    CaptionStyle = myProps.bCaptionStyle

End Property

Public Property Let CaptionStyle(nStyle As CaptionEffectConstants)

  ' Sets the style, raised/sunken or flat (default)

    If nStyle < lv_default Or nStyle > lv_Raised Then
        Exit Property
    End If
    myProps.bCaptionStyle = nStyle
    PropertyChanged "CapStyle"
    If Len(myProps.bCaption) Then
        CalculateBoundingRects
        Refresh
    End If

End Property

Private Function ConvertColor(tColor As Long) As Long

  ' Converts VB color constants to real color values

    If tColor < 0 Then
        ConvertColor = GetSysColor(tColor And &HFF&)
     Else
        ConvertColor = tColor
    End If

End Function

Private Sub CreateButtonRegion()

  ' this function creates the regions for the specific type of button style
  
  Dim rgnA          As Long
  Dim rgnB          As Long
  Dim rgn2Use       As Long
  Dim i             As Integer

    '<:-):UPDATED: Multiple Dim line separated
  Dim lRatio        As Single
  Dim lEdge         As Long
  Dim rEdge         As Long
  Dim Wd            As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim ptTRI(0 To 9) As POINTAPI
    myProps.bSegPts.X = 0
    myProps.bSegPts.Y = UserControl.ScaleWidth
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    rgnB = CreateRectRgn(0, 0, 0, 0)
    '<:-)Auto-inserted With End...With Structure
    With ButtonDC
        If .ClipRgn Then
            ' this was set for round buttons
            SelectClipRgn .hDC, 0
            DeleteObject .ClipRgn
            .ClipRgn = 0
        End If
    End With 'ButtonDC
    Select Case myProps.bShape
     Case lv_Round3D, lv_Round3DFixed, lv_RoundFlat
        rgn2Use = CreateEllipticRgn(0, 0, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1)
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        '<:-)Auto-inserted With End...With Structure
        With myProps
            If .bBackStyle <> 5 Then
                If .bShape < lv_RoundFlat Then
                    i = .bGradient
                    .bGradient = lv_Top2Bottom
                    DrawGradient vbWhite, vbGray
                    .bGradient = i
                End If
                ButtonDC.ClipRgn = CreateEllipticRgn(1, 1, UserControl.ScaleWidth, UserControl.ScaleHeight)
                '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                SelectClipRgn ButtonDC.hDC, ButtonDC.ClipRgn
            End If
        End With 'myProps
     Case lv_Rectangular
        rgn2Use = CreateRectRgn(0, 0, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1)
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        Select Case myProps.bBackStyle
         Case 1 'Windows 16-bit
            GoSub LopOffCorners1
            '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
            GoSub LopOffCorners2
            '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
         Case 2, 7
            GoSub LopOffCorners3
            '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
            GoSub LopOffCorners4
            '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
         Case 3    'Java
            If UserControl.Enabled Then
                GoSub LopOffCorners1
                '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
                GoSub LopOffCorners2
                '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
            End If
        End Select
     Case Else
        ' here is my trick for ensuring a sharp edge on diagonal buttons.
        ' Basically a bastardized carpenters formula for right angles
        ' (i.e., 3+4=5 < the hypoteneus). Here I want a 60 degree angle,
        ' and not a 45 degree angle. The difference is sharp or choppy.
        ' Based off of the button height, I need to figure how much of
        ' the opposite end I need to cutoff for the diagonal edge
        lRatio = (ScaleHeight + 1) / 4
        Wd = UserControl.ScaleWidth
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        lEdge = (4 * lRatio)
        ' here we ensure a width of at least 5 pixels wide
        Do While Wd - lEdge < 5
            Wd = Wd + 5
        Loop
        If Wd <> UserControl.ScaleWidth Then
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            ' resize the control if necessary
            DelayDrawing True
            UserControl.Width = ScaleX(Wd, vbPixels, Parent.ScaleMode)
            myProps.bSegPts.Y = UserControl.ScaleWidth
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            bNoRefresh = False
        End If
        rEdge = UserControl.ScaleWidth - lEdge
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        ' initial dimensions of our rectangle
        ptTRI(0).X = 0
        ptTRI(0).Y = 0
        ptTRI(1).X = 0
        ptTRI(1).Y = UserControl.ScaleHeight + 1
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        ptTRI(2).X = UserControl.ScaleWidth + 1
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        ptTRI(2).Y = UserControl.ScaleHeight + 1
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        ptTRI(3).X = UserControl.ScaleWidth + 1
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        ptTRI(3).Y = 0
        ' now modify the left/right side as needed
        If myProps.bShape = lv_FullDiagonal Or myProps.bShape = lv_LeftDiagonal Then
            ptTRI(1).X = lEdge  ' left portion
            myProps.bSegPts.X = lEdge
        End If
        If myProps.bShape = lv_FullDiagonal Or myProps.bShape = lv_RightDiagonal Then
            ptTRI(3).X = rEdge + 1        ' bottom right
            myProps.bSegPts.Y = rEdge
        End If
        ' for rounded corner buttons, we'll take of the corner pixels where appropriate when the
        ' diagonal button is not a fully-segmeneted type. Diagonal edges are always sharp,
        ' never rounded.
        rgn2Use = CreatePolygonRgn(ptTRI(0), 4, 2)
        Select Case myProps.bBackStyle
         Case 1
            If myProps.bShape = lv_RightDiagonal Then
                GoSub LopOffCorners1
                '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
            End If
            If myProps.bShape = lv_LeftDiagonal Then
                GoSub LopOffCorners2
                '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
            End If
         Case 2, 7
            If myProps.bShape = lv_RightDiagonal Then
                GoSub LopOffCorners3
                '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
            End If
            If myProps.bShape = lv_LeftDiagonal Then
                GoSub LopOffCorners4
                '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
            End If
         Case 3
            If UserControl.Enabled Then
                If myProps.bShape = lv_RightDiagonal Then
                    GoSub LopOffCorners1
                    '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
                End If
                If myProps.bShape = lv_LeftDiagonal Then
                    GoSub LopOffCorners2
                    '<:-):SUGGESTION: Obsolete Code GoSub Create a real Sub instead, it is much safer as you have better control over variables.
                End If
            End If
        End Select
    End Select
    Erase ptTRI
    If rgnA Then
        DeleteObject rgnA
    End If
    If rgnB Then
        DeleteObject rgnB
    End If
    SetWindowRgn UserControl.hWnd, rgn2Use, True
    If myProps.bSegPts.Y = 0 Then
        myProps.bSegPts.Y = UserControl.ScaleWidth
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    End If

Exit Sub

LopOffCorners1:
    ' left side top/bottom corners (Java/Win3.x)
    If myProps.bBackStyle = 3 Then
        rgnA = CreateRectRgn(0, UserControl.ScaleHeight, 1, UserControl.ScaleHeight - 1)
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
     Else
        rgnA = CreateRectRgn(0, 0, 1, 1)
    End If
    CombineRgn rgnB, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    rgnA = CreateRectRgn(0, UserControl.ScaleHeight, 1, UserControl.ScaleHeight - 1)
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    CombineRgn rgn2Use, rgnB, rgnA, RGN_DIFF
    DeleteObject rgnA
    Return
    '<:-):SUGGESTION: Obsolete Code Return Create a real Sub instead, it is much safer as you have better control over variables.
LopOffCorners2:
    ' right side top/bottom corners (Java/Win3.x)
    If myProps.bBackStyle = 3 Then
        rgnA = CreateRectRgn(ScaleWidth, 0, UserControl.ScaleWidth - 1, 1)
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
     Else
        rgnA = CreateRectRgn(ScaleWidth, UserControl.ScaleHeight, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1)
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    End If
    CombineRgn rgnB, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    rgnA = CreateRectRgn(ScaleWidth, 0, UserControl.ScaleWidth - 1, 1)
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    CombineRgn rgn2Use, rgnB, rgnA, RGN_DIFF
    DeleteObject rgnA
    Return
    '<:-):SUGGESTION: Obsolete Code Return Create a real Sub instead, it is much safer as you have better control over variables.
LopOffCorners3:
    ' left side top/bottom corners (XP/Mac)
    ptTRI(0).X = 0
    ptTRI(0).Y = 0
    ptTRI(1).X = 2
    ptTRI(1).Y = 0
    ptTRI(2).X = 0
    ptTRI(2).Y = 2
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgnB, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    ptTRI(0).X = 0
    ptTRI(0).Y = UserControl.ScaleHeight
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    ptTRI(1).X = 3
    ptTRI(1).Y = UserControl.ScaleHeight
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    ptTRI(2).X = 0
    ptTRI(2).Y = UserControl.ScaleHeight - 3
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgn2Use, rgnB, rgnA, RGN_DIFF
    DeleteObject rgnA
    Return
    '<:-):SUGGESTION: Obsolete Code Return Create a real Sub instead, it is much safer as you have better control over variables.
LopOffCorners4:
    ' right side top/bottom corners (XP/Mac)
    ptTRI(0).X = UserControl.ScaleWidth
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    ptTRI(0).Y = 0
    ptTRI(1).X = UserControl.ScaleWidth - 2
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    ptTRI(1).Y = 0
    ptTRI(2).X = UserControl.ScaleWidth
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    ptTRI(2).Y = 2
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgnB, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    ptTRI(0).X = UserControl.ScaleWidth
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    ptTRI(0).Y = UserControl.ScaleHeight
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    ptTRI(1).X = UserControl.ScaleWidth - 3
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    ptTRI(1).Y = UserControl.ScaleHeight
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    ptTRI(2).X = UserControl.ScaleWidth
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    ptTRI(2).Y = UserControl.ScaleHeight - 3
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgn2Use, rgnB, rgnA, RGN_DIFF
    DeleteObject rgnA
    Return
    '<:-):SUGGESTION: Obsolete Code Return Create a real Sub instead, it is much safer as you have better control over variables.

End Sub

Public Sub DelayDrawing(bDelay As Boolean)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'bDelay'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' Used to prevent redrawing button until all properties are set.
  ' Should you want to set multiple properties of the control during runtime
  ' call this function first with a TRUE parameter. Set your button
  ' attributes and then call it again with a FALSE property to update the
  ' button.   IMPORTANT: If called with a TRUE parameter you must
  ' also release it with a call and a FALSE parameter
  ' NOTE: this function will prevent flicker when several properties
  ' are being changed at once during run time. It is similar to
  ' the BeginPaint & EndPaint API functionality

    bNoRefresh = bDelay
    If bDelay = False Then
        Refresh
    End If

End Sub

Private Sub DrawButton_Flat(polyPts() As POINTAPI, _
                            polyColors() As Long, _
                            ActiveStatus As Integer)

  '==========================================================================
  ' If not used in your project, replace this entire routine from the
  ' Dim statements to the last line before the End Sub with
  ' a simple Exit Sub
  '==========================================================================
  
  Dim darkShade As Long
  Dim liteShade As Long
  Dim backShade As Long

    '<:-):UPDATED: Multiple Dim line separated
  Dim i         As Integer
  Dim lColor    As Long
    '<:-):UPDATED: Multiple Dim line separated
    If myProps.bMode = lv_CommandButton Or myProps.bValue = False Then
        backShade = adjBackColorUp
     Else
        backShade = cCheckBox
    End If
    DrawButtonBackground backShade, ActiveStatus, ConvertColor(myProps.bGradientColor), adjHoverColor
    If myProps.bMode > lv_CommandButton And UserControl.Enabled = False Then
        lColor = vbGray
     Else
        lColor = -1
    End If
    DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton
    If myProps.bShape < lv_Round3D Then
        darkShade = vbGray
        liteShade = vbWhite
        ' inner rectangle & outer edges
        For i = 1 To 8
            polyColors(i) = -1
        Next i
        ' Outer Rectangle
        If (bKeyDown And myProps.bMode > lv_CommandButton) Or myProps.bValue Then
            '<:-)Pleonasm Removed
            ActiveStatus = 2
            lColor = vbBlack
         Else
            If myProps.bShape > lv_Rectangular Then
                lColor = ShadeColor(backShade, -&H30, False, False)
            End If
            If ActiveStatus < 2 Then
                ActiveStatus = 1
            End If
        End If
        polyColors(9) = Choose(ActiveStatus, liteShade, darkShade)
        polyColors(10) = polyColors(9)
        polyColors(11) = Choose(ActiveStatus, darkShade, liteShade)
        polyColors(12) = polyColors(11)
        DrawButtonBorder polyPts(), polyColors(), ActiveStatus
    End If
    DrawFocusRectangle lColor, myProps.bValue, False, polyPts()

End Sub

Private Sub DrawButton_Hover(polyPts() As POINTAPI, _
                             polyColors() As Long, _
                             ActiveStatus As Integer)

  '==========================================================================
  ' If not used in your project, replace this entire routine from the
  ' Dim statements to the last line before the End Sub with
  ' a simple Exit Sub
  '==========================================================================
  
  Dim backShade   As Long
  Dim i           As Integer
  Dim lColor      As Long
  Dim lFocusColor As Long

    '<:-):UPDATED: Multiple Dim line separated
    If myProps.bMode = lv_CommandButton Or myProps.bValue = False Then
        backShade = adjBackColorUp
        If myProps.bShape > lv_Rectangular Then
            lFocusColor = ShadeColor(cParentBC, -&H20, False)
        End If
     Else
        backShade = cCheckBox
    End If
    DrawButtonBackground backShade, ActiveStatus, ConvertColor(myProps.bGradientColor), adjHoverColor
    If myProps.bMode > lv_CommandButton And UserControl.Enabled = False Then
        lColor = vbGray
     Else
        lColor = -1
    End If
    DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton
    If myProps.bShape > lv_FullDiagonal And UserControl.Ambient.UserMode = False Then
        If ButtonDC.ClipRgn Then
            DeleteObject ButtonDC.ClipRgn
        End If
        '<:-)Auto-inserted With End...With Structure
        With ButtonDC
            SelectClipRgn .hDC, 0
            SetButtonColors True, .hDC, cObj_Pen, vbWhite, , , , 2
            Arc .hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 0, 0, 0, 0
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            .ClipRgn = CreateEllipticRgn(1, 1, UserControl.ScaleWidth, UserControl.ScaleHeight)
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            SelectClipRgn .hDC, ButtonDC.ClipRgn
        End With 'ButtonDC
     Else
        For i = 1 To 8
            polyColors(i) = -1
        Next i
        If ((myProps.bStatus And 4) = 4) Or myProps.bValue Then
            '<:-)Pleonasm Removed
            If (bKeyDown And myProps.bMode > lv_CommandButton) Or myProps.bValue Then
                '<:-)Pleonasm Removed
                ActiveStatus = 2
            End If
            If ActiveStatus < 2 Then
                ActiveStatus = 1
            End If
            polyColors(9) = Choose(ActiveStatus, vbWhite, vbGray)
            polyColors(10) = polyColors(9)
            polyColors(11) = Choose(ActiveStatus, vbGray, vbWhite)
            polyColors(12) = polyColors(11)
         Else
            If Ambient.UserMode = False Then
                lColor = ShadeColor(cParentBC, -&H40, False)
                If lColor = vbBlack Then
                    lColor = vbWhite
                End If
             Else
                lColor = -1
            End If
            For i = 9 To 12
                polyColors(i) = lColor
            Next i
        End If
        DrawButtonBorder polyPts(), polyColors(), ActiveStatus, Abs(UserControl.Ambient.UserMode = False) * 2
    End If
    DrawFocusRectangle lFocusColor, myProps.bValue, False, polyPts()

End Sub

Private Sub DrawButton_Java(polyPts() As POINTAPI, _
                            polyColors() As Long, _
                            ActiveStatus As Integer)

  '==========================================================================
  ' If not used in your project, replace this entire routine from the
  ' Dim statements to the last line before the End Sub with
  ' a simple Exit Sub
  '==========================================================================
  
  Dim backShade As Long
  Dim darkShade As Long
  Dim liteShade As Long

    '<:-):UPDATED: Multiple Dim line separated
  Dim i         As Integer
  Dim lColor    As Long
    '<:-):UPDATED: Multiple Dim line separated
    backShade = adjBackColorUp
    If myProps.bMode = lv_CommandButton Then
        If ((myProps.bStatus And 6) = 6) Then
            backShade = adjBackColorDn
        End If
     Else
        If myProps.bValue Then
            backShade = cCheckBox
        End If
    End If
    DrawButtonBackground backShade, ActiveStatus, ConvertColor(myProps.bGradientColor), adjHoverColor
    If UserControl.Enabled Then
        lColor = vbGray
     Else
        lColor = ShadeColor(vbGray, -&H10, False)
    End If
    DrawCaptionIcon backShade, lColor, , True
    If myProps.bShape < lv_Round3D Then
        darkShade = ShadeColor(vbGray, -&H1A, False)
        liteShade = vbWhite
        If UserControl.Enabled Or myProps.bMode > lv_CommandButton Then
            For i = 1 To 4
                polyColors(i) = -1
            Next i
            If myProps.bMode > lv_CommandButton Then
                If (myProps.bValue Or bKeyDown) Then
                    ActiveStatus = 2
                 Else
                    ActiveStatus = 3
                End If
            End If
            If ActiveStatus < 2 Then
                ActiveStatus = 1
            End If
            polyColors(5) = Choose(ActiveStatus, liteShade, backShade, liteShade)
            polyColors(6) = polyColors(5)
            polyColors(7) = darkShade
            polyColors(8) = darkShade
         Else
            For i = 1 To 8
                polyColors(i) = backShade
            Next i
            liteShade = darkShade
        End If
        polyColors(9) = darkShade
        polyColors(10) = darkShade
        polyColors(11) = liteShade
        polyColors(12) = liteShade
        DrawButtonBorder polyPts(), polyColors(), ActiveStatus
    End If
    DrawFocusRectangle &HCC9999, True, True, polyPts()

End Sub

Private Sub DrawButton_Macintosh(polyPts() As POINTAPI, _
                                 polyColors() As Long, _
                                 ActiveStatus As Integer)

  '==========================================================================
  ' If not used in your project, replace this entire routine from the
  ' Dim statements to the last line before the End Sub with
  ' a simple Exit Sub
  '==========================================================================
  
  Dim backShade      As Long
  Dim darkShade      As Long
  Dim liteShade      As Long
  Dim midShade       As Long

    '<:-):UPDATED: Multiple Dim line separated
  Dim lGradientColor As Long
  Dim lFocusColor    As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim i              As Integer
  Dim lColor         As Long
    '<:-):UPDATED: Multiple Dim line separated
    backShade = adjBackColorUp
    If myProps.bMode = lv_CommandButton Then
        If ((myProps.bStatus And 6) = 6) Then
            backShade = adjBackColorDn
        End If
        If myProps.bShape > lv_Rectangular Then
            lFocusColor = ShadeColor(backShade, -&H40, False)
        End If
     Else
        If myProps.bValue Then
            backShade = cCheckBox
            lFocusColor = ShadeColor(vbGray, -&H20, False)
        End If
    End If
    '<:-)Auto-inserted With End...With Structure
    With myProps
        If .bGradient Then
            lGradientColor = ShadeColor(ConvertColor(.bGradientColor), &H1F, False)
            If ((.bStatus And 6) = 6) Then
                backShade = adjBackColorUp
            End If
         Else
            lGradientColor = backShade
        End If
    End With 'myProps
    DrawButtonBackground backShade, ActiveStatus, lGradientColor, adjHoverColor
    If ((myProps.bStatus And 6) = 6 And myProps.bMode = lv_CommandButton) Or (myProps.bMode > lv_CommandButton And myProps.bValue) Then
        '<:-)Pleonasm Removed
        If (myProps.bValue And myProps.bMode > lv_CommandButton) And UserControl.Enabled = False Then
            '<:-)Pleonasm Removed
            lColor = ShadeColor(backShade, -&H20, True)
         Else
            lColor = adjBackColorUp
        End If
     Else
        If (myProps.bValue = False And myProps.bMode > lv_CommandButton) And UserControl.Enabled = False Then
            lColor = vbGray
         Else
            If UserControl.Enabled Then
                lColor = ConvertColor(UserControl.ForeColor)
             Else
                lColor = -1
            End If
        End If
    End If
    If UserControl.ForeColor = myProps.bForeHover And myProps.bValue Then
        '<:-)Pleonasm Removed
        lFocusColor = myProps.bForeHover
        myProps.bForeHover = lColor
     Else
        lFocusColor = -1
    End If
    DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton
    If lFocusColor <> -1 Then
        myProps.bForeHover = lFocusColor
    End If
    '<:-)Auto-inserted With End...With Structure
    With myProps
        If .bShape < lv_Round3D Then
            If (bKeyDown And .bMode > lv_CommandButton) Or .bValue Then
                '<:-)Pleonasm Removed
                If .bValue Then
                    ActiveStatus = 2
                 Else
                    ActiveStatus = 0
                End If
            End If
            midShade = ShadeColor(backShade, &H1F, True)
            darkShade = ShadeColor(backShade, -&H40, True)
            liteShade = vbWhite
            If ActiveStatus = 2 Then
                If .bGradient = lv_NoGradient Then
                    lColor = vbGray
                 Else
                    lColor = adjBackColorUp
                End If
                backShade = lColor
                liteShade = ShadeColor(lColor, -&H20, False)
                midShade = ShadeColor(lColor, -&H40, False)
                darkShade = ShadeColor(lColor, -&H10, False)
            End If
            polyColors(1) = liteShade
            polyColors(2) = liteShade
            polyColors(3) = backShade
            polyColors(4) = backShade
            ' middle Rectangle
            polyColors(5) = midShade
            polyColors(6) = midShade
            polyColors(7) = darkShade
            polyColors(8) = darkShade
            ' Outer Rectangle
            For i = 9 To 12
                polyColors(i) = vbBlack
            Next i
            DrawButtonBorder polyPts(), polyColors(), ActiveStatus
            If .bSegPts.X = 0 Then
                SetPixel ButtonDC.hDC, 3, 3, liteShade
                SetPixel ButtonDC.hDC, 1, ScaleHeight - 3, backShade
                SetPixel ButtonDC.hDC, 2, 2, midShade
                SetPixel ButtonDC.hDC, 1, ScaleHeight - 2, 0
                SetPixel ButtonDC.hDC, 1, 1, 0
            End If
            If .bSegPts.Y = UserControl.ScaleWidth Then
                '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                SetPixel ButtonDC.hDC, ScaleWidth - 4, ScaleHeight - 4, backShade
                SetPixel ButtonDC.hDC, ScaleWidth - 3, 1, backShade
                SetPixel ButtonDC.hDC, ScaleWidth - 3, ScaleHeight - 3, darkShade
                SetPixel ButtonDC.hDC, ScaleWidth - 2, ScaleHeight - 2, 0
                SetPixel ButtonDC.hDC, ScaleWidth - 2, 1, 0
            End If
        End If
    End With 'myProps
    DrawFocusRectangle lFocusColor, myProps.bValue, True, polyPts()

End Sub

Private Sub DrawButton_Netscape(polyPts() As POINTAPI, _
                                polyColors() As Long, _
                                ActiveStatus As Integer)

  '==========================================================================
  ' If not used in your project, replace this entire routine from the
  ' Dim statements to the last line before the End Sub with
  ' a simple Exit Sub
  '==========================================================================
  
  Dim backShade As Long
  Dim darkShade As Long
  Dim liteShade As Long

    '<:-):UPDATED: Multiple Dim line separated
  Dim i         As Integer
  Dim lColor    As Long
    '<:-):UPDATED: Multiple Dim line separated
    If myProps.bMode = lv_CommandButton Or myProps.bValue = False Then
        backShade = adjBackColorUp
     Else
        backShade = cCheckBox
    End If
    DrawButtonBackground backShade, ActiveStatus, ConvertColor(myProps.bGradientColor), adjHoverColor
    If myProps.bMode > lv_CommandButton And UserControl.Enabled = False Then
        lColor = vbGray
     Else
        lColor = -1
    End If
    DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton
    If myProps.bShape < lv_Round3D Then
        darkShade = vbGray
        liteShade = ShadeColor(&HDFDFDF, &H8, False)
        For i = 1 To 4
            polyColors(i) = -1
        Next i
        If (bKeyDown And myProps.bMode > lv_CommandButton) Or myProps.bValue Then
            '<:-)Pleonasm Removed
            ActiveStatus = 1
            lColor = vbBlack
         Else
            ActiveStatus = Abs((myProps.bStatus And 6) = 6)
            If myProps.bShape > lv_Rectangular Then
                lColor = ShadeColor(backShade, -&H30, False, False)
            End If
        End If
        polyColors(5) = Choose(ActiveStatus + 1, liteShade, darkShade)
        polyColors(6) = polyColors(5)
        polyColors(9) = polyColors(5)
        polyColors(10) = polyColors(5)
        polyColors(7) = Choose(ActiveStatus + 1, darkShade, liteShade)
        polyColors(8) = polyColors(7)
        polyColors(11) = polyColors(7)
        polyColors(12) = polyColors(7)
        DrawButtonBorder polyPts(), polyColors(), ActiveStatus
    End If
    DrawFocusRectangle lColor, myProps.bValue, False, polyPts()

End Sub

Private Sub DrawButton_Win31(polyPts() As POINTAPI, _
                             polyColors() As Long, _
                             ActiveStatus As Integer)

  '==========================================================================
  ' If not used in your project, replace this entire routine from the
  ' Dim statements to the last line before the End Sub with
  ' a simple Exit Sub
  '==========================================================================
  
  Dim backShade As Long
  Dim darkShade As Long
  Dim liteShade As Long

    '<:-):UPDATED: Multiple Dim line separated
  Dim i         As Integer
  Dim lColor    As Long
    '<:-):UPDATED: Multiple Dim line separated
    If myProps.bMode = lv_CommandButton Or myProps.bValue = False Then
        backShade = adjBackColorUp
     Else
        backShade = cCheckBox
    End If
    DrawButtonBackground backShade, ActiveStatus, ConvertColor(myProps.bGradientColor), adjHoverColor
    If myProps.bMode > lv_CommandButton And UserControl.Enabled = False Then
        lColor = vbGray
     Else
        lColor = -1
    End If
    DrawCaptionIcon backShade, lColor, myProps.bValue = True, myProps.bMode > lv_CommandButton
    If myProps.bShape < lv_Round3D Then
        darkShade = vbGray
        liteShade = vbWhite
        If (bKeyDown And myProps.bMode > lv_CommandButton) Or myProps.bValue Then
            '<:-)Pleonasm Removed
            ActiveStatus = 2
            lColor = vbBlack
         Else
            If myProps.bShape > lv_Rectangular Then
                lColor = ShadeColor(backShade, -&H30, False, False)
            End If
            If ActiveStatus < 2 Then
                ActiveStatus = 1
            End If
            If myProps.bShape > lv_Rectangular Then
                lColor = ShadeColor(backShade, -&H30, False, False)
            End If
        End If
        If ActiveStatus = 2 Then
            polyColors(1) = darkShade
            polyColors(3) = liteShade
         Else
            polyColors(1) = liteShade
            polyColors(3) = darkShade
        End If
        polyColors(2) = polyColors(1)
        polyColors(4) = polyColors(3)
        For i = 5 To 6
            polyColors(i) = polyColors(1)
            polyColors(i + 2) = polyColors(3)
        Next i
        For i = 9 To 12
            polyColors(i) = vbBlack
        Next i
        DrawButtonBorder polyPts(), polyColors(), ActiveStatus
    End If
    DrawFocusRectangle lColor, myProps.bValue, False, polyPts()

End Sub

Private Sub DrawButton_Win95(polyPts() As POINTAPI, _
                             polyColors() As Long, _
                             ActiveStatus As Integer)

  '==========================================================================
  ' If not used in your project, replace this entire routine from the
  ' Dim statements to the last line before the End Sub with
  ' a simple Exit Sub
  '==========================================================================
  
  Dim midShade  As Long
  Dim darkShade As Long
  Dim liteShade As Long
  Dim backShade As Long

    '<:-):UPDATED: Multiple Dim line separated
  Dim lColor    As Long
    'Dim fRect As RECT
    '<:-):WARNING: Unused Dim commented out
    'Dim I As Integer
    '<:-):WARNING: Unused Dim commented out
    '<:-):UPDATED: Multiple Dim line separated
    If myProps.bMode = lv_CommandButton Or myProps.bValue = False Then
        backShade = adjBackColorUp
     Else
        backShade = cCheckBox
    End If
    DrawButtonBackground backShade, ActiveStatus, ConvertColor(myProps.bGradientColor), adjHoverColor
    If myProps.bMode > lv_CommandButton And UserControl.Enabled = False Then
        lColor = vbGray
     Else
        lColor = -1
    End If
    '<:-)Auto-inserted With End...With Structure
    With myProps
        DrawCaptionIcon backShade, lColor, .bValue = True, .bMode > lv_CommandButton
        If .bShape < lv_Round3D Then
            If (((.bStatus And 6) = 6)) Then
                midShade = backShade
             Else
                midShade = RGB(233, 233, 233)
            End If
            darkShade = vbGray
            liteShade = vbWhite
            If (bKeyDown And .bMode > lv_CommandButton) Or .bValue Then
                '<:-)Pleonasm Removed
                ActiveStatus = 3
                midShade = RGB(233, 233, 233)
                lColor = vbBlack
             Else
                If .bShape > lv_Rectangular Then
                    lColor = ShadeColor(backShade, -&H30, False, False)
                End If
            End If
            ' inner rectangle
            polyColors(1) = Choose(ActiveStatus + 1, -1, midShade, -1, -1)
            polyColors(2) = polyColors(1)
            polyColors(3) = Choose(ActiveStatus + 1, -1, darkShade, -1, -1)
            polyColors(4) = polyColors(3)
            ' middle rectangle
            polyColors(5) = Choose(ActiveStatus + 1, midShade, liteShade, -1, -1)
            polyColors(6) = polyColors(5)
            polyColors(7) = Choose(ActiveStatus + 1, darkShade, vbBlack, darkShade, midShade)
            polyColors(8) = polyColors(7)
            ' Outer Rectangle
            If .bValue Or (.bMode > lv_CommandButton And bKeyDown) Then
                '<:-)Pleonasm Removed
                polyColors(9) = vbBlack
                polyColors(10) = vbBlack
                polyColors(11) = vbWhite
                polyColors(12) = vbWhite
             Else
                If Abs((.bStatus And 1) = 1) Then
                    polyColors(9) = vbBlack
                 Else
                    polyColors(9) = liteShade
                End If
                polyColors(10) = polyColors(9)
                polyColors(11) = vbBlack
                polyColors(12) = vbBlack
            End If
            DrawButtonBorder polyPts(), polyColors(), ActiveStatus
        End If
    End With 'myProps
    DrawFocusRectangle lColor, myProps.bValue, False, polyPts()

End Sub

Private Sub DrawButton_WinXP(polyPts() As POINTAPI, _
                             polyColors() As Long, _
                             ActiveStatus As Integer)

  '==========================================================================
  ' If not used in your project, replace this entire routine from the
  ' Dim statements to the last line before the End Sub with
  ' a simple Exit Sub
  '==========================================================================
  
  Dim backShade      As Long

    'Dim darkShade As Long
    '<:-):WARNING: Unused Dim commented out
  Dim liteShade      As Long
    'Dim midShade As Long
    '<:-):WARNING: Unused Dim commented out
    '<:-):UPDATED: Multiple Dim line separated
  Dim i              As Integer
  Dim lColor         As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim cDisabled      As Long
  Dim lGradientColor As Long
    '<:-):UPDATED: Multiple Dim line separated
    If myProps.bMode > lv_CommandButton Then
        If myProps.bValue Then
            lColor = cCheckBox
         Else
            lColor = adjBackColorUp
        End If
        backShade = lColor
     Else
        lColor = adjBackColorUp
        If ((myProps.bStatus And 6) = 6) And myProps.bGradient = lv_NoGradient Then
            backShade = adjBackColorDn
         Else
            If UserControl.Enabled = False Then
                backShade = ShadeColor(lColor, -&H18, True)
             Else
                backShade = lColor
            End If
        End If
    End If
    If Not UserControl.Enabled Then
        cDisabled = ShadeColor(backShade, -&H68, True)
    End If
    If myProps.bGradient Then
        lGradientColor = ShadeColor(ConvertColor(myProps.bGradientColor), &H30, True)
     Else
        lGradientColor = backShade
    End If
    DrawButtonBackground backShade, ActiveStatus, lGradientColor, adjHoverColor
    '<:-)Auto-inserted With End...With Structure
    With myProps
        DrawCaptionIcon backShade, cDisabled, .bValue = True, True
        If .bShape > lv_FullDiagonal Then
            If ((.bStatus And 1) = 1) And bTimerActive = False Then
                DrawFocusRectangle &HEF826B, False, False, polyPts()
             Else
                If bTimerActive Then
                    '<:-)Pleonasm Removed
                    DrawFocusRectangle &H96E7&, False, False, polyPts()
                End If
            End If
         Else
            If UserControl.Enabled Then
                If (bKeyDown And .bMode > lv_CommandButton) Or .bValue Then
                    '<:-)Pleonasm Removed
                    If ((.bStatus And 1) = 1) Then
                        ActiveStatus = 1
                     Else
                        ActiveStatus = 2
                    End If
                End If
                If ((.bStatus And 4) = 4) And ((.bStatus And 2) <> 2) Then
                    ActiveStatus = 3
                 Else
                    If ActiveStatus = 1 Then
                        If .bShowFocus = False Then
                            '<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
                            ActiveStatus = 0
                        End If
                    End If '<:-)Short Circuit inserted this line
                End If
                liteShade = lColor
                ' inner Rectangle
                polyColors(1) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&HA, True), &HF0D1B5, ShadeColor(liteShade, -&H16, True), &H6BCBFF)
                polyColors(2) = Choose(ActiveStatus + 1, ShadeColor(lColor, &HA, True), &HF7D7BD, ShadeColor(liteShade, -&H18, True), &H8CDBFF)
                polyColors(3) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H18, True), &HF0D1B5, lColor, &H6BCBFF)
                polyColors(4) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H20, True), &HE7AE8C, ShadeColor(liteShade, &HA, True), &H31B2FF)
                ' middle Rectangle
                polyColors(5) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H5, True), &HE7AE8C, ShadeColor(liteShade, -&H20, True), &H31B2FF)
                polyColors(6) = Choose(ActiveStatus + 1, ShadeColor(lColor, &H10, True), &HFFDFBF, ShadeColor(liteShade, -&H20, True), &HA6E9FF)
                polyColors(7) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H24, True), &HE7AE8C, ShadeColor(liteShade, &H5, True), &H31B2FF)
                polyColors(8) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H30, True), &HEF826B, ShadeColor(liteShade, &H10, True), &H96E7&)
                lColor = &H733C00
             Else
                For i = 1 To 8
                    polyColors(i) = -1
                Next i
                lColor = ShadeColor(lColor, -&H54, True)
            End If
            For i = 9 To 12
                polyColors(i) = lColor
            Next i
            DrawButtonBorder polyPts(), polyColors(), ActiveStatus
            If .bSegPts.X = 0 Then
                SetPixel ButtonDC.hDC, 1, ScaleHeight - 2, lColor
                SetPixel ButtonDC.hDC, 1, 1, lColor
            End If
            If .bSegPts.Y = UserControl.ScaleWidth Then
                '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                SetPixel ButtonDC.hDC, ScaleWidth - 2, ScaleHeight - 2, lColor
                SetPixel ButtonDC.hDC, ScaleWidth - 2, 1, lColor
            End If
        End If
    End With 'myProps

End Sub

Private Sub DrawButtonBackground(bColor As Long, _
                                 ActiveStatus As Integer, _
                                 Optional bGradientColor As Long = -1, _
                                 Optional bHoverColor As Long = -1)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'ActiveStatus'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' Fill the button with the appropriate backcolor
  
  Dim i           As Integer
  Dim bColor2Use  As Long

    '<:-):UPDATED: Multiple Dim line separated
  Dim focusOffset As Byte
  Dim isDown      As Byte
    '<:-):UPDATED: Multiple Dim line separated
    focusOffset = Abs(((myProps.bStatus And 1) = 1))
    isDown = Abs((myProps.bStatus And 6) = 6)
    If isDown Then
        ActiveStatus = 2
     Else
        ActiveStatus = focusOffset
    End If
    If bHoverColor < 0 Then
        bHoverColor = bColor
    End If
    If bTimerActive And (((myProps.bMode = lv_CommandButton And isDown = 0) Or (myProps.bValue = False And myProps.bMode > lv_CommandButton))) Then
        bColor2Use = bHoverColor
     Else
        bColor2Use = bColor
    End If
    '<:-)Auto-inserted With End...With Structure
    With myProps
        If .bShape > lv_FullDiagonal Then
            If bTimerActive Or ((.bStatus And 6) = 6) Or .bValue Or .bBackStyle <> 5 Then
                '<:-)Pleonasm Removed
                If .bShape < lv_RoundFlat Then
                    ' this little trick gives us a good edge to our round button
                    ' Too simple really--draw over the entire button a gradient background
                    ' then set a clipping region excluding the border size and draw the
                    ' rest of the button. Text/images can't overlap the border this way
                    i = .bGradient
                    SelectClipRgn ButtonDC.hDC, 0
                    If ButtonDC.ClipRgn Then
                        DeleteObject ButtonDC.ClipRgn
                    End If
                    If .bBackStyle = 5 Then
                        .bGradient = lv_Top2Bottom
                        DrawGradient vbWhite, vbGray
                        ButtonDC.ClipRgn = 0
                    End If
                    If ((.bStatus And 6) = 6) Or .bValue Then
                        '<:-)Pleonasm Removed
                        ButtonDC.ClipRgn = CreateEllipticRgn(1, 1, UserControl.ScaleWidth, UserControl.ScaleHeight)
                        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                        SelectClipRgn ButtonDC.hDC, ButtonDC.ClipRgn
                        DeleteObject ButtonDC.ClipRgn
                        .bGradient = lv_Top2Bottom
                        DrawGradient vbGray, vbWhite
                        ButtonDC.ClipRgn = CreateEllipticRgn(2, 2, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1)
                        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                        SelectClipRgn ButtonDC.hDC, ButtonDC.ClipRgn
                     Else
                        ButtonDC.ClipRgn = CreateEllipticRgn(1, 1, UserControl.ScaleWidth - 0, UserControl.ScaleHeight - 0)
                        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                        SelectClipRgn ButtonDC.hDC, ButtonDC.ClipRgn
                    End If
                    .bGradient = i
                 Else
                    SelectClipRgn ButtonDC.hDC, 0
                    If ButtonDC.ClipRgn Then
                        DeleteObject ButtonDC.ClipRgn
                    End If
                    SetButtonColors True, ButtonDC.hDC, cObj_Pen, vbBlack, , 2
                    Arc ButtonDC.hDC, 0, 0, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1, 0, 0, 0, 0
                    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                    ButtonDC.ClipRgn = CreateEllipticRgn(1, 1, UserControl.ScaleWidth, UserControl.ScaleHeight)
                    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                    SelectClipRgn ButtonDC.hDC, ButtonDC.ClipRgn
                End If
             Else
                If .bValue = False Then
                    If .bBackStyle = 5 Then
                        If bTimerActive = False Then
                            '<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
                            SelectClipRgn ButtonDC.hDC, 0
                        End If
                    End If '<:-)Short Circuit inserted this line
                End If '<:-)Short Circuit inserted this line
            End If
        End If
    End With 'myProps
    If myProps.bGradient And myProps.bValue = False Then
        If bTimerActive And ((myProps.bStatus And 6) = 6) = False And (myProps.bGradientColor <> myProps.bBackHover) Then
            '<:-)Pleonasm Removed
            DrawRect ButtonDC.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, bHoverColor
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
         Else
            If bGradientColor < 0 Then
                bGradientColor = bColor
            End If
            DrawGradient bColor, bGradientColor
        End If
     Else
        If myProps.bBackStyle = 2 And (UserControl.Enabled Or myProps.bMode > lv_CommandButton) Then
            '<:-)Pleonasm Removed
            For i = 1 To ScaleHeight - 1
                DrawRect ButtonDC.hDC, 1, i, UserControl.ScaleWidth - 1, i + 1, ShadeColor(bColor2Use, -(25 / UserControl.ScaleHeight) * i, True)
                '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            Next i
         Else
            DrawRect ButtonDC.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, bColor2Use
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        End If
    End If

End Sub

Private Sub DrawButtonBorder(polyPts() As POINTAPI, _
                             polyColors() As Long, _
                             ActiveStatus As Integer, _
                             Optional OuterBorderStyle As Long = -1)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'ActiveStatus'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  '<:-):SUGGESTION: Unused Parameter  'ActiveStatus As Integer' could be removed.
  ' This routine draws the border depending on the button style
  
  Dim i            As Integer

    'Dim J As Integer
    '<:-):WARNING: Unused Dim commented out
  Dim xColorRef    As Integer
    '<:-):UPDATED: Multiple Dim line separated
  Dim lBorderStyle As Long
  Dim lastColor    As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim polyOffset   As POINTAPI
    ' need to run special calculations for diagonal buttons
    '<:-)Auto-inserted With End...With Structure
    With myProps
        If .bShape > lv_Rectangular Then
            If .bShape < lv_Round3D Then
                '<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
                polyOffset.X = Abs(CInt(.bShape <> lv_RightDiagonal))
                polyOffset.Y = Abs(CInt(.bShape <> lv_LeftDiagonal))
            End If
        End If '<:-)Short Circuit inserted this line
    End With 'myProps
    ' calculate X,Y points for all three levels of borders
    polyPts(0).X = 2 + myProps.bSegPts.X - polyOffset.X * 4
    polyPts(0).Y = UserControl.ScaleHeight - 3
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    polyPts(1).X = 2 + polyOffset.X * 2
    polyPts(1).Y = 2
    polyPts(2).X = myProps.bSegPts.Y - 3 + polyOffset.Y * 3
    polyPts(2).Y = 2
    polyPts(3).X = UserControl.ScaleWidth - 3 - polyOffset.Y * 3
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    polyPts(3).Y = UserControl.ScaleHeight - 3
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    polyPts(4).X = 1 + myProps.bSegPts.X - polyOffset.X * 4
    polyPts(4).Y = UserControl.ScaleHeight - 3
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    For i = 5 To 9
        polyPts(i).X = polyPts(i - 5).X + Choose(i - 4, polyOffset.X - 1, -1 - polyOffset.X, 1 - polyOffset.Y, 1 + polyOffset.Y, -1, -1)
        polyPts(i).Y = polyPts(i - 5).Y + Choose(i - 4, 1, -1, -1, 1, 1, 1)
    Next i
    polyPts(10).X = myProps.bSegPts.X - polyOffset.X
    polyPts(10).Y = UserControl.ScaleHeight - 1 + polyOffset.X
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    polyPts(11).X = 0
    polyPts(11).Y = 0
    polyPts(12).X = myProps.bSegPts.Y - 1 + polyOffset.Y
    polyPts(12).Y = 0
    polyPts(13).X = UserControl.ScaleWidth - 1
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    polyPts(13).Y = UserControl.ScaleHeight - 1 + polyOffset.Y
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    polyPts(14).X = myProps.bSegPts.X - 1 - polyOffset.X * 2
    polyPts(14).Y = UserControl.ScaleHeight - 1
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    lastColor = -1
    For i = 0 To 13
        Select Case i
         Case Is < 4
            xColorRef = i + 1
         Case Is > 8
            xColorRef = i - 1   ' next line used for dashed borders
            If OuterBorderStyle > -1 Then
                lBorderStyle = OuterBorderStyle
            End If
         Case Else
            xColorRef = i
        End Select
        If (i <> 4 And i <> 9) Then
            ' if -1 is the color, we skip that level
            If polyColors(xColorRef) > -1 Then
                ' change the pen color if needed
                If lastColor <> polyColors(xColorRef) Then
                    SetButtonColors True, ButtonDC.hDC, cObj_Pen, polyColors(xColorRef), , , , lBorderStyle
                End If
                Polyline ButtonDC.hDC, polyPts(i), 2
                lastColor = polyColors(xColorRef)
            End If
        End If
    Next i
    If polyOffset.Y <> UserControl.ScaleWidth Then
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        ' tweak to ensure bottom, outer border draws correctly on diagonal buttons
        polyPts(15).X = UserControl.ScaleWidth - 1
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        polyPts(15).Y = UserControl.ScaleHeight - 1
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        Polyline ButtonDC.hDC, polyPts(14), 2
    End If

End Sub

Private Sub DrawButtonIcon(iRect As RECT)

  ' Routine will draw the button image
  
  Dim lImgCopy      As Long
  Dim imgWidth      As Long
  Dim imgHeight     As Long

    '<:-):UPDATED: Multiple Dim line separated
  Dim rcImage       As RECT
  Dim dRect         As RECT
    '<:-):UPDATED: Multiple Dim line separated
  Const MAGICROP    As Long = &HB8074A
    '<:-):UPDATED: Un-Typed Const with Hex(&H) value  changed to As Long
  Dim hMemDC        As Long
  Dim hBitmap       As Long
  Dim hOldBitmap    As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim hOldBrush     As Long
  Dim hOldBackColor As Long
  Dim hbrShadow     As Long
  Dim hbrHilite     As Long
    '<:-):UPDATED: Multiple Dim line separated
    If (myImage.SourceSize.X + myImage.SourceSize.Y) = 0 Then
        Exit Sub
    End If
    imgWidth = iRect.Right - iRect.Left
    imgHeight = iRect.Bottom - iRect.Top
    lImgCopy = CopyImage(myImage.Image.HANDLE, myImage.Type, imgWidth, imgHeight, 0)
    If lImgCopy = 0 Then
        Exit Sub
    End If
    ' destination rectangle for drawing on the DC
    dRect = iRect
    If UserControl.Enabled Then
        hMemDC = ButtonDC.hDC
     Else
        ' Create a temporary DC and bitmap to hold the image
        hMemDC = CreateCompatibleDC(ButtonDC.hDC)
        hBitmap = CreateCompatibleBitmap(ButtonDC.hDC, imgWidth + 1, imgHeight + 1)
        hOldBitmap = SelectObject(hMemDC, hBitmap)
        PatBlt hMemDC, 0, 0, imgWidth, imgHeight, WHITENESS
        OffsetRect dRect, -dRect.Left, -dRect.Top
    End If
    If myImage.Type = CI_ICON Then
        ' draw icon directly onto the temporary DC
        ' for icons, we can draw directly on the destination DC
        DrawIconEx hMemDC, dRect.Left, dRect.Top, lImgCopy, 0, 0, 0, 0, &H3
     Else
        ' draw transparent bitmap onto the temporary DC
        DrawTransparentBitmap hMemDC, dRect, lImgCopy, rcImage, , CLng(imgWidth), CLng(imgHeight)
    End If
    If UserControl.Enabled = False Then
        hOldBackColor = SetBkColor(ButtonDC.hDC, vbWhite)
        hbrHilite = CreateSolidBrush(ShadeColor(&HC0C0C0, 36, False))
        hbrShadow = CreateSolidBrush(ShadeColor(&HC0C0C0, -36, False))
        '<:-)Auto-inserted With End...With Structure
        With ButtonDC
            hOldBrush = SelectObject(.hDC, hbrHilite)
            BitBlt .hDC, iRect.Left - 1, iRect.Top - 1, imgWidth, imgHeight, hMemDC, 0, 0, MAGICROP
            SelectObject .hDC, hbrShadow
            BitBlt .hDC, iRect.Left, iRect.Top, imgWidth, imgHeight, hMemDC, 0, 0, MAGICROP
            SetBkColor .hDC, hOldBackColor
            SelectObject .hDC, hOldBrush
        End With 'ButtonDC
        SelectObject hMemDC, hOldBitmap
        DeleteObject hbrHilite
        If hbrShadow Then
            DeleteObject hbrShadow
        End If
        DeleteObject hBitmap
        DeleteDC hMemDC
    End If
    If myImage.Type = CI_ICON Then
        DestroyIcon lImgCopy
     Else
        DeleteObject lImgCopy
    End If

End Sub

Private Sub DrawCaptionIcon(bColor As Long, _
                            Optional tColorDisabled As Long = -1, _
                            Optional bOffsetTextDown As Boolean = False, _
                            Optional bSingleDisableColor As Boolean = False)

  ' Routine draws the caption & calls the DrawButtonIcon routine
  
  Dim tRect       As RECT
  Dim iRect       As RECT

    '<:-):UPDATED: Multiple Dim line separated
  Dim lColor      As Long
  Dim sCaption    As String               ' note Replace$ not compatible with VB5
  Dim shadeOffset As Integer
    ' set these rectangles & they may be adjusted a little later
    tRect = myProps.bRect
    iRect = myImage.iRect
    ' if the button is in a down position, we'll offset the image/text rects by 1
    If (((myProps.bStatus And 6) = 6) Or bOffsetTextDown) And myProps.bBackStyle <> 3 Then
        OffsetRect tRect, 1 + Int(myProps.bShape > lv_FullDiagonal), 1
        OffsetRect iRect, 1 + Int(myProps.bShape > lv_FullDiagonal), 1
    End If
    DrawButtonIcon iRect
    If Len(myProps.bCaption) = 0 Then
        Exit Sub
    End If
    sCaption = Replace$(myProps.bCaption, "||", vbNewLine)
    ' Setting text colors and offsets
    If UserControl.Enabled = False Then
        If tColorDisabled > -1 And myProps.bGradient = lv_NoGradient Then
            lColor = tColorDisabled
         Else
            lColor = vbWhite
            OffsetRect tRect, 1, 1
            bSingleDisableColor = False
        End If
     Else
        ' get the right forecolor to use
        If bTimerActive And ((myProps.bStatus And 6) = 6) = False Then
            '<:-)Pleonasm Removed
            lColor = ConvertColor(myProps.bForeHover)
         Else
            If myProps.bGradient Then
                lColor = ConvertColor(UserControl.ForeColor)
             Else
                If myProps.bBackStyle = 7 Then
                    lColor = tColorDisabled
                 Else
                    lColor = ConvertColor(UserControl.ForeColor)
                End If
            End If
        End If
        If (myProps.bCaptionStyle And UserControl.Enabled) Then
            '<:-)Pleonasm Removed
            ' drawing raised/sunken caption styles
            If myProps.bCaptionStyle = lv_Raised Then
                shadeOffset = 40
             Else
                shadeOffset = -40
            End If
            SetButtonColors True, ButtonDC.hDC, cObj_Text, ShadeColor(bColor, shadeOffset, False)
            OffsetRect tRect, -1, 0
            DrawText ButtonDC.hDC, sCaption, Len(sCaption), tRect, DT_WORDBREAK Or Choose(myProps.bCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
            SetButtonColors True, ButtonDC.hDC, cObj_Text, ShadeColor(bColor, -shadeOffset, False)
            OffsetRect tRect, 2, 2
            DrawText ButtonDC.hDC, sCaption, Len(sCaption), tRect, DT_WORDBREAK Or Choose(myProps.bCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
            OffsetRect tRect, -1, -1
        End If
    End If
    SetButtonColors True, ButtonDC.hDC, cObj_Text, lColor
    DrawText ButtonDC.hDC, sCaption, Len(sCaption), tRect, DT_WORDBREAK Or Choose(myProps.bCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
    If UserControl.Enabled = False Then
        If bSingleDisableColor = False Then
            '<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
            ' finish drawing the disabled caption
            SetButtonColors True, ButtonDC.hDC, cObj_Text, vbGray
            OffsetRect tRect, -1, -1
            DrawText ButtonDC.hDC, sCaption, Len(sCaption), tRect, DT_WORDBREAK Or Choose(myProps.bCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
        End If
    End If '<:-)Short Circuit inserted this line

End Sub

Private Sub DrawFocusRectangle(fColor As Long, _
                               bSolid As Boolean, _
                               bOnText As Boolean, _
                               polyPts() As POINTAPI)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameters 'bSolid, bOnText'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' Draws focus rectangles for the button style & button mode
  
  Dim focusOffset As Byte
  Dim bDownOffset As Byte

    '<:-):UPDATED: Multiple Dim line separated
  Dim polyOffset  As POINTAPI
  Dim fRect       As RECT
    If ((myProps.bStatus And 1) <> 1) Then
        Exit Sub
    End If
    ' round button
    If myProps.bShape > lv_FullDiagonal Then
        bOnText = True
    End If
    If myProps.bShape > lv_Rectangular And myProps.bShape < lv_Round3D Then
        ' diagonal buttons
        If myProps.bSegPts.X Then
            polyOffset.X = 1
        End If
        If myProps.bSegPts.Y < UserControl.ScaleWidth Then
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            polyOffset.Y = 1
        End If
        polyPts(0).X = 4 + myProps.bSegPts.X - polyOffset.X * 6
        polyPts(0).Y = UserControl.ScaleHeight - 5
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        polyPts(1).X = 4 + polyOffset.X * 4
        polyPts(1).Y = 4
        polyPts(2).X = myProps.bSegPts.Y - 5 + polyOffset.Y * 4
        polyPts(2).Y = 4
        polyPts(3).X = UserControl.ScaleWidth - 5 - polyOffset.Y * 6
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        polyPts(3).Y = UserControl.ScaleHeight - 5
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        polyPts(4).X = 3 + myProps.bSegPts.X - polyOffset.X * 4
        polyPts(4).Y = UserControl.ScaleHeight - 5
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        SetButtonColors True, ButtonDC.hDC, cObj_Pen, fColor
        Polyline ButtonDC.hDC, polyPts(0), 5
     Else
        If bOnText Then
            '<:-)Pleonasm Removed
            If Len(myProps.bCaption) Then
                fRect = myProps.bRect
             Else
                fRect = myImage.iRect
                If fRect.Bottom > UserControl.ScaleHeight - 4 Then
                    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                    fRect.Bottom = UserControl.ScaleHeight - 4
                    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                End If
                If fRect.Right > UserControl.ScaleWidth - 4 Then
                    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                    fRect.Right = UserControl.ScaleWidth - 4
                    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
                End If
            End If
            If myProps.bRect.Left > 4 + myProps.bSegPts.X And myProps.bRect.Right < myProps.bSegPts.Y - 4 Then
                focusOffset = 2
             Else
                focusOffset = 1
            End If
            bDownOffset = Abs((((myProps.bStatus And 6) = 6) Or myProps.bValue) And myProps.bBackStyle <> 3)
            '<:-)Pleonasm Removed
            OffsetRect fRect, -focusOffset + bDownOffset * Abs(myProps.bShape < lv_Round3D), -focusOffset + bDownOffset
            fRect.Right = fRect.Right + focusOffset * 2 + bDownOffset * Abs(myProps.bShape < lv_Round3D)
            fRect.Bottom = fRect.Bottom + focusOffset * 2 + bDownOffset
            ' for now, only used on Java buttons & round buttons
            If bSolid Then
                SetButtonColors True, ButtonDC.hDC, cObj_Pen, fColor
                polyPts(0).X = fRect.Left
                polyPts(0).Y = fRect.Top
                polyPts(1).X = fRect.Right - 1
                polyPts(1).Y = fRect.Top
                polyPts(2).X = fRect.Right - 1
                polyPts(2).Y = fRect.Bottom - 1
                polyPts(3).X = fRect.Left
                polyPts(3).Y = fRect.Bottom - 1
                polyPts(4).X = fRect.Left
                polyPts(4).Y = fRect.Top
                Polyline ButtonDC.hDC, polyPts(0), 5
             Else            ' for now, only used on Macintosh buttons
                DrawFocusRect ButtonDC.hDC, fRect
            End If
         Else
            '<:-)Auto-inserted With End...With Structure
            With fRect
                .Top = 0
                .Left = 0
                .Right = myProps.bSegPts.Y - (myProps.bSegPts.X + 8)
                .Bottom = UserControl.ScaleHeight - 8
                '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            End With 'fRect
            OffsetRect fRect, 4 + myProps.bSegPts.X, 4
            ' used when option buttons/checkboxes have focus if Value=True
            If bSolid Then
                polyPts(0).X = fRect.Left
                polyPts(0).Y = fRect.Bottom
                polyPts(1).X = fRect.Left
                polyPts(1).Y = fRect.Top
                polyPts(2).X = fRect.Right
                polyPts(2).Y = fRect.Top
                polyPts(3).X = fRect.Right
                polyPts(3).Y = fRect.Bottom
                polyPts(4).X = fRect.Left
                polyPts(4).Y = fRect.Bottom
                SetButtonColors True, ButtonDC.hDC, cObj_Pen, fColor
                Polyline ButtonDC.hDC, polyPts(0), 5
             Else
                DrawFocusRect ButtonDC.hDC, fRect
            End If
        End If
    End If

End Sub

Private Sub DrawGradient(ByVal Color1 As Long, _
                         ByVal Color2 As Long)

  Dim mRect     As RECT
  Dim i         As Long
  Dim rctOffset As Integer

    '<:-):UPDATED: Multiple Dim line separated
  Dim PixelStep As Long
  Dim rIndex    As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim Colors()  As Long
    ' The gist is to draw 1 pixel rectangles of various colors to create
    ' the gradient effect. If the size of the rectangle is greater than a
    ' quarter of the screen size, we'll step it up to 2 pixel rectangles
    ' to speed things up a bit
    On Error Resume Next
    mRect.Right = UserControl.ScaleWidth
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    mRect.Bottom = UserControl.ScaleHeight
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    rctOffset = 1
    If myProps.bGradient < 3 Then
        If (Screen.Width \ Screen.TwipsPerPixelX) \ UserControl.ScaleWidth < 4 Then
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            PixelStep = UserControl.ScaleWidth \ 2
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            rctOffset = 2
         Else
            PixelStep = UserControl.ScaleWidth
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        End If
     Else
        If (Screen.Height \ Screen.TwipsPerPixelY) \ UserControl.ScaleHeight < 4 Then
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            PixelStep = UserControl.ScaleHeight \ 2
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            rctOffset = 2
         Else
            PixelStep = UserControl.ScaleHeight
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        End If
    End If
    ReDim Colors(0 To PixelStep - 1) As Long
    LoadGradientColors Colors(), Color1, Color2
    If myProps.bGradient > 2 Then
        mRect.Bottom = rctOffset
     Else
        mRect.Right = rctOffset
    End If
    For i = 0 To PixelStep - 1
        If myProps.bGradient Mod 2 Then
            rIndex = i
         Else
            rIndex = PixelStep - i - 1
        End If
        DrawRect ButtonDC.hDC, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, Colors(rIndex)
        If myProps.bGradient > 2 Then
            OffsetRect mRect, 0, rctOffset
         Else
            OffsetRect mRect, rctOffset, 0
        End If
    Next i
    On Error GoTo 0 '<:-):RISK: Turns off 'On Error Resume Next' in routine( Good coding but may not be what you want)

End Sub

Private Sub DrawRect(m_hDC As Long, _
                     ByVal X1 As Long, _
                     ByVal Y1 As Long, _
                     ByVal X2 As Long, _
                     ByVal Y2 As Long, _
                     tColor As Long, _
                     Optional pColor As Long = -1, _
                     Optional PenWidth As Long = 0, _
                     Optional PenStyle As Long = 0)

  ' Simple routine to draw a rectangle

    If pColor <> -1 Then
        SetButtonColors True, m_hDC, cObj_Pen, pColor, , PenWidth, , PenStyle
    End If
    SetButtonColors True, m_hDC, cObj_Brush, tColor, (pColor = -1)
    Call Rectangle(m_hDC, X1, Y1, X2, Y2)

End Sub

Private Sub DrawTransparentBitmap(lHDCdest As Long, _
                                  destRect As RECT, _
                                  lBMPsource As Long, _
                                  bmpRect As RECT, _
                                  Optional lMaskColor As Long = -1, _
                                  Optional lNewBmpCx As Long, _
                                  Optional lNewBmpCy As Long)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameters 'lHDCdest, lBMPsource'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  '<:-):SUGGESTION: Unused Parameter  'Optional lMaskColor As Long = -1' could be removed.
  
  Const DSna         As Long = &H220326            '0x00220326

    '<:-):UPDATED: Un-Typed Const with Hex(&H) value  changed to As Long
  Dim lMask2Use      As Long                       'COLORREF
  Dim lBmMask        As Long
  Dim lBmAndMem      As Long
  Dim lBmColor       As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim lBmObjectOld   As Long
  Dim lBmMemOld      As Long
  Dim lBmColorOld    As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim lHDCMem        As Long
  Dim lHDCscreen     As Long
  Dim lHDCsrc        As Long
  Dim lHDCMask       As Long
  Dim lHDCcolor      As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim X              As Long
  Dim Y              As Long
  Dim srcX           As Long
  Dim srcY           As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim lRatio(0 To 1) As Single
  Dim hPalOld        As Long
  Dim hPalMem        As Long
    '<:-):UPDATED: Multiple Dim line separated
    ' =====================================================================
    ' A pretty good transparent bitmap maker I use in several projects
    ' Modified here to remove stuff I wont use (i.e., Flipping/Rotating images)
    ' =====================================================================
    lHDCscreen = GetDC(0&)
    lHDCsrc = CreateCompatibleDC(lHDCscreen)     'Create a temporary HDC compatible to the Destination HDC
    SelectObject lHDCsrc, lBMPsource             'Select the bitmap
    srcX = lNewBmpCx                  'Get width of bitmap
    srcY = lNewBmpCy                 'Get height of bitmap
    If bmpRect.Right = 0 Then
        bmpRect.Right = srcX
     Else
        srcX = bmpRect.Right - bmpRect.Left
    End If
    If bmpRect.Bottom = 0 Then
        bmpRect.Bottom = srcY
     Else
        srcY = bmpRect.Bottom - bmpRect.Top
    End If
    If (destRect.Right) = 0 Then
        X = lNewBmpCx
     Else
        X = (destRect.Right - destRect.Left)
    End If
    If (destRect.Bottom) = 0 Then
        Y = lNewBmpCy
     Else
        Y = (destRect.Bottom - destRect.Top)
    End If
    If lNewBmpCx > X Or lNewBmpCy > Y Then
        lRatio(0) = (X / lNewBmpCx)
        lRatio(1) = (Y / lNewBmpCy)
        If lRatio(1) < lRatio(0) Then
            lRatio(0) = lRatio(1)
        End If
        lNewBmpCx = lRatio(0) * lNewBmpCx
        lNewBmpCy = lRatio(0) * lNewBmpCy
        Erase lRatio
    End If
    lMask2Use = ConvertColor(GetPixel(lHDCsrc, 0, 0))
    'Create some DCs & bitmaps
    lHDCMask = CreateCompatibleDC(lHDCscreen)
    lHDCMem = CreateCompatibleDC(lHDCscreen)
    lHDCcolor = CreateCompatibleDC(lHDCscreen)
    lBmColor = CreateCompatibleBitmap(lHDCscreen, srcX, srcY)
    lBmAndMem = CreateCompatibleBitmap(lHDCscreen, X, Y)
    lBmMask = CreateBitmap(srcX, srcY, 1&, 1&, ByVal 0&)
    lBmColorOld = SelectObject(lHDCcolor, lBmColor)
    lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
    lBmObjectOld = SelectObject(lHDCMask, lBmMask)
    ReleaseDC 0&, lHDCscreen
    ' ====================== Start working here ======================
    SetMapMode lHDCMem, GetMapMode(lHDCdest)
    hPalMem = SelectPalette(lHDCMem, 0, True)
    RealizePalette lHDCMem
    BitBlt lHDCMem, 0&, 0&, X, Y, lHDCdest, destRect.Left, destRect.Top, vbSrcCopy
    hPalOld = SelectPalette(lHDCcolor, 0, True)
    RealizePalette lHDCcolor
    SetBkColor lHDCcolor, GetBkColor(lHDCsrc)
    SetTextColor lHDCcolor, GetTextColor(lHDCsrc)
    BitBlt lHDCcolor, 0&, 0&, srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, vbSrcCopy
    SetBkColor lHDCcolor, lMask2Use
    SetTextColor lHDCcolor, vbWhite
    BitBlt lHDCMask, 0&, 0&, srcX, srcY, lHDCcolor, 0&, 0&, vbSrcCopy
    SetTextColor lHDCcolor, vbBlack
    SetBkColor lHDCcolor, vbWhite
    BitBlt lHDCcolor, 0, 0, srcX, srcY, lHDCMask, 0, 0, DSna
    StretchBlt lHDCMem, 0, 0, lNewBmpCx, lNewBmpCy, lHDCMask, 0&, 0&, srcX, srcY, vbSrcAnd
    StretchBlt lHDCMem, 0&, 0&, lNewBmpCx, lNewBmpCy, lHDCcolor, 0, 0, srcX, srcY, vbSrcPaint
    BitBlt lHDCdest, destRect.Left, destRect.Top, X, Y, lHDCMem, 0&, 0&, vbSrcCopy
    'Delete memory bitmaps & DCs
    DeleteObject SelectObject(lHDCcolor, lBmColorOld)
    DeleteObject SelectObject(lHDCMask, lBmObjectOld)
    DeleteObject SelectObject(lHDCMem, lBmMemOld)
    DeleteDC lHDCMem
    DeleteDC lHDCMask
    DeleteDC lHDCcolor
    DeleteDC lHDCsrc

End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Determines if events are fired for this button."
Attribute Enabled.VB_UserMemId = -514

    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(bEnabled As Boolean)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'bEnabled'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' Enables or disables the button

    If bEnabled = UserControl.Enabled Then
        Exit Property
    End If
    UserControl.Enabled = bEnabled
    If myProps.bBackStyle = 3 And myProps.bMode = lv_CommandButton Then
        ' java disabled seems to not have the lower-left/upper-right pixels transparent
        DelayDrawing True
        CreateButtonRegion
        CalculateBoundingRects
        DelayDrawing False
     Else
        Refresh
    End If
    PropertyChanged "Enabled"

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Font used to display the caption."

    Set Font = UserControl.Font

End Property

Public Property Set Font(nFont As StdFont)

  ' Sets the control's font & also the logical font to use on off-screen DC

    Set UserControl.Font = nFont
    GetGDIMetrics "Font"
    CalculateBoundingRects          ' recalculate caption's text/image bounding rects
    Refresh
    PropertyChanged "Font"

End Property

Public Property Get FontStyle() As FontStyles
Attribute FontStyle.VB_Description = "Various font attributes that can be changed directly."

  Dim nStyle As Integer

    '<:-)Auto-inserted With End...With Structure
    With UserControl
        nStyle = nStyle Or Abs(.Font.Bold) * 2
        nStyle = nStyle Or Abs(.Font.Italic) * 4
        nStyle = nStyle Or Abs(.Font.Underline) * 8
    End With 'UserControl
    FontStyle = nStyle

End Property

Public Property Let FontStyle(nStyle As FontStyles)

  ' Allows direct changes to font attributes

    With UserControl.Font
        .Bold = ((nStyle And lv_Bold) = lv_Bold)
        .Italic = ((nStyle And lv_Italic) = lv_Italic)
        .Underline = ((nStyle And lv_Underline) = lv_Underline)
    End With
    GetGDIMetrics "Font"
    CalculateBoundingRects
    PropertyChanged "Font"
    Refresh

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "The color of the caption's font ."

    ForeColor = UserControl.ForeColor

End Property

Public Property Let ForeColor(nColor As OLE_COLOR)

  ' Sets the caption text color

    If nColor = UserControl.ForeColor Then
        Exit Property
    End If
    UserControl.ForeColor = nColor
    If myProps.bLockHover = lv_LockTextandBackColor Or myProps.bLockHover = lv_LockTextColorOnly Then
        Me.HoverForeColor = UserControl.ForeColor
    End If
    bNoRefresh = False
    RedrawButton True
    PropertyChanged "cFore"

End Property

Private Sub GetGDIMetrics(sObject As String)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'sObject'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' This routine caches information we don't want to keep gathering every time a button is redrawn.
  
  Dim newFont As LOGFONT
  Dim bmpInfo As BITMAP
  Dim icoInfo As ICONINFO

    '<:-):UPDATED: Multiple Dim line separated
    Select Case sObject
     Case "Font"
        ' called when font is changed or control is initialized
        '<:-)Auto-inserted With End...With Structure
        With newFont
            .lfCharSet = 1
            .lfFaceName = UserControl.Font.Name & vbNullChar
            .lfHeight = (UserControl.Font.Size * -20) / Screen.TwipsPerPixelY
            .lfWeight = UserControl.Font.Weight
            .lfItalic = Abs(CInt(UserControl.Font.Italic))
            .lfStrikeOut = Abs(CInt(UserControl.Font.Strikethrough))
            .lfUnderline = Abs(CInt(UserControl.Font.Underline))
        End With 'newFont
        If ButtonDC.OldFont Then
            DeleteObject SelectObject(ButtonDC.hDC, CreateFontIndirect(newFont))
         Else
            ButtonDC.OldFont = SelectObject(ButtonDC.hDC, CreateFontIndirect(newFont))
        End If
     Case "Picture"
        ' get key image information
        If Not myImage.Image Is Nothing Then
            GetGDIObject myImage.Image.HANDLE, LenB(bmpInfo), bmpInfo
            If bmpInfo.bmBits = 0 Then
                GetIconInfo myImage.Image.HANDLE, icoInfo
                If icoInfo.hbmColor <> 0 Then
                    ' downside... API creates 2 bitmaps that we need to destroy since they aren't used in this
                    ' routine & are not destroyed automatically. To prevent memory leak, we destroy them here
                    '<:-)Auto-inserted With End...With Structure
                    With icoInfo
                        GetGDIObject .hbmColor, LenB(bmpInfo), bmpInfo
                        DeleteObject .hbmColor
                        If .hbmMask <> 0 Then
                            DeleteObject .hbmMask
                        End If
                    End With 'icoInfo
                    myImage.Type = CI_ICON        ' flag indicating image is an icon
                End If
             Else
                myImage.Type = CI_BITMAP     ' flag indicating image is a bitmap
            End If
        End If
        myImage.SourceSize.X = bmpInfo.bmWidth
        myImage.SourceSize.Y = bmpInfo.bmHeight
     Case "BackColor"
        adjBackColorUp = ConvertColor(curBackColor)
        adjBackColorDn = adjBackColorUp
        adjHoverColor = ConvertColor(myProps.bBackHover)
        If myProps.bBackStyle = 7 Then
            adjBackColorUp = ShadeColor(adjBackColorUp, &H1F, False)
            adjBackColorDn = ShadeColor(vbGray, -&H10, False)
            adjHoverColor = ShadeColor(adjHoverColor, &H1F, False)
            cCheckBox = ShadeColor(vbGray, &H10, True)
         ElseIf myProps.bBackStyle = 2 Then
            adjBackColorUp = ShadeColor(adjBackColorUp, &H30, True)
            adjHoverColor = ShadeColor(adjHoverColor, &H30, True)
            adjBackColorDn = ShadeColor(adjBackColorUp, -&H20, True)
            cCheckBox = ShadeColor(vbWhite, -&H20, True)
         Else
            If myProps.bBackStyle = 3 Then
                adjBackColorDn = ShadeColor(vbGray, &HC, False)
                cCheckBox = ShadeColor(adjBackColorDn, &H1F, False)
             Else
                cCheckBox = ShadeColor(vbWhite, -&H20, False)
            End If
        End If
    End Select

End Sub

Private Sub GetSetOffDC(bSet As Boolean)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'bSet'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' This sets up our off screen DC & pastes results onto our control.
  
  Dim hBmp As Long

    If bSet Then
        '<:-)Pleonasm Removed
        '<:-)Auto-inserted With End...With Structure
        With ButtonDC
            If .hDC = 0 Then
                .hDC = CreateCompatibleDC(UserControl.hDC)
                SetBkMode .hDC, 3&
                ' by pulling these objects now, we ensure no memory leaks &
                ' changing the objects as needed can be done in 1 line of code
                ' in the SetButtonColors routine
                .OldBrush = SelectObject(.hDC, CreateSolidBrush(0&))
                .OldPen = SelectObject(.hDC, CreatePen(0&, 1&, 0&))
                GetGDIMetrics "Font"
            End If
        End With 'ButtonDC
        If ButtonDC.OldBitmap = 0 Then
            hBmp = CreateCompatibleBitmap(UserControl.hDC, UserControl.ScaleWidth, UserControl.ScaleHeight)
            '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
            ButtonDC.OldBitmap = SelectObject(ButtonDC.hDC, hBmp)
        End If
     Else
        BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, ButtonDC.hDC, 0, 0, vbSrcCopy
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    End If

End Sub

Public Property Get GradientColor() As OLE_COLOR
Attribute GradientColor.VB_Description = "Secondary color used for gradient shades. The BackColor property is the primary color."

    GradientColor = myProps.bGradientColor

End Property

Public Property Let GradientColor(nColor As OLE_COLOR)

  ' Sets the gradient color. Gradients are used this way...
  ' Shade from BackColor to GradientColor
  ' GradientMode must be set
  '<:-)Auto-inserted With End...With Structure

    With myProps
        If (.bLockHover = lv_LockTextandBackColor Or .bLockHover = lv_LockBackColorOnly) And .bGradient > lv_NoGradient Then
            .bBackHover = nColor
            .bBackHover = Me.HoverBackColor
        End If
    End With 'myProps
    myProps.bGradientColor = nColor
    GetGDIMetrics "BackColor"
    If myProps.bGradient Then
        Refresh
    End If
    PropertyChanged "cGradient"

End Property

Public Property Get GradientMode() As GradientConstants
Attribute GradientMode.VB_Description = "Various directions to draw the gradient shading."

    GradientMode = myProps.bGradient

End Property

Public Property Let GradientMode(nOpt As GradientConstants)

  ' Sets the direction of gradient shading

    If nOpt < lv_NoGradient Or nOpt > lv_Bottom2Top Then
        Exit Property
    End If
    myProps.bGradient = nOpt
    If myProps.bLockHover = lv_LockBackColorOnly Or myProps.bLockHover = lv_LockTextandBackColor Then
        If nOpt > lv_NoGradient Then
            myProps.bBackHover = myProps.bGradientColor
         Else
            myProps.bBackHover = curBackColor
        End If
        myProps.bBackHover = Me.HoverBackColor
        GetGDIMetrics "BackColor"
    End If
    Refresh
    PropertyChanged "Gradient"

End Property

Public Property Get hDC() As Long

  ' Makes the control's hDC availabe at runtime

    hDC = UserControl.hDC

End Property

Public Property Get HoverBackColor() As OLE_COLOR
Attribute HoverBackColor.VB_Description = "Color of button background when mouse is hovering over it. Affects the HoverLockColors property."

    HoverBackColor = myProps.bBackHover

End Property

Public Property Let HoverBackColor(nColor As OLE_COLOR)

  ' Changes the backcolor when mouse is over the button
  ' Changing this property will affect the type of HoverLock

    If myProps.bBackHover = nColor Then
        Exit Property
    End If
    myProps.bBackHover = nColor
    If nColor <> curBackColor Then
        If myProps.bLockHover = lv_LockTextandBackColor Then
            myProps.bLockHover = lv_LockTextColorOnly
         Else
            If myProps.bLockHover = lv_LockBackColorOnly Then
                myProps.bLockHover = lv_NoLocks
            End If
        End If
    End If
    myProps.bLockHover = Me.HoverColorLocks
    GetGDIMetrics "BackColor"
    PropertyChanged "cBHover"

End Property

Public Property Get HoverColorLocks() As HoverLockConstants
Attribute HoverColorLocks.VB_Description = "Can ensure the hover colors match the caption and back colors. Click for more options."

    HoverColorLocks = myProps.bLockHover

End Property

Public Property Let HoverColorLocks(nLock As HoverLockConstants)

  ' Has two purposes.
  ' 1. If the lock wasn't set but is now set, then setting it will
  ' force HoverForeColor=ForeColor & HoverBackColor=Backcolor
  ' If gradeints in use, then HoverBackColor=GradientColor
  ' 2. If the lock was already set, then changing BackColor
  ' will force HoverBackColor to match. If gradients are used then
  ' it will force HoverBackColor to match GradientColor
  ' It will also force HoverForeColor to match ForeColor.
  ' After the locks have been set, manually changing the
  ' HoverForeColor, HoverBackColor will adjust/remove the lock
  '<:-)Auto-inserted With End...With Structure

    With myProps
        .bLockHover = nLock
        If .bLockHover = lv_LockTextandBackColor Or .bLockHover = lv_LockBackColorOnly Then
            If .bGradient Then
                .bBackHover = .bGradientColor
             Else
                .bBackHover = curBackColor
            End If
            PropertyChanged "cBHover"
        End If
    End With 'myProps
    If myProps.bLockHover = lv_LockTextandBackColor Or myProps.bLockHover = lv_LockTextColorOnly Then
        myProps.bForeHover = UserControl.ForeColor
        PropertyChanged "cFHover"
    End If
    myProps.bBackHover = Me.HoverBackColor
    myProps.bForeHover = Me.HoverForeColor
    GetGDIMetrics "BackColor"
    PropertyChanged "LockHover"

End Property

Public Property Get HoverForeColor() As OLE_COLOR
Attribute HoverForeColor.VB_Description = "Color of button caption's text when mouse is hovering over it. Affects the HoverLockColors property."

    HoverForeColor = myProps.bForeHover

End Property

Public Property Let HoverForeColor(nColor As OLE_COLOR)

  ' Changes the text color when mouse is over the button
  ' Changing this property will affect the type of HoverLock

    If myProps.bForeHover = nColor Then
        Exit Property
    End If
    myProps.bForeHover = nColor
    PropertyChanged "cFHover"
    If nColor <> UserControl.ForeColor Then
        If myProps.bLockHover = lv_LockTextandBackColor Then
            myProps.bLockHover = lv_LockBackColorOnly
         Else
            If myProps.bLockHover = lv_LockTextColorOnly Then
                myProps.bLockHover = lv_NoLocks
            End If
        End If
    End If
    myProps.bLockHover = Me.HoverColorLocks
    PropertyChanged "cFHover"

End Property

Public Property Get hWnd() As Long

  ' Makes the control's hWnd available at runtime

    hWnd = UserControl.hWnd

End Property

Private Sub LoadGradientColors(Colors() As Long, _
                               ByVal Color1 As Long, _
                               ByVal Color2 As Long)

  Dim i              As Integer
  Dim j              As Integer

    '<:-):UPDATED: Multiple Dim line separated
  Dim sBase(0 To 2)  As Single
  Dim xBase(0 To 2)  As Long
  Dim lRatio(0 To 2) As Single
    ' routine adds/removes colors between a range of two colors
    ' Used by the DrawGradient routine. A variation of the ShadeColor routine
    sBase(0) = (Color1 And &HFF)
    sBase(1) = (Color1 And &HFF00&) / 255&
    sBase(2) = (Color1 And &HFF0000) / &HFF00&
    xBase(0) = (Color2 And &HFF)
    xBase(1) = (Color2 And &HFF00&) / 255&
    xBase(2) = (Color2 And &HFF0000) / &HFF00&
    For j = 0 To 2
        lRatio(j) = (xBase(j) - sBase(j)) / UBound(Colors)
    Next j
    Colors(0) = Color1
    For j = 1 To UBound(Colors)
        For i = 0 To 2
            sBase(i) = sBase(i) + lRatio(i)
            If sBase(i) > 255 Then
                sBase(i) = 255
            End If
            If sBase(i) < 0 Then
                sBase(i) = 0
            End If
        Next i
        Colors(j) = Int(sBase(0)) + 256& * Int(sBase(1)) + 65536 * Int(sBase(2))
    Next j
    Erase sBase
    Erase xBase
    Erase lRatio

End Sub

Public Property Get Mode() As ButtonModeConstants
Attribute Mode.VB_Description = "Command button, check box or option button mode"

    Mode = myProps.bMode

End Property

Public Property Let Mode(nMode As ButtonModeConstants)

  ' Sets the button function/mode

    If nMode < lv_CommandButton Or nMode > lv_OptionButton Then
        Exit Property
    End If
    If myProps.bMode = lv_OptionButton Then
        ' option buttons. Need to remove references if the Mode changed
        If nMode < lv_OptionButton Then
            Call ToggleOptionButtons(-1)
        End If
    End If
    If myProps.bMode < lv_OptionButton Then
        If nMode = lv_OptionButton Then
            '<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
            Call ToggleOptionButtons(1) ' add this instance to optionbutton collection
        End If
    End If '<:-)Short Circuit inserted this line
    If nMode = lv_CommandButton Then
        If myProps.bMode > lv_CommandButton Then
            '<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
            Me.Value = False
        End If
    End If '<:-)Short Circuit inserted this line
    myProps.bMode = nMode
    Refresh
    PropertyChanged "Mode"

End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_Description = "Icon or cursor used to display when mouse is over the button. MousePointer must be set to Custom."

    Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(nIcon As StdPicture)

  ' Sets the mouse icon for the button, MousePointer must be vbCustom

    On Error GoTo ShowPropertyError
    Set UserControl.MouseIcon = nIcon
    If Not nIcon Is Nothing Then
        Me.MousePointer = vbCustom
        PropertyChanged "mIcon"
    End If

Exit Property

ShowPropertyError:
    If Ambient.UserMode = False Then
        MsgBox Err.Description, vbInformation + vbOKOnly, "Select .ico Or .cur Files Only"
    End If

End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Various optional mouse pointers to use when mouse is over the button"

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(nPointer As MousePointerConstants)

  ' Sets the mouse pointer for the button

    UserControl.MousePointer = nPointer
    PropertyChanged "mPointer"

End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "The image used to display on the button."

    Set Picture = myImage.Image

End Property

Public Property Set Picture(xPic As StdPicture)

  ' Sets the button image which to display
  '<:-)Auto-inserted With End...With Structure

    With myImage
        Set .Image = xPic
        If .Size = 0 Then
            .Size = 16
        End If
    End With 'myImage
    GetGDIMetrics "Picture"
    CalculateBoundingRects              ' recalculate button's text/image bounding rects
    Refresh
    PropertyChanged "Image"

End Property

Public Property Get PictureAlign() As ImagePlacementConstants
Attribute PictureAlign.VB_Description = "Alignment of the button image in relation to the caption and/or button."

    PictureAlign = myImage.Align

End Property

Public Property Let PictureAlign(ImgAlign As ImagePlacementConstants)

  ' Image alignment options for button (6 different positions)

    If ImgAlign < lv_LeftEdge Or ImgAlign > lv_BottomCenter Then
        Exit Property
    End If
    myImage.Align = ImgAlign
    If ImgAlign = lv_BottomCenter Or ImgAlign = lv_TopCenter Then
        CaptionAlign = vbCenter
    End If
    CalculateBoundingRects              ' recalculate button's text/image bounding rects
    Refresh
    PropertyChanged "ImgAlign"

End Property

Public Property Get PictureSize() As ImageSizeConstants
Attribute PictureSize.VB_Description = "Size of button image. Last two options automatically center image."

    If myImage.Size = 0 Then
        myImage.Size = 16
    End If
    ' parameters are 0,1,2,3,4 & 5, but we store them as 16,24,32,40, & 44
    PictureSize = Choose(myImage.Size / 8 - 1, lv_16x16, lv_24x24, lv_32x32, lv_Fill_Stretch, lv_Fill_ScaleUpDown)

End Property

Public Property Let PictureSize(nSize As ImageSizeConstants)

  ' Sets up to 5 picture sizes

    If PictureSize < lv_16x16 Or PictureSize > lv_Fill_ScaleUpDown Then
        Exit Property
    End If
    myImage.Size = (nSize + 2) * 8      ' I just want the pixel size
    CalculateBoundingRects              ' recalculate text/image bounding rects
    Refresh
    PropertyChanged "ImgSize"

End Property

Private Sub RedrawButton(bDrawEntireButton As Boolean)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'bDrawEntireButton'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  '<:-):SUGGESTION: Unused Parameter  'bDrawEntireButton As Boolean' could be removed.
  ' ==================================================
  ' Main switchboard routine for redrawing a button
  ' ==================================================
  
  Dim polyPts(0 To 15)    As POINTAPI
  Dim polyColors(1 To 12) As Long

    '<:-):UPDATED: Multiple Dim line separated
  Dim ActiveStatus        As Integer
    If bNoRefresh Then
        '<:-)Pleonasm Removed
        Exit Sub
    End If
    Select Case myProps.bBackStyle
     Case 0
        DrawButton_Win95 polyPts(), polyColors(), ActiveStatus
     Case 1
        DrawButton_Win31 polyPts(), polyColors(), ActiveStatus
     Case 2
        DrawButton_WinXP polyPts(), polyColors(), ActiveStatus
     Case 3
        DrawButton_Java polyPts(), polyColors(), ActiveStatus
     Case 4
        DrawButton_Flat polyPts(), polyColors(), ActiveStatus
     Case 5
        DrawButton_Hover polyPts(), polyColors(), ActiveStatus
     Case 6
        DrawButton_Netscape polyPts(), polyColors(), ActiveStatus
     Case 7
        DrawButton_Macintosh polyPts(), polyColors(), ActiveStatus
    End Select
    Erase polyPts()
    Erase polyColors()
    GetSetOffDC False    ' copy the offscreen DC onto the control
    UserControl.Refresh

End Sub

Public Sub Refresh()

  ' //////////////////// GENERAL FUNCTIONS, PUBLIC \\\\\\\\\\\\\\\\\\\\\
  ' Refreshes the button & can be called from any form/module

    RedrawButton True

End Sub

Public Property Get ResetDefaultColors() As Boolean
Attribute ResetDefaultColors.VB_Description = "Resets button's back color and text color to Window's standard. The hover properties are also reset."

    ResetDefaultColors = False

End Property

Public Property Let ResetDefaultColors(nDefault As Boolean)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'nDefault'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' Resets the BackColor, ForeColor, GradientColor,
  ' HoverBackColor & HoverForeColor to defaults

    If Ambient.UserMode Or nDefault = False Then
        Exit Property
    End If
    DelayDrawing True
    curBackColor = vbButtonFace
    '<:-)Auto-inserted With End...With Structure
    With Me
        .ForeColor = vbButtonText
        .GradientColor = vbButtonFace
        .GradientMode = lv_NoGradient
        .HoverColorLocks = lv_LockTextandBackColor
        myProps.bGradientColor = .GradientColor
    End With 'Me
    GetGDIMetrics "BackColor"
    DelayDrawing False
    PropertyChanged "cGradient"
    PropertyChanged "cBack"

End Property

Private Sub SetButtonColors(bSet As Boolean, _
                            m_hDC As Long, _
                            TypeObject As ColorObjects, _
                            lColor As Long, _
                            Optional bSamePenColor As Boolean = True, _
                            Optional PenWidth As Long = 1, _
                            Optional bSwapPens As Boolean = False, _
                            Optional PenStyle As Long = 0)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameters 'bSet, m_hDC, lColor'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  '<:-):SUGGESTION: Unused Parameter  'Optional bSwapPens As Boolean = False' could be removed.
  ' This is the basic routine that sets a DC's pen, brush or font color
  ' here we store the most recent "sets" so we can reset when needed
  'Dim tBrush As Long
  '<:-):WARNING: Unused Dim commented out
  'Dim tPen As Long
  '<:-):WARNING: Unused Dim commented out
  '<:-):UPDATED: Multiple Dim line separated
  ' changing a DC's setting

    If bSet Then
        Select Case TypeObject
         Case cObj_Brush         ' brush is being changed
            DeleteObject SelectObject(ButtonDC.hDC, CreateSolidBrush(lColor))
            ' if the pen color will be the same
            If bSamePenColor Then
                DeleteObject SelectObject(ButtonDC.hDC, CreatePen(PenStyle, PenWidth, lColor))
            End If
         Case cObj_Pen   ' pen is being changed (mostly for drawing lines)
            DeleteObject SelectObject(ButtonDC.hDC, CreatePen(PenStyle, PenWidth, lColor))
         Case cObj_Text  ' text color is changing
            SetTextColor m_hDC, ConvertColor(lColor)
        End Select
     Else            ' resetting the DC back to the way it was
        DeleteObject SelectObject(ButtonDC.hDC, ButtonDC.OldBrush)
        DeleteObject SelectObject(ButtonDC.hDC, ButtonDC.OldPen)
    End If

End Sub

Private Function ShadeColor(lColor As Long, _
                            shadeOffset As Integer, _
                            lessBlue As Boolean, _
                            Optional bFocusRect As Boolean, _
                            Optional bInvert As Boolean) As Long

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameters 'lColor, lessBlue, bFocusRect, bInvert'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' Basically supply a value between -255 and +255. Positive numbers make
  ' the passed color lighter and negative numbers make the color darker
  
  Dim valRGB(0 To 2) As Integer
  Dim i              As Integer

    '<:-):UPDATED: Multiple Dim line separated
CalcNewColor:
    valRGB(0) = (lColor And &HFF) + shadeOffset
    valRGB(1) = ((lColor And &HFF00&) / 255&) + shadeOffset
    If lessBlue Then
        valRGB(2) = (lColor And &HFF0000) / &HFF00&
        valRGB(2) = valRGB(2) + ((valRGB(2) * CLng(shadeOffset)) \ &HC0)
     Else
        valRGB(2) = (lColor And &HFF0000) / &HFF00& + shadeOffset
    End If
    For i = 0 To 2
        If valRGB(i) > 255 Then
            valRGB(i) = 255
        End If
        If valRGB(i) < 0 Then
            valRGB(i) = 0
        End If
        If bInvert Then
            '<:-)Pleonasm Removed
            valRGB(i) = Abs(255 - valRGB(i))
        End If
    Next i
    ShadeColor = valRGB(0) + 256& * valRGB(1) + 65536 * valRGB(2)
    Erase valRGB
    If bFocusRect And (ShadeColor = vbBlack Or ShadeColor = vbWhite) Then
        '<:-)Pleonasm Removed
        shadeOffset = -shadeOffset
        If shadeOffset = 0 Then
            shadeOffset = 64
        End If
        GoTo CalcNewColor
    End If

End Function

Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_Description = "Allows or prevents a focus rectangle from being displayed. In design mode, this may always be displayed for button set as Default."

    ShowFocusRect = myProps.bShowFocus

End Property

Public Property Let ShowFocusRect(bShow As Boolean)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'bShow'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' Shows/hides the focus rectangle when button comes into focus

    myProps.bShowFocus = bShow
    If ((myProps.bStatus And 1) = 1) Then
        ' if currently has the focus, then we take it off
        If Ambient.UserMode Then
            myProps.bStatus = myProps.bStatus And Not 1
            Refresh
         Else
            ' however, we don't if it is the default button
            MsgBox "The focus rectangle may appear on default buttons ONLY while in design mode, " & vbNewLine & "but will not appear when the form is running.", vbInformation + vbOKOnly
        End If
     Else
        Refresh
    End If
    PropertyChanged "Focus"

End Property

Friend Sub TimerUpdate(lvTimerID As Long)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'lvTimerID'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' pretty good way to determine when cursor moves outside of any shape region
  ' especially useful for my diagonal/round buttons since they are not your typical
  ' rectangular shape.
  
  Dim mousePt As POINTAPI

    'Dim cRect As RECT
    '<:-):WARNING: Unused Dim commented out
    '<:-):UPDATED: Multiple Dim line separated
    GetCursorPos mousePt
    If WindowFromPoint(mousePt.X, mousePt.Y) <> UserControl.hWnd Then
        ' when exits button area, kill the timer
        KillTimer UserControl.hWnd, lvTimerID
        myProps.bStatus = myProps.bStatus And Not 4
        bTimerActive = False
        bNoRefresh = False
        RaiseEvent MouseOnButton(False)
        bKeyDown = False
        Refresh
    End If

End Sub

Private Function ToggleOptionButtons(nMode As Integer) As Boolean

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'nMode'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' Function tracks option buttons for each container they are placed on
  ' It will 1) Toggle others to false when one is set to true
  '         2) Add or remove option buttons from a collection
  '         3) Query option buttons to see if one is set to true
  
  Dim i          As Integer
  Dim NrCtrls    As Integer

    '<:-):UPDATED: Multiple Dim line separated
  Dim myObjRef   As Long
  Dim tgtObjRef  As Long
    '<:-):UPDATED: Multiple Dim line separated
  Dim optControl As lvButtons_H
  Dim bOffset    As Boolean
    NrCtrls = GetProp(CLng(Tag), "lv_OptCount")
    On Error GoTo OptionToggleError
    If myProps.bValue And (NrCtrls > 0 Or nMode = 1) Then
        ' called when an option button is set to True; set others to false
        myObjRef = ObjPtr(Me)
        For i = 1 To NrCtrls
            tgtObjRef = GetProp(CLng(Tag), "lv_Obj" & i)
            If tgtObjRef <> myObjRef Then
                CopyMemory optControl, tgtObjRef, &H4
                optControl.Value = False
                CopyMemory optControl, 0&, &H4
            End If
        Next i
    End If
    Select Case nMode
     Case 1 ' Add instance to window db
        SetProp CLng(Tag), "lv_OptCount", NrCtrls + nMode
        SetProp CLng(Tag), "lv_Obj" & NrCtrls + nMode, ObjPtr(Me)
     Case -1 ' Remove instance from window db
        myObjRef = ObjPtr(Me)
        For i = 1 To NrCtrls
            tgtObjRef = GetProp(CLng(Tag), "lv_Obj" & i)
            If tgtObjRef = myObjRef Then
                bOffset = -1
             Else
                If bOffset Then
                    SetProp CLng(Tag), "lv_Obj" & i, tgtObjRef
                End If
            End If
        Next i
        RemoveProp CLng(Tag), "lv_Obj" & i - 1
        If NrCtrls = 0 Then
            RemoveProp CLng(Tag), "lv_OptCount"
        End If
     Case 2 ' See if any option buttons have True values
        For i = 1 To NrCtrls
            tgtObjRef = GetProp(CLng(Tag), "lv_Obj" & i)
            CopyMemory optControl, tgtObjRef, &H4
            If optControl.Value Then
                '<:-)Pleonasm Removed
                i = NrCtrls + 1
                ToggleOptionButtons = True
            End If
            CopyMemory optControl, 0&, &H4
        Next i
    End Select

Exit Function

OptionToggleError:
    Debug.Print "Err in OptionToggle: " & Err.Description
    '<:-):SUGGESTION: Active Debug should be removed from final code.

End Function

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

  ' This happens when hot key is pressed or button is default/cancel and
  ' Enter/Escape key is pressed. Basically, we need to fire a click event

    If (KeyAscii = 13 Or KeyAscii = 27) And myProps.bMode > lv_CommandButton Then
        Exit Sub
    End If
    If ((myProps.bStatus And 1) <> 1) And (KeyAscii <> 13 And KeyAscii <> 27) Then
        Refresh
    End If
    ' flag that needs to be set in order to fire a click event
    mButton = vbLeftButton
    Call UserControl_Click  ' now trigger a click event

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

  ' //////////////////// USER CONTROL EVENTS  \\\\\\\\\\\\\\\\\\\\\\\\
  ' something on the parent container changed

    On Error GoTo AbortCheck
    Select Case PropertyName
     Case "DisplayAsDefault" 'changing focus
        If Ambient.DisplayAsDefault And myProps.bShowFocus Then
            '<:-)Pleonasm Removed
            myProps.bStatus = myProps.bStatus Or 1
         Else
            myProps.bStatus = myProps.bStatus And Not 1
        End If
        Refresh
     Case "BackColor"
        cParentBC = ConvertColor(Ambient.BackColor)
        If myProps.bShape > lv_FullDiagonal Or myProps.bBackStyle = 5 Then
            Refresh
        End If
    End Select
AbortCheck:

End Sub

Private Sub UserControl_Click()

  ' Again, only allow left mouse button to fire click events. Keyboard
  ' actions may set the mButton variable to ensure event is fired

    If mButton = vbLeftButton Then
        If myProps.bMode > lv_CommandButton Then
            If myProps.bValue And myProps.bMode = lv_OptionButton Then
                '<:-)Pleonasm Removed
                Exit Sub
            End If
            Me.Value = Not myProps.bValue
        End If
        RaiseEvent Click
    End If

End Sub

Private Sub UserControl_DblClick()

  ' Typical Window buttons do not have a double click event. Each
  ' double click event on a typical button is registered as 2 clicks
  ' with 2 sets of MouseDown & MouseUp events. We simulate that too
  
  Dim mousePt As POINTAPI

    ' another plus... other button routines out there may not pass the
    ' true X,Y coordinates when firing a fake 2nd click event
    GetCursorPos mousePt
    ScreenToClient UserControl.hWnd, mousePt
    RaiseEvent DoubleClick(CInt(mButton))   ' added benefit/information
    If mButton = vbLeftButton Then
        ' double clicked with left mouse button fire a mouse down event
        Call UserControl_MouseDown(vbLeftButton, 0, CSng(mousePt.X), CSng(mousePt.Y))
        ' key variable. This flag indicates we will be sending a fake click event
        mButton = -1
     Else
        ' double clicked with middle/right mouse button, send this event only
        RaiseEvent MouseDown(vbLeftButton, 0, CSng(mousePt.X), CSng(mousePt.Y))
    End If

End Sub

Private Sub UserControl_GotFocus()

  ' If no option button in the group is set to True, then the first one that
  ' gets the focus is set to True by default

    If myProps.bMode = lv_OptionButton Then
        If myProps.bValue = False Then
            '<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
            If ToggleOptionButtons(2) = False Then
                mButton = vbLeftButton
                Call UserControl_Click
            End If
        End If
    End If '<:-)Short Circuit inserted this line

End Sub

Private Sub UserControl_InitProperties()

  ' Initial properties for a new button

    UserControl.Tag = UserControl.ContainerHwnd
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    With myProps
        .bCaption = Ambient.DisplayName
        .bCaptionAlign = vbCenter
        .bShowFocus = True
        .bForeHover = vbButtonText
        .bBackHover = vbButtonFace
    End With
    cParentBC = ConvertColor(Ambient.BackColor)
    curBackColor = vbButtonFace         ' this will be the button's initial backcolor
    GetGDIMetrics "Font"
    GetGDIMetrics "BackColor"
    PropertyChanged "Caption"
    PropertyChanged "CapAlign"
    PropertyChanged "Focus"
    PropertyChanged "cFHover"
    PropertyChanged "cBHover"

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

  ' forward arrow keys as next/previous controls

    Select Case KeyCode
     Case vbKeyRight
        KeyCode = 0             ' simulate a tab key
        PostMessage CLng(Tag), WM_KEYDOWN, ByVal &H27, ByVal &H4D0001
     Case vbKeyDown
        KeyCode = 0
        PostMessage CLng(Tag), WM_KEYDOWN, ByVal &H28, ByVal &H500001
     Case vbKeyLeft
        KeyCode = 0             ' simulate a shift+tab key
        PostMessage CLng(Tag), WM_KEYDOWN, ByVal &H25, ByVal &H4B0001
     Case vbKeyUp
        KeyCode = 0
        PostMessage CLng(Tag), WM_KEYDOWN, ByVal &H26, ByVal &H480001
     Case vbKeySpace
        ' space key on a button is same as enter, but shows the button state changes
        If ((myProps.bStatus And 2) <> 2) Then
            bKeyDown = True
            ' we only want to do this once. Subsequent space keys will still fire
            ' a KeyDown event, but won't keep changing button state
            ' tell routines that mouse is over button & it is "down"
            myProps.bStatus = myProps.bStatus Or 4
            myProps.bStatus = myProps.bStatus Or 2
            Refresh
            ' reset the mouse hover status if needed
            If Not bTimerActive Then
                myProps.bStatus = myProps.bStatus And Not 4
            End If
        End If
    End Select
    If KeyCode Then
        RaiseEvent KeyDown(KeyCode, Shift)
    End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

  ' not used by me, but we'll send the event

    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

  ' Key up events.

    bKeyDown = False
    Select Case KeyCode
     Case vbKeySpace
        ' if space bar released & button state is "down", we make button "normal"
        If ((myProps.bStatus And 2) = 2) Then
            mButton = vbLeftButton
            myProps.bStatus = myProps.bStatus And Not 2
            If myProps.bMode > lv_CommandButton Then
                RaiseEvent KeyUp(KeyCode, Shift)
                KeyCode = 0
             Else
                Refresh
            End If
            Call UserControl_Click      ' simulate a click event
        End If
     Case vbKeyRight, vbKeyDown, vbKeyLeft, vbKeyUp
        KeyCode = 0
        If myProps.bMode = lv_OptionButton Then
            If myProps.bValue = False Then
                '<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
                mButton = vbLeftButton
                Call UserControl_Click
            End If
        End If '<:-)Short Circuit inserted this line
     Case vbKeyShift
        KeyCode = 0
    End Select
    If KeyCode Then
        RaiseEvent KeyUp(KeyCode, Shift)
    End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

  ' Only allow left clicks to fire a click event
  ' key variable... this tells our mouse routines & the click event
  ' whether or not the left button is doing the clicking

    mButton = Button
    If Button = vbLeftButton Then
        bKeyDown = True
        myProps.bStatus = myProps.bStatus Or 2      ' simulate a "down" state
        bNoRefresh = False
        Refresh
        ' we need this in case the user clicks & drags mouse off of the control
        ' Without it, we may never get the mouse up event
        SetCapture UserControl.hWnd
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

  ' Here we may fire 2 events: MouseMove & MouseOnButton

    On Error GoTo RaiseTheMouseEvent
    ' if we are already over the button, simply fire the MouseMove event
    If bTimerActive Then
    On Error Resume Next
        Err.Raise 5
    End If
    ' if we are outside of the mouse we fire the MouseMove event only.
    ' Note. We don't use SetCapture/ReleaseCapture, except in one special
    ' case, because it affects the actual control (not my button control)
    ' that should rightfully have the focus. However, should the mouse
    ' be down on this button & the use drags mouse off of the button,
    ' we will continue to get mouse move events & will fire them accordingly;
    ' and this is the only appropriate exception for using SetCapture
    If X < 0 Or Y < 0 Or X > UserControl.ScaleWidth Or Y > UserControl.ScaleHeight Then
        '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
        Err.Raise 5
    End If
    ' An improvement over most other button routines out there....
    ' A soft timer. No timer control needed. The trick is to get the timer
    ' to fire back to this instance of the button control. We do that by
    ' setting a reference to this instance in our Window properties. When
    ' the timer routine (see modLvTimer) gets an event, the hWnd is passed
    ' along & with that, the timer routine can retrieve the property we set.
    ' All this allows the timer routine to positively identify this instance.
    myProps.bStatus = myProps.bStatus Or 4      ' set a mouse hover state
    RaiseEvent MouseOnButton(True)              ' fire this event
    ' The MouseOnButton event allows users to change the properties of the
    ' button while the mouse is over it. For instance, you can supply a different
    ' image/font/etc & replace it when the mouse leaves the button area
    '<:-)Auto-inserted With End...With Structure
    With UserControl
        SetProp .hWnd, "lv_ClassID", ObjPtr(Me)
        ' the next line is used with expandability in mind. May use multiple timers in future upgrade
        SetProp .hWnd, "lv_TimerID", 237
        SetTimer .hWnd, 237, 50, AddressOf lv_TimerCallBack
    End With 'UserControl
    bTimerActive = True                         ' flag used for drawing
    bNoRefresh = False                         ' ensure flag is reset
    If Button = vbLeftButton Then
        bKeyDown = True
    End If
    Refresh
RaiseTheMouseEvent:
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

  ' The only tweak here is to trigger a fake click event if user
  ' double clicked on this button

    bKeyDown = False
    If Button = vbLeftButton Then
        ReleaseCapture
        myProps.bStatus = myProps.bStatus And Not 2     ' "normal" state
        bNoRefresh = False                              ' ensure flag is reset
        If myProps.bMode = lv_CommandButton Then
            Refresh
        End If
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)            ' fire event
    ' key flag. Update
    mButton = Button

End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)

  ' not used by me, but we'll send the event

    RaiseEvent OLECompleteDrag(Effect)

End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, _
                                    Effect As Long, _
                                    Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

  ' not used by me, but we'll send the event

    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, _
                                    Effect As Long, _
                                    Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single, _
                                    State As Integer)

  ' not used by me, but we'll send the event

    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)

End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, _
                                        DefaultCursors As Boolean)

  ' not used by me, but we'll send the event

    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)

End Sub

Private Sub UserControl_OLESetData(Data As DataObject, _
                                   DataFormat As Integer)

  ' not used by me, but we'll send the event

    RaiseEvent OLESetData(Data, DataFormat)

End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, _
                                     AllowedEffects As Long)

  ' not used by me, but we'll send the event

    RaiseEvent OLEStartDrag(Data, AllowedEffects)

End Sub

Private Sub UserControl_Paint()

  ' this routine typically called by Windows when another window covering
  ' this button is removed, or when the parent is moved/minimized/etc.

    bNoRefresh = False
    RedrawButton False

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  ' Write properties

    UserControl.Tag = UserControl.ContainerHwnd
    '<:-):WARNING: It is clearer to use the Me/UserControl reference than depend on VB's auto behaviour.
    cParentBC = ConvertColor(Ambient.BackColor)
    DelayDrawing True
    With PropBag
        myProps.bCaption = .ReadProperty("Caption", "")
        myProps.bCaptionAlign = .ReadProperty("CapAlign", 2)
        myProps.bBackStyle = .ReadProperty("BackStyle", 0)
        myProps.bShape = .ReadProperty("Shape", 0)
        myProps.bGradient = .ReadProperty("Gradient", 0)
        myProps.bGradientColor = .ReadProperty("cGradient", vbButtonFace)
        UserControl.ForeColor = .ReadProperty("cFore", vbButtonText)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        myProps.bShowFocus = .ReadProperty("Focus", True)
        myProps.bMode = .ReadProperty("Mode", 0)
        myProps.bValue = .ReadProperty("Value", False)
        Set myImage.Image = .ReadProperty("Image", Nothing)
        myImage.Size = .ReadProperty("ImgSize", 16)
        myImage.Align = .ReadProperty("ImgAlign", 0)
        myProps.bForeHover = .ReadProperty("cFHover", vbButtonText)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        curBackColor = .ReadProperty("cBack", Parent.BackColor)
        myProps.bBackHover = .ReadProperty("cBHover", curBackColor)
        myProps.bLockHover = .ReadProperty("LockHover", 0)
        myProps.bCaptionStyle = .ReadProperty("CapStyle", 0)
        Set Me.MouseIcon = .ReadProperty("mIcon", Nothing)
        Me.MousePointer = .ReadProperty("mPointer", 0)
    End With
    GetGDIMetrics "Picture"
    GetGDIMetrics "Font"
    GetGDIMetrics "BackColor"
    Me.Caption = myProps.bCaption      ' sets the hot key if needed
    If myProps.bMode = lv_OptionButton Then
        ToggleOptionButtons (1)
    End If
    bNoRefresh = False
    Call UserControl_Resize

End Sub

Private Sub UserControl_Resize()

  ' since we are using a separate DC for drawing, we need to resize the
  ' bitmap in that DC each time the control resizes
  '<:-)Auto-inserted With End...With Structure

    With ButtonDC
        If .hDC Then
            DeleteObject SelectObject(.hDC, .OldBitmap)
            .OldBitmap = 0  ' this will force a new bitmap for existing DC
        End If
    End With 'ButtonDC
    GetSetOffDC True
    If bNoRefresh Then
        Exit Sub
    End If
    CreateButtonRegion
    CalculateBoundingRects
    Refresh

End Sub

Private Sub UserControl_Show()

  ' interesting, NT won't send the DisplayAsDefault (while in IDE) until after the button is shown
  ' Win98 fires this regardless. So fix is to put the test here also.

    If Ambient.UserMode = False Then
        If Ambient.DisplayAsDefault And myProps.bShowFocus Then
            '<:-)Pleonasm Removed
            myProps.bStatus = myProps.bStatus Or 1
            Refresh
        End If
    End If

End Sub

Private Sub UserControl_Terminate()

  ' Button is ending, let's clean up
  ' should never happen that we have a timer left over; but just in case

    If bTimerActive Then
        KillTimer UserControl.hWnd, 1
    End If
    ' circular buttons have a clipping region, kill that too
    If ButtonDC.ClipRgn Then
        DeleteObject ButtonDC.ClipRgn
    End If
    '<:-)Auto-inserted With End...With Structure
    With ButtonDC
        If .hDC Then
            ' get rid of left over pen & brush
            SetButtonColors False, .hDC, cObj_Pen, 0
            ' get rid of logical font
            DeleteObject SelectObject(.hDC, .OldFont)
            ' destroy the separate Bitmap & select original back into DC
            DeleteObject SelectObject(.hDC, .OldBitmap)
            ' destroy the temporary DC
            DeleteDC .hDC
        End If
    End With 'ButtonDC

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  ' Store Properties

    With PropBag
        .WriteProperty "Caption", myProps.bCaption, ""
        .WriteProperty "CapAlign", myProps.bCaptionAlign, 0
        .WriteProperty "BackStyle", myProps.bBackStyle, 0
        .WriteProperty "Shape", myProps.bShape, 0
        .WriteProperty "Font", UserControl.Font, Nothing
        .WriteProperty "cFore", UserControl.ForeColor, vbButtonText
        .WriteProperty "cFHover", myProps.bForeHover, vbButtonText
        .WriteProperty "cBhover", myProps.bBackHover, curBackColor
        .WriteProperty "Focus", myProps.bShowFocus, True
        .WriteProperty "LockHover", myProps.bLockHover, 0
        .WriteProperty "cGradient", myProps.bGradientColor, vbButtonFace
        .WriteProperty "Gradient", myProps.bGradient, 0
        .WriteProperty "CapStyle", myProps.bCaptionStyle, 0
        .WriteProperty "Mode", myProps.bMode
        .WriteProperty "Value", myProps.bValue
        .WriteProperty "ImgAlign", myImage.Align, 0
        .WriteProperty "Image", myImage.Image, Nothing
        .WriteProperty "ImgSize", myImage.Size, 16
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "cBack", curBackColor
        .WriteProperty "mPointer", UserControl.MousePointer, 0
        .WriteProperty "mIcon", UserControl.MouseIcon, Nothing
    End With

End Sub

Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Applicable to only check box or option button modes: True or False"
Attribute Value.VB_UserMemId = 0

    Value = myProps.bValue

End Property

Public Property Let Value(bValue As Boolean)

  '<:-):SUGGESTION:  Insert 'ByVal' for Parameter  'bValue'
  '<:-)WARNING NEW FIX : This is still experimental (Testing is very conservative).
  '<:-) List may be incomplete or contain members in error, test carefully.
  '<:-) User created Events can use ByVal but you must edit the Declaration as well.
  '<:-) otherwise you will get the Compile error message 'Procedure declaration does not match description of event or procedure having the same name'
  '<:-) The Rule is: If the routine doesn't change the variable (This is what Code Fixer looks for)
  '<:-) OR you don't want any changes returned (You have to hand code this) make the parameter ByVal.
  '<:-) Find this message in the code (Sub ByValParameter in ParameterMod) and there is an alternate version to use the Verbose Message version.
  ' For option button & check box modes

    If myProps.bMode = lv_CommandButton And bValue Then
        '<:-)Pleonasm Removed
        ' TRUE values for command buttons not allowed
        If Not UserControl.Ambient.UserMode Then
            MsgBox "This property is not applicable for command button modes.", vbInformation + vbOKOnly
        End If
        Exit Property
    End If
    myProps.bValue = bValue
    ' if optionbutton now true, need to toggle the other options buttons off
    If bValue And myProps.bMode = lv_OptionButton Then
        Call ToggleOptionButtons(0)
    End If
    Refresh
    PropertyChanged "Value"

End Property

':) Roja's VB Code Fixer V1.0.97 (6/20/03 2:46:01 AM) 343 + 3526 = 3869 Lines Thanks Ulli for inspiration and lots of code.


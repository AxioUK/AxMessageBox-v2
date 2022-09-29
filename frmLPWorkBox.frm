VERSION 5.00
Begin VB.Form frmLPWBox 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "frmMessage"
   ScaleHeight     =   2055
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin AxMsgBox2.axProgBar axProg 
      Height          =   330
      Left            =   1575
      TabIndex        =   6
      Top             =   1305
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   255
      ShowText        =   -1  'True
      Text            =   ""
   End
   Begin AxMsgBox2.ucProgressCircular ucProg 
      Height          =   1065
      Left            =   150
      TabIndex        =   5
      Top             =   795
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1879
      Caption1_ForeColor=   11565097
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StepSpaceSize   =   3
      PF_Steps        =   20
      PB_Color1       =   11565097
      PB_Color1Opacity=   10
      PB_Color2       =   16777215
      PB_ColorGradient=   -1  'True
      PB_Steps        =   20
      Value           =   0
      DisplayInPercent=   0   'False
      PF_ForeColor    =   11565097
      AnimationInterval=   100
   End
   Begin AxMsgBox2.axLabelPlus picX 
      Height          =   330
      Left            =   5940
      TabIndex        =   3
      Top             =   165
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   582
      BackColorPress  =   255
      Border          =   -1  'True
      BorderColor     =   14737632
      ColorOnMouseOver=   12632319
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption1        =   "frmLPWorkBox.frx":0000
      Caption2        =   "frmLPWorkBox.frx":0020
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
      ChangeColorOnClick=   -1  'True
      ChangeOnMouseOver=   6
      GradientColorP1 =   0
      GradientColorP1Opacity=   0
      GradientColorP2 =   0
      GradientColorP2Opacity=   0
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61153
      IconForeColor   =   0
      IconPaddingY    =   1
      IconAlignmentH  =   1
      GlowSpeed       =   0
      GlowColor       =   0
      GlowTiks        =   0
      PictureArr      =   0
   End
   Begin AxMsgBox2.axLabelPlus axButton2 
      Height          =   405
      Left            =   4905
      TabIndex        =   2
      Top             =   885
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   714
      Border          =   -1  'True
      BorderColor     =   14737632
      ColorOnMouseOver=   4210752
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption1        =   "frmLPWorkBox.frx":0040
      Caption2        =   "frmLPWorkBox.frx":0070
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
      ChangeColorOnClick=   -1  'True
      GradientColorP1 =   0
      GradientColorP1Opacity=   0
      GradientColorP2 =   0
      GradientColorP2Opacity=   0
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
      IconForeColor   =   0
      GlowSpeed       =   0
      GlowColor       =   0
      GlowTiks        =   0
      PictureArr      =   0
   End
   Begin AxMsgBox2.axLabelPlus axButton1 
      Height          =   405
      Left            =   4905
      TabIndex        =   1
      Top             =   1440
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   714
      Border          =   -1  'True
      BorderColor     =   14737632
      ColorOnMouseOver=   4210752
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption1        =   "frmLPWorkBox.frx":0090
      Caption2        =   "frmLPWorkBox.frx":00BE
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
      ChangeColorOnClick=   -1  'True
      GradientColorP1 =   0
      GradientColorP1Opacity=   0
      GradientColorP2 =   0
      GradientColorP2Opacity=   0
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
      IconForeColor   =   0
      GlowSpeed       =   0
      GlowColor       =   0
      GlowTiks        =   0
      PictureArr      =   0
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje........................................................................."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1485
      TabIndex        =   0
      Top             =   750
      Width           =   4635
   End
   Begin AxMsgBox2.axLabelPlus lblTitle 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1085
      Border          =   -1  'True
      BorderColor     =   14737632
      BorderCornerLeftTop=   9
      BorderCornerRightTop=   9
      BorderWidth     =   1
      CaptionAlignmentV=   1
      Caption1        =   "frmLPWorkBox.frx":00DE
      Caption2        =   "frmLPWorkBox.frx":010A
      CaptionAngle    =   270
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
      Caption1WordWrap=   -1  'True
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   61098
      IconForeColor   =   49152
      IconPaddingX    =   12
      IconAlignmentV  =   1
      GlowSpeed       =   0
      GlowColor       =   0
      GlowTiks        =   0
      PictureArr      =   0
   End
End
Attribute VB_Name = "frmLPWBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function CreateStyle(tPos As TStyle) As Boolean
Dim MsgAncho As Integer, TtlAncho As Integer
Dim FrmAncho As Integer, FrmAlto As Integer
Dim minH As Integer, MsgH As Integer

Select Case tPos
  Case Is = TitleSide
      
    lblMsg.Move 1500, 500
    MsgAncho = lblMsg.Left + lblMsg.Width '+ 450
    TtlAncho = lblTitle.Left + lblTitle.Width '+ 450
    FrmAncho = 6850
    
    If MsgAncho > TtlAncho And MsgAncho > FrmAncho Then
      Me.Width = lblMsg.Left + lblMsg.Width + 1650
    ElseIf TtlAncho > MsgAncho And TtlAncho > FrmAncho Then
      Me.Width = lblTitle.Left + lblTitle.Width + 1000
    Else
      Me.Width = FrmAncho
    End If
    
    minH = 2050 'IIf(tPos = TitleSide, picX.Top + picX.Height + (axButton1.Height * 2) + 500, picX.Top + picX.Height + (axButton1.Height * 2) + 600)
    MsgH = lblMsg.Top + lblMsg.Height + 200
    
    Me.Height = IIf(MsgH < minH, minH + 300, MsgH + 500)
    axButton1.Move Me.ScaleWidth - 1550, Me.ScaleHeight - 600
    axButton2.Move Me.ScaleWidth - 1550, Me.ScaleHeight - 1110
      
    ucProg.Visible = False
    axProg.Visible = True
    axProg.Move lblMsg.Left + 50, axButton1.Top, (Me.ScaleWidth - (lblMsg.Left + axButton1.Width + 500))

    With lblTitle
      .Move 0, 0, 1200, Me.ScaleHeight
      .BorderCornerBottomLeft = 6
      .BorderCornerBottomRight = 0
      .BorderCornerLeftTop = 6
      .BorderCornerRightTop = 0
      .CaptionAngle = 270
      .Caption1PaddingX = 5 'IIf(bPos = ButtonBottom, 20, 0)
      .IconAlignmentH = cCenter
      .IconAlignmentV = cTop
      .IconPaddingX = 0
      .IconPaddingY = 15
    End With
    
    picX.Move (Me.ScaleWidth - (picX.Width + (picX.Width / 2))), picX.Height / 2

  Case Is = TitleTop
      
    lblMsg.Move 1600, 850
    MsgAncho = lblMsg.Left + lblMsg.Width + 300
    FrmAncho = 5000
          
    If MsgAncho > FrmAncho Then
      Me.Width = lblMsg.Left + lblMsg.Width + 500
    Else
      Me.Width = FrmAncho
    End If
    
    minH = 2050
    MsgH = lblMsg.Top + lblMsg.Height + 650
    Me.Height = IIf(MsgH < minH, minH + 300, MsgH)
    axButton1.Move Me.ScaleWidth - 1450, Me.ScaleHeight - 540
    axButton2.Move Me.ScaleWidth - 2940, Me.ScaleHeight - 540
    
    ucProg.Visible = True
    axProg.Visible = False
    ucProg.Move 150, 800
    
    With lblTitle
      .Move 0, 0, Me.ScaleWidth, 650
      .BorderCornerBottomLeft = 0
      .BorderCornerBottomRight = 0
      .BorderCornerLeftTop = 6
      .BorderCornerRightTop = 6
      .CaptionAngle = 0
      .Caption1PaddingX = 50
      .IconAlignmentH = cLeft
      .IconAlignmentV = cMiddle
      .IconPaddingX = 12
      .IconPaddingY = 0
    End With
    
    picX.Move Me.ScaleWidth - ((lblTitle.Height / 2) + (picX.Width / 2)), ((lblTitle.Height / 2) - (picX.Width / 2))
End Select

'-----------------------
'Call CreateTitle(Me, 9, picX.BackColor)
Call RoundCorner(Me, 8)

End Function

Private Sub axButton1_Click()
asClicked = 1
Unload Me
End Sub

Private Sub axButton2_Click()
asClicked = 2
Unload Me
End Sub

Private Sub Form_Load()
lblMsg.BorderStyle = 0
Me.BorderStyle = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormBar Me.hwnd
End Sub

Private Sub Form_Resize()
'Nothing
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormBar Me.hwnd
End Sub

Private Sub picX_Click()
asClicked = 2
Unload Me
End Sub


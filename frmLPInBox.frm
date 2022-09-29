VERSION 5.00
Begin VB.Form frmLPIBox 
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
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1500
      TabIndex        =   5
      Top             =   1365
      Width           =   3015
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
      Caption1        =   "frmLPInBox.frx":0000
      Caption2        =   "frmLPInBox.frx":0020
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
      Caption1        =   "frmLPInBox.frx":0040
      Caption2        =   "frmLPInBox.frx":0070
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
      Caption1        =   "frmLPInBox.frx":0090
      Caption2        =   "frmLPInBox.frx":00BE
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
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1500
      TabIndex        =   0
      Top             =   495
      Width           =   2835
   End
   Begin AxMsgBox2.axLabelPlus lblTitle 
      Height          =   2010
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   3545
      Border          =   -1  'True
      BorderColor     =   14737632
      BorderCornerLeftTop=   9
      BorderCornerRightTop=   9
      BorderWidth     =   1
      CaptionAlignmentV=   1
      Caption1        =   "frmLPInBox.frx":00DE
      Caption2        =   "frmLPInBox.frx":010A
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
      IconPaddingY    =   15
      IconAlignmentH  =   1
      GlowSpeed       =   0
      GlowColor       =   0
      GlowTiks        =   0
      PictureArr      =   0
   End
End
Attribute VB_Name = "frmLPIBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function CreateStyle(tPos As TStyle, bPos As BStyle) As Boolean
Dim MsgAncho As Integer, TtlAncho As Integer
Dim FrmAncho As Integer, FrmAlto As Integer
Dim minH As Integer, MsgH As Integer

Select Case tPos
  Case Is = TitleSide
    If bPos = ButtonBottom Then
      lblMsg.Move 1500, 300
      MsgAncho = lblMsg.Left + lblMsg.Width '+ 450
      TtlAncho = lblTitle.Left + lblTitle.Width '+ 450
      FrmAncho = 6850
      
      If MsgAncho > TtlAncho And MsgAncho > FrmAncho Then
        Me.Width = lblMsg.Left + lblMsg.Width + 650
      ElseIf TtlAncho > MsgAncho And TtlAncho > FrmAncho Then
        Me.Width = lblTitle.Left + lblTitle.Width + 500
      Else
        Me.Width = FrmAncho
      End If
      
      Me.Height = lblMsg.Top + lblMsg.Height + 850 + Text1.Height
      axButton1.Move Me.ScaleWidth - 1450, Me.ScaleHeight - 540
      axButton2.Move Me.ScaleWidth - 2940, Me.ScaleHeight - 540
      

    ElseIf bPos = ButtonRight Then
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
      
      minH = IIf(tPos = TitleSide, picX.Top + picX.Height + (axButton1.Height * 2) + 500, picX.Top + picX.Height + (axButton1.Height * 2) + 600)
      MsgH = lblMsg.Top + lblMsg.Height + 200
      
      Me.Height = IIf(MsgH < minH, minH + 300, MsgH + 500)
      axButton1.Move Me.ScaleWidth - 1550, Me.ScaleHeight - 600
      axButton2.Move Me.ScaleWidth - 1550, Me.ScaleHeight - 1110
      
    End If

  Case Is = TitleTop
    If bPos = ButtonBottom Then
      lblMsg.Move 650, 850
      MsgAncho = lblMsg.Left + lblMsg.Width + 300
      FrmAncho = 5000
            
      If MsgAncho > FrmAncho Then
        Me.Width = lblMsg.Left + lblMsg.Width + 500
      Else
        Me.Width = FrmAncho
      End If
      
      Me.Height = lblMsg.Top + lblMsg.Height + 850 + Text1.Height
      axButton1.Move Me.ScaleWidth - 1450, Me.ScaleHeight - 540
      axButton2.Move Me.ScaleWidth - 2940, Me.ScaleHeight - 540
    
    ElseIf bPos = ButtonRight Then
      lblMsg.Move 650, 850
      MsgAncho = lblMsg.Left + lblMsg.Width + 300
      FrmAncho = 5000
            
      If MsgAncho > FrmAncho Then
        Me.Width = MsgAncho + 1500
      Else
        Me.Width = FrmAncho
      End If
      
      minH = IIf(tPos = TitleSide, picX.Top + picX.Height + (axButton1.Height * 2) + 500, picX.Top + picX.Height + (axButton1.Height * 2) + 600)
      MsgH = lblMsg.Top + lblMsg.Height + 200
      
      Me.Height = IIf(MsgH < minH, minH + 300, MsgH + 500)
      axButton1.Move Me.ScaleWidth - 1550, Me.ScaleHeight - 600
      axButton2.Move Me.ScaleWidth - 1550, Me.ScaleHeight - 1110
    
    End If
    
End Select

'Borders and Details
Select Case tPos
  Case Is = TitleSide
      With lblTitle
        .Move 0, 0, 1200, Me.ScaleHeight
        .BorderCornerBottomLeft = 6
        .BorderCornerBottomRight = 0
        .BorderCornerLeftTop = 6
        .BorderCornerRightTop = 0
        .CaptionAngle = 270
        .CaptionAlignmentH = cLeft
        .CaptionAlignmentV = cMiddle
        .Caption1PaddingX = 0 'IIf(bPos = ButtonBottom, 20, 0)
        .IconAlignmentH = cCenter
        .IconAlignmentV = cTop
        .IconPaddingX = 0
        .IconPaddingY = 15
      End With
      
      picX.Move (Me.ScaleWidth - (picX.Width + (picX.Width / 2))), picX.Height / 2
  
  Case Is = TitleTop
      With lblTitle
        .Move 0, 0, Me.ScaleWidth, 650
        .BorderCornerBottomLeft = 0
        .BorderCornerBottomRight = 0
        .BorderCornerLeftTop = 6
        .BorderCornerRightTop = 6
        .CaptionAngle = 0
        .CaptionAlignmentH = cLeft
        .CaptionAlignmentV = cMiddle
        .Caption1PaddingX = 65
        .IconAlignmentH = cLeft
        .IconAlignmentV = cMiddle
        .IconPaddingX = 12
        .IconPaddingY = 0
      End With
      
      picX.Move Me.ScaleWidth - ((lblTitle.Height / 2) + (picX.Width / 2)), ((lblTitle.Height / 2) - (picX.Width / 2))
End Select

Text1.Move lblMsg.Left + 50, IIf(bPos = ButtonRight, axButton1.Top, axButton1.Top - (axButton1.Height + 100)), IIf(bPos = ButtonRight, (Me.ScaleWidth - (lblMsg.Left + axButton1.Width + 500)), (Me.ScaleWidth - (lblMsg.Left + 500)))
'-----------------------
'Call CreateTitle(Me, 9, picX.BackColor)
Call RoundCorner(Me, 6)

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


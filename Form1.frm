VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test AxMessageBox DLL"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test Working Msg"
      Height          =   765
      Index           =   1
      Left            =   1275
      TabIndex        =   34
      Top             =   3915
      Width           =   840
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   180
      Left            =   0
      TabIndex        =   33
      Top             =   4710
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   4350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Working Msg"
      Height          =   765
      Index           =   0
      Left            =   2160
      TabIndex        =   32
      Top             =   3915
      Width           =   840
   End
   Begin VB.PictureBox ColorActual 
      Height          =   210
      Index           =   4
      Left            =   2790
      ScaleHeight     =   150
      ScaleWidth      =   180
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3495
      Width           =   240
   End
   Begin VB.PictureBox ColorActual 
      Height          =   210
      Index           =   3
      Left            =   2790
      ScaleHeight     =   150
      ScaleWidth      =   180
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3234
      Width           =   240
   End
   Begin VB.PictureBox ColorActual 
      Height          =   210
      Index           =   2
      Left            =   2790
      ScaleHeight     =   150
      ScaleWidth      =   180
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2976
      Width           =   240
   End
   Begin VB.PictureBox ColorActual 
      Height          =   210
      Index           =   1
      Left            =   2790
      ScaleHeight     =   150
      ScaleWidth      =   180
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2718
      Width           =   240
   End
   Begin VB.PictureBox ColorActual 
      Height          =   210
      Index           =   0
      Left            =   2790
      ScaleHeight     =   150
      ScaleWidth      =   180
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2460
      Width           =   240
   End
   Begin VB.PictureBox ColorNuevo 
      Height          =   210
      Index           =   4
      Left            =   2535
      ScaleHeight     =   150
      ScaleWidth      =   180
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3495
      Width           =   240
   End
   Begin VB.PictureBox ColorNuevo 
      Height          =   210
      Index           =   3
      Left            =   2535
      ScaleHeight     =   150
      ScaleWidth      =   180
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3234
      Width           =   240
   End
   Begin VB.PictureBox ColorNuevo 
      Height          =   210
      Index           =   2
      Left            =   2535
      ScaleHeight     =   150
      ScaleWidth      =   180
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2976
      Width           =   240
   End
   Begin VB.PictureBox ColorNuevo 
      Height          =   210
      Index           =   1
      Left            =   2535
      ScaleHeight     =   150
      ScaleWidth      =   180
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2718
      Width           =   240
   End
   Begin VB.PictureBox ColorNuevo 
      Height          =   210
      Index           =   0
      Left            =   2535
      ScaleHeight     =   150
      ScaleWidth      =   180
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2460
      Width           =   240
   End
   Begin VB.OptionButton OptionColor 
      Caption         =   "Color Form"
      Height          =   210
      Index           =   0
      Left            =   1215
      TabIndex        =   20
      Top             =   2460
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.OptionButton OptionColor 
      Caption         =   "Color Text"
      Height          =   210
      Index           =   1
      Left            =   1215
      TabIndex        =   19
      Top             =   2718
      Width           =   1305
   End
   Begin VB.OptionButton OptionColor 
      Caption         =   "Color Buttons"
      Height          =   210
      Index           =   2
      Left            =   1215
      TabIndex        =   18
      Top             =   2976
      Width           =   1305
   End
   Begin VB.OptionButton OptionColor 
      Caption         =   "Color Title"
      Height          =   210
      Index           =   3
      Left            =   1215
      TabIndex        =   17
      Top             =   3234
      Width           =   1305
   End
   Begin VB.OptionButton OptionColor 
      Caption         =   "Color Border"
      Height          =   210
      Index           =   4
      Left            =   1215
      TabIndex        =   16
      Top             =   3495
      Width           =   1305
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1830
      Left            =   135
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   1800
      ScaleWidth      =   960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2445
      Width           =   990
   End
   Begin VB.Frame Frame2 
      Caption         =   "AxInputBox"
      Height          =   2430
      Left            =   3090
      TabIndex        =   5
      Top             =   2280
      Width           =   4065
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   150
         TabIndex        =   22
         Top             =   465
         Width           =   3600
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   2880
         TabIndex        =   14
         Top             =   1140
         Width           =   675
      End
      Begin VB.CommandButton Command13 
         Caption         =   "InputBox"
         Height          =   780
         Left            =   2880
         TabIndex        =   7
         Top             =   1530
         Width           =   1005
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         Width           =   2670
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Texto Retornado"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estilo                                                    Icono"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   930
         Width           =   3120
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AxMsgBox"
      Height          =   2235
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7125
      Begin VB.CommandButton CommandLP 
         Caption         =   "Tittle Left / Buttons Right"
         Height          =   375
         Index           =   3
         Left            =   4395
         TabIndex        =   13
         Top             =   1650
         Width           =   2460
      End
      Begin VB.CommandButton CommandLP 
         Caption         =   "Tittle Left / Buttons Bottom"
         Height          =   375
         Index           =   2
         Left            =   4395
         TabIndex        =   12
         Top             =   1245
         Width           =   2460
      End
      Begin VB.CommandButton CommandLP 
         Caption         =   "Tittle Top / Buttons Right"
         Height          =   375
         Index           =   1
         Left            =   4395
         TabIndex        =   11
         Top             =   825
         Width           =   2460
      End
      Begin VB.CommandButton CommandLP 
         Caption         =   "Tittle Top / Buttons Bottom"
         Height          =   375
         Index           =   0
         Left            =   4410
         TabIndex        =   10
         Top             =   420
         Width           =   2460
      End
      Begin VB.TextBox txtTitle 
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   450
         Width           =   4065
      End
      Begin VB.TextBox txtMsg 
         Height          =   990
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1050
         Width           =   4080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   840
         Width           =   600
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xBox As New AxMsgBox

Dim sCadena As String, sTitle As String
Dim C As Integer
Dim I As Integer
Dim T As Integer

Dim ColorForm As OLE_COLOR
Dim ColorText As OLE_COLOR
Dim ColorTitle As OLE_COLOR
Dim ColorBorder As OLE_COLOR
Dim ColorButtons As OLE_COLOR


Private Sub Command1_Click(index As Integer)

Timer1.Enabled = True

Select Case index
  Case Is = 0
    If xBox.AxWrkBox("test", sTitle, "ef01", TitleTop, ColorForm, ColorTitle, ColorButtons, ColorText, ColorBorder, , , bCustom, "Seguir", "Detener") = bButton2 Then
      Timer1.Enabled = False
    End If

  Case Is = 1
    If xBox.AxWrkBox("test", sTitle, "ef01", TitleSide, ColorForm, ColorTitle, ColorButtons, ColorText, ColorBorder, , , bCustom, "Seguir", "Detener") = bButton2 Then
      Timer1.Enabled = False
    End If
End Select
End Sub

Private Sub Command13_Click()
Dim TS As Long, BS As Long
Select Case List1.ListIndex
  Case Is = 0: TS = 0: BS = 0
  Case Is = 1: TS = 0: BS = 1
  Case Is = 2: TS = 1: BS = 0
  Case Is = 3: TS = 1: BS = 1
End Select
Text2.Text = xBox.AxInputBox(sCadena, sTitle, "EEAA", TS, BS, ColorForm, ColorTitle, ColorButtons, ColorText, ColorBorder, bAcceptCancel)
End Sub

Private Sub CommandLP_Click(index As Integer)
Select Case index

  Case Is = 0
    C = xBox.AxMsgBox(sCadena, sTitle, "ef01", TitleTop, ButtonBottom, ColorForm, ColorTitle, ColorButtons, ColorText, ColorBorder)

  Case Is = 1
    C = xBox.AxMsgBox(sCadena, sTitle, "ef01", TitleTop, ButtonRight, ColorForm, ColorTitle, ColorButtons, ColorText, ColorBorder)
  
  Case Is = 2
    C = xBox.AxMsgBox(sCadena, sTitle, "ef01", TitleSide, ButtonBottom, ColorForm, ColorTitle, ColorButtons, ColorText, ColorBorder)
    
  Case Is = 3
    C = xBox.AxMsgBox(sCadena, sTitle, "ef01", TitleSide, ButtonRight, ColorForm, ColorTitle, ColorButtons, ColorText, ColorBorder)
  
End Select

End Sub

Private Sub Form_Load()
sTitle = "Etiam iaculis dui non commodo elementum."
sCadena = "Lorem ipsum dolor sit amet, consectetur adipiscing elit." & vbCrLf & _
          "Fusce vel scelerisque mauris. Donec eu odio et sem vulputate egestas." & vbCrLf & _
          "Vivamus tincidunt elit ipsum, vitae ultricies turpis semper sit amet." & vbCrLf & _
          "Donec congue velit non neque sodales, sit amet sagittis felis pellentesque." & vbCrLf & _
          "Ut nec mi felis. Duis a dignissim dolor, non venenatis erat." & vbCrLf & _
          "Curabitur lobortis, odio ac pellentesque egestas, dui nisi fringilla nibh, sed tincidunt ligula felis vitae elit." & vbCrLf & _
          "Sed tristique a mi non congue."
          
txtTitle.Text = sTitle
txtMsg.Text = sCadena
T = 0

ColorForm = &HFFFFFF
ColorText = &H80000012
ColorButtons = &HFFFFFF
ColorTitle = &HFFFFFF
ColorBorder = &HE0E0E0

ColorActual(0).BackColor = ColorForm
ColorActual(1).BackColor = ColorText
ColorActual(2).BackColor = ColorButtons
ColorActual(3).BackColor = ColorTitle
ColorActual(4).BackColor = ColorBorder

With List1
  For C = 0 To 3
    .AddItem CommandLP(C).Caption, C
  Next C
End With

List1.ListIndex = 0
I = 0
End Sub

Private Sub OptionColor_Click(index As Integer)
Dim Color As OLE_COLOR
I = index
Select Case index
  Case Is = 0: Color = ColorForm
  Case Is = 1: Color = ColorText
  Case Is = 2: Color = ColorButtons
  Case Is = 3: Color = ColorTitle
  Case Is = 4: Color = ColorBorder
End Select
End Sub

Private Sub picColor_Click()
If OptionColor(0).Value = True Then
   'Form
   ColorActual(0).BackColor = ColorNuevo(0).BackColor
    ColorForm = ColorNuevo(0).BackColor
    
ElseIf OptionColor(1).Value = True Then
    'Text
    ColorActual(1).BackColor = ColorNuevo(1).BackColor
    ColorText = ColorNuevo(1).BackColor
    
ElseIf OptionColor(2).Value = True Then
    'Button
    ColorActual(2).BackColor = ColorNuevo(2).BackColor
    ColorButtons = ColorNuevo(2).BackColor
    
ElseIf OptionColor(3).Value = True Then
    'Title
    ColorActual(3).BackColor = ColorNuevo(3).BackColor
    ColorTitle = ColorNuevo(3).BackColor
    
ElseIf OptionColor(4).Value = True Then
    'Border
    ColorActual(4).BackColor = ColorNuevo(4).BackColor
    ColorBorder = ColorNuevo(4).BackColor

End If

End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Color As Long
Color = picColor.Point(X, Y)
ColorNuevo(I).BackColor = Color
End Sub

Private Sub Timer1_Timer()
T = IIf(T = 100, T = 0, T + 1)
xBox.wMessage T, IIf(T > 50, "faltan..." & 100 - T, "Aumentando en ..." & T)
pBar.Value = T
End Sub

Private Sub txtMsg_Change()
sCadena = txtMsg.Text
End Sub

Private Sub txtTitle_Change()
sTitle = txtTitle.Text
End Sub

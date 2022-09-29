VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AxMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum BStyle
  [ButtonBottom] = 0
  [ButtonRight] = 1
End Enum

Public Enum TStyle
  [TitleTop] = 0
  [TitleSide] = 1
End Enum

Public Enum StyleButtons
    bYesNo
    bAcceptCancel
    bOKOnly
    bAcceptOnly
    bCustom
End Enum

Public Enum ClickedButton
    bYes
    bNo
    bAccept
    bCancel
    bOK
    bButton1
    bButton2
End Enum

Public Function AxMsgBox(strMessage As String, strTitle As String, hIcon As String, TitlePos As TStyle, _
                         ButtonPos As BStyle, Optional ColorForm As OLE_COLOR = &HFFFFFF, _
                         Optional ColorTitle As OLE_COLOR = &HFFFFFF, Optional ColorButtons As OLE_COLOR = &HFFFFFF, _
                         Optional ColorText As OLE_COLOR = &H80000012, Optional ColorBorder As OLE_COLOR = &HE0E0E0, _
                         Optional iButtons As StyleButtons = bAcceptCancel, Optional strButton1 As String, _
                         Optional strButton2 As String) As ClickedButton

With frmLPMBox
  
    Select Case iButtons
      Case bYesNo
        .axButton1.Caption1 = "Si"
        .axButton2.Caption1 = "No"
      Case bAcceptCancel
        .axButton1.Caption1 = "Aceptar"
        .axButton2.Caption1 = "Cancelar"
      Case bOKOnly
        .axButton1.Caption1 = "OK"
        .axButton2.Visible = False
      Case bAcceptOnly
        .axButton1.Caption1 = "Aceptar"
        .axButton2.Visible = False
      Case bCustom
        .axButton1.Caption1 = strButton1
        .axButton2.Caption1 = strButton2
    End Select
     
  ' Seteo Estilo ...
    .lblTitle.BackColor = ColorTitle
    .lblTitle.BorderColor = ColorBorder
    .lblTitle.Caption1 = strTitle
    .lblTitle.Caption1Forecolor = ColorText
    .lblTitle.IconCharCode = hIcon
    .lblMsg.Caption = strMessage
    .lblMsg.ForeColor = ColorText
    .BackColor = ColorForm
    .axButton1.BackColor = ColorButtons
    .axButton2.BackColor = ColorButtons
    .axButton1.BorderColor = ColorBorder
    .axButton2.BorderColor = ColorBorder
    .axButton1.Caption1Forecolor = ColorText
    .axButton2.Caption1Forecolor = ColorText
    .picX.BackColor = ColorButtons
    .picX.BorderColor = ColorBorder
    .picX.IconForeColor = ColorText
        
    .CreateStyle TitlePos, ButtonPos
    .Show vbModal
End With

    ' Retorna Botón Presionado
Select Case iButtons
  Case Is = bAcceptCancel
      If asClicked = 1 Then AxMsgBox = bAccept
      If asClicked = 2 Then AxMsgBox = bCancel
  Case Is = bYesNo
      If asClicked = 1 Then AxMsgBox = bYes
      If asClicked = 2 Then AxMsgBox = bNo
  Case Is = bCustom
      If asClicked = 1 Then AxMsgBox = bButton1
      If asClicked = 2 Then AxMsgBox = bButton2
  Case Is = bAcceptOnly
      If asClicked = 1 Then AxMsgBox = bAccept
  Case Is = bOKOnly
      If asClicked = 1 Then AxMsgBox = bOK
End Select

End Function

Public Function AxInputBox(strQuestion As String, strTitle As String, hIcon As String, TitlePos As TStyle, _
                           ButtonPos As BStyle, Optional ColorForm As OLE_COLOR = &HFFFFFF, _
                           Optional ColorTitle As OLE_COLOR = &HFFFFFF, Optional ColorButtons As OLE_COLOR = &HFFFFFF, _
                           Optional ColorText As OLE_COLOR = &H80000012, Optional ColorBorder As OLE_COLOR = &HE0E0E0, _
                           Optional iButtons As StyleButtons = bAcceptCancel, Optional strButton1 As String, _
                           Optional strButton2 As String) As String

With frmLPIBox
  
    Select Case iButtons
      Case bYesNo
        .axButton1.Caption1 = "Si"
        .axButton2.Caption1 = "No"
      Case bAcceptCancel
        .axButton1.Caption1 = "Aceptar"
        .axButton2.Caption1 = "Cancelar"
      Case bOKOnly
        .axButton1.Caption1 = "OK"
        .axButton2.Visible = False
      Case bAcceptOnly
        .axButton1.Caption1 = "Aceptar"
        .axButton2.Visible = False
      Case bCustom
        .axButton1.Caption1 = strButton1
        .axButton2.Caption1 = strButton2
    End Select
     
  ' Seteo Estilo ...
    .lblMsg.Caption = strQuestion
    .lblMsg.ForeColor = ColorText
    .lblTitle.BackColor = ColorTitle
    .lblTitle.BorderColor = ColorBorder
    .lblTitle.Caption1 = strTitle
    .lblTitle.Caption1Forecolor = ColorText
    .lblTitle.IconCharCode = hIcon
    .BackColor = ColorForm
    .axButton1.BackColor = ColorButtons
    .axButton2.BackColor = ColorButtons
    .axButton1.BorderColor = ColorBorder
    .axButton2.BorderColor = ColorBorder
    .axButton1.Caption1Forecolor = ColorText
    .axButton2.Caption1Forecolor = ColorText
    .picX.BackColor = ColorButtons
    .picX.BorderColor = ColorBorder
    .picX.IconForeColor = ColorText
    
    ' Seteo Caption Botones
    If strButton1 <> vbNullString Then
      .axButton1.Caption1 = strButton1
    End If
    If strButton2 <> vbNullString Then
      .axButton2.Caption1 = strButton2
      .axButton2.Visible = True
    End If
    
    .CreateStyle TitlePos, ButtonPos
    .Show vbModal
End With

    ' Retorna Mensaje segun Botón Presionado
Select Case iButtons
  Case Is = bAcceptCancel
      If asClicked = 1 Then AxInputBox = strResp
      If asClicked = 2 Then AxInputBox = vbNullString
  Case Is = bYesNo
      If asClicked = 1 Then AxInputBox = strResp
      If asClicked = 2 Then AxInputBox = vbNullString
  Case Is = bCustom
      If asClicked = 1 Then AxInputBox = strResp
      If asClicked = 2 Then AxInputBox = vbNullString
  Case Is = bAcceptOnly
      If asClicked = 1 Then AxInputBox = strResp
  Case Is = bOKOnly
      If asClicked = 1 Then AxInputBox = strResp
End Select
    
End Function

Public Function AxWrkBox(strMessage As String, strTitle As String, hIcon As String, TitlePos As TStyle, _
                         Optional ColorForm As OLE_COLOR = &HFFFFFF, Optional ColorTitle As OLE_COLOR = &HFFFFFF, _
                         Optional ColorButtons As OLE_COLOR = &HFFFFFF, Optional ColorText As OLE_COLOR = &H80000012, _
                         Optional ColorBorder As OLE_COLOR = &HE0E0E0, Optional MinValue As Integer = 0, _
                         Optional MaxValue As Integer = 100, Optional iButtons As StyleButtons = bAcceptCancel, _
                         Optional strButton1 As String, Optional strButton2 As String) As ClickedButton

With frmLPWBox
  
    Select Case iButtons
      Case bYesNo
        .axButton1.Caption1 = "Si"
        .axButton2.Caption1 = "No"
      Case bAcceptCancel
        .axButton1.Caption1 = "Aceptar"
        .axButton2.Caption1 = "Cancelar"
      Case bOKOnly
        .axButton1.Caption1 = "OK"
        .axButton2.Visible = False
      Case bAcceptOnly
        .axButton1.Caption1 = "Aceptar"
        .axButton2.Visible = False
      Case bCustom
        .axButton1.Caption1 = strButton1
        .axButton2.Caption1 = strButton2
    End Select
     
  ' Seteo Estilo ...
    .lblTitle.BackColor = ColorTitle
    .lblTitle.BorderColor = ColorBorder
    .lblTitle.Caption1 = strTitle
    .lblTitle.Caption1Forecolor = ColorText
    .lblTitle.IconCharCode = hIcon
    .lblMsg.Caption = strMessage
    .lblMsg.ForeColor = ColorText
    .BackColor = ColorForm
    .axButton1.BackColor = ColorButtons
    .axButton2.BackColor = ColorButtons
    .axButton1.BorderColor = ColorBorder
    .axButton2.BorderColor = ColorBorder
    .axButton1.Caption1Forecolor = ColorText
    .axButton2.Caption1Forecolor = ColorText
    .picX.BackColor = ColorButtons
    .picX.BorderColor = ColorBorder
    .picX.IconForeColor = ColorText
    .axProg.Color = ColorTitle
    .axProg.Min = MinValue
    .axProg.Max = MaxValue
    .ucProg.Caption1_ForeColor = ColorText
    .ucProg.PB_Color1 = IIf(ColorTitle = ColorForm, ColorBorder, ColorTitle)
    .ucProg.PB_Color1Opacity = 10
    .ucProg.PF_ForeColor = IIf(ColorTitle = ColorForm, ColorBorder, ColorTitle)
    .ucProg.PF_ForeColorOpacity = 100
    .ucProg.Min = MinValue
    .ucProg.Max = MaxValue
    
    .CreateStyle TitlePos
    .Show vbModal
End With

    ' Retorna Botón Presionado
Select Case iButtons
  Case Is = bAcceptCancel
      If asClicked = 1 Then AxWrkBox = bAccept
      If asClicked = 2 Then AxWrkBox = bCancel
  Case Is = bYesNo
      If asClicked = 1 Then AxWrkBox = bYes
      If asClicked = 2 Then AxWrkBox = bNo
  Case Is = bCustom
      If asClicked = 1 Then AxWrkBox = bButton1
      If asClicked = 2 Then AxWrkBox = bButton2
  Case Is = bAcceptOnly
      If asClicked = 1 Then AxWrkBox = bAccept
  Case Is = bOKOnly
      If asClicked = 1 Then AxWrkBox = bOK
End Select

End Function

Public Sub wMessage(Optional ByVal Value As Integer = -1, Optional ByRef strMsg As String = vbNullString)
With frmLPWBox
  If strMsg <> vbNullString Then .lblMsg.Caption = strMsg
  If Value <> -1 Then
    .axProg.Value = Value
    .ucProg.Value = Value
  End If
End With
End Sub

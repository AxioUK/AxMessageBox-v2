Attribute VB_Name = "mdlForm"
Option Explicit

Public asClicked As Byte
Public strResp   As String

' constantes para capturar Movimiento de Form desde un Control
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

' Crea la región Redondeada
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' Establece la región Redondeada
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
' Dibuja una figura con bordes Redondeados
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    
Private Pic As IPictureDisp

Sub MoveFormBar(fHwnd As Long)
    Call ReleaseCapture
    Call SendMessage(fHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Public Sub RoundCorner(sForm As Form, Radio As Long)
Dim Region As Long
Dim Ret As Long
Dim Ancho As Long
Dim Alto As Long
Dim oScale As Integer
      
    ' guardar la escala
    oScale = sForm.ScaleMode
    ' cambiar la escala a pixeles
    sForm.ScaleMode = vbPixels
    'Obtenemos el ancho y alto de la region del Form
    Ancho = sForm.ScaleWidth '/ Screen.TwipsPerPixelX
    Alto = sForm.ScaleHeight '/ Screen.TwipsPerPixelY
    'Pasar el ancho alto del formualrio y el valor de redondeo .. es decir el radio
    Region = CreateRoundRectRgn(0, 0, Ancho, Alto, Radio, Radio)
    ' Aplica la región al formulario (Crear Puntas Redondeadas)
    Ret = SetWindowRgn(sForm.hwnd, Region, True)
    
  ' Dibujo bordes al Form
  sForm.ForeColor = &HE0E0E0
  sForm.DrawWidth = 1
  RoundRect sForm.hdc, 0, 0, sForm.ScaleWidth - 1, sForm.ScaleHeight - 1, Radio, Radio
  
  ' restaurar la escala
  sForm.ScaleMode = oScale
End Sub

Public Sub CreateTitle(sForm As Form, Radio As Long, OleColor As OLE_COLOR)
Dim oScale As Integer, I As Integer
Dim Rgn As Long, tBar As Long
  
  ' guardar la escala
  oScale = sForm.ScaleMode
  ' cambiar la escala a pixeles
  sForm.ScaleMode = vbPixels

  ' Dibujo bordes al Form
  sForm.ForeColor = OleColor
  sForm.DrawWidth = 1
  
  For I = 45 To 0 Step -1
    tBar = RoundRect(sForm.hdc, 1, 1, sForm.ScaleWidth - 2, I, Radio, Radio)
  Next I
  
  ' restaurar la escala
  sForm.ScaleMode = oScale
End Sub

'Subrutina que dibuja el gráfico en el control Picture en forma centrada y a escala
 '*******************************************************
Sub DrawPic(Objeto As Object, lngImagen As Long)
On Error GoTo ErrSub
   Dim Pos_x As Single
   Dim Pos_y As Single
   Dim Ancho_IMG As Single
   Dim Alto_IMG As Single
   Dim Ancho_Obj As Single
   Dim Alto_Obj As Single
   Dim old_Scale As Single

Set Pic.Picture = lngImagen

With Objeto
    .AutoRedraw = True
    .Cls
    old_Scale = .ScaleMode
    .ScaleMode = vbPixels
    Ancho_IMG = .ScaleX(Pic.Width, vbHimetric, vbPixels)
    Alto_IMG = .ScaleY(Pic.Height, vbHimetric, vbPixels)
    Ancho_Obj = .ScaleWidth
    Alto_Obj = .ScaleHeight
    
    If Ancho_IMG > Ancho_Obj Then
        Alto_IMG = Alto_IMG * Ancho_Obj / Ancho_IMG
        Ancho_IMG = Ancho_Obj
    End If
    If Alto_IMG > Alto_Obj Then
        Ancho_IMG = Ancho_IMG * Alto_Obj / Alto_IMG
        Alto_IMG = Alto_Obj
    End If
    Pos_x = (Ancho_Obj - Ancho_IMG) / 2
    Pos_y = (Alto_Obj - Alto_IMG) / 2
End With
    
   Objeto.PaintPicture Pic, Pos_x, Pos_y, Ancho_IMG, Alto_IMG
   Objeto.ScaleMode = old_Scale
   Exit Sub
    
'Error
ErrSub:
    If Err.Number = 76 Then
       Objeto.Cls
       Exit Sub
    End If
End Sub


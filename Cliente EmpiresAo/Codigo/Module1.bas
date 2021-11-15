Attribute VB_Name = "DrawingPJ"
Sub DibujaElemental(Surface As DirectDrawSurface7, grh As grh, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next
Dim r1           As RECT, r2 As RECT, auxr As RECT
Dim iGrhIndex As Integer
If grh.GrhIndex <= 0 Then Exit Sub
iGrhIndex = GrhData(grh.GrhIndex).Frames(grh.FrameCounter)

With r1
    .Right = GrhData(iGrhIndex).pixelWidth
    .Bottom = GrhData(iGrhIndex).pixelHeight
End With

With r2
   .Left = GrhData(iGrhIndex).sX
   .Top = GrhData(iGrhIndex).sY
   .Right = .Left + GrhData(iGrhIndex).pixelWidth
   .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With


Surface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.BltToDC frmMascota.MascotaView.hdc, auxr, auxr
frmMascota.MascotaView.Refresh

End Sub

Sub DibujaPJ(Surface As DirectDrawSurface7, grh As grh, ByVal X As Integer, ByVal Y As Integer, index As Integer)
On Error Resume Next
Dim r1           As RECT, r2 As RECT, auxr As RECT
Dim iGrhIndex As Integer
If grh.GrhIndex <= 0 Then Exit Sub
iGrhIndex = GrhData(grh.GrhIndex).Frames(grh.FrameCounter)

With r1
    .Right = GrhData(iGrhIndex).pixelWidth
    .Bottom = GrhData(iGrhIndex).pixelHeight
End With

With r2
   .Left = GrhData(iGrhIndex).sX
   .Top = GrhData(iGrhIndex).sY
   .Right = .Left + GrhData(iGrhIndex).pixelWidth
   .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With

Surface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.BltToDC frmCuent.PJ(index).hdc, auxr, auxr

frmCuent.PJ(index).Refresh

End Sub
Sub dibujaban(Surface As DirectDrawSurface7, index As Integer)

Dim r2 As RECT, auxr As RECT

With r2
   .Left = 0
   .Top = 0
   .Right = 20
   .Bottom = 20
End With

With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With

Surface.SetFontTransparency True
Surface.SetForeColor vbRed
frmCuent.font.Size = 15
Surface.SetFont frmMain.font
Surface.BltFast X, Y, Surface, r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.DrawText 6, 60, "Banned", False
Surface.BltToDC frmCuent.PJ(index).hdc, auxr, auxr

End Sub

Sub dibujamuerto(Surface As DirectDrawSurface7, index As Integer)

Dim r2 As RECT, auxr As RECT

With r2
   .Left = 0
   .Top = 0
   .Right = 20
   .Bottom = 20
End With

With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With

Surface.SetFontTransparency True
Surface.SetForeColor vbWhite
frmCuent.font.Size = 6
Surface.SetFont frmCuent.font
Surface.BltFast X, Y, Surface, r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.DrawText 5, 10, "MUERTO", False
Surface.BltToDC frmCuent.PJ(index).hdc, auxr, auxr

End Sub

Sub DibujarTodo(ByVal index As Integer, Body As Integer, Head As Integer, Casco As Integer, Shield As Integer, Weapon As Integer, Baned As Integer, Nombre As String, LVL As Integer, Clase As String, Muerto As Integer)

Dim grh As grh
Dim Pos As Integer
Dim loopc As Integer
Dim r As RECT
Dim r2 As RECT

Dim YBody As Integer
Dim YYY As Integer
Dim XBody As Integer
Dim BBody As Integer


With r2
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With

BackBufferSurface.BltColorFill r, 0

If Baned = 1 Then
    Call dibujaban(BackBufferSurface, index)
End If

frmCuent.Nombre(index).Caption = Nombre

frmCuent.Label1(index).font = frmMain.font
frmCuent.Label1(index).font = frmMain.font

frmCuent.Label1(index).Caption = LVL
frmCuent.Label2(index).Caption = Clase

XBody = 12
YBody = 15
BBody = 17

If Muerto = 1 Then
    Body = 8
    Head = 500
    Arma = 2
    Shield = 2
    Weapon = 2
    XBody = 10
    YBody = 35
    BBody = 16
    Call dibujamuerto(BackBufferSurface, index)
End If

grh = BodyData(Body).Walk(3)
    
Call DibujaPJ(BackBufferSurface, grh, XBody, YBody, index)

If Muerto = 0 Then YYY = BodyData(Body).HeadOffset.Y
If Muerto = 1 Then YYY = -9

Pos = YYY + GrhData(GrhData(grh.GrhIndex).Frames(grh.FrameCounter)).pixelHeight
grh = HeadData(Head).Head(3)
    
Call DibujaPJ(BackBufferSurface, grh, BBody, Pos, index)
    
If Casco <> 2 And Casco > 0 Then
    grh = CascoAnimData(Casco).Head(3)
    Call DibujaPJ(BackBufferSurface, grh, BBody, Pos, index)
End If

If Weapon <> 2 And Weapon > 0 Then
    grh = WeaponAnimData(Weapon).WeaponWalk(3)
    Call DibujaPJ(BackBufferSurface, grh, XBody, YBody, index)
End If

If Shield <> 2 And Shield > 0 Then
    grh = ShieldAnimData(Shield).ShieldWalk(3)
    Call DibujaPJ(BackBufferSurface, grh, XBody, BBody, index)
End If
    
End Sub

Sub DibujarTodoE(Body As Integer)

Dim grh As grh
Dim Pos As Integer
Dim loopc As Integer
Dim r As RECT
Dim r2 As RECT

Dim YBody As Integer
Dim YYY As Integer
Dim XBody As Integer
Dim BBody As Integer


With r2
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With

BackBufferSurface.BltColorFill r, 0

grh = BodyData(Body).Walk(3)
    
Call DibujaElemental(BackBufferSurface, grh, 0, 0)
    
    
End Sub


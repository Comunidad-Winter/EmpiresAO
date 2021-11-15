Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez



Option Explicit

'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'    C       O       N       S      T
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'    T       I      P      O      S
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'Encabezado bmp
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Integer
    Y As Integer
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh
'tama�o y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(1 To 4) As grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(1 To 4) As grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(1 To 4) As grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(1 To 4) As grh
End Type


'Lista de cuerpos
Public Type FxData
    fX As grh
    OffsetX As Long
    OffsetY As Long
End Type

'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As Byte ' As E_Heading ?
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    fX As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    
    Nombre As String
    
    Moving As Byte
    MoveOffset As Position
    ServerIndex As Integer
    
    pie As Boolean
    Muerto As Boolean
    invisible As Boolean
    'PuedoVerlo As Byte
    priv As Byte
    aCaballo As Byte
    
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As grh
    CharIndex As Integer
    ObjGrh As grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
End Type


Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Tama�o del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Offset del desde 0,0 del main view
Public MainViewTop As Integer
Public MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tama�o muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tama�o de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'?�?�?�?�?�?�?�?�?�?�Totales?�?�?�?�?�?�?�?�?�?�?

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

'�?�?�?�?�?�?�?�?�?�Graficos�?�?�?�?�?�?�?�?�?�?�?

Public LastTime As Long 'Para controlar la velocidad


'[CODE]:MatuX'
Public MainDestRect   As RECT
'[END]'
Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer




'�?�?�?�?�?�?�?�?�?�Graficos�?�?�?�?�?�?�?�?�?�?�?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public grh() As grh 'Animaciones publicas
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'�?�?�?�?�?�?�?�?�?�Mapa?�?�?�?�?�?�?�?�?�?�?�?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'�?�?�?�?�?�?�?�?�?�Usuarios?�?�?�?�?�?�?�?�?�?�?�?�?
'
'epa ;)
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'�?�?�?�?�?�?�?�?�?�?�API?�?�?�?�?�?�?�?�?�?�?�?�?�?
'Blt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?


'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'       [CODE 000]: MatuX
'
Public bRain        As Boolean 'est� raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private Rmapa As RECT
Private Rperga As RECT
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer
Private LTMapa As Integer

Public charlist(1 To 10000) As Char

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

'estados internos del surface (read only)
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

#If conalfab Then

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
        
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long
    

#End If

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Sub CargarCabezas()

Dim N As Integer, I As Integer, Numheads As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Numheads

'Resize array
ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For I = 1 To Numheads
    Get #N, , Miscabezas(I)
    InitGrh HeadData(I).Head(1), Miscabezas(I).Head(1), 0
    InitGrh HeadData(I).Head(2), Miscabezas(I).Head(2), 0
    InitGrh HeadData(I).Head(3), Miscabezas(I).Head(3), 0
    InitGrh HeadData(I).Head(4), Miscabezas(I).Head(4), 0
Next I

Close #N

If Numheads < 500 Then
    MsgBox "ERROR: se ha detectado un error en la carga de archivos criticos.", vbCritical
    End
End If

End Sub

Sub CargarCascos()
On Error Resume Next
Dim N As Integer, I As Integer, NumCascos As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCascos

'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For I = 1 To NumCascos
    Get #N, , Miscabezas(I)
    InitGrh CascoAnimData(I).Head(1), Miscabezas(I).Head(1), 0
    InitGrh CascoAnimData(I).Head(2), Miscabezas(I).Head(2), 0
    InitGrh CascoAnimData(I).Head(3), Miscabezas(I).Head(3), 0
    InitGrh CascoAnimData(I).Head(4), Miscabezas(I).Head(4), 0
Next I

Close #N

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim N As Integer, I As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

N = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCuerpos

'Resize array
ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For I = 1 To NumCuerpos
    Get #N, , MisCuerpos(I)
    InitGrh BodyData(I).Walk(1), MisCuerpos(I).Body(1), 0
    InitGrh BodyData(I).Walk(2), MisCuerpos(I).Body(2), 0
    InitGrh BodyData(I).Walk(3), MisCuerpos(I).Body(3), 0
    InitGrh BodyData(I).Walk(4), MisCuerpos(I).Body(4), 0
    BodyData(I).HeadOffset.X = MisCuerpos(I).HeadOffsetX
    BodyData(I).HeadOffset.Y = MisCuerpos(I).HeadOffsetY
Next I

Close #N

End Sub
Sub CargarFxs()
On Error Resume Next
Dim N As Integer, I As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

N = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumFxs

'Resize array
ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For I = 1 To NumFxs
    Get #N, , MisFxs(I)
    Call InitGrh(FxData(I).fX, MisFxs(I).Animacion, 1)
    FxData(I).OffsetX = MisFxs(I).OffsetX
    FxData(I).OffsetY = MisFxs(I).OffsetY
Next I

Close #N

End Sub

Sub CargarTips()
On Error Resume Next
Dim N As Integer, I As Integer
Dim NumTips As Integer

N = FreeFile
Open App.Path & "\init\Tips.ayu" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumTips

'Resize array
ReDim Tips(1 To NumTips) As String * 255

For I = 1 To NumTips
    Get #N, , Tips(I)
Next I

Close #N

End Sub

Sub CargarArrayLluvia()
On Error Resume Next
Dim N As Integer, I As Integer
Dim Nu As Integer

N = FreeFile
Open App.Path & "\init\fk.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Nu

'Resize array
ReDim bLluvia(1 To Nu) As Byte

For I = 1 To Nu
    Get #N, , bLluvia(I)
Next I

Close #N

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.X + CX
tY = UserPos.Y + CY

End Sub






Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer, ByVal Invii As Boolean)

Dim NombreSaved As String

On Error Resume Next

NombreSaved = charlist(CharIndex).Nombre

If charlist(CharIndex).Active = 1 Then Call EraseChar(CharIndex)

'Apuntamos al ultimo Char
If CharIndex > LastChar Then LastChar = CharIndex

If charlist(CharIndex).Active = 0 Then

NumChars = NumChars + 1

If Arma = 0 Then Arma = 2
If Escudo = 0 Then Escudo = 2
If Casco = 0 Then Casco = 2

charlist(CharIndex).Nombre = NombreSaved
charlist(CharIndex).iHead = Head
charlist(CharIndex).iBody = Body
charlist(CharIndex).Head = HeadData(Head)
charlist(CharIndex).Body = BodyData(Body)
charlist(CharIndex).Arma = WeaponAnimData(Arma)
'[ANIM ATAK]
charlist(CharIndex).Arma.WeaponAttack = 0

charlist(CharIndex).Escudo = ShieldAnimData(Escudo)
charlist(CharIndex).Casco = CascoAnimData(Casco)

charlist(CharIndex).Heading = Heading

'Reset moving stats
charlist(CharIndex).Moving = 0
charlist(CharIndex).MoveOffset.X = 0
charlist(CharIndex).MoveOffset.Y = 0

'Update position
charlist(CharIndex).Pos.X = X
charlist(CharIndex).Pos.Y = Y

charlist(CharIndex).invisible = Invii

'Make active
charlist(CharIndex).Active = 1

End If

'Plot on map
MapData(X, Y).CharIndex = CharIndex


End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    charlist(CharIndex).Active = 0
    charlist(CharIndex).Criminal = 0
    charlist(CharIndex).fX = 0
    charlist(CharIndex).FxLoopTimes = 0
    charlist(CharIndex).invisible = False
#If SeguridadAlkon Then
    Call MI(CualMI).ResetInvisible(CharIndex)
#End If

    charlist(CharIndex).Moving = 0
    charlist(CharIndex).Muerto = False
    charlist(CharIndex).Nombre = ""
    charlist(CharIndex).pie = False
    charlist(CharIndex).Pos.X = 0
    charlist(CharIndex).Pos.Y = 0
    charlist(CharIndex).UsandoArma = False

End Sub


Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

charlist(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until charlist(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0

Call ResetCharInfo(CharIndex)

'Update NumChars
NumChars = NumChars - 1


End Sub

Sub InitGrh(ByRef grh As grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(grh.GrhIndex).NumFrames > 1 Then
        grh.Started = 1
    Else
        grh.Started = 0
    End If
Else
    grh.Started = Started
End If

grh.FrameCounter = 1
'[CODE 000]:MatuX
'
'  La linea generaba un error en la IDE, (no ocurr�a debido al
' on error)
'
'   Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
If grh.GrhIndex <> 0 Then grh.SpeedCounter = GrhData(grh.GrhIndex).Speed
'
'[END]'

End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = charlist(CharIndex).Pos.X
Y = charlist(CharIndex).Pos.Y

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        addY = -1

    Case E_Heading.EAST
        addX = 1

    Case E_Heading.SOUTH
        addY = 1
    
    Case E_Heading.WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).CharIndex = CharIndex
charlist(CharIndex).Pos.X = nX
charlist(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

charlist(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
charlist(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

charlist(CharIndex).Moving = 1
charlist(CharIndex).Heading = nHeading

If UserEstado <> 1 Then Call DoPasosFx(CharIndex)

'areas viejos
If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Call EraseChar(CharIndex)
End If

End Sub

Public Sub DoFogataFx()
If Sound Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata()
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim X As Integer, Y As Integer

For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
  For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1
            
            If MapData(X, Y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next X
Next Y

EstaPCarea = False

End Function

Public Function TickON(Cual As Integer, Cont As Integer) As Boolean
Static TickCount(200) As Integer
If Cont = 999 Then Exit Function
TickCount(Cual) = TickCount(Cual) + 1
If TickCount(Cual) < Cont Then
    TickON = False
Else
    TickCount(Cual) = 0
    TickON = True
End If
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
Static pie As Boolean

If Not Sound Then Exit Sub

If charlist(CharIndex).priv = 25 Then Exit Sub

If Not UserNavegando Then
    If userMontando = True And CharIndex = UserCharIndex And EstaPCarea(CharIndex) Then
        If TickON(0, 3) Then Call Audio.PlayWave(SND_GALOPE)
    Else
    If Not charlist(CharIndex).Muerto And EstaPCarea(CharIndex) Then
        charlist(CharIndex).pie = Not charlist(CharIndex).pie
        If charlist(CharIndex).pie Then
            Call Audio.PlayWave(SND_PASOS1)
        Else
            Call Audio.PlayWave(SND_PASOS2)
        End If
    End If
    End If
Else
    Call Audio.PlayWave(SND_NAVEGANDO)
End If

End Sub


Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As E_Heading



X = charlist(CharIndex).Pos.X
Y = charlist(CharIndex).Pos.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = E_Heading.EAST
End If

If Sgn(addX) = -1 Then
    nHeading = E_Heading.WEST
End If

If Sgn(addY) = -1 Then
    nHeading = E_Heading.NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = E_Heading.SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex


charlist(CharIndex).Pos.X = nX
charlist(CharIndex).Pos.Y = nY

charlist(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
charlist(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

charlist(CharIndex).Moving = 1
charlist(CharIndex).Heading = nHeading

'parche para que no medite cuando camina
Dim fxCh As Integer
fxCh = charlist(CharIndex).fX
If fxCh = FxMeditar.CHICO Or fxCh = FxMeditar.GRANDE Or fxCh = FxMeditar.MEDIANO Or fxCh = FxMeditar.XGRANDE Then
    charlist(CharIndex).fX = 0
    charlist(CharIndex).FxLoopTimes = 0
End If

If Not EstaPCarea(CharIndex) Then Dialogos.QuitarDialogo (CharIndex)

If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Call EraseChar(CharIndex)
End If

End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        Y = -1

    Case E_Heading.EAST
        X = 1

    Case E_Heading.SOUTH
        Y = 1
    
    Case E_Heading.WEST
        X = -1
        
End Select

'Fill temp pos
tX = UserPos.X + X
tY = UserPos.Y + Y

'Check to see if its out of bounds
If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    'Start moving... MainLoop does the rest
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
   
End If


    

End Sub


Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.X - 8 To UserPos.X + 8
    For k = UserPos.Y - 6 To UserPos.Y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim loopc As Integer
Dim Dale As Boolean

loopc = 1
Do While charlist(loopc).Active And Dale
    loopc = loopc + 1
    Dale = (loopc <= UBound(charlist))
Loop

NextOpenChar = loopc

End Function


Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim grh As Integer
Dim Frame As Integer
Dim tempint As Integer




'Resize arrays
ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

'Open files
Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

'Fill Grh List

'Get first Grh Number
Get #1, , grh

Do Until grh <= 0
        
    'Get number of frames
    Get #1, , GrhData(grh).NumFrames
    If GrhData(grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(grh).NumFrames
        
            Get #1, , GrhData(grh).Frames(Frame)
            If GrhData(grh).Frames(Frame) <= 0 Or GrhData(grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(grh).Speed
        If GrhData(grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(grh).pixelHeight = GrhData(GrhData(grh).Frames(1)).pixelHeight
        If GrhData(grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(grh).pixelWidth = GrhData(GrhData(grh).Frames(1)).pixelWidth
        If GrhData(grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(grh).TileWidth = GrhData(GrhData(grh).Frames(1)).TileWidth
        If GrhData(grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(grh).TileHeight = GrhData(GrhData(grh).Frames(1)).TileHeight
        If GrhData(grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(grh).FileNum
        If GrhData(grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(grh).sX
        If GrhData(grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(grh).sY
        If GrhData(grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(grh).pixelWidth
        If GrhData(grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(grh).pixelHeight
        If GrhData(grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(grh).TileWidth = GrhData(grh).pixelWidth / TilePixelHeight
        GrhData(grh).TileHeight = GrhData(grh).pixelHeight / TilePixelWidth
        
        GrhData(grh).Frames(1) = grh
            
    End If

    'Get Next Grh Number
    Get #1, , grh

Loop
'************************************************

Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & grh

End Sub

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Limites del mapa
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    '�Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function




Function InMapLegalBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, grh As grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)

Dim CurrentGrh As grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If grh.Started = 1 Then
        If grh.SpeedCounter > 0 Then
            grh.SpeedCounter = grh.SpeedCounter - 1
            If grh.SpeedCounter = 0 Then
                grh.SpeedCounter = GrhData(grh.GrhIndex).Speed
                grh.FrameCounter = grh.FrameCounter + 1
                If grh.FrameCounter > GrhData(grh.GrhIndex).NumFrames Then
                    grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If
'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(grh.GrhIndex).Frames(grh.FrameCounter)
'Center Grh over X,Y pos
If center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
Surface.BltFast X, Y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
End Sub

Sub DDrawTransGrhIndextoSurface(Surface As DirectDrawSurface7, grh As Integer, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = X
    .Top = Y
    .Right = .Left + GrhData(grh).pixelWidth
    .Bottom = .Top + GrhData(grh).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(grh).sX
        .Top = GrhData(grh).sY
        .Right = .Left + GrhData(grh).pixelWidth
        .Bottom = .Top + GrhData(grh).pixelHeight
    End With
    
    Surface.BltFast destRect.Left, destRect.Top, SurfaceDB.Surface(GrhData(grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub

'Sub DDrawTransGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[CODE 000]:MatuX
    Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, grh As grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean


If Animate Then
    If grh.Started = 1 Then
        If grh.SpeedCounter > 0 Then
            grh.SpeedCounter = grh.SpeedCounter - 1
            If grh.SpeedCounter = 0 Then
                grh.SpeedCounter = GrhData(grh.GrhIndex).Speed
                grh.FrameCounter = grh.FrameCounter + 1
                If grh.FrameCounter > GrhData(grh.GrhIndex).NumFrames Then
                    grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).fX = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(grh.GrhIndex).Frames(grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With


Surface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub

#If conalfab = 1 Then
    Sub DDrawTransGrhtoSurfaceAlpha(Surface As DirectDrawSurface7, grh As grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)

'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean


If Animate Then
    If grh.Started = 1 Then
        If grh.SpeedCounter > 0 Then
            grh.SpeedCounter = grh.SpeedCounter - 1
            If grh.SpeedCounter = 0 Then
                grh.SpeedCounter = GrhData(grh.GrhIndex).Speed
                grh.FrameCounter = grh.FrameCounter + 1
                If grh.FrameCounter > GrhData(grh.GrhIndex).NumFrames Then
                    grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then

                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).fX = 0
                                Exit Sub
                            End If

                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(grh.GrhIndex).Frames(grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

'surface.BltFast X, Y, SurfaceDB.surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Dim Src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long

Set Src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)

Src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest

With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
    
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With

' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 3
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 4
Else
    'Modo = 2 '16 bits raro ?
    Surface.BltFast X, Y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False

On Local Error GoTo HayErrorAlpha

Src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
Surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()
Src.GetLockedArray sArray()

Call BltAlphaFast(ByVal VarPtr(dArray(X + X, Y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)

Surface.Unlock rDest
DstLock = False
Src.Unlock SourceRect
SrcLock = False


Exit Sub

HayErrorAlpha:
If SrcLock Then Src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest

End Sub
#End If 'ConAlfaB = 1

Sub DrawBackBufferSurface()
    PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(hwnd As Long, hdc As Long, grh As Integer, SourceRect As RECT, destRect As RECT)
    If grh <= 0 Then Exit Sub
    
    SecundaryClipper.SetHWnd hwnd
    SurfaceDB.Surface(GrhData(grh).FileNum).BltToDC hdc, SourceRect, destRect
End Sub

Sub RenderScreen(tilex As Integer, tiley As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
On Error Resume Next


If UserCiego Then Exit Sub

Dim Y        As Integer 'Keeps track of where on map we are
Dim X        As Integer 'Keeps track of where on map we are
Dim minY     As Integer 'Start Y pos on current map
Dim maxY     As Integer 'End Y pos on current map
Dim minX     As Integer 'Start X pos on current map
Dim maxX     As Integer 'End X pos on current map
Dim ScreenX  As Integer 'Keeps track of where to place tile on screen
Dim ScreenY  As Integer 'Keeps track of where to place tile on screen
Dim Moved    As Byte
Dim grh      As grh     'Temp Grh for show tile and blocked
Dim TempChar As Char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer 'Usado en el Layer de Chars
Dim iPPy     As Integer 'Usado en el Layer de Chars
Dim rSourceRect      As RECT    'Usado en el Layer 1
Dim iGrhIndex        As Integer 'Usado en el Layer 1
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim nX As Integer
Dim nY As Integer

'Figure out Ends and Starts of screen
' Hardcodeado para speed!
minY = (tiley - 15)
maxY = (tiley + 15)
minX = (tilex - 17)
maxX = (tilex + 17)


'Draw floor layer
ScreenY = 8
For Y = (minY + 8) To maxY - 8
    ScreenX = 8
    For X = minX + 8 To maxX - 8
        If X > 100 Or Y < 1 Then Exit For
        'Layer 1 **********************************
        With MapData(X, Y).Graphic(1)
            If (.Started = 1) Then
                If (.SpeedCounter > 0) Then
                    .SpeedCounter = .SpeedCounter - 1
                    If (.SpeedCounter = 0) Then
                        .SpeedCounter = GrhData(.GrhIndex).Speed
                        .FrameCounter = .FrameCounter + 1
                        If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then _
                            .FrameCounter = 1
                    End If
                End If
            End If

            'Figure out what frame to draw (always 1 if not animated)
            iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
        End With

        rSourceRect.Left = GrhData(iGrhIndex).sX
        rSourceRect.Top = GrhData(iGrhIndex).sY
        rSourceRect.Right = rSourceRect.Left + GrhData(iGrhIndex).pixelWidth
        rSourceRect.Bottom = rSourceRect.Top + GrhData(iGrhIndex).pixelHeight

        'El width fue hardcodeado para speed!
        Call BackBufferSurface.BltFast( _
                ((32 * ScreenX) - 32) + PixelOffsetX, _
                ((32 * ScreenY) - 32) + PixelOffsetY, _
                SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), _
                rSourceRect, _
                DDBLTFAST_WAIT)
        '******************************************
        'Layer 2 **********************************
        If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
            Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    MapData(X, Y).Graphic(2), _
                    ((32 * ScreenX) - 32) + PixelOffsetX, _
                    ((32 * ScreenY) - 32) + PixelOffsetY, _
                    1, _
                    1)
        End If
        '******************************************
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y > 100 Then Exit For
Next Y


'busco que nombre dibujar
Call ConvertCPtoTP(frmMain.MainViewShp.Left, frmMain.MainViewShp.Top, frmMain.MouseX, frmMain.MouseY, nX, nY)


'Draw Transparent Layers  (Layer 2, 3)
ScreenY = 8
For Y = minY + 8 To maxY - 1
    ScreenX = 5
    For X = minX + 5 To maxX - 5
        If X > 100 Or X < -3 Then Exit For
        iPPx = 32 * ScreenX - 32 + PixelOffsetX
        iPPy = 32 * ScreenY - 32 + PixelOffsetY

        'Object Layer **********************************
        If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then
'            If Y > UserPos.Y Then
'                Call DDrawTransGrhtoSurfaceAlpha( _
'                        BackBufferSurface, _
'                        MapData(X, Y).ObjGrh, _
'                        iPPx, iPPy, 1, 1)
'            Else
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).ObjGrh, _
                        iPPx, iPPy, 1, 1)
'            End If
        End If
        '***********************************************
        'Char layer ************************************
        If MapData(X, Y).CharIndex <> 0 Then
            TempChar = charlist(MapData(X, Y).CharIndex)
            PixelOffsetXTemp = PixelOffsetX
            PixelOffsetYTemp = PixelOffsetY

            Moved = 0
            'If needed, move left and right
            If TempChar.MoveOffset.X <> 0 Then
                TempChar.Body.Walk(TempChar.Heading).Started = 1
                TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 1
                TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 1
                PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
                TempChar.MoveOffset.X = TempChar.MoveOffset.X - (8 * Sgn(TempChar.MoveOffset.X))
                Moved = 1
            End If
            'If needed, move up and down
            If TempChar.MoveOffset.Y <> 0 Then
                TempChar.Body.Walk(TempChar.Heading).Started = 1
                TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 1
                TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 1
                PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
                TempChar.MoveOffset.Y = TempChar.MoveOffset.Y - (8 * Sgn(TempChar.MoveOffset.Y))
                Moved = 1
            End If
            'If done moving stop animation
            If Moved = 0 And TempChar.Moving = 1 Then
                TempChar.Moving = 0
                TempChar.Body.Walk(TempChar.Heading).FrameCounter = 1
                TempChar.Body.Walk(TempChar.Heading).Started = 0
                TempChar.Arma.WeaponWalk(TempChar.Heading).FrameCounter = 1
                TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 0
                TempChar.Escudo.ShieldWalk(TempChar.Heading).FrameCounter = 1
                TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 0
            End If
            
            '[ANIM ATAK]
            If TempChar.Arma.WeaponAttack > 0 Then
                TempChar.Arma.WeaponAttack = TempChar.Arma.WeaponAttack - 1
                If TempChar.Arma.WeaponAttack = 0 Then
                    TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 0
                End If
            End If
            '[/ANIM ATAK]
            
            'Dibuja solamente players
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
            If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                If Not charlist(MapData(X, Y).CharIndex).invisible Or MapData(X, Y).CharIndex = UserCharIndex Or SoyGM = 1 Or MiClan(charlist(MapData(X, Y).CharIndex).Nombre) Then
#If SeguridadAlkon Then
                    If Not MI(CualMI).IsInvisible(MapData(X, Y).CharIndex) Then
#End If
                        #If conalfab = 1 Then
                            If Not charlist(MapData(X, Y).CharIndex).invisible Then
                        #End If
                        '[CUERPO]'
                            Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                    1, 1)
                        '[CABEZA]'
                            Call DDrawTransGrhtoSurface( _
                                    BackBufferSurface, _
                                    TempChar.Head.Head(TempChar.Heading), _
                                    iPPx + TempChar.Body.HeadOffset.X, _
                                    iPPy + TempChar.Body.HeadOffset.Y, _
                                    1, 0)
                        '[Casco]'
                            If TempChar.Casco.Head(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Casco.Head(TempChar.Heading), _
                                        iPPx + TempChar.Body.HeadOffset.X, _
                                        iPPy + TempChar.Body.HeadOffset.Y, _
                                        1, 0)
                            End If
                        '[ARMA]'
                            If TempChar.Arma.WeaponWalk(TempChar.Heading).GrhIndex <> 0 Then
                                If TempChar.aCaballo = 0 Then
                                        Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Arma.WeaponWalk(TempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                                Else
                                        Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Arma.WeaponWalk(TempChar.Heading), _
                                        iPPx, iPPy - 15, 1, 1)
                                        
                                End If
                            End If
                        '[Escudo]'
                            If TempChar.Escudo.ShieldWalk(TempChar.Heading).GrhIndex <> 0 Then
                                If TempChar.aCaballo = 0 Then
                                    Call DDrawTransGrhtoSurface( _
                                            BackBufferSurface, _
                                            TempChar.Escudo.ShieldWalk(TempChar.Heading), _
                                            iPPx, iPPy, 1, 1)
                                Else
                                    Call DDrawTransGrhtoSurface( _
                                            BackBufferSurface, _
                                            TempChar.Escudo.ShieldWalk(TempChar.Heading), _
                                            iPPx + 2, iPPy - 15, 1, 1)
                                End If
                            End If
                        #If conalfab = 1 Then
                            Else
                            Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                    1, 1)
                        '[CABEZA]'
                            Call DDrawTransGrhtoSurfaceAlpha( _
                                    BackBufferSurface, _
                                    TempChar.Head.Head(TempChar.Heading), _
                                    iPPx + TempChar.Body.HeadOffset.X, _
                                    iPPy + TempChar.Body.HeadOffset.Y, _
                                    1, 0)
                        '[Casco]'
                            If TempChar.Casco.Head(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurfaceAlpha( _
                                        BackBufferSurface, _
                                        TempChar.Casco.Head(TempChar.Heading), _
                                        iPPx + TempChar.Body.HeadOffset.X, _
                                        iPPy + TempChar.Body.HeadOffset.Y, _
                                        1, 0)
                            End If
                        '[ARMA]'
                            If TempChar.Arma.WeaponWalk(TempChar.Heading).GrhIndex <> 0 Then
                                If TempChar.aCaballo = 0 Then
                                        Call DDrawTransGrhtoSurfaceAlpha( _
                                        BackBufferSurface, _
                                        TempChar.Arma.WeaponWalk(TempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                                Else
                                        Call DDrawTransGrhtoSurfaceAlpha( _
                                        BackBufferSurface, _
                                        TempChar.Arma.WeaponWalk(TempChar.Heading), _
                                        iPPx, iPPy - 15, 1, 1)
                                        
                                End If
                            End If
                        '[Escudo]'
                            If TempChar.Escudo.ShieldWalk(TempChar.Heading).GrhIndex <> 0 Then
                                If TempChar.aCaballo = 0 Then
                                    Call DDrawTransGrhtoSurfaceAlpha( _
                                            BackBufferSurface, _
                                            TempChar.Escudo.ShieldWalk(TempChar.Heading), _
                                            iPPx, iPPy, 1, 1)
                                Else
                                    Call DDrawTransGrhtoSurfaceAlpha( _
                                            BackBufferSurface, _
                                            TempChar.Escudo.ShieldWalk(TempChar.Heading), _
                                            iPPx + 2, iPPy - 15, 1, 1)
                                End If
                            End If
                            End If
                        #End If
                    
                             'If Nombres And Abs(nX - X) < 2 And (Abs(nY - Y)) < 2 Then
                                'ya estoy dibujando SOLO si esta visible
                               ' If TempChar.invisible = False Then 'And Not MI(CualMI).IsInvisible(MapData(X, Y).CharIndex) Then
                                    If TempChar.Nombre <> "" Then
                                        Dim lCenter As Long
                                        'Call Dialogos.DrawText(iPPx - 30, iPPy + 60, "mi:" & IIf(MI(CualMI).IsInvisible(MapData(X, Y).CharIndex), "1", "0") & " .i:" & IIf(TempChar.invisible, "1", "0") & "  X,Y:" & X & "," & Y, RGB(ColoresPJ(5).r, ColoresPJ(5).G, ColoresPJ(5).B))
                                        If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
                                            lCenter = (frmMain.TextWidth(Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                            Dim sClan As String
                                            sClan = mid(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
                                            
                                            If TempChar.invisible = False Then
                                            Select Case TempChar.priv
                                            Case 0
                                                If TempChar.Criminal Then
                                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b))
                                                    lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(51).r, ColoresPJ(51).g, ColoresPJ(51).b))
                                                Else
                                                   Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b))
                                                    lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(51).r, ColoresPJ(51).g, ColoresPJ(51).b))
                                                End If
                                                    
                                            Case 25  'admin
                                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b))
                                                    lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(51).r, ColoresPJ(51).g, ColoresPJ(51).b))
                                            Case Else 'el resto
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).g, ColoresPJ(TempChar.priv).b))
                                                lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(51).r, ColoresPJ(51).g, ColoresPJ(51).b))
                                            End Select
                                            Else 'else de invisibilidad
                                            Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(52).r, ColoresPJ(52).g, ColoresPJ(52).b))
                                                lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(52).r, ColoresPJ(52).g, ColoresPJ(52).b))
                                                End If 'termina if de invi
                                        Else
                                            lCenter = (frmMain.TextWidth(TempChar.Nombre) / 2) - 16
                                            If TempChar.invisible = False Then
                                            Select Case TempChar.priv
                                            Case 0
                                                If TempChar.Criminal Then
                                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b))
                                                Else
                                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b))
                                                End If
                                            Case 7
                                                Call Dialogos.DrawTextBig(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).g, ColoresPJ(TempChar.priv).b))
                                            Case Else
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).g, ColoresPJ(TempChar.priv).b))
                                            End Select
                                            Else 'else de invi sin clan
                                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(52).r, ColoresPJ(52).g, ColoresPJ(52).b))
                                            End If 'fin de if sin clan
                                        End If
                                    End If
                               ' End If  'enidf nI
                            'End If
#If SeguridadAlkon Then
                    Else
                        Do While True
                            Call MsgBox("WOAAAAA CHEATER!!! Ahora te deben estar matando de lo lindo ;)" & vbNewLine & "Aprieta OK para salir", vbCritical + vbOKOnly, ":D")
                            Call MsgBox("no, mejor no salimos")
                        Loop
                    End If  'end if not mi.isi
#End If
                End If  'end if ~in

                If Dialogos.CantidadDialogos > 0 Then
                    Call Dialogos.Update_Dialog_Pos( _
                            (iPPx + TempChar.Body.HeadOffset.X), _
                            (iPPy + TempChar.Body.HeadOffset.Y), _
                            MapData(X, Y).CharIndex)
                End If
                
                
            Else '<-> If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                If Dialogos.CantidadDialogos > 0 Then
                    Call Dialogos.Update_Dialog_Pos( _
                            (iPPx + TempChar.Body.HeadOffset.X), _
                            (iPPy + TempChar.Body.HeadOffset.Y), _
                            MapData(X, Y).CharIndex)
                End If

                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        TempChar.Body.Walk(TempChar.Heading), _
                        iPPx, iPPy, 1, 1)
            End If '<-> If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then


            'Refresh charlist
            charlist(MapData(X, Y).CharIndex) = TempChar

            'BlitFX (TM)
            If charlist(MapData(X, Y).CharIndex).fX <> 0 Then
#If (conalfab = 1) Then
                Call DDrawTransGrhtoSurfaceAlpha( _
                        BackBufferSurface, _
                        FxData(TempChar.fX).fX, _
                        iPPx + FxData(TempChar.fX).OffsetX, _
                        iPPy + FxData(TempChar.fX).OffsetY, _
                        1, 1, MapData(X, Y).CharIndex)
#Else
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        FxData(TempChar.fX).fX, _
                        iPPx + FxData(TempChar.fX).OffsetX, _
                        iPPy + FxData(TempChar.fX).OffsetY, _
                        1, 1, MapData(X, Y).CharIndex)
#End If
            End If
        End If '<-> If MapData(X, Y).CharIndex <> 0 Then
        '*************************************************
        'Layer 3 *****************************************
        If MapData(X, Y).Graphic(3).GrhIndex <> 0 Then
            'Draw
            Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    MapData(X, Y).Graphic(3), _
                    ((32 * ScreenX) - 32) + PixelOffsetX, _
                    ((32 * ScreenY) - 32) + PixelOffsetY, _
                    1, 1)
        End If
        '************************************************
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y >= 100 Or Y < 1 Then Exit For
Next Y

If Not bTecho Then
    'Draw blocked tiles and grid
    ScreenY = 5
    For Y = minY + 5 To maxY - 1
        ScreenX = 5
        For X = minX + 5 To maxX
            'Check to see if in bounds
            If X < 101 And X > 0 And Y < 101 And Y > 0 Then
                If MapData(X, Y).Graphic(4).GrhIndex <> 0 Then
                    'Draw
                    #If conalfab Then
                    Call DDrawTransGrhtoSurfaceAlpha( _
                        BackBufferSurface, _
                        MapData(X, Y).Graphic(4), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, 1)
                    #Else
                    Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).Graphic(4), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, 1)
                    #End If
                End If
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
End If

If bLluvia(UserMap) = 1 Then
    If bRain Then
                'Figure out what frame to draw
                If llTick < DirectX.TickCount - 50 Then
                    iFrameIndex = iFrameIndex + 1
                    If iFrameIndex > 7 Then iFrameIndex = 0
                    llTick = DirectX.TickCount
                End If
    
                For Y = 0 To 4
                    For X = 0 To 4
                        Call BackBufferSurface.BltFast(LTLluvia(Y), LTLluvia(X), SurfaceDB.Surface(5556), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
                    Next X
                Next Y
    End If
End If

If VeoMapa Then Call BackBufferSurface.BltFast(390, 330, SurfaceDB.Surface(15000), Rmapa, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
If VeoPergamino Then Call BackBufferSurface.BltFast(370, 260, SurfaceDB.Surface(88894), Rperga, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)


Dim PP As RECT

PP.Left = 0
PP.Top = 0
PP.Right = WindowTileWidth * TilePixelWidth
PP.Bottom = WindowTileHeight * TilePixelHeight

'Call BackBufferSurface.BltFast(2, 2, SurfaceDB.Surface(15000), Rmapa, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)

#If conalfab Then
If nightEffect = True Then Call EfectoNoche(BackBufferSurface)
#End If

'[USELESS]:El codigo para llamar a la noche, nublado, etc.
           ' If bTecho Then
              '  Dim bbarray() As Byte, nnarray() As Byte
              '  Dim ddsdBB As DDSURFACEDESC2 'backbuffer
              '  Dim ddsdNN As DDSURFACEDESC2 'nnublado
              '  Dim r As RECT, r2 As RECT
              '  Dim retVal As Long
                '[LOCK]:BackBufferSurface
                 '   BackBufferSurface.GetSurfaceDesc ddsdBB
                '    BackBufferSurface.Lock r, ddsdBB, DDLOCK_NOSYSLOCK + DDLOCK_WRITEONLY + DDLOCK_WAIT, 0
               '     BackBufferSurface.Lock r, ddsdBB, DDLOCK_WRITEONLY + DDLOCK_WAIT, 0
              '      BackBufferSurface.GetLockedArray bbarray()
'                '[LOCK]:BBMask
'                    SurfaceXU(2).GetSurfaceDesc ddsdNN
'                    SurfaceXU(2).Lock r2, ddsdNN, DDLOCK_READONLY + DDLOCK_NOSYSLOCK + DDLOCK_WAIT, 0
'                    SurfaceXU(2).Lock r2, ddsdNN, DDLOCK_READONLY + DDLOCK_WAIT, 0
'                    SurfaceXU(2).GetLockedArray nnarray()
                '[BLIT]'
                    'retVal = BlitNoche(bbarray(0, 0), ddsdBB.lHeight, ddsdBB.lWidth, 0)
                    'retval = BlitNublar(bbarray(0, 0), ddsdBB.lHeight, ddsdBB.lWidth)
                    'retVal = BlitNublarMMX(bbarray(0, 0), nnarray(0, 0), ddsdBB.lHeight, ddsdBB.lWidth, ddsdBB.lPitch, ddsdNN.lHeight, ddsdNN.lWidth, ddsdNN.lPitch)
                '[UNLOCK]'
            '        BackBufferSurface.Unlock r
                    'SurfaceXU(2).Unlock r2
                '[END]'
             '   If retVal = -1 Then MsgBox "error!"
      '      End If
'[END]'
End Sub
Public Function RenderSounds()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 4/22/2006
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bLluvia(UserMap) = 1 And Sound Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviain.wav", LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
    
    DoFogataFx
End Function


Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
        And charlist(UserCharIndex).Pos.Y <= Y
        
End If
End Function

Function PixelPos(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
    PixelPos = (TilePixelWidth * X) - TilePixelWidth
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
          
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128

    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
    
    Rmapa.Top = 0
    Rmapa.Left = 0
    Rmapa.Right = 280
    Rmapa.Bottom = 280
    
    Rperga.Top = 0
    Rperga.Left = 0
    Rperga.Right = 309
    Rperga.Bottom = 413
    
    'We are done!
    AddtoRichTextBox frmCargando.status, "Hecho.", , 128, 128, 1, , False
End Sub

'[END]'
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\Init\"

'Set intial user position
UserPos.X = MinXBorder
UserPos.Y = MinYBorder

'Fill startup variables

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)


ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock





DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With



Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hwnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
    .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
End With

With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If ClientSetup.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck



Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs

LTMapa = 50

LTLluvia(0) = 224
LTLluvia(1) = 352
LTLluvia(2) = 480
LTLluvia(3) = 608
LTLluvia(4) = 736

AddtoRichTextBox frmCargando.status, "Cargando Gr�ficos....", 0, 128, 128, , , True
Call LoadGraphics

InitTileEngine = True

End Function

Sub ShowNextFrame()
'***********************************************
'Updates and draws next frame to screen
'***********************************************
    Static OffsetCounterX As Integer
    Static OffsetCounterY As Integer
    Static I As Integer
    
    '****** Set main view rectangle ******
    GetWindowRect DisplayFormhWnd, MainViewRect
    
    With MainViewRect
        .Left = .Left + MainViewLeft
        .Top = .Top + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    If EngineRun Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.X)))
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
                UserMoving = 0
            End If
        '****** Move screen Up and Down if needed ******
        ElseIf AddtoUserPos.Y <> 0 Then
            OffsetCounterY = (OffsetCounterY - (8 * Sgn(AddtoUserPos.Y)))
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                UserMoving = 0
            End If
        End If

        '****** Update screen ******
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        
        If IScombate Then
            Call Dialogos.DrawText(260, 260, "MODO COMBATE", vbRed)
            frmMain.Label2(0).ForeColor = &H8000&
        Else
            frmMain.Label2(0).ForeColor = vbRed
        End If
        
        Call Dialogos.DrawText(260, 630, "X: " & UserPos.X, vbWhite)
        Call Dialogos.DrawText(260, 640, "Y: " & UserPos.Y, vbWhite)
        Call Dialogos.DrawText(260, 650, "Mapa: " & UserMap & " - " & UserMapName, vbWhite)
        
        Call Dialogos.MostrarTexto
        Call DibujarCartel
        
        Call DialogosClanes.Draw(Dialogos)
        
        Call DrawBackBufferSurface
        
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

Sub CrearGrh(GrhIndex As Integer, index As Integer)
ReDim Preserve grh(1 To index) As grh
grh(index).FrameCounter = 1
grh(index).GrhIndex = GrhIndex
grh(index).SpeedCounter = GrhData(GrhIndex).Speed
grh(index).Started = 1
End Sub

Sub CargarAnimsExtra()
Call CrearGrh(6580, 1) 'Anim Invent
Call CrearGrh(534, 2) 'Animacion de teleport
End Sub

Function ControlVelocidad(ByVal LastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - LastTime > 20)
End Function


#If conalfab Then

Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
    
    Surface.GetSurfaceDesc ddsdDest
    
    With rRect
        .Left = 0
        .Top = 0
        .Right = ddsdDest.lWidth
        .Bottom = ddsdDest.lHeight
    End With
    
   If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
   'If ddsdDest.ddpfPixelFormat.lGBitMask = &H0& Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
        Modo = 1
    Else
        Modo = 2
    End If
    
    Dim DstLock As Boolean
    DstLock = False
    
    On Local Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
    
    Surface.GetLockedArray dArray()
    
    Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
        ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
        Modo)
    
HayErrorAlpha:
    If DstLock = True Then
        Surface.Unlock rRect
        DstLock = False
    End If
End Sub

#End If

Public Function MiClan(ByVal Nombre As String) As Boolean

Dim sClan As String
Dim lCenter As Long

If InStr(Nombre, "<") > 0 And InStr(Nombre, ">") > 0 Then
    lCenter = (frmMain.TextWidth(Left(Nombre, InStr(Nombre, "<") - 1)) / 2) - 16
    sClan = mid(Nombre, InStr(Nombre, "<"))
    
    If sClan = "<" & MyGuildName & ">" Then
        MiClan = True
    Else
        MiClan = False
    End If

End If


End Function

VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCrearPersonaje.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMascota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   285
      Left            =   6390
      TabIndex        =   12
      Top             =   5625
      Width           =   2265
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   360
      Left            =   2160
      TabIndex        =   11
      Top             =   2220
      Width           =   7455
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0CCA
      Left            =   6285
      List            =   "frmCrearPersonaje.frx":0CDD
      MouseIcon       =   "frmCrearPersonaje.frx":0D0A
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3120
      Width           =   2385
   End
   Begin VB.ComboBox lstProfesion 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":19D4
      Left            =   6285
      List            =   "frmCrearPersonaje.frx":1A0E
      MouseIcon       =   "frmCrearPersonaje.frx":1AB4
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3930
      Width           =   2385
   End
   Begin VB.ComboBox lstGenero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":277E
      Left            =   6285
      List            =   "frmCrearPersonaje.frx":2788
      MouseIcon       =   "frmCrearPersonaje.frx":279B
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4770
      Width           =   2385
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   11280
      Picture         =   "frmCrearPersonaje.frx":3465
      Top             =   360
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   180
      Index           =   4
      Left            =   4800
      MouseIcon       =   "frmCrearPersonaje.frx":3993
      MousePointer    =   99  'Custom
      Top             =   5205
      Width           =   150
   End
   Begin VB.Image Image3 
      Height          =   180
      Index           =   3
      Left            =   4800
      MouseIcon       =   "frmCrearPersonaje.frx":465D
      MousePointer    =   99  'Custom
      Top             =   4815
      Width           =   150
   End
   Begin VB.Image Image3 
      Height          =   180
      Index           =   2
      Left            =   4800
      MouseIcon       =   "frmCrearPersonaje.frx":5327
      MousePointer    =   99  'Custom
      Top             =   4425
      Width           =   150
   End
   Begin VB.Image Image3 
      Height          =   180
      Index           =   1
      Left            =   4800
      MouseIcon       =   "frmCrearPersonaje.frx":5FF1
      MousePointer    =   99  'Custom
      Top             =   4035
      Width           =   150
   End
   Begin VB.Image Image3 
      Height          =   180
      Index           =   0
      Left            =   4800
      MouseIcon       =   "frmCrearPersonaje.frx":6CBB
      MousePointer    =   99  'Custom
      Top             =   3645
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   5625
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   4
      Left            =   4170
      MouseIcon       =   "frmCrearPersonaje.frx":7985
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   3
      Left            =   4170
      MouseIcon       =   "frmCrearPersonaje.frx":864F
      MousePointer    =   99  'Custom
      Top             =   4785
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   2
      Left            =   4170
      MouseIcon       =   "frmCrearPersonaje.frx":9319
      MousePointer    =   99  'Custom
      Top             =   4395
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   4170
      MouseIcon       =   "frmCrearPersonaje.frx":9FE3
      MousePointer    =   99  'Custom
      Top             =   3975
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   4155
      MouseIcon       =   "frmCrearPersonaje.frx":ACAD
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   810
      Left            =   3795
      MouseIcon       =   "frmCrearPersonaje.frx":B977
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   6195
      Width           =   4275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4425
      TabIndex        =   5
      Top             =   4800
      Width           =   225
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4425
      TabIndex        =   3
      Top             =   4410
      Width           =   210
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4425
      TabIndex        =   2
      Top             =   5145
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4425
      TabIndex        =   1
      Top             =   4050
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4425
      TabIndex        =   0
      Top             =   3675
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public SkillPoints As Byte

Function CheckData() As Boolean

CheckData = False

If Username = "" Then
    MsgBox "Seleccione un nombre para el personaje"
    Exit Function
End If

If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If PetSelected = "" Then
    MsgBox "Seleccione un nombre para su mascota."
    Exit Function
End If

If Label1.Caption <> 0 Then
    MsgBox "Los atributos del personaje son invalidos."
    Exit Function
End If

CheckData = True


End Function

Private Sub boton_Click(index As Integer)



End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub TirarDados()

lbFuerza.Caption = "10"
lbInteligencia.Caption = "10"
lbAgilidad.Caption = "10"
lbCarisma.Caption = "10"
lbConstitucion.Caption = "10"

End Sub

Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\graficos\bckCP.bmp")
Image2.Picture = LoadPicture(App.Path & "\Graficos\cmdcp.bmp")

Dim I As Integer
lstProfesion.Clear
For I = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(I)
Next I

lstProfesion.listIndex = 1

Call TirarDados
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2.Tag = "1" Then
    Call Audio.PlayWave(SND_OVER)
    Image2.Tag = "0"
    Image2.Picture = LoadPicture(App.Path & "\Graficos\cmdcp.bmp")
End If
End Sub

Private Sub Image1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

If Label1.Caption > 0 Then

    Select Case index
    Case 0
    
        If lbFuerza.Caption < 18 Then
            lbFuerza.Caption = lbFuerza.Caption + 1
            Label1.Caption = Label1.Caption - 1
        End If
        
    Case 1
    
        If lbAgilidad.Caption < 18 Then
            lbAgilidad.Caption = lbAgilidad.Caption + 1
            Label1.Caption = Label1.Caption - 1
        End If
        
    Case 2
    
        If lbInteligencia.Caption < 18 Then
            lbInteligencia.Caption = lbInteligencia.Caption + 1
            Label1.Caption = Label1.Caption - 1
        End If
        
    Case 3
        
        If lbCarisma.Caption < 18 Then
            lbCarisma.Caption = lbCarisma.Caption + 1
            Label1.Caption = Label1.Caption - 1
        End If
        
    Case 4
        
        If lbConstitucion.Caption < 18 Then
            lbConstitucion.Caption = lbConstitucion.Caption + 1
            Label1.Caption = Label1.Caption - 1
        End If
        
    End Select
    
End If

End Sub

Private Sub Image2_Click()

Username = txtNombre.Text
        
If Right$(Username, 1) = " " Then
    Username = RTrim$(Username)
    MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
End If

UserRaza = lstRaza.List(lstRaza.listIndex)
UserSexo = lstGenero.List(lstGenero.listIndex)
UserClase = lstProfesion.List(lstProfesion.listIndex)
        
UserAtributos(1) = Val(lbFuerza.Caption) - 10
UserAtributos(2) = Val(lbAgilidad.Caption) - 10
UserAtributos(3) = Val(lbInteligencia.Caption) - 10
UserAtributos(4) = Val(lbCarisma.Caption) - 10
UserAtributos(5) = Val(lbConstitucion.Caption) - 10

PetSelected = txtMascota.Text
        
If CheckData() Then

#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
#End If

    'SendNewChar = True
    EstadoLogin = CrearNuevoPj
    
    Me.MousePointer = 99

    EstadoLogin = CrearNuevoPj

#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State <> sckConnected Then
#End If
        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
        
    Else
        Call Login(ValidarLoginMSG(CInt(bRK)))
    End If
End If



End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2.Tag = "0" Then
    Call Audio.PlayWave(SND_OVER)
    Image2.Tag = "1"
    Image2.Picture = LoadPicture(App.Path & "\Graficos\cmdcpA.bmp")
End If
End Sub

Private Sub Image4_Click()

Unload frmCrearPersonaje

End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

Private Sub Image3_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

If Label1.Caption >= 0 Then

    Select Case index
    Case 0
    
        If lbFuerza.Caption > 10 Then
            lbFuerza.Caption = lbFuerza.Caption - 1
            Label1.Caption = Label1.Caption + 1
        End If
        
    Case 1
    
        If lbAgilidad.Caption > 10 Then
            lbAgilidad.Caption = lbAgilidad.Caption - 1
            Label1.Caption = Label1.Caption + 1
        End If
        
    Case 2
    
        If lbInteligencia.Caption > 10 Then
            lbInteligencia.Caption = lbInteligencia.Caption - 1
            Label1.Caption = Label1.Caption + 1
        End If
        
    Case 3
        
        If lbCarisma.Caption > 10 Then
            lbCarisma.Caption = lbCarisma.Caption - 1
            Label1.Caption = Label1.Caption + 1
        End If
        
    Case 4
        
        If lbConstitucion.Caption > 10 Then
            lbConstitucion.Caption = lbConstitucion.Caption - 1
            Label1.Caption = Label1.Caption + 1
        End If
        
    End Select
    
End If

End Sub


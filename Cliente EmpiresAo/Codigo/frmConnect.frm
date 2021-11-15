VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmConnect.frx":000C
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox PassBox 
      Alignment       =   2  'Center
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
      IMEMode         =   3  'DISABLE
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   5460
      Width           =   3570
   End
   Begin VB.TextBox UserBox 
      Alignment       =   2  'Center
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
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   1620
      TabIndex        =   4
      Top             =   4560
      Width           =   3570
   End
   Begin VB.ListBox lst_servers 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1620
      ItemData        =   "frmConnect.frx":0CD6
      Left            =   -600
      List            =   "frmConnect.frx":0CDD
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3120
      TabIndex        =   0
      Text            =   "7666"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   5040
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   1800
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estado: Nulo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      MouseIcon       =   "frmConnect.frx":0CEE
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   705
      Index           =   5
      Left            =   7425
      MouseIcon       =   "frmConnect.frx":19B8
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   5625
      Width           =   3045
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1185
      Index           =   0
      Left            =   8640
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   3090
   End
   Begin VB.Image Image1 
      Height          =   705
      Index           =   4
      Left            =   7425
      MouseIcon       =   "frmConnect.frx":2682
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   4605
      Width           =   3045
   End
   Begin VB.Image Image1 
      Height          =   705
      Index           =   1
      Left            =   7425
      MouseIcon       =   "frmConnect.frx":334C
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2550
      Width           =   3045
   End
   Begin VB.Image Image1 
      Height          =   705
      Index           =   3
      Left            =   7425
      MouseIcon       =   "frmConnect.frx":4016
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   3585
      Width           =   3045
   End
   Begin VB.Image FONDO 
      Height          =   9000
      Left            =   0
      MouseIcon       =   "frmConnect.frx":4CE0
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image imgServEspana 
      Height          =   435
      Left            =   225
      MousePointer    =   99  'Custom
      Top             =   6495
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   -195
      MousePointer    =   99  'Custom
      Top             =   6765
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Image imgGetPass 
      Height          =   495
      Left            =   3600
      MousePointer    =   99  'Custom
      Top             =   8220
      Width           =   4575
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   2
      Left            =   8610
      MousePointer    =   99  'Custom
      Top             =   8025
      Width           =   3120
   End
End
Attribute VB_Name = "frmConnect"
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
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Public Sub CargarLst()

Dim I As Integer

lst_servers.Clear

If ServersRecibidos Then
'    Call WriteVar(App.Path & "\init\sinfo.dat", "INIT", "Cant", UBound(ServersLst))
    'For I = 1 To UBound(ServersLst)
    '    Call WriteVar(App.Path & "\init\sinfo.dat", "S" & I, "Desc", ServersLst(I).desc)
    '    Call WriteVar(App.Path & "\init\sinfo.dat", "S" & I, "IP", ServersLst(I).Ip)
    '    Call WriteVar(App.Path & "\init\sinfo.dat", "S" & I, "PJ", Str(ServersLst(I).Puerto))
    '    Call WriteVar(App.Path & "\init\sinfo.dat", "S" & I, "P2", Str(ServersLst(I).PassRecPort))
    '    lst_servers.AddItem ServersLst(I).Ip & ":" & ServersLst(I).Puerto & " - Desc:" & ServersLst(I).desc
   ' Next I
End If

End Sub

Private Sub Command2_Click()

'frmMain.Inet1.url = "http://www.empiresao.com.ar/server/eao.ip"
RawServersList = frmMain.Inet1.OpenURL


If RawServersList = "" Then
    'Mithrandir - Tu IP abajo :$
    frmConnect.IPTxt.Text = "127.0.0.1"
    frmConnect.PortTxt.Text = "7666"
Else
    ServersRecibidos = True
End If

Call InitServersList(RawServersList)
Call CargarLst

End Sub


Private Sub FONDO_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(1).Tag = "1" Then
    Image1(1).Tag = "0"
    Image1(1).Picture = LoadPicture(App.Path & "\Graficos\cmdConectar.bmp")
End If
If Image1(3).Tag = "1" Then
    Image1(3).Tag = "0"
    Image1(3).Picture = LoadPicture(App.Path & "\Graficos\cmdCrear.bmp")
End If
If Image1(4).Tag = "1" Then
    Image1(4).Tag = "0"
    Image1(4).Picture = LoadPicture(App.Path & "\Graficos\cmdVerificar.bmp")
End If
If Image1(5).Tag = "1" Then
    Image1(5).Tag = "0"
    Image1(5).Picture = LoadPicture(App.Path & "\Graficos\cmdRecuperar.bmp")
End If
End Sub

Private Sub Form_Activate()
'On Error Resume Next

If ServersRecibidos Then
    If CurServer <> 0 Then
        IPTxt = ServerEAO
        PortTxt = PuertoEAO
    Else
        IPTxt = IPdelServidor
        PortTxt = PuertoDelServidor
    End If
    
    Call CargarLst
Else
    lst_servers.Clear
End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Argentum Online.", 0, 128, 128, 1, 0, 1
        
        Call SaveGameini
        frmConnect.MousePointer = 99
        frmMain.MousePointer = 99
        prgRun = False
        
        AddtoRichTextBox frmCargando.status, "Liberando recursos...", 0, 128, 128, 1, 0, 1
        frmCargando.Refresh
        LiberarObjetosDX
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 128, 128, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Argentum Online!!", 0, 128, 128, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
End If

Select Case KeyCode
            Case vbKeyReturn:
                #If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
#Else
        If frmMain.Winsock1.State <> sckClosed Then _
            frmMain.Winsock1.Close
#End If
      '  If frmConnect.MousePointer = 99 Then
      '      Exit Sub
     '   End If
        
        
        'update user info
        Username = UserBox.Text
        Dim aux As String
        aux = PassBox.Text
#If SeguridadAlkon Then
        UserPassword = md5.GetMD5String(aux)
        Call md5.MD5Reset
#Else
        UserPassword = aux
#End If
        If CheckUserData(False) = True Then
            'SendNewChar = False
            'EstadoLogin = Normal
            EstadoLogin = loginaccount
            Me.MousePointer = 99
#If UsarWrench = 1 Then
            frmMain.Socket1.HostAddress = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
#Else
            'If frmMain.Winsock1.State <> sckClosed Then _
               ' frmMain.Winsock1.Close
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        End If
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
'If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    'Port
 '   PortTxt.Visible = True
    'Label4.Visible = True
    
    'Server IP
  '  PortTxt.Text = "7666"
  '  IPTxt.Text = "192.168.0.2"
  '  IPTxt.Visible = True
    'Label5.Visible = True
    
  '  KeyCode = 0
  '  Exit Sub
'End If



End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]
    
 Dim J
 For Each J In Image1()
    J.Tag = "0"
 Next
 PortTxt.Text = Config_Inicio.Puerto
 
IntervaloPaso = INTERVALOWALK
 
 FONDO.Picture = LoadPicture(App.Path & "\Graficos\Login.bmp")
 
Image1(1).Picture = LoadPicture(App.Path & "\Graficos\cmdConectar.bmp")
Image1(3).Picture = LoadPicture(App.Path & "\Graficos\cmdCrear.bmp")
Image1(4).Picture = LoadPicture(App.Path & "\Graficos\cmdVerificar.bmp")
Image1(5).Picture = LoadPicture(App.Path & "\Graficos\cmdRecuperar.bmp")

 '[CODE]:MatuX
 '
 '  El código para mostrar la versión se genera acá para
 ' evitar que por X razones luego desaparezca, como suele
 ' pasar a veces :)
    version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
 '[END]'

'Call setWindowTransparent(rchTxtNews.hWnd)

End Sub



Private Sub Image1_Click(index As Integer)

If ServersRecibidos Then
    If Not IsIp(IPTxt) And CurServer <> 0 Then
        If MsgBox("Atencion, está intentando conectarse a un servidor no oficial, NoLand Studios no se hace responsable de los posibles problemas que estos servidores presenten. ¿Desea continuar?", vbYesNo) = vbNo Then
            If CurServer <> 0 Then
                IPTxt = ServersLst(CurServer).IP
                PortTxt = ServersLst(CurServer).Puerto
            Else
                IPTxt = IPdelServidor
                PortTxt = PuertoDelServidor
            End If
            Exit Sub
        End If
    End If
End If
CurServer = 0
IPdelServidor = IPTxt
PuertoDelServidor = PortTxt


Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
        
        If Musica Then
            Call Audio.PlayMIDI("7.mid")
        End If
        
        
        
        'frmCrearPersonaje.Show vbModal
        EstadoLogin = Dados
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
        End If
        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
#Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
        End If
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        Me.MousePointer = 99

        
    Case 1
    
  #If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
#Else
        If frmMain.Winsock1.State <> sckClosed Then _
            frmMain.Winsock1.Close
#End If
      '  If frmConnect.MousePointer = 99 Then
      '      Exit Sub
     '   End If
        
        
        'update user info
        Username = UserBox.Text
        Dim aux As String
        aux = PassBox.Text
#If SeguridadAlkon Then
        UserPassword = md5.GetMD5String(aux)
        Call md5.MD5Reset
#Else
        UserPassword = aux
#End If
        If CheckUserData(False) = True Then
            'SendNewChar = False
            'EstadoLogin = Normal
            EstadoLogin = loginaccount
            Me.MousePointer = 99
#If UsarWrench = 1 Then
            frmMain.Socket1.HostAddress = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
#Else
            'If frmMain.Winsock1.State <> sckClosed Then _
               ' frmMain.Winsock1.Close
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        End If
        
        
    Case 2
        On Error GoTo errH
        Call Shell(App.Path & "\RECUPERAR.EXE", vbNormalFocus)
        
    Case 3
        
        EstadoLogin = CrearAccount
        
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
        End If
        
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
        
        Me.MousePointer = 99
        
    Case 4
    
        EstadoLogin = verificaraccount
        
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
        End If
    
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
        
        Me.MousePointer = 99
        
    Case 5
        
        MsgBox "Opción inhabilitada temporalmente.", vbCritical
        
        
        'EstadoLogin = RecuperarAccount
        
        'If frmMain.Winsock1.State <> sckClosed Then
            'frmMain.Winsock1.Close
        'End If
    
       'frmMain.Winsock1.Connect CurServerIp, CurServerPort
        
        'Me.MousePointer = 99

End Select
Exit Sub

errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub Image1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case index
    Case 1
        If Image1(index).Tag = "0" Then
            Call Audio.PlayWave(SND_OVER)
            Image1(index).Tag = "1"
            Image1(index).Picture = LoadPicture(App.Path & "\Graficos\cmdConectara.bmp")
        End If
    Case 3
        If Image1(index).Tag = "0" Then
            Call Audio.PlayWave(SND_OVER)
            Image1(index).Tag = "1"
            Image1(index).Picture = LoadPicture(App.Path & "\Graficos\cmdCreara.bmp")
        End If
    Case 4
        If Image1(index).Tag = "0" Then
            Call Audio.PlayWave(SND_OVER)
            Image1(index).Tag = "1"
            Image1(index).Picture = LoadPicture(App.Path & "\Graficos\cmdVerificara.bmp")
        End If
    Case 5
        If Image1(index).Tag = "0" Then
            Call Audio.PlayWave(SND_OVER)
            Image1(index).Tag = "1"
            Image1(index).Picture = LoadPicture(App.Path & "\Graficos\cmdRecuperara.bmp")
        End If
End Select
End Sub

Private Sub imgGetPass_Click()
On Error GoTo errH

    Call Audio.PlayWave(SND_CLICK)
    Call Shell(App.Path & "\RECUPERAR.EXE", vbNormalFocus)
    'Call frmRecuperar.Show(vbModal, frmConnect)
    Exit Sub
errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub imgServArgentina_Click()
    Call Audio.PlayWave(SND_CLICK)
    IPTxt.Text = IPdelServidor
    PortTxt.Text = PuertoDelServidor
End Sub

Private Sub imgServEspana_Click()
    Call Audio.PlayWave(SND_CLICK)
    'Mithrandir
    IPTxt.Text = "127.0.0.1"
    PortTxt.Text = "7666"
End Sub



Private Sub lblStatus_Click()

Call ConnectMod

End Sub

Private Sub lst_servers_Click()
If ServersRecibidos Then
    CurServer = lst_servers.listIndex + 1
    IPTxt = ServersLst(CurServer).IP
    PortTxt = ServersLst(CurServer).Puerto
End If

End Sub


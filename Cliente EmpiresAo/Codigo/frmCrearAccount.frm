VERSION 5.00
Begin VB.Form frmCrearAccount 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCrearAccount.frx":0000
   MousePointer    =   99  'Custom
   PaletteMode     =   2  'Custom
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Mail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   295
      Left            =   6480
      TabIndex        =   4
      Top             =   5295
      Width           =   4730
   End
   Begin VB.TextBox Pass2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   295
      IMEMode         =   3  'DISABLE
      Left            =   6480
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4350
      Width           =   4730
   End
   Begin VB.TextBox Pass1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   295
      IMEMode         =   3  'DISABLE
      Left            =   6480
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3280
      Width           =   4730
   End
   Begin VB.TextBox AccountName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   295
      Left            =   6480
      TabIndex        =   1
      Top             =   2330
      Width           =   4730
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   3615
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmCrearAccount.frx":0CCA
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   6900
      MouseIcon       =   "frmCrearAccount.frx":2449
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   6080
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   930
      MouseIcon       =   "frmCrearAccount.frx":3113
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   6070
      Width           =   345
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   11280
      MouseIcon       =   "frmCrearAccount.frx":3DDD
      MousePointer    =   99  'Custom
      Picture         =   "frmCrearAccount.frx":4AA7
      Top             =   240
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   3900
      MouseIcon       =   "frmCrearAccount.frx":4FD5
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   6690
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   480
      MouseIcon       =   "frmCrearAccount.frx":5C9F
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   1935
   End
End
Attribute VB_Name = "frmCrearAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NombredeCuenta As String

Private Sub Form_Load()
frmCrearAccount.Picture = LoadPicture(App.Path & "\Graficos\bckCrearCuenta.bmp")
frmCrearAccount.Image3.Picture = LoadPicture(App.Path & "\Graficos\cmdCrearCuenta.bmp")
'AccountName.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image3.Tag = "1" Then
    Image3.Tag = "0"
    Image3.Picture = LoadPicture(App.Path & "\Graficos\cmdCrearCuenta.bmp")
End If

End Sub

Private Sub TextBox1_Change(index As Integer)

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image3_Click()

If frmMain.Winsock1.State <> sckConnected Then
    MsgBox "Error: se ha perdido conexión con el servidor.", vbCritical
    Unload Me
Else
    If AccountName.Text = "" Or Pass1.Text = "" Or Pass2.Text = "" Or Mail.Text = "" Then
    
    MsgBox "Debe completar todos los casilleros para proseguir con la registración."
    Exit Sub
    
    Else
        If Pass1.Text <> Pass2.Text Then
            MsgBox "La contraseña de confirmación no coincide con la original."
            Exit Sub
        Else
            cuentaNom = AccountName.Text
            cuentaPass = Pass1.Text
            cuentaEmail = Mail.Text
            
            If NombreAccountCheck() Then
            
                If Image5.Tag = "1" Then
                    If Image6.Tag = "1" Then Call GoWeb("http://www.empiresao.com.ar/foro/")
                    EstadoLogin = CrearAccount
                    Call Login(ValidarLoginMSG(CInt(bRK)))
                Else
                    MsgBox "Debes leer y aceptar la reglamentacion vigente en EmpiresAO para crear una cuenta."
                End If
                
            
            End If
            
        End If

    End If
    
End If

End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image3.Tag = "0" Then
    Image3.Tag = "1"
    Image3.Picture = LoadPicture(App.Path & "\Graficos\cmdCrearCuentaa.bmp")
End If
End Sub

Private Sub Image4_Click()

If Musica Then
    Call Audio.PlayMIDI("2.mid")
End If
        
frmMain.Winsock1.Close

frmConnect.FONDO.Picture = LoadPicture(App.Path & "\Graficos\Login.bmp")
Unload Me

End Sub

Private Sub Image5_Click()

If Image5.Tag = "0" Then
    Image5.Tag = "1"
    Image5.Picture = LoadPicture(App.Path & "\Graficos\Acepto1.bmp")
    Exit Sub
End If

If Image5.Tag = "1" Then
    Image5.Tag = "0"
    Image5.Picture = LoadPicture()
End If

End Sub

Private Sub Image6_Click()

If Image6.Tag = "0" Then
    Image6.Tag = "1"
    Image6.Picture = LoadPicture(App.Path & "\Graficos\Acepto2.bmp")
    Exit Sub
End If

If Image6.Tag = "1" Then
    Image6.Tag = "0"
    Image6.Picture = LoadPicture()
End If

End Sub




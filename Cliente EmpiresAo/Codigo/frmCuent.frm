VERSION 5.00
Begin VB.Form frmCuent 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCuent.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   5690
      MouseIcon       =   "frmCuent.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":1994
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   35
      Top             =   2500
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   5690
      MouseIcon       =   "frmCuent.frx":1C2F
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":28F9
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   34
      Top             =   2500
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   5690
      MouseIcon       =   "frmCuent.frx":2B94
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":385E
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   33
      Top             =   2500
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   5690
      MouseIcon       =   "frmCuent.frx":3AF9
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":47C3
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   2500
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   5690
      MouseIcon       =   "frmCuent.frx":4A5E
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":5728
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   2500
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   5690
      MouseIcon       =   "frmCuent.frx":59C3
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":668D
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   2500
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   5690
      MouseIcon       =   "frmCuent.frx":6928
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":75F2
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   2500
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   5690
      MouseIcon       =   "frmCuent.frx":788D
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":8557
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   2500
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   7
      Left            =   4990
      TabIndex        =   32
      Top             =   4890
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   6
      Left            =   4990
      TabIndex        =   31
      Top             =   4890
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   4990
      TabIndex        =   30
      Top             =   4890
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   4990
      TabIndex        =   29
      Top             =   4530
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   4990
      TabIndex        =   28
      Top             =   4530
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4990
      TabIndex        =   27
      Top             =   4530
      Width           =   1455
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "hhhh"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   4990
      TabIndex        =   26
      Top             =   4160
      Width           =   1815
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "hhhh"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   4990
      TabIndex        =   25
      Top             =   4160
      Width           =   1815
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "hhhh"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   4990
      TabIndex        =   24
      Top             =   4160
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   11280
      MouseIcon       =   "frmCuent.frx":87F2
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   705
      Left            =   4020
      MouseIcon       =   "frmCuent.frx":94BC
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   6040
      Width           =   4155
   End
   Begin VB.Image Image4 
      Height          =   705
      Left            =   4020
      MouseIcon       =   "frmCuent.frx":A186
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   5260
      Width           =   4155
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7885
      TabIndex        =   23
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7810
      TabIndex        =   22
      Top             =   3720
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   930
      Index           =   1
      Left            =   4810
      MouseIcon       =   "frmCuent.frx":AE50
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2640
      Width           =   555
   End
   Begin VB.Image Image3 
      Height          =   945
      Index           =   0
      Left            =   6780
      MouseIcon       =   "frmCuent.frx":BB1A
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7320
      TabIndex        =   21
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Erikkunete"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5210
      TabIndex        =   20
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4990
      TabIndex        =   19
      Top             =   4530
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4990
      TabIndex        =   18
      Top             =   4530
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4990
      TabIndex        =   17
      Top             =   4530
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4990
      TabIndex        =   16
      Top             =   4530
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   4990
      TabIndex        =   15
      Top             =   4890
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   4990
      TabIndex        =   14
      Top             =   4890
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   4990
      TabIndex        =   13
      Top             =   4890
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   4990
      TabIndex        =   12
      Top             =   4890
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   4990
      TabIndex        =   11
      Top             =   4890
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4990
      TabIndex        =   10
      Top             =   4530
      Width           =   1455
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "hhhh"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   4990
      TabIndex        =   9
      Top             =   4160
      Width           =   1815
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "gggggg"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   4990
      TabIndex        =   8
      Top             =   4160
      Width           =   1815
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "hhh"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   4990
      TabIndex        =   7
      Top             =   4160
      Width           =   1815
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "hhh"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   4990
      TabIndex        =   6
      Top             =   4160
      Width           =   1815
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "hhhh"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   4990
      TabIndex        =   5
      Top             =   4160
      Width           =   1815
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmCuent.PJ(7).Visible = True
End Sub

Private Sub Form_Load()
Dim I As Integer
'Call dibujapj(BackBufferSurface, BodyData(57).Walk(3), 12, 15, 1)
'Call dibujapj(BackBufferSurface, HeadData(1).Head(3), 16, 6, 1)
'Call dibujapj(BackBufferSurface, ShieldAnimData(1).ShieldWalk(3), 12, 12, 1)
'Call dibujapj(BackBufferSurface, CascoAnimData(3).Head(3), 16, 6, 1)
'Call dibujapj(BackBufferSurface, WeaponAnimData(1).WeaponWalk(3), 12, 18, 1)
For I = 0 To 7
    frmCuent.nombre(I).Visible = False
    frmCuent.Label1(I).Visible = False
    frmCuent.Label2(I).Visible = False
    frmCuent.PJ(I).Visible = False
    frmCuent.nombre(I).ForeColor = vbBlack
    frmCuent.Label1(I).FontBold = True
    frmCuent.Label2(I).FontBold = True
    frmCuent.Label1(I).FontSize = 8
    frmCuent.Label2(I).FontSize = 8
    frmCuent.nombre(I).Caption = "Nada"
    frmCuent.Label1(I).Caption = ""
    frmCuent.Label2(I).Caption = ""
    If I = 0 Then
        frmCuent.nombre(I).Visible = True
        frmCuent.PJ(I).Visible = True
        frmCuent.Label1(I).Visible = True
        frmCuent.Label2(I).Visible = True
    End If
Next I

'Label3.Caption = Username
frmCuent.Picture = LoadPicture(App.Path & "\Graficos\Dentrocuenta.bmp")
frmCuent.Image4.Picture = LoadPicture(App.Path & "\Graficos\BotonConectar.bmp")
frmCuent.Image5.Picture = LoadPicture(App.Path & "\Graficos\BotonCrearPersonaje.bmp")
frmCuent.Image1.Picture = LoadPicture(App.Path & "\Graficos\Exit.bmp")
frmCuent.Image3(1).Picture = LoadPicture(App.Path & "\Graficos\FlechaIzquierda.bmp")
frmCuent.Image3(0).Picture = LoadPicture(App.Path & "\Graficos\FlechaDerecha.bmp")
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)
Username = frmCuent.nombre(VerVisible).Caption
UserPassword = "i"


        'If frmConnect.MousePointer = 11 Then
            'Exit Sub
        'End If
        
        
        'update user info
        Username = frmCuent.nombre(VerVisible).Caption
        Dim aux As String
        aux = "i"

        UserPassword = aux

        If CheckUserData(False) = True Then
            'SendNewChar = False
            'EstadoLogin = Normal
            EstadoLogin = Normal
            Me.MousePointer = 11
            
            EstadoLogin = Normal
            Call Login(ValidarLoginMSG(CInt(bRK)))

        End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image4.Tag = "1" Then
    Image4.Tag = "0"
    Image4.Picture = LoadPicture(App.Path & "\Graficos\BotonConectar.bmp")
End If
If Image5.Tag = "1" Then
    Image5.Tag = "0"
    Image5.Picture = LoadPicture(App.Path & "\Graficos\BotonCrearPersonaje.bmp")
End If
If Image3(0).Tag = "1" Then
    Image3(0).Tag = "0"
    Image3(0).Picture = LoadPicture(App.Path & "\Graficos\FlechaDerecha.bmp")
End If
If Image3(1).Tag = "1" Then
    Image3(1).Tag = "0"
    Image3(1).Picture = LoadPicture(App.Path & "\Graficos\FlechaIzquierda.bmp")
End If
End Sub

Private Sub Image1_Click()
If Musica Then
    Call Audio.PlayMIDI("2.mid")
End If
        
frmMain.Winsock1.Close

frmConnect.FONDO.Picture = LoadPicture(App.Path & "\Graficos\Login.bmp")
Unload Me

End Sub

Private Sub Image3_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
    
    If Label4.Caption + 1 <= Label6.Caption Then
        frmCuent.nombre(Label4.Caption - 1).Visible = False
        frmCuent.Label1(Label4.Caption - 1).Visible = False
        frmCuent.Label2(Label4.Caption - 1).Visible = False
        frmCuent.PJ(Label4.Caption - 1).Visible = False
        Label4.Caption = Label4.Caption + 1
        frmCuent.nombre(Label4.Caption - 1).Visible = True
        frmCuent.Label1(Label4.Caption - 1).Visible = True
        frmCuent.Label2(Label4.Caption - 1).Visible = True
        frmCuent.PJ(Label4.Caption - 1).Visible = True
    End If
    
    Case 1
    
    If Label4.Caption - 1 > 0 Then
        frmCuent.nombre(Label4.Caption - 1).Visible = False
        frmCuent.Label1(Label4.Caption - 1).Visible = False
        frmCuent.Label2(Label4.Caption - 1).Visible = False
        frmCuent.PJ(Label4.Caption - 1).Visible = False
        Label4.Caption = Label4.Caption - 1
        frmCuent.nombre(Label4.Caption - 1).Visible = True
        frmCuent.Label1(Label4.Caption - 1).Visible = True
        frmCuent.Label2(Label4.Caption - 1).Visible = True
        frmCuent.PJ(Label4.Caption - 1).Visible = True
    End If
        
End Select

End Sub

Public Function VerVisible() As Integer

Dim I As Integer

For I = 0 To 7
    If frmCuent.nombre(I).Visible = True Then VerVisible = I
Next I

End Function

Private Sub Image3_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case index
    Case 0
        If Image3(index).Tag = "0" Then
            Image3(index).Tag = "1"
            Call Audio.PlayWave(SND_OVER)
            Image3(index).Picture = LoadPicture(App.Path & "\Graficos\FlechaDerechaApretada.bmp")
        End If
    Case 1
        If Image3(index).Tag = "0" Then
            Image3(index).Tag = "1"
            Call Audio.PlayWave(SND_OVER)
            Image3(index).Picture = LoadPicture(App.Path & "\Graficos\FlechaIzquierdaApretada.bmp")
        End If
End Select
End Sub

Private Sub Image4_Click()
Call Audio.PlayWave(SND_CLICK)
Username = frmCuent.nombre(VerVisible).Caption
UserPassword = "i"


        'If frmConnect.MousePointer = 11 Then
            'Exit Sub
        'End If
        
        
        'update user info
        Username = frmCuent.nombre(VerVisible).Caption
        Dim aux As String
        aux = "i"

        UserPassword = aux

        If CheckUserData(False) = True Then
            'SendNewChar = False
            'EstadoLogin = Normal
            EstadoLogin = Normal
            Me.MousePointer = 11
            
            EstadoLogin = Normal
            Call Login(ValidarLoginMSG(CInt(bRK)))

        End If
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image4.Tag = "0" Then
    Image4.Tag = "1"
    Call Audio.PlayWave(SND_OVER)
    Image4.Picture = LoadPicture(App.Path & "\Graficos\BotonConectarApretado.bmp")
End If
End Sub

Private Sub Image5_Click()
Call Audio.PlayWave(SND_CLICK)

If "8" <= Label6.Caption Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

If Musica Then
         Call Audio.PlayMIDI("7.mid")
    End If
    
    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    Me.MousePointer = 11
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image5.Tag = "0" Then
    Image5.Tag = "1"
    Call Audio.PlayWave(SND_OVER)
    Image5.Picture = LoadPicture(App.Path & "\Graficos\BotonCrearPersonajeApretado.bmp")
End If
End Sub

Private Sub PJ_Click(index As Integer)

PJClickeado = frmCuent.nombre(index).Caption

End Sub

Private Sub PJ_DblClick(index As Integer)

Username = PJClickeado
UserPassword = "i"

If Username = "" Then Exit Sub


        'If frmConnect.MousePointer = 11 Then
            'Exit Sub
        'End If
        
        
        'update user info
        Username = PJClickeado
        Dim aux As String
        aux = "i"

        UserPassword = aux

        If CheckUserData(False) = True Then
            'SendNewChar = False
            'EstadoLogin = Normal
            EstadoLogin = Normal
            Me.MousePointer = 11
            
            EstadoLogin = Normal
            Call Login(ValidarLoginMSG(CInt(bRK)))

        End If
        
End Sub

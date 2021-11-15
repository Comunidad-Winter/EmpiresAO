VERSION 5.00
Begin VB.Form frmVerificarAccount 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmVerificarAccount.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5190
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextBox2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   1880
      TabIndex        =   1
      Top             =   2000
      Width           =   3040
   End
   Begin VB.TextBox TextBox1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   1880
      TabIndex        =   0
      Top             =   1595
      Width           =   3040
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   5280
      MouseIcon       =   "frmVerificarAccount.frx":0CCA
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   1750
      MouseIcon       =   "frmVerificarAccount.frx":1994
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   2100
   End
End
Attribute VB_Name = "frmVerificarAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
frmVerificarAccount.Picture = LoadPicture(App.Path & "\Graficos\iVerificar.bmp")
Image2.Picture = LoadPicture(App.Path & "\Graficos\Exit.bmp")
End Sub

Private Sub Image1_Click()
verifUser = TextBox1.Text
verifCode = TextBox2.Text

If frmMain.Winsock1.State <> sckConnected Then
    MsgBox "Error: se ha perdido conexión con el servidor.", vbCritical
    Unload Me
Else
    EstadoLogin = verificaraccount
    Call Login(ValidarLoginMSG(CInt(bRK)))
End If
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

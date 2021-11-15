VERSION 5.00
Begin VB.Form frmRecovery 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
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
      Index           =   1
      Left            =   1850
      TabIndex        =   1
      Top             =   2000
      Width           =   3070
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
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
      Index           =   0
      Left            =   1850
      TabIndex        =   0
      Top             =   1590
      Width           =   3070
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   5040
      MouseIcon       =   "frmRecovery.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmRecovery.frx":0CCA
      Top             =   240
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1780
      MouseIcon       =   "frmRecovery.frx":11F8
      MousePointer    =   99  'Custom
      Top             =   2440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   5265
      Left            =   0
      MouseIcon       =   "frmRecovery.frx":1EC2
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   5820
   End
End
Attribute VB_Name = "frmRecovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Image1.Picture = LoadPicture(App.Path & "\Graficos\Recuperar.bmp")

End Sub

Private Sub Image2_Click()

If CheckMailString(Text1(1).Text) Then
    If Text1(0).Text <> "" Then
        If frmMain.Winsock1.State <> sckConnected Then
            MsgBox "Error: se ha perdido conexión con el servidor.", vbCritical
            Unload Me
        Else
            recoveryAccount = Text1(0).Text
            recoveryMail = Text1(1).Text
            EstadoLogin = RecuperarAccount
            Call Login(ValidarLoginMSG(CInt(bRK)))
        End If
    Else
        MsgBox "Debes poner el nombre de la cuenta a recuperar."
    End If
Else
    MsgBox "Dirección de mail incorrecta."
End If

End Sub

Private Sub Image3_Click()
Unload Me
End Sub

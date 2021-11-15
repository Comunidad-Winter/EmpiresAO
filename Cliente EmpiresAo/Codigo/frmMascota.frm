VERSION 5.00
Begin VB.Form frmMascota 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMascota.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5235
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MascotaView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   840
      MouseIcon       =   "frmMascota.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "frmMascota.frx":1994
      ScaleHeight     =   1095
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   5040
      MouseIcon       =   "frmMascota.frx":1C2F
      MousePointer    =   99  'Custom
      Picture         =   "frmMascota.frx":28F9
      Top             =   240
      Width           =   420
   End
   Begin VB.Label LVL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Exp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200/300"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   1660
      Width           =   1695
   End
   Begin VB.Label Defensa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   3320
      Width           =   1095
   End
   Begin VB.Label Ataque 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200/210"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Vida 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "25/500"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4090
      TabIndex        =   1
      Top             =   2250
      Width           =   1095
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pepelepe"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3580
      TabIndex        =   0
      Top             =   1160
      Width           =   1575
   End
End
Attribute VB_Name = "frmMascota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmMascota.Picture = LoadPicture(App.Path & "\Graficos\PanelMascota.bmp")
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

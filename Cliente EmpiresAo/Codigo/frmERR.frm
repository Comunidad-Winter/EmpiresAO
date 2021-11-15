VERSION 5.00
Begin VB.Form frmERR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmERR.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmERR.frx":0CCA
   ScaleHeight     =   1635
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   375
      Left            =   840
      MouseIcon       =   "frmERR.frx":D90F
      MousePointer    =   99  'Custom
      Top             =   1090
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmERR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.MousePointer = frmMain.MousePointer
End Sub

Private Sub Image1_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
End Sub

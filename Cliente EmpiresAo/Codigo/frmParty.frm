VERSION 5.00
Begin VB.Form frmParty 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   6195
      Left            =   0
      Picture         =   "frmParty.frx":0000
      Top             =   0
      Width           =   4635
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\Graficos\Party.bmp")
End Sub

Private Sub Image1_Click()

End Sub

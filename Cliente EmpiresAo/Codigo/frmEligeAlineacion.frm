VERSION 5.00
Begin VB.Form frmEligeAlineacion 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alineaci�n del clan"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmEligeAlineacion.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5625
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   315
      Left            =   2520
      MouseIcon       =   "frmEligeAlineacion.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEligeAlineacion.frx":1994
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   4
      Left            =   120
      MouseIcon       =   "frmEligeAlineacion.frx":1A69
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4260
      Width           =   6465
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEligeAlineacion.frx":2733
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   3
      Left            =   120
      MouseIcon       =   "frmEligeAlineacion.frx":280F
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3360
      Width           =   6465
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEligeAlineacion.frx":34D9
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   2
      Left            =   120
      MouseIcon       =   "frmEligeAlineacion.frx":3585
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2415
      Width           =   6465
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEligeAlineacion.frx":424F
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmEligeAlineacion.frx":4318
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1470
      Width           =   6465
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEligeAlineacion.frx":4FE2
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   0
      Left            =   120
      MouseIcon       =   "frmEligeAlineacion.frx":5108
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   390
      Width           =   6465
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineaci�n del mal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   4035
      Width           =   1680
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineaci�n criminal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   3135
      Width           =   1680
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineaci�n neutral"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2190
      Width           =   1635
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineaci�n legal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1245
      Width           =   1455
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineaci�n Real"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   1455
   End
End
Attribute VB_Name = "frmEligeAlineacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

'odio programar sin tiempo (c) el oso

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

    For i = 0 To 4
        lblDescripcion(i).BorderStyle = 0
        lblDescripcion(i).BackStyle = 0
    Next i
    
End Sub

Private Sub lblDescripcion_Click(index As Integer)
Dim s As String
    
    Select Case index
        Case 0
            s = "armada"
        Case 1
            s = "legal"
        Case 2
            s = "neutro"
        Case 3
            s = "criminal"
        Case 4
            s = "mal"
    End Select
    
    s = "/fundarclan " & s
    Call SendData(s)
    Unload Me
End Sub

Private Sub lblDescripcion_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDescripcion(index).BorderStyle = 1
    lblDescripcion(index).BackStyle = 1
    Select Case index
        Case 0
            lblDescripcion(index).BackColor = &H400000
        Case 1
            lblDescripcion(index).BackColor = &H800000
        Case 2
            lblDescripcion(index).BackColor = 4194368
        Case 3
            lblDescripcion(index).BackColor = &H80&
        Case 4
            lblDescripcion(index).BackColor = &H40&
    End Select
End Sub


Private Sub lblNombre_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

    For i = 0 To 4
        lblDescripcion(i).BorderStyle = 0
        lblDescripcion(i).BackStyle = 0
    Next i
    

End Sub

VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmOpciones.frx":0152
   MousePointer    =   99  'Custom
   ScaleHeight     =   3660
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Otras opciones"
      Height          =   855
      Left            =   360
      TabIndex        =   9
      Top             =   2280
      Width           =   4215
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000A&
         Caption         =   "Activar ""Mostrar lo que estoy jugando"" (MSN)"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Diálogos de clan"
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   375
      TabIndex        =   4
      Top             =   1425
      Width           =   4230
      Begin VB.TextBox txtCantMensajes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2925
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "5"
         Top             =   315
         Width           =   450
      End
      Begin VB.OptionButton optPantalla 
         BackColor       =   &H8000000A&
         Caption         =   "En pantalla,"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1770
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optConsola 
         BackColor       =   &H8000000A&
         Caption         =   "En consola"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   105
         TabIndex        =   5
         Top             =   315
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mensajes"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3480
         TabIndex        =   8
         Top             =   345
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   1080
      MouseIcon       =   "frmOpciones.frx":0E1C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3240
      Width           =   2790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sonidos Activados"
      Height          =   345
      Index           =   1
      Left            =   1080
      MouseIcon       =   "frmOpciones.frx":1AE6
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   960
      Width           =   2790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Musica Activada"
      Height          =   345
      Index           =   0
      Left            =   1080
      MouseIcon       =   "frmOpciones.frx":27B0
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   540
      Width           =   2790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
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

Private Sub Command1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
        If Musica Then
            Musica = False
            Command1(0).Caption = "Musica Desactivada"
            Audio.StopMidi
        Else
            Musica = True
            Command1(0).Caption = "Musica Activada"
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
        End If
    Case 1
    
        If Sound Then
            Sound = False
            Command1(1).Caption = "Sonidos Desactivados"
            Call Audio.StopWave
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        Else
            Sound = True
            Command1(1).Caption = "Sonidos Activados"
        End If
End Select
End Sub

Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Check1_Click()

If Check1.value = 1 Then
    Call WriteVar(App.Path & "\Init\opts.eao", "INIT", "Active", "1")
    MSNshow = 1
Else
    Call WriteVar(App.Path & "\Init\opts.eao", "INIT", "Active", "0")
    MSNshow = 0
End If

End Sub

Private Sub Form_Load()
    If Musica Then
        Command1(0).Caption = "Musica Activada"
    Else
        Command1(0).Caption = "Musica Desactivada"
    End If
    
    If Sound Then
        Command1(1).Caption = "Sonidos Activados"
    Else
        Command1(1).Caption = "Sonidos Desactivados"
    End If
    
    If MSNshow = 1 Then
        Check1.value = 1
    Else
        Check1.value = 0
    End If
    
End Sub

Private Sub optConsola_Click()
    DialogosClanes.Activo = False
End Sub

Private Sub optPantalla_Click()
    DialogosClanes.Activo = True
End Sub

Private Sub txtCantMensajes_LostFocus()
    txtCantMensajes.Text = Trim$(txtCantMensajes.Text)
    If IsNumeric(txtCantMensajes.Text) Then
        DialogosClanes.CantidadDialogos = Trim$(txtCantMensajes.Text)
    Else
        txtCantMensajes.Text = 5
    End If
End Sub

VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1320
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   2190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCantidad.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   1320
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TextBox1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   240
      Left            =   305
      TabIndex        =   0
      Text            =   "0"
      Top             =   510
      Width           =   1470
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1950
      MouseIcon       =   "frmCantidad.frx":0CCA
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   1200
      MouseIcon       =   "frmCantidad.frx":1994
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmCantidad.frx":265E
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmCantidad"
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

Private Sub Form_Deactivate()
'Unload Me
End Sub

Private Sub Form_Load()
frmCantidad.Picture = LoadPicture(App.Path & "\Graficos\Tirar.bmp")
End Sub

Private Sub Image1_Click()
frmCantidad.Visible = False
SendData "TI" & Inventario.SelectedItem & "," & frmCantidad.TextBox1.Text
frmCantidad.TextBox1.Text = "0"
End Sub

Private Sub Image2_Click()
frmCantidad.Visible = False
If Inventario.SelectedItem <> FLAGORO Then
    SendData "TI" & Inventario.SelectedItem & "," & Inventario.Amount(Inventario.SelectedItem)
Else
    SendData "TI" & Inventario.SelectedItem & "," & UserGLD
End If

frmCantidad.TextBox1.Text = "0"
End Sub

Private Sub TextBox1_Change()
On Error GoTo ErrHandler
    If Val(TextBox1.Text) < 0 Then
        TextBox1.Text = MAX_INVENTORY_OBJS
    End If
    
    If Val(TextBox1.Text) > MAX_INVENTORY_OBJS Then
        If Inventario.SelectedItem <> FLAGORO Or Val(TextBox1.Text) > UserGLD Then
            TextBox1.Text = "1"
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    TextBox1.Text = "1"
End Sub



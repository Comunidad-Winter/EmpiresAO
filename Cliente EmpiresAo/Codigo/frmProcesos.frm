VERSION 5.00
Begin VB.Form frmProcesos 
   Caption         =   "EmpiresAO 2 - Procesos"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6975
   Icon            =   "frmProcesos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "frmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

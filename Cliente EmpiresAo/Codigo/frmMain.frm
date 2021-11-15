VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "EmpiresAO 2.0"
   ClientHeight    =   9000
   ClientLeft      =   345
   ClientTop       =   315
   ClientWidth     =   12000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":1CCA
   MousePointer    =   99  'Custom
   Picture         =   "frmMain.frx":2994
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6960
      Top             =   2520
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer checkSM 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   1440
      Top             =   4320
   End
   Begin MSWinsockLib.Winsock sckMod 
      Left            =   6360
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2400
      Top             =   2520
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1920
      Left            =   9060
      ScaleHeight     =   128
      ScaleMode       =   0  'User
      ScaleWidth      =   160
      TabIndex        =   23
      Top             =   2640
      Width           =   2400
   End
   Begin VB.CommandButton DespInv 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   11040
      MouseIcon       =   "frmMain.frx":1622D6
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton DespInv 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   10800
      MouseIcon       =   "frmMain.frx":162428
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2730
      Left            =   8940
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   6000
      Top             =   2520
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000C000&
      Height          =   210
      Left            =   165
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1920
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   5520
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6480
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   3600
      Top             =   2520
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3120
      Top             =   2520
   End
   Begin VB.Timer Trabajo 
      Enabled         =   0   'False
      Left            =   4080
      Top             =   2520
   End
   Begin VB.Timer FPS 
      Interval        =   1000
      Left            =   5040
      Top             =   2520
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7440
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.Timer Attack 
      Enabled         =   0   'False
      Left            =   4560
      Top             =   2520
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   165
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1920
      Visible         =   0   'False
      Width           =   8160
   End
   Begin RichTextLib.RichTextBox RecTxt 
      CausesValidation=   0   'False
      Height          =   1410
      Left            =   135
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   480
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   2487
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":16257A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image11 
      Height          =   375
      Left            =   1290
      MouseIcon       =   "frmMain.frx":1625F7
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   150
      Index           =   1
      Left            =   6960
      TabIndex        =   30
      Top             =   8730
      Width           =   120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   150
      Index           =   0
      Left            =   6960
      TabIndex        =   29
      Top             =   8535
      Width           =   120
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmMain.frx":1632C1
      MousePointer    =   99  'Custom
      Top             =   8535
      Width           =   855
   End
   Begin VB.Image Image9 
      Height          =   405
      Index           =   6
      Left            =   10650
      MouseIcon       =   "frmMain.frx":163F8B
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   405
      Index           =   5
      Left            =   10080
      MouseIcon       =   "frmMain.frx":164C55
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   405
      Index           =   4
      Left            =   9525
      MouseIcon       =   "frmMain.frx":16591F
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   405
      Index           =   3
      Left            =   9000
      MouseIcon       =   "frmMain.frx":1665E9
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   405
      Index           =   2
      Left            =   8400
      MouseIcon       =   "frmMain.frx":1672B3
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   405
      Index           =   1
      Left            =   7860
      MouseIcon       =   "frmMain.frx":167F7D
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   405
      Index           =   0
      Left            =   7290
      MouseIcon       =   "frmMain.frx":168C47
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   495
   End
   Begin VB.Label lblAG 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8790
      TabIndex        =   28
      Top             =   8010
      Width           =   1380
   End
   Begin VB.Label lblCOM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8790
      TabIndex        =   27
      Top             =   7725
      Width           =   1380
   End
   Begin VB.Label lblSTA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8790
      TabIndex        =   26
      Top             =   6855
      Width           =   1380
   End
   Begin VB.Label lblMANA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8790
      TabIndex        =   25
      Top             =   7440
      Width           =   1380
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8790
      TabIndex        =   24
      Top             =   7155
      Width           =   1380
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   11160
      MouseIcon       =   "frmMain.frx":169911
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   540
   End
   Begin VB.Image Image7 
      Height          =   390
      Left            =   10470
      MouseIcon       =   "frmMain.frx":16A5DB
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   3885
      MouseIcon       =   "frmMain.frx":16B2A5
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Image AGUAsp 
      Height          =   135
      Left            =   8790
      Picture         =   "frmMain.frx":16BF6F
      Top             =   8025
      Width           =   1380
   End
   Begin VB.Image COMIDAsp 
      Height          =   135
      Left            =   8790
      Picture         =   "frmMain.frx":16C965
      Top             =   7755
      Width           =   1380
   End
   Begin VB.Image Hpshp 
      Height          =   135
      Left            =   8790
      Picture         =   "frmMain.frx":16D35B
      Top             =   7170
      Width           =   1380
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   11205
      MouseIcon       =   "frmMain.frx":16DD51
      MousePointer    =   99  'Custom
      Top             =   45
      Width           =   255
   End
   Begin VB.Image MANShp 
      Height          =   135
      Left            =   8790
      Picture         =   "frmMain.frx":16EA1B
      Top             =   7455
      Width           =   1380
   End
   Begin VB.Image STAShp 
      Height          =   135
      Left            =   8790
      Picture         =   "frmMain.frx":16F411
      Top             =   6885
      Width           =   1380
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   11475
      MouseIcon       =   "frmMain.frx":16FE07
      MousePointer    =   99  'Custom
      Top             =   45
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2400
      MouseIcon       =   "frmMain.frx":170AD1
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label Arma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   135
      Index           =   1
      Left            =   10590
      TabIndex        =   22
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Arma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   135
      Index           =   0
      Left            =   10320
      TabIndex        =   21
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Escudo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   135
      Index           =   1
      Left            =   10590
      TabIndex        =   20
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Escudo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   135
      Index           =   0
      Left            =   10320
      TabIndex        =   19
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Torso 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   135
      Index           =   1
      Left            =   10590
      TabIndex        =   18
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Torso 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   135
      Index           =   0
      Left            =   10320
      TabIndex        =   17
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Cabeza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   135
      Index           =   1
      Left            =   10590
      TabIndex        =   16
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Cabeza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   135
      Index           =   0
      Left            =   10320
      TabIndex        =   15
      Top             =   4680
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   2
      Left            =   10440
      MouseIcon       =   "frmMain.frx":17179B
      MousePointer    =   99  'Custom
      Top             =   6510
      Width           =   1365
   End
   Begin VB.Image Image3 
      Height          =   255
      Index           =   0
      Left            =   9600
      MouseIcon       =   "frmMain.frx":172465
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   195
      Index           =   2
      Left            =   9480
      MouseIcon       =   "frmMain.frx":17312F
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   195
      Index           =   1
      Left            =   9600
      MouseIcon       =   "frmMain.frx":173DF9
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   1
      Left            =   10440
      MouseIcon       =   "frmMain.frx":174AC3
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   0
      Left            =   5325
      MouseIcon       =   "frmMain.frx":17578D
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   1410
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   9720
      TabIndex        =   14
      Top             =   5790
      Width           =   930
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   150
      Left            =   9675
      TabIndex        =   13
      Top             =   1335
      Width           =   525
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp: 0/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   9570
      TabIndex        =   12
      Top             =   1140
      Width           =   705
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Imanol"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8640
      TabIndex        =   11
      Top             =   765
      Width           =   2025
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   255
      Index           =   0
      Left            =   11520
      MouseIcon       =   "frmMain.frx":176457
      MousePointer    =   99  'Custom
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   255
      Index           =   1
      Left            =   11520
      MouseIcon       =   "frmMain.frx":177121
      MousePointer    =   99  'Custom
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image cmdInfo 
      Height          =   315
      Left            =   10200
      MouseIcon       =   "frmMain.frx":177DEB
      MousePointer    =   99  'Custom
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image CmdLanzar 
      Height          =   315
      Left            =   8760
      MouseIcon       =   "frmMain.frx":178AB5
      MousePointer    =   99  'Custom
      Top             =   5640
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   10260
      MouseIcon       =   "frmMain.frx":17977F
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8760
      MouseIcon       =   "frmMain.frx":17A449
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1815
      Width           =   1485
   End
   Begin VB.Label lbCRIATURA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   120
      Left            =   8925
      TabIndex        =   8
      Top             =   2475
      Width           =   30
   End
   Begin VB.Image InvEqu 
      Height          =   4395
      Left            =   8580
      Top             =   1665
      Width           =   3240
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "58"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   10875
      TabIndex        =   4
      Top             =   735
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11415
      MouseIcon       =   "frmMain.frx":17B113
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1230
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image PicAU 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   11520
      Picture         =   "frmMain.frx":17BDDD
      Stretch         =   -1  'True
      Top             =   9480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image PicMH 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   12480
      Picture         =   "frmMain.frx":17D04F
      Stretch         =   -1  'True
      Top             =   9240
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label Coord 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(000,00,00)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   10080
      TabIndex        =   2
      Top             =   9240
      Width           =   1035
   End
   Begin VB.Image PicSeg 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   12600
      Picture         =   "frmMain.frx":17DE61
      Stretch         =   -1  'True
      Top             =   8880
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Shape MainViewShp 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   6240
      Left            =   135
      Top             =   2205
      Width           =   8175
   End
End
Attribute VB_Name = "frmMain"
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
'Foto con el yeti=
Private Declare Sub keybd_event Lib "user32" ( _
ByVal bVk As Byte, _
ByVal bScan As Byte, _
ByVal dwFlags As Long, _
ByVal dwExtraInfo As Long)
Private Const VK_SNAPSHOT = &H2C

'Jeje
Public ActualSecond As Long
Public lastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long

Dim gDSB As DirectSoundBuffer
Dim gD As DSBUFFERDESC
Dim gW As WAVEFORMATEX
Dim gFileName As String
Dim dsE As DirectSoundEnum
Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte

Dim endEvent As Long
Dim PuedeMacrear As Boolean

Implements DirectXEvent

Private Sub cmdMoverHechi_Click(index As Integer)
If hlst.listIndex = -1 Then Exit Sub

Select Case index
Case 0 'subir
    If hlst.listIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.listIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & index + 1 & "," & hlst.listIndex + 1)

Select Case index
Case 0 'subir
    hlst.listIndex = hlst.listIndex - 1
Case 1 'bajar
    hlst.listIndex = hlst.listIndex + 1
End Select

End Sub

Private Sub checkSM_Timer()

'Static Tried As Integer

'Tried = Tried + 1

'If Tried = 1 Then
 '   If sckMod.State <> sckClosed Then sckMod.Close
  '  sckMod.Connect EAOmodsrv, EAOmodport
   ' AddtoRichTextBox frmCargando.status, "Intentando conectar al servidor MOD....", 0, 128, 128, 0, 0, 0
   ' frmConnect.lblStatus.Caption = "Estado: re-conectando..."
'ElseIf Tried > 1 Then
'    checkSM.Enabled = False
 '   sckMod.Close
 '   AddtoRichTextBox frmCargando.status, "No se ha podido conectar al servidor MOD.", 0, 128, 128, 0, 0, 0
  '  frmConnect.lblStatus.Caption = "Estado: servidor offline."
  '  Tried = 0
'End If
    

End Sub

Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

End Sub

Private Sub CreateEvent()
     endEvent = DirectX.CreateEvent(Me)
End Sub

Public Sub ActivarMacroHechizos()
    If Not hlst.Visible Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, False)
        Exit Sub
    End If
    TrainingMacro.Interval = 2788
    TrainingMacro.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, False)
    PicMH.Visible = True
End Sub

Public Sub DesactivarMacroHechizos()
        PicMH.Visible = False
        TrainingMacro.Enabled = False
        SecuenciaMacroHechizos = 0
        Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, False)
End Sub
Public Sub DibujarMH()
PicMH.Visible = True
End Sub

Public Sub DesDibujarMH()
PicMH.Visible = False
End Sub

Public Sub DibujarSeguro()
PicSeg.Visible = True
End Sub

Public Sub DesDibujarSeguro()
PicSeg.Visible = False
End Sub

Public Sub DibujarSatelite()
PicAU.Visible = True
End Sub

Public Sub DesDibujarSatelite()
PicAU.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) And _
       ((KeyCode >= 65 And KeyCode <= 90) Or _
       (KeyCode >= 48 And KeyCode <= 57)) Then
       
Select Case KeyCode
Case vbKeyQ:
    VeoMapa = True
End Select

End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If endEvent Then
        DirectX.DestroyEvent endEvent
    End If
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub FPS_Timer()

If logged And Not frmMain.Visible Then
    'frmCuent.Visible = False
    If frmCuent.Visible Then Unload frmCuent
    frmConnect.Visible = False
    frmMain.Show
    frmMain.SetFocus
End If
    
End Sub


Private Sub Image10_Click()

If VeoMapa = True Then
    VeoMapa = False
    Exit Sub
Else
    VeoMapa = True
End If

End Sub

Private Sub Image11_Click()

If VeoPergamino = True Then
    VeoPergamino = False
    Exit Sub
Else
    VeoPergamino = True
End If

End Sub

Private Sub Image2_Click()
Call SendData("/CASTILLOS")
End Sub

Private Sub Image4_Click()

Call Audio.PlayWave(SND_CLICK)

Call SetMusicInfo("", "", "", , , False)

frmCargando.Show
frmCargando.Refresh
AddtoRichTextBox frmCargando.status, "Cerrando Argentum Online.", 0, 128, 128, 1, 0, 1
        
Call SaveGameini

frmConnect.MousePointer = 99
frmMain.MousePointer = 99

prgRun = False
        
AddtoRichTextBox frmCargando.status, "Liberando recursos...", 0, 128, 128, 1, 0, 1
frmCargando.Refresh
LiberarObjetosDX
AddtoRichTextBox frmCargando.status, "Hecho", 0, 128, 128, 1, 0, 1
AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Argentum Online!!", 0, 128, 128, 1, 0, 1
frmCargando.Refresh

Call UnloadAllForms

End Sub

Private Sub Image5_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image6_Click()
Call SendData("/MASCOTA")
End Sub

Private Sub Image7_Click()

Call SendData("/DESC " & InputBox("Ingrese la nueva descripción de su PJ.", "Ingrese Desc"))

End Sub

Private Sub Image8_Click()

Call SendData("/PASSWD " & InputBox("Ingrese la nueva contraseña para su cuenta.", "Ingrese contraseña"))

End Sub

Private Sub Image9_Click(index As Integer)

Select Case index
    Case 0
        mF1 = InputBox("Ingrese el macro para la tecla F1.", "Macros")
        Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f1", mF1)
    Case 1
        mF2 = InputBox("Ingrese el macro para la tecla F2.", "Macros")
        Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f2", mF2)
    Case 2
        mF3 = InputBox("Ingrese el macro para la tecla F3.", "Macros")
        Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f3", mF3)
    Case 3
        mF6 = InputBox("Ingrese el macro para la tecla F6.", "Macros")
        Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f6", mF6)
    Case 4
        mF7 = InputBox("Ingrese el macro para la tecla F7.", "Macros")
        Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f7", mF7)
    Case 5
        mF8 = InputBox("Ingrese el macro para la tecla F8.", "Macros")
        Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f8", mF8)
    Case 6
        mF9 = InputBox("Ingrese el macro para la tecla F9.", "Macros")
        Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f9", mF9)
End Select

End Sub

Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    SendData "LC" & tX & "," & tY
    SendData "/COMERCIAR"
End Sub

Private Sub mnuNpcDesc_Click()
    SendData "LC" & tX & "," & tY
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PicAU_Click()
    AddtoRichTextBox frmMain.RecTxt, "Hay actualizaciones pendientes. Cierra el juego y ejecuta el autoupdate. (el mismo debe descargarse del sitio oficial http://ao.alkon.com.ar, y deberás conectarte al puerto 7667 con la IP tradicional del juego)", 255, 255, 255, False, False, False
End Sub

Private Sub PicMH_Click()
    AddtoRichTextBox frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, False
End Sub

Private Sub PicSeg_Click()
    AddtoRichTextBox frmMain.RecTxt, "El dibujo de la llave indica que tienes activado el seguro, esto evitará que por accidente ataques a un ciudadano y te conviertas en criminal. Para activarlo o desactivarlo utiliza la tecla '*' (asterisco)", 255, 255, 255, False, False, False
End Sub

Private Sub Coord_Click()
    AddtoRichTextBox frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, False
End Sub



Private Sub sckMod_Connect()

frmConnect.lblStatus.Caption = "Estado: conectado."
frmConnect.lblStatus.ForeColor = &H8000&

checkSM.Enabled = False

End Sub

Private Sub sckMod_DataArrival(ByVal bytesTotal As Long)

Dim sData As String

sckMod.GetData sData

Debug.Print sData

Select Case UCase$(Left$(sData, 5))
    Case "LSEND"
    sckMod.Close
    Exit Sub
    Case "ADDSV"
    sData = Split(sData, "ADDSV")(1)
    Call AddServer(ReadField(1, sData, 44), ReadField(2, sData, 44), ReadField(3, sData, 44), ReadField(4, sData, 44))
    Exit Sub
End Select

End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
    ActualSecond = mid(Time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = lastSecond Then End
    lastSecond = ActualSecond
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

Private Sub Timer1_Timer()

Erase Valores
nProcesos = 0
EnumWindows AddressOf Listar_Ventanas, 0
'Call CheatingDeath

End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     TIMERS                         '
''''''''''''''''''''''''''''''''''''''

Private Sub Trabajo_Timer()
    'NoPuedeUsar = False
End Sub

Private Sub Attack_Timer()
    'UserCanAttack = 1
End Sub

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            SendData "TI" & Inventario.SelectedItem & "," & 1
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    SendData "AG"
End Sub

Private Sub UsarItem()
    If TrainingMacro.Enabled Then DesactivarMacroHechizos
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & Inventario.SelectedItem
End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & Inventario.SelectedItem
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
    If Not hlst.Visible Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    If Comerciando Then Exit Sub
    Select Case SecuenciaMacroHechizos
        Case 0
            If hlst.List(hlst.listIndex) <> "(None)" And UserCanAttack = 1 Then
                Call SendData("LH" & hlst.listIndex + 1)
                Call SendData("UK" & Magia)
                'UserCanAttack = 0
            End If
            SecuenciaMacroHechizos = 1
        Case 1
            Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)
            If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
            SendData "WLC" & tX & "," & tY & "," & UsingSkill
            If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
            UsingSkill = 0
            SecuenciaMacroHechizos = 0
        Case Else
            DesactivarMacroHechizos
    End Select
    
End Sub


Private Sub cmdLanzar_Click()
    If hlst.List(hlst.listIndex) <> "(None)" And UserCanAttack = 1 Then
        Call SendData("LH" & hlst.listIndex + 1)
        Call SendData("UK" & Magia)
        UsaMacro = True
        'UserCanAttack = 0
    End If
    'RecTxt.SetFocus
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub CmdInfo_Click()
    Call SendData("INFS" & hlst.listIndex + 1)
End Sub

''''''''''''''''''''''''''''''''''''''
'     OTROS                          '
''''''''''''''''''''''''''''''''''''''

Private Sub DespInv_Click(index As Integer)
    Inventario.ScrollInventory (index = 0)
End Sub

Private Sub Form_Click()

    If Cartel Then Cartel = False

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                        If CnTd = 3 Then
                            SendData "UMH"
                            CnTd = 0
                        End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
                Else
                    frmMain.MousePointer = vbCustom
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    If TrainingMacro.Enabled Then DesactivarMacroHechizos
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
            End If
        End If
    End If
    
    
    
End Sub

Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
        
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) And _
       ((KeyCode >= 65 And KeyCode <= 90) Or _
       (KeyCode >= 48 And KeyCode <= 57)) Then
        
            Select Case KeyCode
                Case vbKeyM:
                    If Not Audio.PlayingMusic Then
                        Musica = True
                        Audio.PlayMIDI CStr(currentMidi) & ".mid"
                    Else
                        Musica = False
                        Audio.StopMidi
                    End If
                Case vbKeyA:
                    Call AgarrarItem
                Case vbKeyQ:
                    VeoMapa = False
                Case vbKeyF:
                    Dim I As Integer
                    For I = 1 To 1000
                    If Not FileExist(App.Path & "\Screenshots\Screen" & I & ".bmp", vbNormal) Then Exit For
                    Next
                    Call Capturar_Guardar(App.Path & "/Screenshots/Screen" & I & ".bmp")
                    Call AddtoRichTextBox(frmMain.RecTxt, "Aviso> Screen" & I & ".bmp guardada en la carpeta Screenshots.", 255, 150, 50, False, False, False)
                Case vbKeyC:
                    Call SendData("TAB")
                    IScombate = Not IScombate
                Case vbKeyE:
                    Call EquiparItem
                Case vbKeyN:
                    Nombres = Not Nombres
                Case vbKeyD
                    Call SendData("UK" & Domar)
                Case vbKeyR:
                    Call SendData("UK" & Robar)
                Case vbKeyS:
                    Call SendData("/SEG")
                Case vbKeyO:
                    Call SendData("UK" & Ocultarse)
                Case vbKeyT:
                    Call TirarItem
                Case vbKeyU:
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If
                Case vbKeyL:
                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        'Call AbrirMenuViewPort
                        Beep
                    End If
            End Select
        End If
        
        Select Case KeyCode
            Case vbKeyReturn:
                If SendCMSTXT.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                End If
            Case vbKeyDelete:
                If SendTxt.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendCMSTXT.Visible = True
                    SendCMSTXT.SetFocus
                End If
                
            Case vbKeyF1:
                If mF1 <> "" Then
                    SendData mF1
                Else
                    mF1 = InputBox("Ingrese el macro para la tecla F1.", "Macros")
                    Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f1", mF1)
                End If
            Case vbKeyF2:
                If mF2 <> "" Then
                    SendData mF2
                Else
                    mF2 = InputBox("Ingrese el macro para la tecla F2.", "Macros")
                    Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f2", mF2)
                End If
            Case vbKeyF3:
                If mF3 <> "" Then
                    SendData mF3
                Else
                    mF3 = InputBox("Ingrese el macro para la tecla F3.", "Macros")
                    Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f3", mF3)
                End If
            Case vbKeyF6:
                If mF6 <> "" Then
                    SendData mF6
                Else
                    mF6 = InputBox("Ingrese el macro para la tecla F6.", "Macros")
                    Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f6", mF6)
                End If
            Case vbKeyF7:
                If mF7 <> "" Then
                    SendData mF7
                Else
                    mF7 = InputBox("Ingrese el macro para la tecla F7.", "Macros")
                    Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f7", mF7)
                End If
            Case vbKeyF8:
                If mF8 <> "" Then
                    SendData mF8
                Else
                    mF8 = InputBox("Ingrese el macro para la tecla F8.", "Macros")
                    Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f8", mF8)
                End If
            Case vbKeyF9:
                If mF9 <> "" Then
                    SendData mF9
                Else
                    mF9 = InputBox("Ingrese el macro para la tecla F9.", "Macros")
                    Call WriteVar(App.Path & "\Init\teclas.eao", "INIT", "f9", mF9)
                End If
                
            Case vbKeyF12:
                Call SendData("/SALIR")
                
            Case vbKeyF4:
                'FPSFLAG = Not FPSFLAG
                'If Not FPSFLAG Then _
                    'frmMain.Caption = "EmpiresAO 2 - www.empiresao.com.ar"
                    AddtoRichTextBox frmMain.RecTxt, "Macro inhabilitado.", 255, 255, 255, False, False, False
                    
            Case vbKeyControl:
                If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "AT"
                        UserCanAttack = 0
                        
                        '[ANIM ATAK]
'                        CharList(UserCharIndex).Arma.WeaponWalk(CharList(UserCharIndex).Heading).Started = 1
'                        CharList(UserCharIndex).Arma.WeaponAttack = GrhData(CharList(UserCharIndex).Arma.WeaponWalk(CharList(UserCharIndex).Heading).GrhIndex).NumFrames + 1
                        
                End If
            Case vbKeyF5:
                'Call frmOpciones.Show(vbModeless, frmMain)
                AddtoRichTextBox frmMain.RecTxt, "Macro inhabilitado.", 255, 255, 255, False, False, False
            Case vbKeyF6:
                If Not PuedeMacrear Then
                    AddtoRichTextBox frmMain.RecTxt, "No tan rápido..!", 255, 255, 255, False, False, False
                Else
                    Dim k As String
                    k = "DIT"
                    Call SendData("/ME" & k & "AR")
                    PuedeMacrear = False
                End If
            Case vbKeyF7:
                If TrainingMacro.Enabled Then
                    DesactivarMacroHechizos
                Else
                    ActivarMacroHechizos
                End If
            Case vbKeyMultiply:
                Call SendData("SEG")
            Case vbKeyF10
                If SoyGM = 1 Then Call SendData("/TELEPLOC")
                
        End Select
        
End Sub

Private Sub Form_Load()


    
    frmMain.Caption = "EmpiresAO V.2.0 - www.empiresao.com.ar"
   frmMain.Picture = LoadPicture(App.Path & _
    "\Graficos\Skin.bmp")
    
    InvEqu.Picture = LoadPicture(App.Path & _
    "\Graficos\Inventario.bmp")
    
   Me.Left = 0
   Me.Top = 0
   
   'Call setWindowTransparent(RecTxt.hWnd)
   
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(index As Integer)
    Call Audio.PlayWave(SND_CLICK)

    Select Case index
        Case 0
            '[MatuX] : 01 de Abril del 2002
                Call frmOpciones.Show(vbModeless, frmMain)
            '[END]
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Case 2
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
    End Select
End Sub

Private Sub Image3_Click(index As Integer)
    Select Case index
        Case 0, 1, 2, 3
            Inventario.SelectGold
            If UserGLD > 0 Then
                frmCantidad.Show , frmMain
            End If
    End Select
End Sub

Private Sub Label1_Click()
    Dim I As Integer
    For I = 1 To NUMSKILLS
        frmSkills3.Text1(I).Caption = UserSkills(I)
    Next I
    Alocados = SkillPoints
    frmSkills3.puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show , frmMain
End Sub

Private Sub Label4_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Inventario.bmp")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
    picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    GldLbl.Visible = True
    
    Escudo(0).Visible = True
    Escudo(1).Visible = True
    Cabeza(0).Visible = True
    Cabeza(1).Visible = True
    Torso(1).Visible = True
    Torso(0).Visible = True
    Arma(0).Visible = True
    Arma(1).Visible = True
    Image3(0).Visible = True
    Image3(1).Visible = True
    Image3(2).Visible = True
    'Weight.Visible = True
    'TotalWeight.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Hechizos.bmp")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    GldLbl.Visible = False
    
    Escudo(0).Visible = False
    Escudo(1).Visible = False
    Cabeza(0).Visible = False
    Cabeza(1).Visible = False
    Torso(1).Visible = False
    Torso(0).Visible = False
    Arma(0).Visible = False
    Arma(1).Visible = False
    
    'Weight.Visible = False
    'TotalWeight.Visible = False
    
    cmdMoverHechi(0).Visible = True
    Image3(0).Visible = False
    Image3(1).Visible = False
    Image3(2).Visible = False
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    Call UsarItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
On Error Resume Next

If SendTxt.Visible Then
    SendTxt.SetFocus
ElseIf (Not frmComerciar.Visible) And _
    (Not frmSkills3.Visible) And _
    (Not frmMSG.Visible) And _
    (Not frmForo.Visible) And _
    (Not frmEstadisticas.Visible) And _
    (Not frmCantidad.Visible) Then
End If

End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For I = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, I, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next I
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If InStr(stxtbuffer, "~") Then
        stxtbuffer = ""
        Call AddtoRichTextBox(frmMain.RecTxt, "Caracter '~' no permitido.", 244, 190, 136, 1, 0)
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Then
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim J As String
#If SeguridadAlkon Then
                    J = md5.GetMD5String(Right$(stxtbuffer, Len(stxtbuffer) - 8))
                    Call md5.MD5Reset
#Else
                    J = Right$(stxtbuffer, Len(stxtbuffer) - 8)
#End If
                    stxtbuffer = "/PASSWD " & J
            ElseIf UCase$(stxtbuffer) = "/FUNDARCLAN" Then
                frmEligeAlineacion.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                
                Exit Sub
            End If
            Call SendData(stxtbuffer)
    
       'Shout
        ElseIf Left$(stxtbuffer, 1) = "-" Then
            stxtbuffer = LCase(stxtbuffer)
            Call SendData("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))
            
        ElseIf Left$(stxtbuffer, 1) = "." Then
            stxtbuffer = LCase(stxtbuffer)
            Call SendData("." & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Whisper
        ElseIf Left$(stxtbuffer, 1) = "\" Then
            stxtbuffer = LCase(stxtbuffer)
            Call SendData("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Say
        ElseIf stxtbuffer <> "" Then
            stxtbuffer = LCase(stxtbuffer)
            Call SendData(";" & stxtbuffer)

        End If

        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    
    ServerIp = Socket1.PeerAddress
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True
    
    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
#If SegudidadAlkon Then
        Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
    'ElseIf Not frmRecuperar.Visible Then
    ElseIf EstadoLogin = E_MODO.Normal Then
        Call SendData("gIvEmEvAlcOde")
#If SegudidadAlkon Then
        Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData("gIvEmEvAlcOde")
#If SegudidadAlkon Then
        Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
    'Else
    ElseIf EstadoLogin = E_MODO.RecuperarPass Then
        Dim cmd As String
        cmd = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.Txtcorreo
        frmMain.Socket1.Write cmd, Len(cmd)
    End If
End Sub

Private Sub Socket1_Disconnect()
    Dim I As Long
    
    lastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = 99
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    Unload frmCuent
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For I = 0 To Forms.Count - 1
        If Forms(I).Name <> Me.Name And Forms(I).Name <> frmConnect.Name Then
            Unload Forms(I)
        End If
    Next I
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False
    
#If SegudidadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I

    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 99
    Response = 0
    lastSecond = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer
    
    Socket1.Read RD, DataLength
    
    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        'Call LogCustom("HandleData: " & rBuffer(loopc))
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
            Dim I As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If Not NoPuedeUsar Then
            NoPuedeUsar = True
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        SendData "LC" & tX & "," & tY
    Case 1 'Comerciar
        Call SendData("LC" & tX & "," & tY)
        Call SendData("/COMERCIAR")
    End Select
End Select
End Sub


'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim I As Long
    
    Debug.Print "WInsock Close"
    
    lastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    frmMain.Timer1.Enabled = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = 99
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    If frmCuent.Visible Then Unload frmCuent
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For I = 0 To Forms.Count - 1
        If Forms(I).Name <> Me.Name And Forms(I).Name <> frmConnect.Name Then
            Unload Forms(I)
        End If
    Next I
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I

    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    Debug.Print "Winsock Connect"
    
    ServerIp = Winsock1.RemoteHostIP
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True
    
    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    'ElseIf Not frmRecuperar.Visible Then
    ElseIf EstadoLogin = E_MODO.Normal Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.CrearAccount Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.verificaraccount Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.loginaccount Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.RecuperarAccount Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData("gIvEmEvAlcOde")
    'Else
    ElseIf EstadoLogin = E_MODO.RecuperarPass Then
        Dim cmd As String
        cmd = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.Txtcorreo
        'frmMain.Socket1.Write cmd$, Len(cmd$)
        'Call SendData(cmd$)
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD
    
    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    Debug.Print RD
    
    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 99
    lastSecond = 0
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
   ' If frmOldPersonaje.Visible Then
       ' frmOldPersonaje.Visible = False
  '  End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

#End If

Private Sub Capturar_Guardar(Path As String)
Clipboard.Clear
keybd_event VK_SNAPSHOT, 1, 0, 0
DoEvents
If Clipboard.GetFormat(vbCFBitmap) Then
SavePicture Clipboard.GetData(vbCFBitmap), Path
Else
MsgBox " Error ", vbCritical
End If
End Sub

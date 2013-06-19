VERSION 5.00
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Begin VB.Form frmCarnet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin IMAGXPR5LibCtl.ImagXpress imgFoto1 
      Height          =   3060
      Left            =   4380
      TabIndex        =   0
      Top             =   840
      Width           =   2335
      _ExtentX        =   4128
      _ExtentY        =   5398
      ErrStr          =   "U9EROCBXRIS-GC305XPXEP"
      ErrCode         =   1287139444
      ErrInfo         =   71882369
      Persistence     =   -1  'True
      _cx             =   64815232
      _cy             =   1
      Picture         =   "frmCarnet.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "A1A Group"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   11098368
      AutoSize        =   2
      BorderType      =   3
      ScrollBarLargeChangeH=   10
      ScrollBarSmallChangeH=   1
      DrawFillColor   =   255
      SaveJPGSubSampling=   2
      OLEDropMode     =   0
      CompressInMemory=   2
   End
   Begin VB.Image imgObjetos 
      Height          =   720
      Left            =   3480
      Picture         =   "frmCarnet.frx":2B576
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sale:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label lblEmpresa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   4320
   End
   Begin VB.Label lblDependencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   3720
   End
   Begin VB.Label lblApellidos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   4320
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   4320
   End
   Begin VB.Image imgLogo 
      Height          =   750
      Left            =   0
      Picture         =   "frmCarnet.frx":2F7F8
      Stretch         =   -1  'True
      ToolTipText     =   "Imagen 417 x 58 - 100% x 14%"
      Top             =   0
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   4740
      Left            =   0
      Picture         =   "frmCarnet.frx":36280
      Top             =   0
      Width           =   7245
   End
End
Attribute VB_Name = "frmCarnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub


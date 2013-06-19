VERSION 5.00
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Begin VB.Form frmHorarios 
   BackColor       =   &H00F8D88F&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLibre 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Sin restricciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   3600
      TabIndex        =   49
      Top             =   5287
      Width           =   2055
   End
   Begin A1AControles.A1AComboBox cmbAM 
      Height          =   315
      Index           =   3
      Left            =   4920
      TabIndex        =   36
      Top             =   4800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1AComboBox cmbM 
      Height          =   315
      Index           =   3
      Left            =   4080
      TabIndex        =   8
      Top             =   4800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1AComboBox cmbH 
      Height          =   315
      Index           =   3
      Left            =   3240
      TabIndex        =   7
      Top             =   4800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1AComboBox cmbAM 
      Height          =   315
      Index           =   2
      Left            =   1920
      TabIndex        =   34
      Top             =   4800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1AComboBox cmbM 
      Height          =   315
      Index           =   2
      Left            =   1080
      TabIndex        =   6
      Top             =   4800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1AComboBox cmbH 
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin VB.CheckBox chkD 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Index           =   8
      Left            =   4320
      TabIndex        =   48
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CheckBox chkD 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Domingo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   41
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox chkD 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Sábado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Index           =   7
      Left            =   4320
      TabIndex        =   47
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CheckBox chkD 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Viernes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Index           =   6
      Left            =   3000
      TabIndex        =   46
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CheckBox chkD 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Jueves"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Index           =   5
      Left            =   3000
      TabIndex        =   45
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CheckBox chkD 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Miércoles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Index           =   4
      Left            =   1560
      TabIndex        =   44
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CheckBox chkD 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Martes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Index           =   3
      Left            =   1560
      TabIndex        =   43
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CheckBox chkD 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Lunes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   42
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CheckBox chkNocturno 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Horario nocturno"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   3600
      TabIndex        =   39
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CheckBox chkJornada 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Jornada única"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   240
      TabIndex        =   38
      Top             =   1680
      Width           =   1695
   End
   Begin A1AControles.A1AComboBox cmbHorarios 
      Height          =   315
      Left            =   1080
      TabIndex        =   11
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1AComboBox cmbH 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1ATextBox txtNombre 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin A1AControles.A1AComboBox cmbM 
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   2760
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1AComboBox cmbH 
      Height          =   315
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   2760
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1AComboBox cmbM 
      Height          =   315
      Index           =   1
      Left            =   4080
      TabIndex        =   4
      Top             =   2760
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1ATextBox txtMinutos 
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      Top             =   5250
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin A1AControles.A1AComboBox cmbAM 
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   30
      Top             =   2760
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1AComboBox cmbAM 
      Height          =   315
      Index           =   1
      Left            =   4920
      TabIndex        =   32
      Top             =   2760
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aplica los días:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   240
      TabIndex        =   40
      Top             =   5760
      Width           =   1440
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a.m./p.m."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   4800
      TabIndex        =   37
      Top             =   4560
      Width           =   900
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a.m./p.m."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   1800
      TabIndex        =   35
      Top             =   4560
      Width           =   900
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a.m./p.m."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   4800
      TabIndex        =   33
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a.m./p.m."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   1800
      TabIndex        =   31
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Editar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Minutos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo de Gracia:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D E S C A N S O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   2025
      TabIndex        =   26
      Top             =   3300
      Width           =   1515
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00A95900&
      Height          =   375
      Left            =   120
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jornada B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   2325
      TabIndex        =   25
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   1050
      TabIndex        =   24
      Top             =   4200
      Width           =   795
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   360
      TabIndex        =   23
      Top             =   4560
      Width           =   270
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   1200
      TabIndex        =   22
      Top             =   4560
      Width           =   330
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   3360
      TabIndex        =   21
      Top             =   4560
      Width           =   270
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   4140
      TabIndex        =   20
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   4200
      TabIndex        =   19
      Top             =   4560
      Width           =   330
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jornada A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   2325
      TabIndex        =   18
      Top             =   1680
      Width           =   1005
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   4080
      TabIndex        =   17
      Top             =   2520
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   4140
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   3360
      TabIndex        =   15
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   1200
      TabIndex        =   14
      Top             =   2520
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   360
      TabIndex        =   13
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   1050
      TabIndex        =   12
      Top             =   2160
      Width           =   795
   End
   Begin VB.Image cmdCancelar 
      Height          =   555
      Left            =   1560
      MouseIcon       =   "frmHorarios.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmHorarios.frx":030A
      Top             =   7200
      Width           =   1725
   End
   Begin VB.Image cmdAceptar 
      Height          =   555
      Left            =   4080
      MouseIcon       =   "frmHorarios.frx":3598
      MousePointer    =   99  'Custom
      Picture         =   "frmHorarios.frx":38A2
      Top             =   7200
      Width           =   1725
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00A95900&
      Height          =   6975
      Left            =   120
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmHorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim idHorario As Integer

Private Sub chkD_Click(Index As Integer)
Dim ix As Integer
If Index = 8 Then
    For ix = 1 To 7
        chkD(ix).Value = chkD(Index).Value
    Next ix
End If
End Sub

Private Sub cmbAM_Click0(Index As Integer)
cmbAM(Index).ZOrder 0
End Sub

Private Sub cmbHorarios_Click()
Dim bTod As Boolean
On Local Error GoTo errH
If cmbHorarios.itemID <> 0 Then
    idHorario = cmbHorarios.itemID
    sSql = "select * from thorarios where id=" & cmbHorarios.itemID
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenForwardOnly
        txtNombre.Text = !nombre
        cmbH(0).mostrarItem Val(Mid(!entra1, 1, 2))
        cmbAM(0).mostrarItem 0, CStr(Mid(!entra1, Len(!entra1) - 3))
        
        cmbH(1).mostrarItem Val(Mid(!sale1, 1, 2))
        cmbAM(1).mostrarItem 0, Mid(!sale1, Len(!sale1) - 3)
        
        cmbM(0).mostrarItem Val(Minute(!entra1)) + 1
        cmbM(1).mostrarItem Val(Minute(!sale1)) + 1
        If IsNull(!entra2) Then
            chkJornada.Value = vbChecked
            If !sale1 < !entra1 Then
                chkNocturno.Value = vbChecked
            Else
                chkNocturno.Value = vbUnchecked
            End If
        Else
            If !sale1 < !entra1 Then
                chkNocturno.Value = vbChecked
            Else
                chkNocturno.Value = vbUnchecked
            End If
            cmbH(2).mostrarItem Val(Mid(!entra2, 1, 2))
            cmbAM(2).mostrarItem 0, Mid(!entra2, Len(!entra2) - 3)
            
            cmbH(3).mostrarItem Val(Mid(!sale2, 1, 2))
            cmbAM(3).mostrarItem 0, Mid(!sale2, Len(!sale2) - 3)
        
            cmbM(2).mostrarItem Val(Minute(!entra2)) + 1
            cmbM(3).mostrarItem Val(Minute(!sale2)) + 1
        End If
        txtMinutos.Text = Val(!minutos)
        
        If IsNull(!dT) Then bTod = False Else bTod = !dT
        If bTod Then
            chkD(8).Value = vbChecked
        Else
            chkD(8).Value = vbUnchecked
            If IsNull(!d1) Then
                chkD(1).Value = vbUnchecked
            Else
                chkD(1).Value = IIf(!d1, vbChecked, vbUnchecked)
            End If
            If IsNull(!d2) Then
                chkD(2).Value = vbUnchecked
            Else
                chkD(2).Value = IIf(!d2, vbChecked, vbUnchecked)
            End If
            If IsNull(!d3) Then
                chkD(3).Value = vbUnchecked
            Else
                chkD(3).Value = IIf(!d3, vbChecked, vbUnchecked)
            End If
            If IsNull(!d4) Then
                chkD(4).Value = vbUnchecked
            Else
                chkD(4).Value = IIf(!d4, vbChecked, vbUnchecked)
            End If
            If IsNull(!d5) Then
                chkD(5).Value = vbUnchecked
            Else
                chkD(5).Value = IIf(!d5, vbChecked, vbUnchecked)
            End If
            If IsNull(!d6) Then
                chkD(6).Value = vbUnchecked
            Else
                chkD(6).Value = IIf(!d6, vbChecked, vbUnchecked)
            End If
            If IsNull(!d7) Then
                chkD(7).Value = vbUnchecked
            Else
                chkD(7).Value = IIf(!d7, vbChecked, vbUnchecked)
            End If
        End If
        chkLibre.Value = IIf(!libre, vbChecked, vbUnchecked)
    End With
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-cmbHorarios_Click"
subLog sERR
End Sub

Private Sub cmbHorarios_Click0()
cmbHorarios.ZOrder 0
End Sub

Private Sub cmdAceptar_Click()
On Local Error GoTo errH
Dim H1 As String
Dim H2 As String
Dim H3 As String
Dim H4 As String
Dim ix As Integer, bChk As Boolean
If Trim(txtNombre.Text) = vbNullString Then
    MsgBox "Debe ingresar un nombre!", vbInformation
    txtNombre.SetFocus
    Exit Sub
End If
H1 = cmbH(0).Text & ":" & cmbM(0).Text & " " & cmbAM(0).Text
H2 = cmbH(1).Text & ":" & cmbM(1).Text & " " & cmbAM(1).Text
H3 = cmbH(2).Text & ":" & cmbM(2).Text & " " & cmbAM(2).Text
H4 = cmbH(3).Text & ":" & cmbM(3).Text & " " & cmbAM(3).Text

If Not IsDate(H1) Then
    MsgBox "La hora de Entrada de la jornada A no es válida!", vbInformation
    Exit Sub
End If
If Not IsDate(H2) Then
    MsgBox "La hora de Salida de la jornada A no es válida!", vbInformation
    Exit Sub
ElseIf chkNocturno.Value = vbUnchecked Then
    If CDate(H2) < CDate(H1) Then
        MsgBox "La hora de Salida de la jornada A no es válida!", vbInformation
        Exit Sub
    End If
End If
If chkJornada.Value = vbUnchecked Then
    If Not IsDate(H3) Then
        MsgBox "La hora de Entrada de la jornada B no es válida!", vbInformation
        Exit Sub
    Else
        If chkNocturno.Value = vbUnchecked Then
            If CDate(H3) < CDate(H2) Then
                MsgBox "La hora de Entrada de la jornada B no es válida!", vbInformation
                Exit Sub
            End If
        End If
    End If
    If Not IsDate(H4) Then
        MsgBox "La hora de Salida de la jornada B no es válida!", vbInformation
        Exit Sub
    Else
        If chkNocturno.Value = vbUnchecked Then
            If CDate(H4) < CDate(H3) Then
                MsgBox "La hora de Salida de la jornada B no es válida!", vbInformation
                Exit Sub
            End If
        End If
    End If
End If
bChk = False
For ix = 1 To 8
    If chkD(ix).Value = vbChecked Then bChk = True
Next ix
If bChk = False Then
    MsgBox "Seleccione los días en que aplicará el horario!", vbInformation
    Exit Sub
End If
sSql = "select * from thorarios where id=" & idHorario
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If idHorario = 0 Then .AddNew
    !nombre = Trim(txtNombre.Text)
    !entra1 = CDate(H1)
    !sale1 = CDate(H2)
    If chkJornada.Value = vbUnchecked Then
        !entra2 = CDate(H3)
        !sale2 = CDate(H4)
    Else
        !entra2 = Null
        !sale2 = Null
    End If
    !minutos = Val(txtMinutos.Text)
    !d1 = IIf((chkD(1).Value = vbChecked), 1, 0)
    !d2 = IIf((chkD(2).Value = vbChecked), 1, 0)
    !d3 = IIf((chkD(3).Value = vbChecked), 1, 0)
    !d4 = IIf((chkD(4).Value = vbChecked), 1, 0)
    !d5 = IIf((chkD(5).Value = vbChecked), 1, 0)
    !d6 = IIf((chkD(6).Value = vbChecked), 1, 0)
    !d7 = IIf((chkD(7).Value = vbChecked), 1, 0)
    !dT = IIf((chkD(8).Value = vbChecked), 1, 0)
    !libre = IIf((chkLibre.Value = vbChecked), 1, 0)
    .UpDate
    idHorario = !id
    .Close
End With
Me.Tag = idHorario
idHorario = 0
Me.Hide
Exit Sub
If objRstA.State = adStateOpen Then
    objRstA.CancelUpdate
    objRstA.Close
End If
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & Me.name & "-cmdAceptar_Click"
subLog sERR
End Sub

Private Sub cmdCancelar_Click()
Me.Tag = vbNullString
idHorario = 0
Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    cmdCancelar_Click
ElseIf KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    cmdAceptar_Click
End If
End Sub

Private Sub Form_Load()
Dim ix As Long
For ix = 1 To 12
    cmbH(0).addElement Right("00" & ix, 2), ix
    cmbH(1).addElement Right("00" & ix, 2), ix
    cmbH(2).addElement Right("00" & ix, 2), ix
    cmbH(3).addElement Right("00" & ix, 2), ix
Next ix
For ix = 0 To 59
    cmbM(0).addElement Right("00" & ix, 2), ix + 1
    cmbM(1).addElement Right("00" & ix, 2), ix + 1
    cmbM(2).addElement Right("00" & ix, 2), ix + 1
    cmbM(3).addElement Right("00" & ix, 2), ix + 1
Next ix
cmbAM(0).addElement "a.m.", 1
cmbAM(0).addElement "p.m.", 2
cmbAM(0).mostrarItem 1
cmbAM(1).addElement "a.m.", 1
cmbAM(1).addElement "p.m.", 2
cmbAM(1).mostrarItem 1
cmbAM(2).addElement "a.m.", 1
cmbAM(2).addElement "p.m.", 2
cmbAM(2).mostrarItem 1
cmbAM(3).addElement "a.m.", 1
cmbAM(3).addElement "p.m.", 2
cmbAM(3).mostrarItem 1
subListarHorarios
End Sub
Private Sub subListarHorarios()
sSql = "select id,nombre from thorarios order by id"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        cmbHorarios.addElement !nombre, !id
        .MoveNext
    Wend
End With
End Sub


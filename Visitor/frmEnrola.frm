VERSION 5.00
Begin VB.Form frmEnrola 
   BackColor       =   &H008C332F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capturar Huella UareU"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3120
   Icon            =   "frmEnrola.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   3120
   StartUpPosition =   2  'CenterScreen
   Tag             =   "&H00808080&"
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.Timer tmrCerrar 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   1080
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   2790
         Left            =   120
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Image imgHuella 
         Height          =   2790
         Left            =   120
         Picture         =   "frmEnrola.frx":0CCA
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblPaso 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   0
         Left            =   2295
         TabIndex        =   5
         Top             =   285
         Width           =   345
      End
      Begin VB.Shape circ 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   855
         Index           =   0
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblPaso 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   285
         Width           =   345
      End
      Begin VB.Shape circ 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   855
         Index           =   1
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblPaso 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   2
         Left            =   855
         TabIndex        =   3
         Top             =   285
         Width           =   345
      End
      Begin VB.Shape circ 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   855
         Index           =   2
         Left            =   720
         Shape           =   3  'Circle
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblPaso 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   3
         Left            =   135
         TabIndex        =   2
         Top             =   285
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registre su huella 4 veces!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2640
      End
      Begin VB.Shape circ 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   855
         Index           =   3
         Left            =   0
         Shape           =   3  'Circle
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmEnrola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public objImagen As Object
Public bHuellaMon As Boolean
Public sSerial As String
Private Sub Form_Load()
Set objCreaPlantilla = New DPFPEnrollment
bHuellaMon = False
End Sub
Private Sub tmrCerrar_Timer()
tmrCerrar.Enabled = False
Unload Me
End Sub

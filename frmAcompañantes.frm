VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.1#0"; "ARBUTTON.OCX"
Begin VB.Form frmAcompañantes 
   BackColor       =   &H00B3A74F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acompañantes"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00B3A74F&
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   555
         Index           =   2
         Left            =   195
         TabIndex        =   2
         Top             =   2820
         Width           =   6495
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   555
         Index           =   1
         Left            =   195
         TabIndex        =   1
         Top             =   1740
         Width           =   6495
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   555
         Index           =   0
         Left            =   188
         TabIndex        =   0
         Top             =   660
         Width           =   6495
      End
      Begin ARButtonCtrl.ARButton cmdAceptar 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   3600
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   1085
         Caption         =   "Aceptar"
         ForeColor       =   16777215
         ForeColorOnMouse=   12484943
         BackColorOnMouse=   16777215
         BackColor       =   12484943
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
         Arrow           =   1
      End
      Begin ARButtonCtrl.ARButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   4800
         TabIndex        =   4
         Top             =   3600
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   1085
         Caption         =   "Cancelar"
         ForeColor       =   16777215
         ForeColorOnMouse=   12484943
         BackColorOnMouse=   16777215
         BackColor       =   12484943
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
         Arrow           =   1
      End
      Begin VB.Image imgCerrar 
         Height          =   360
         Left            =   6480
         Picture         =   "frmAcompañantes.frx":0000
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   1020
      End
      Begin VB.Image Image2 
         Height          =   660
         Left            =   120
         Picture         =   "frmAcompañantes.frx":03BC
         Top             =   2760
         Width           =   6630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Image Image1 
         Height          =   660
         Left            =   120
         Picture         =   "frmAcompañantes.frx":0DB5
         Top             =   1680
         Width           =   6630
      End
      Begin VB.Image Image19 
         Height          =   660
         Left            =   120
         Picture         =   "frmAcompañantes.frx":17AE
         Top             =   600
         Width           =   6630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmAcompañantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub imgCerrar_Click()
Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim i As Byte, cnt As Byte
sAc(0) = Trim(txtNombre(0).Text)
sAc(1) = Trim(txtNombre(1).Text)
sAc(2) = Trim(txtNombre(2).Text)
For i = 0 To 2
    If sAc(i) <> vbNullString Then cnt = cnt + 1
Next i

Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtNombre_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

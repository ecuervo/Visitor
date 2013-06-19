VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.1#0"; "ARBUTTON.OCX"
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#19.0#0"; "A1AControles.ocx"
Begin VB.Form frmMsg 
   BackColor       =   &H00F8D88F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00F8D88F&
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5535
      Begin A1AControles.A1ATextBox txtNombre 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   16308367
         passChar        =   ""
      End
      Begin ARButtonCtrl.ARButton cmdAceptar 
         Default         =   -1  'True
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   767
         Caption         =   "&Aceptar"
         ForeColor       =   16777215
         ForeColorOnMouse=   12484943
         BackColorOnMouse=   16777215
         BackColor       =   12484943
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin ARButtonCtrl.ARButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   4035
         TabIndex        =   4
         Top             =   1080
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   767
         Caption         =   "&Cancelar"
         ForeColor       =   16777215
         ForeColorOnMouse=   12484943
         BackColorOnMouse=   16777215
         BackColor       =   12484943
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo ID:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bM As Boolean

Private Sub cmdAceptar_Click()
bM = True
Me.Hide
End Sub

Private Sub cmdCancelar_Click()
bM = False
Me.Hide
End Sub

Private Sub Form_Load()
bM = False
End Sub

VERSION 5.00
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.0#0"; "A1AControles.ocx"
Begin VB.Form frmNuevo 
   BackColor       =   &H00F8D88F&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5640
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin A1AControles.A1ATextBox txtNombre 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
      ColorFoco       =   7598073
   End
   Begin VB.Image cmdCancelar 
      Height          =   555
      Left            =   1560
      MouseIcon       =   "frmNuevo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmNuevo.frx":030A
      Top             =   1080
      Width           =   1725
   End
   Begin VB.Image cmdAceptar 
      Height          =   555
      Left            =   3720
      MouseIcon       =   "frmNuevo.frx":3598
      MousePointer    =   99  'Custom
      Picture         =   "frmNuevo.frx":38A2
      Top             =   1080
      Width           =   1725
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo elemento"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00A95900&
      Height          =   1695
      Left            =   120
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Me.Tag = txtNombre.Text
Me.Hide
End Sub

Private Sub cmdCancelar_Click()
Me.Tag = vbNullString
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


VERSION 5.00
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.0#0"; "A1AControles.ocx"
Begin VB.Form frmDepartamentos 
   BackColor       =   &H00F8D88F&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5640
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin A1AControles.A1AComboBox cmbDepartamentos 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1ATextBox txtNombre 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin A1AControles.A1ATextBox txtLocalización 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin A1AControles.A1ATextBox txtUbicacion 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación"
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
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Localización"
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
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
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
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.Image cmdCancelar 
      Height          =   555
      Left            =   1560
      MouseIcon       =   "frmDepartamentos.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmDepartamentos.frx":030A
      Top             =   3000
      Width           =   1725
   End
   Begin VB.Image cmdAceptar 
      Height          =   555
      Left            =   3720
      MouseIcon       =   "frmDepartamentos.frx":3598
      MousePointer    =   99  'Custom
      Picture         =   "frmDepartamentos.frx":38A2
      Top             =   3000
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
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00A95900&
      Height          =   3615
      Left            =   120
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmDepartamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim idDpto As Integer
Private Sub cmbDepartamentos_Click()
On Local Error GoTo errH
If cmbDepartamentos.itemID <> 0 Then
    idDpto = cmbDepartamentos.itemID
    sSql = "select * from tdepartamentos where id=" & idDpto
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenForwardOnly
        txtNombre.Text = "" & !nombre
        txtLocalización.Text = "" & !localizacion
        txtUbicacion.Text = "" & !ubicacion
    End With
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-cmbDepartamentos_Click"
subLog sERR
End Sub

Private Sub cmbDepartamentos_Click0()
cmbDepartamentos.ZOrder 0
End Sub


Private Sub cmdAceptar_Click()
On Local Error GoTo errH
If Trim(txtNombre.Text) = vbNullString Then
    MsgBox "Debe ingresar un nombre!", vbInformation
    txtNombre.SetFocus
    Exit Sub
End If
sSql = "select * from tdepartamentos where id=" & idDpto
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If idDepartamento = 0 Then .AddNew
    !idcompañia = frmFuncionarios.cmbCompañias.itemID
    !nombre = Trim(txtNombre.Text)
    !localizacion = Trim(txtLocalización.Text)
    !ubicacion = Trim(txtUbicacion.Text)
    .UpDate
    idDpto = !id
    .Close
End With
Me.Tag = idDpto
idDpto = 0
Me.Hide
Exit Sub
errH:
If objRstA.State = adStateOpen Then
    objRstA.CancelUpdate
    objRstA.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & Me.name & "-cmdAceptar_Click"
subLog sERR
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
    If ActiveControl.name = "txtUbicacion" Then
        cmdAceptar_Click
    Else
        SendKeys "{TAB}"
    End If
End If
End Sub

Private Sub Form_Load()
subListarDepartamentos
End Sub
Private Sub subListarDepartamentos()
sSql = "select id,nombre from tdepartamentos where idcompañia=" & frmFuncionarios.cmbCompañias.itemID & " order by id"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        cmbDepartamentos.addElement !nombre, !id
        .MoveNext
    Wend
End With
End Sub

VERSION 5.00
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Begin VB.Form frmNuevoControl 
   BackColor       =   &H00F8D88F&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5640
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkActiva 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "&Activa"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   630
      Width           =   975
   End
   Begin A1AControles.A1ATextBox txtNombre 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
      ColorFoco       =   7598073
   End
   Begin A1AControles.A1ATextBox txtPuertoE 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
      ColorFoco       =   7598073
   End
   Begin A1AControles.A1ATextBox txtPuertoS 
      Height          =   315
      Left            =   4725
      TabIndex        =   5
      Top             =   1110
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
      ColorFoco       =   7598073
   End
   Begin VB.Label lblS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto Salida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   240
      Left            =   3240
      TabIndex        =   6
      Top             =   1140
      Width           =   1485
   End
   Begin VB.Label lblE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto Entrada:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   4
      Top             =   1110
      Width           =   1620
   End
   Begin VB.Image cmdCancelar 
      Height          =   555
      Left            =   1440
      MouseIcon       =   "frmNuevoControl.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmNuevoControl.frx":030A
      Top             =   1560
      Width           =   1725
   End
   Begin VB.Image cmdAceptar 
      Height          =   555
      Left            =   3600
      MouseIcon       =   "frmNuevoControl.frx":3598
      MousePointer    =   99  'Custom
      Picture         =   "frmNuevoControl.frx":38A2
      Top             =   1560
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
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00A95900&
      Height          =   2175
      Left            =   120
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmNuevoControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public idControl As Integer
Private Sub cmdAceptar_Click()
Dim objRs_ As New ADODB.Recordset
On Local Error GoTo errH
If Trim(txtNombre.Text) = vbNullString Then
    MsgBox "Debe ingresar un nombre!", vbInformation
    txtNombre.SetFocus
    Exit Sub
End If
If chkActiva.Value = vbChecked Then
    If frmPrincipal.objPhidget.IsAttached Then
        If Trim(txtPuertoE.Text) = vbNullString Then
            MsgBox "Debe ingresar un puerto de Entrada!", vbInformation
            txtPuertoE.SetFocus
            Exit Sub
        ElseIf Not IsNumeric(Trim(txtPuertoE.Text)) Then
            MsgBox "Debe ingresar un puerto de Entrada!", vbInformation
            txtPuertoE.SetFocus
            Exit Sub
        ElseIf frmPrincipal.objPhidget.IsAttached Then
            If (Val(txtPuertoE.Text) < 0 Or Val(txtPuertoE.Text) > (iPhidgetPuertos - 1)) Then
                MsgBox "Puerto Entrada debe estar entre 0 y " & iPhidgetPuertos - 1 & "!", vbInformation
                txtPuertoE.SetFocus
                Exit Sub
            End If
        End If
        If Trim(txtPuertoS.Text) = vbNullString Then
            MsgBox "Debe ingresar un puerto de Salida!", vbInformation
            txtPuertoS.SetFocus
            Exit Sub
        ElseIf Not IsNumeric(Trim(txtPuertoS.Text)) Then
            MsgBox "Debe ingresar un puerto de Salida!", vbInformation
            txtPuertoS.SetFocus
            Exit Sub
        Else
            If (Val(txtPuertoS.Text) < 0 Or Val(txtPuertoS.Text) > (iPhidgetPuertos - 1)) Then
                MsgBox "Puerto Salida debe estar entre 0 y " & iPhidgetPuertos - 1 & "!", vbInformation
                txtPuertoS.SetFocus
                Exit Sub
            End If
        End If
    End If
End If
sSql = "select * from tcontrol where id=" & idControl
With objRs_
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If .EOF Then .AddNew
    !nombre = UCase(Trim(txtNombre.Text))
    !activa = IIf((chkActiva.Value = vbChecked), -1, 0)
    !puerto_e = Val(Trim(txtPuertoE.Text))
    !puerto_s = Val(Trim(txtPuertoS.Text))
    !terminal = sTerminal
    .UpDate
    idControl = !id
    .Close
End With
Set objRs_ = Nothing
Me.Hide
Exit Sub
errH:
If Err.Number = -2147467259 Then
    objRs_.CancelUpdate
    MsgBox "El nombre ingresado ya existe!", vbInformation
Else
    sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdAceptar_Click"
    subLog sERR
End If
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
Public Sub subCargaControl()
Dim objRs_ As New ADODB.Recordset
On Local Error GoTo errH
sSql = "select * from tcontrol where id=" & idControl
Set objRs_ = objCon.Execute(sSql)
If Not objRs_.EOF Then
    txtNombre.Text = "" & objRs_!nombre
    If IsNull(objRs_!activa) Then
        chkActiva.Value = vbUnchecked
    Else
        chkActiva.Value = IIf(objRs_!activa, vbChecked, vbUnchecked)
    End If
    txtPuertoE.Text = Val("" & objRs_!puerto_e)
    txtPuertoS.Text = Val("" & objRs_!puerto_s)
End If
Set objRs_ = Nothing
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdAceptar_Click"
subLog sERR
End Sub

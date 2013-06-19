VERSION 5.00
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.0#0"; "A1AControles.ocx"
Begin VB.Form frmAnotacion 
   BackColor       =   &H00F8D88F&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11040
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin A1AControles.A1ATextBox txtAnotacion 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
      ColorFoco       =   7598073
   End
   Begin VB.Image cmdCancelar 
      Height          =   555
      Left            =   7080
      MouseIcon       =   "frmAnotacion.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmAnotacion.frx":030A
      Top             =   600
      Width           =   1725
   End
   Begin VB.Image cmdAceptar 
      Height          =   555
      Left            =   9120
      MouseIcon       =   "frmAnotacion.frx":3598
      MousePointer    =   99  'Custom
      Picture         =   "frmAnotacion.frx":38A2
      Top             =   600
      Width           =   1725
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00A95900&
      Height          =   1215
      Left            =   120
      Top             =   120
      Width           =   10815
   End
End
Attribute VB_Name = "frmAnotacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAceptar_Click()
On Local Error GoTo errH
Dim objRs_ As New ADODB.Recordset

If Trim(txtAnotacion.Text) <> vbNullString Then
    sSql = "select * from tanotaciones where id=0"
    With objRs_
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenKeyset, adLockOptimistic
        .AddNew
        !documento = frmPrincipal.sDocTmp
        !fecha_hora = fnFecha(Now, True)
        !anotacion = Trim(txtAnotacion.Text)
        .UpDate
        .Close
    End With
    Set objRs_ = Nothing
    Unload Me
End If
Exit Sub
errH:
subLog "Error " & Err.Number & ". " & Err.Description & "_" & Me.name & "_cmdAceptar_Click"
End Sub

Private Sub cmdCancelar_Click()
Unload Me
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

Private Sub Form_Resize()
Me.Top = Me.Top + (Me.Height * 1.3)
End Sub

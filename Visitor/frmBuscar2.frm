VERSION 5.00
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#19.0#0"; "A1AControles.ocx"
Begin VB.Form frmBuscar2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F8D88F&
   BorderStyle     =   0  'None
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   ControlBox      =   0   'False
   Icon            =   "frmBuscar2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstEmerge 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3630
      ItemData        =   "frmBuscar2.frx":0742
      Left            =   120
      List            =   "frmBuscar2.frx":0744
      TabIndex        =   1
      Top             =   480
      Width           =   6975
   End
   Begin A1AControles.A1ATextBox txtNombre 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00A95900&
      BorderWidth     =   3
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Image imgCerrar 
      Height          =   300
      Left            =   6795
      MouseIcon       =   "frmBuscar2.frx":0746
      MousePointer    =   99  'Custom
      Picture         =   "frmBuscar2.frx":0A50
      Top             =   120
      Width           =   300
   End
   Begin VB.Image imgBuscar 
      Height          =   300
      Left            =   120
      Picture         =   "frmBuscar2.frx":0F42
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "frmBuscar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idPer As Long
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    If lstEmerge.Visible Then lstEmerge.ListIndex = 0
ElseIf KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    Me.Tag = vbNullString
    Unload Me
End If
End Sub
Private Sub imgCerrar_Click()
Me.Tag = vbNullString
Unload Me

End Sub

Private Sub lstEmerge_Click()
If lstEmerge.ListIndex <> -1 Then
    Me.Tag = lstEmerge.ItemData(lstEmerge.ListIndex)
    Me.Hide
End If
End Sub
Private Sub subLista()
On Local Error GoTo errH
If Trim(txtNombre.Text) <> vbNullString Then
    lstEmerge.Clear
    sSql = "select id,evento as nombre from tminuta where evento like '%" & Trim(txtNombre.Text) & "%' order by fecha_hora desc"
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenForwardOnly
        While Not .EOF
            lstEmerge.AddItem !nombre
            lstEmerge.ItemData(lstEmerge.NewIndex) = !Id
            .MoveNext
        Wend
        .Close
    End With
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subLista"
subLog sERR
End Sub

Private Sub txtNombre_txtCambio()
subLista
End Sub

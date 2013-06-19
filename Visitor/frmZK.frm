VERSION 5.00
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Begin VB.Form frmZK 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00909890&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relacionar lectores ZK"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   Icon            =   "frmZK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6660
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00909890&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.ListBox lstAsoc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         ItemData        =   "frmZK.frx":0442
         Left            =   240
         List            =   "frmZK.frx":0444
         TabIndex        =   6
         Top             =   960
         Width           =   4935
      End
      Begin VB.CommandButton cmdAsoc 
         Caption         =   "Asociar"
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin A1AControles.A1AComboBox cmbOrigen 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         bkColor         =   9476240
         ColorFoco       =   7598073
      End
      Begin A1AControles.A1AComboBox cmbDestino 
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         bkColor         =   9476240
         ColorFoco       =   7598073
      End
      Begin VB.Image imgBorrarAsoc 
         Height          =   480
         Left            =   5520
         MouseIcon       =   "frmZK.frx":0446
         MousePointer    =   99  'Custom
         Picture         =   "frmZK.frx":0750
         ToolTipText     =   "Eliminar Actual"
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lector destino:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lector origen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmZK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objZK As New ADODB.Recordset

Private Sub cmbDestino_Click0()
cmbDestino.ZOrder 0
End Sub

Private Sub cmbOrigen_Click()
sSql = "select * from tcontrol_disp where tipo=3 and id<>" & cmbOrigen.itemID & " order by id"
cmbDestino.Limpiar
cmbDestino.itemID = 0
With objZK
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        cmbDestino.addElement "" & !nombre, !id
        .MoveNext
    Wend
End With
End Sub

Private Sub cmbOrigen_Click0()
cmbOrigen.ZOrder 0
End Sub

Private Sub cmdAsoc_Click()
On Local Error GoTo errH
If cmbOrigen.itemID > 0 And cmbDestino.itemID > 0 Then
    sSql = "insert into tzk_asoc(idorigen,iddestino) values (" & cmbOrigen.itemID & "," & cmbDestino.itemID & ")"
    objCon.Execute sSql
    cmbOrigen.itemID = 0: cmbOrigen.Text = vbNullString
    cmbDestino.Limpiar: cmbDestino.Text = vbNullString
    subListar
End If

Exit Sub
errH:
If Err.Number = -2147217873 Then
    MsgBox "Ya existente!", vbInformation
End If
End Sub

Private Sub Form_Load()
sSql = "select * from tcontrol_disp where tipo=3 order by id"
With objZK
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        cmbOrigen.addElement "" & !nombre, !id
        .MoveNext
    Wend
End With
subListar
End Sub
Private Sub subListar()
If modoBD = bdSQL Then
    sSql = "select z.id,c1.nombre as o,c2.nombre as d from tzk_asoc as z "
    sSql = sSql & "join tcontrol_disp c1 on z.idorigen=c1.id "
    sSql = sSql & "join tcontrol_disp c2 on z.iddestino=c2.id order by z.id"
ElseIf modoBD = bdACCESS Then
    sSql = "select * from v_zkasoc2"
End If
lstAsoc.Clear
With objZK
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        lstAsoc.AddItem "" & !o & "==>" & "" & !d
        lstAsoc.ItemData(lstAsoc.NewIndex) = !id
        .MoveNext
    Wend
    .Close
End With

End Sub


Private Sub imgBorrarAsoc_Click()
Dim bResp As Byte
If lstAsoc.ListIndex <> -1 Then
    bResp = MsgBox("Eliminar la asociación seleccionada?", vbYesNo + vbQuestion)
    If bResp = vbYes Then
        objCon.Execute "delete from tzk_asoc where id=" & lstAsoc.ItemData(lstAsoc.ListIndex)
        subListar
    End If
End If
End Sub

VERSION 5.00
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Begin VB.Form frmPuertasAsoc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00909890&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relacionar Puertas Dependientes"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmPuertasAsoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00909890&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.OptionButton optES 
         Appearance      =   0  'Flat
         BackColor       =   &H00909890&
         Caption         =   "Sale"
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
         Height          =   195
         Index           =   1
         Left            =   5280
         TabIndex        =   9
         Tag             =   "S"
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton optES 
         Appearance      =   0  'Flat
         BackColor       =   &H00909890&
         Caption         =   "Entra"
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
         Height          =   195
         Index           =   0
         Left            =   5280
         TabIndex        =   8
         Tag             =   "E"
         Top             =   720
         Width           =   975
      End
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
         ItemData        =   "frmPuertasAsoc.frx":0442
         Left            =   240
         List            =   "frmPuertasAsoc.frx":0444
         TabIndex        =   6
         Top             =   960
         Width           =   4935
      End
      Begin VB.CommandButton cmdAsoc 
         Caption         =   "Asociar"
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Top             =   1440
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuando..."
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
         Left            =   5280
         TabIndex        =   7
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgBorrarAsoc 
         Height          =   480
         Left            =   5520
         MouseIcon       =   "frmPuertasAsoc.frx":0446
         MousePointer    =   99  'Custom
         Picture         =   "frmPuertasAsoc.frx":0750
         ToolTipText     =   "Eliminar Actual"
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depende de..."
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
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "La puerta..."
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
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmPuertasAsoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objZK As New ADODB.Recordset
Dim sES As String
Private Sub cmbDestino_Click0()
cmbDestino.ZOrder 0
End Sub

Private Sub cmbOrigen_Click()
sSql = "select id,nombre from tcontrol where id<>" & cmbOrigen.itemID & " order by id"
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
    If sES = vbNullString Then
        MsgBox "Seleccione Entra o Sale!", vbInformation
    Else
        sSql = "insert into tpuertas_asoc(idcontrol,idcontrol_previo,modo) values (" & cmbOrigen.itemID & "," & cmbDestino.itemID & ",'" & sES & "')"
        objCon.Execute sSql
        cmbOrigen.itemID = 0: cmbOrigen.Text = vbNullString
        cmbDestino.Limpiar: cmbDestino.Text = vbNullString
        optES(0).Value = False
        optES(1).Value = False
        sES = vbNullString
        subListar
    End If
End If

Exit Sub
errH:
If Err.Number = -2147217873 Then
    optES(0).Value = False
    optES(1).Value = False
    MsgBox "Ya existente!", vbInformation
End If
End Sub

Private Sub Form_Load()
sSql = "select id,nombre from tcontrol order by id"
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
    sSql = "select z.id,c1.nombre as o,c2.nombre as d,z.modo from tpuertas_asoc as z "
    sSql = sSql & "join tcontrol c1 on z.idcontrol=c1.id "
    sSql = sSql & "join tcontrol c2 on z.idcontrol_previo=c2.id order by z.id"
ElseIf modoBD = bdACCESS Then
    sSql = "select * from v_zkasoc2"
End If
lstAsoc.Clear
With objZK
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        lstAsoc.AddItem "" & !o & ">>>" & "" & !d & ">" & !modo
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
        objCon.Execute "delete from tpuertas_asoc where id=" & lstAsoc.ItemData(lstAsoc.ListIndex)
        subListar
    End If
End If
End Sub

Private Sub optES_Click(Index As Integer)
sES = optES(Index).Tag
End Sub

VERSION 5.00
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#19.0#0"; "A1AControles.ocx"
Begin VB.Form frmMinuta 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F8D88F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minuta Digital"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12255
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMinuta.frx":0000
   ScaleHeight     =   3360
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Tag             =   "3255"
   Begin VB.ListBox lstEmerge 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      ItemData        =   "frmMinuta.frx":10B92
      Left            =   10440
      List            =   "frmMinuta.frx":10B94
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRelación 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   5205
      TabIndex        =   0
      Top             =   525
      Width           =   6735
   End
   Begin A1AControles.A1ATextBox txtFecha 
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin VB.CheckBox chkReloj 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D88F&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   510
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4800
      Top             =   2640
   End
   Begin VB.Timer tmrFecha 
      Interval        =   1000
      Left            =   4080
      Top             =   2640
   End
   Begin A1AControles.A1ATextBox txtHora 
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin A1AControles.A1ATextBox txtEvento 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin A1AControles.A1ATextBox txtReportadoa 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin A1AControles.A1AComboBox cmbTipoComunica 
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   1680
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1AComboBox cmbProceso 
      Height          =   315
      Left            =   8700
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      bkColor         =   16308367
   End
   Begin A1AControles.A1ATextBox txtAnotacion 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   556
      Text            =   ""
      bkColor         =   16308367
      passChar        =   ""
   End
   Begin VB.Image imgBuscar 
      Height          =   300
      Left            =   11760
      MouseIcon       =   "frmMinuta.frx":10B96
      MousePointer    =   99  'Custom
      Picture         =   "frmMinuta.frx":10EA0
      Top             =   1080
      Width           =   300
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00AC5D03&
      Height          =   315
      Index           =   10
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   6900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Relación con:"
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
      Left            =   3840
      TabIndex        =   15
      Top             =   517
      Width           =   1305
   End
   Begin VB.Image cmdCancelar 
      Height          =   555
      Left            =   8160
      MouseIcon       =   "frmMinuta.frx":112A5
      MousePointer    =   99  'Custom
      Picture         =   "frmMinuta.frx":115AF
      Top             =   2640
      Width           =   1725
   End
   Begin VB.Image cmdAceptar 
      Height          =   555
      Left            =   10320
      MouseIcon       =   "frmMinuta.frx":1483D
      MousePointer    =   99  'Custom
      Picture         =   "frmMinuta.frx":14B47
      Top             =   2640
      Width           =   1725
   End
   Begin VB.Image imgFecha 
      Enabled         =   0   'False
      Height          =   240
      Left            =   720
      MouseIcon       =   "frmMinuta.frx":17DD5
      MousePointer    =   99  'Custom
      Picture         =   "frmMinuta.frx":180DF
      ToolTipText     =   "Seleccionar fecha"
      Top             =   510
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   240
      Picture         =   "frmMinuta.frx":1832E
      Top             =   517
      Width           =   240
   End
   Begin VB.Image imgMinuta 
      Height          =   435
      Left            =   4480
      MouseIcon       =   "frmMinuta.frx":18670
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   3195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anotación:"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso:"
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
      Left            =   8700
      TabIndex        =   12
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de comunicación:"
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
      Left            =   5280
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reportado a:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Evento:"
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
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha/Hora del Evento"
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
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   2160
   End
End
Attribute VB_Name = "frmMinuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bModo As Byte
Dim objRstMinuta As New ADODB.Recordset
Dim idRelacion As Variant
Dim idMinuta As Variant
Dim bBuscar As Boolean
Dim txtDestino As Control
Private Sub A1ATextBox1_txtCambio()

End Sub

Private Sub chkReloj_Click()
If chkReloj.Value = vbChecked Then
    tmrFecha.Enabled = True
    imgFecha.Enabled = False
Else
    tmrFecha.Enabled = False
    imgFecha.Enabled = True
End If
End Sub

Private Sub cmbProceso_Click()
On Local Error GoTo errH:
Dim iElem As Long, sElem As String
If cmbProceso.itemID = -1 Then
    frmNuevo.Show vbModal
    sElem = frmNuevo.Tag
    Unload frmNuevo
    If sElem <> vbNullString Then
        sSql = "select * from tproceso where id=0"
        With objRstA
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            .AddNew
            !nombre = sElem
            .UpDate
            iElem = !Id
            .Close
        End With
        subProceso
        cmbProceso.mostrarItem iElem
    End If
End If
Exit Sub
errH:
If objRstA.State = adStateOpen Then
    objRstA.CancelUpdate
    objRstA.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmbProceso_click"
subLog sERR
End Sub

Private Sub cmbProceso_Click0()
cmbProceso.ZOrder 0
End Sub

Private Sub cmbTipoComunica_Click()
On Local Error GoTo errH:
Dim iElem As Long, sElem As String
If cmbTipoComunica.itemID = -1 Then
    frmNuevo.Show vbModal
    sElem = frmNuevo.Tag
    Unload frmNuevo
    If sElem <> vbNullString Then
        sSql = "select * from ttipo_comunica where id=0"
        With objRstA
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            .AddNew
            !nombre = sElem
            .UpDate
            iElem = !Id
            .Close
        End With
        subTipo
        cmbTipoComunica.mostrarItem iElem
    End If
End If
Exit Sub
errH:
If objRstA.State = adStateOpen Then
    objRstA.CancelUpdate
    objRstA.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmbTipoComunica_click"
subLog sERR
End Sub

Private Sub cmbTipoComunica_Click0()
cmbTipoComunica.ZOrder 0
End Sub

Private Sub cmdAceptar_Click()
If Val(idMinuta) = 0 Then
    If Trim(txtEvento.Text) = vbNullString Then
        MsgBox "No hay evento para guardar!", vbInformation
    ElseIf Not IsDate(txtFecha.Text) Then
        MsgBox "La valor del campo Fecha no es válido!", vbInformation
    ElseIf Not IsDate(txtHora.Text) Then
        MsgBox "El valor del campo Hora no es válido!", vbInformation
    Else
        sSql = "select * from tminuta where id=0"
        With objRstMinuta
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            .AddNew
            !idRelacion = idRelacion
            !fecha_hora = fnFecha(txtFecha.Text & " " & txtHora.Text, True)
            !evento = Trim(txtEvento.Text)
            !reportado_a = Trim(txtReportadoa.Text)
            !idcomunica = cmbTipoComunica.itemID
            !idproceso = cmbProceso.itemID
            !anotaciones = Trim(txtAnotacion.Text)
            .UpDate
            .Close
        End With
        imgMinuta_Click
    End If
Else
    imgMinuta_Click
End If
End Sub
Private Sub subLimpiar()
idMinuta = 0
bBuscar = True
chkReloj.Value = vbChecked
idRelacion = 0
txtRelación.Text = vbNullString
txtEvento.Text = vbNullString
txtReportadoa.Text = vbNullString
cmbTipoComunica.Limpiar
cmbProceso.Limpiar
txtAnotacion.Text = vbNullString

End Sub
Private Sub cmdCancelar_Click()
imgMinuta_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstEmerge.Visible Then lstEmerge.ListIndex = 0
    SendKeys "{TAB}"
    KeyAscii = 0
ElseIf KeyAscii = vbKeyEscape Then
    imgMinuta_Click
    KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
bBuscar = True
subTipo
subProceso
End Sub
Private Sub subTipo()
On Local Error GoTo errH
sSql = "select * from ttipo_comunica order by nombre"
cmbTipoComunica.Limpiar
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbTipoComunica.addElement "(Nuevo...)", -1
    While Not .EOF
        cmbTipoComunica.addElement !nombre, !Id
        If !defecto Then cmbTipoComunica.porDefectoElUltimo
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subTipo"
subLog sERR

End Sub
Private Sub subProceso()
On Local Error GoTo errH
sSql = "select * from tproceso order by nombre"
cmbProceso.Limpiar
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbProceso.addElement "(Nuevo...)", -1
    While Not .EOF
        cmbProceso.addElement !nombre, !Id
        If !defecto Then cmbProceso.porDefectoElUltimo
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subProceso"
subLog sERR
End Sub

Private Sub imgBuscar_Click()
frmBuscar2.Show vbModal
If frmBuscar2.Tag <> vbNullString Then
    idMinuta = Val(frmBuscar2.Tag)
    Unload frmBuscar2
    sSql = "select * from tminuta where id=" & idMinuta
    Set objRstMinuta = objCon.Execute(sSql)
    chkReloj.Value = vbUnchecked
    txtFecha.Text = FormatDateTime(objRstMinuta!fecha_hora, vbShortDate)
    txtHora.Text = FormatDateTime(objRstMinuta!fecha_hora, vbLongTime)
    txtEvento.Text = "" & objRstMinuta!evento
    txtReportadoa.Text = "" & objRstMinuta!reportado_a
    cmbTipoComunica.mostrarItem objRstMinuta!idcomunica
    cmbProceso.mostrarItem objRstMinuta!idproceso
    txtAnotacion.Text = "" & objRstMinuta!anotaciones
Else
    idMinuta = 0
    Unload frmBuscar2
End If
End Sub

Private Sub imgFecha_Click()
frmCalendario.Show vbModal
If frmCalendario.Tag <> vbNullString Then
    txtFecha.Text = frmCalendario.Tag
End If
Unload frmCalendario
End Sub

Private Sub imgMinuta_Click()
If frmMinuta.Visible Then
    frmMinuta.bModo = 2
    frmMinuta.Move frmPrincipal.Left + ((frmPrincipal.ScaleWidth / 2) - (frmMinuta.ScaleWidth / 2))
    frmMinuta.tmrScroll.Enabled = True
Else
    'frmMinuta.SetFocus
End If
End Sub

Private Sub lstEmerge_Click()
Dim sDest As String
On Local Error GoTo errH
bBuscar = False
idRelacion = lstEmerge.ItemData(lstEmerge.ListIndex)
txtDestino.Text = lstEmerge.Text
lstEmerge.Visible = False
bBuscar = True
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subLimpiar"
subLog sERR
End Sub

Private Sub tmrFecha_Timer()
txtFecha.Text = Date
txtHora.Text = Time
End Sub

Private Sub tmrScroll_Timer()
If bModo = 1 Then
    frmMinuta.Visible = True
    If frmMinuta.ScaleHeight < Val(frmMinuta.Tag) Then
        frmMinuta.Height = frmMinuta.Height + 200
        frmMinuta.Top = (frmPrincipal.Top + frmPrincipal.ScaleHeight) - frmMinuta.ScaleHeight
    Else
        tmrScroll.Enabled = False
    End If
ElseIf bModo = 2 Then
    frmMinuta.Visible = True
    If frmMinuta.ScaleHeight > 30 Then
        frmMinuta.Height = frmMinuta.Height - 200
        frmMinuta.Top = (frmPrincipal.Top + frmPrincipal.ScaleHeight) - frmMinuta.ScaleHeight
    Else
        tmrScroll.Enabled = False
        Unload Me
    End If
End If

End Sub

Private Sub txtRelación_Change()
subLista txtRelación
End Sub
Private Sub subLista(ByRef txt As Control)
If bBuscar Then
    txt.Tag = vbNullString
    lstEmerge.Tag = vbNullString
    lstEmerge.Visible = False
    lstEmerge.Clear
    lstEmerge.Tag = txt.name
    If Trim(txt.Text) <> vbNullString Then
        Select Case txt.name
            Case "txtRelación"
                sSql = "select id,evento as nombre from tminuta where evento like '%" & txt.Text & "%' order by evento"
        End Select
        Set txtDestino = txt
        With objRst
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenForwardOnly
            While Not .EOF
                lstEmerge.AddItem !nombre
                lstEmerge.ItemData(lstEmerge.NewIndex) = !Id
                .MoveNext
            Wend
            If lstEmerge.ListCount >= 1 Then
                lstEmerge.Move txt.Left, txt.Top + txt.Height, txt.Width
                lstEmerge.Visible = True
            End If
            .Close
        End With
    Else
        lstEmerge.Visible = False
    End If
End If
End Sub

Private Sub txtRelación_Validate(Cancel As Boolean)
lstEmerge.Visible = False
End Sub

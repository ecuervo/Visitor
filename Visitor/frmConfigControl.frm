VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Begin VB.Form frmConfigControl 
   BackColor       =   &H00909890&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controles de Acceso"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9375
   ControlBox      =   0   'False
   Icon            =   "frmConfigControl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9375
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00909890&
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
      Height          =   7215
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9135
      Begin VB.Frame fraDisp 
         BackColor       =   &H00909890&
         Caption         =   "Dispositivos Asociados"
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
         Height          =   5415
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   8895
         Begin VB.CheckBox chkLogin 
            Appearance      =   0  'Flat
            BackColor       =   &H00909890&
            Caption         =   "&Login"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5880
            TabIndex        =   9
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkEnrolaVis 
            Appearance      =   0  'Flat
            BackColor       =   &H00909890&
            Caption         =   "Enrola &Visitantes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3480
            TabIndex        =   8
            Top             =   2280
            Width           =   2655
         End
         Begin VB.CheckBox chkEnrolaFun 
            Appearance      =   0  'Flat
            BackColor       =   &H00909890&
            Caption         =   "Enrola &Funcionarios"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   7
            Top             =   2280
            Width           =   3015
         End
         Begin VB.CheckBox chkActivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00909890&
            Caption         =   "&Activo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5880
            TabIndex        =   6
            Top             =   1830
            Width           =   975
         End
         Begin A1AControles.A1ATextBox txtNombre 
            Height          =   315
            Left            =   840
            TabIndex        =   1
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   9476240
            passChar        =   ""
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1AComboBox cmbTipo 
            Height          =   315
            Left            =   5040
            TabIndex        =   2
            Top             =   600
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            bkColor         =   9476240
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1AComboBox cmbModo 
            Height          =   315
            Left            =   840
            TabIndex        =   4
            Top             =   1800
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            bkColor         =   9476240
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1AComboBox cmbPuerto 
            Height          =   315
            Left            =   840
            TabIndex        =   3
            Top             =   1200
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   556
            bkColor         =   9476240
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1AComboBox cmbPersona 
            Height          =   315
            Left            =   3360
            TabIndex        =   5
            Top             =   1800
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            bkColor         =   9476240
            ColorFoco       =   7598073
         End
         Begin MSDataGridLib.DataGrid objGrid 
            Height          =   2655
            Left            =   120
            TabIndex        =   10
            Top             =   2640
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4683
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   9476240
            ForeColor       =   16777215
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   9226
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   9226
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin A1AControles.A1ATextBox txtZkIp 
            Height          =   315
            Left            =   840
            TabIndex        =   19
            Top             =   1200
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   9476240
            passChar        =   ""
            ColorFoco       =   7598073
         End
         Begin VB.Image imgZk 
            Height          =   450
            Left            =   7515
            MouseIcon       =   "frmConfigControl.frx":70E2
            MousePointer    =   99  'Custom
            Picture         =   "frmConfigControl.frx":73EC
            ToolTipText     =   "Relacionar Lectores ZK"
            Top             =   2040
            Width           =   1260
         End
         Begin VB.Image imgBorrarDisp 
            Height          =   480
            Left            =   8040
            MouseIcon       =   "frmConfigControl.frx":91B6
            MousePointer    =   99  'Custom
            Picture         =   "frmConfigControl.frx":94C0
            ToolTipText     =   "Eliminar Actual"
            Top             =   1305
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image imgGuardar 
            Height          =   480
            Left            =   8040
            MouseIcon       =   "frmConfigControl.frx":A102
            MousePointer    =   99  'Custom
            Picture         =   "frmConfigControl.frx":A40C
            ToolTipText     =   "Guardar"
            Top             =   585
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acceso a"
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
            Left            =   3360
            TabIndex        =   18
            Top             =   1560
            Width           =   990
         End
         Begin VB.Label lblPuerto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puerto"
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
            Left            =   840
            TabIndex        =   17
            Top             =   960
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modo"
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
            Left            =   840
            TabIndex        =   16
            Top             =   1560
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
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
            Left            =   5040
            TabIndex        =   15
            Top             =   360
            Width           =   495
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
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   14
            Top             =   360
            Width           =   1815
         End
      End
      Begin A1AControles.A1AComboBox cmbControl 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         bkColor         =   9476240
         ColorFoco       =   7598073
      End
      Begin VB.Image imgPuertasAsoc 
         Height          =   450
         Left            =   6480
         MouseIcon       =   "frmConfigControl.frx":A8DC6
         MousePointer    =   99  'Custom
         Picture         =   "frmConfigControl.frx":A90D0
         ToolTipText     =   "Relacionar puertas dependientes"
         Top             =   412
         Width           =   2190
      End
      Begin VB.Image cmdAceptar 
         Height          =   555
         Left            =   7200
         MouseIcon       =   "frmConfigControl.frx":AC4A2
         MousePointer    =   99  'Custom
         Picture         =   "frmConfigControl.frx":AC7AC
         Top             =   6480
         Width           =   1725
      End
      Begin VB.Image imgEliminarItem 
         Height          =   300
         Left            =   5760
         MouseIcon       =   "frmConfigControl.frx":AFA3A
         MousePointer    =   99  'Custom
         Picture         =   "frmConfigControl.frx":AFD44
         ToolTipText     =   "Eliminar Actual"
         Top             =   480
         Width           =   300
      End
      Begin VB.Image imgEditaControl 
         Height          =   300
         Left            =   5400
         MouseIcon       =   "frmConfigControl.frx":B0236
         MousePointer    =   99  'Custom
         Picture         =   "frmConfigControl.frx":B0540
         ToolTipText     =   "Editar Nombre actual"
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control de Acceso:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmConfigControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim idDisp As Integer
Private Sub cmbControl_Click()
On Local Error GoTo errH:
Dim iElem As Long
If cmbControl.itemID = -1 Then
    frmNuevoControl.Show vbModal
    iElem = frmNuevoControl.idControl
    frmNuevoControl.idControl = 0
    Unload frmNuevoControl
    If iElem > 0 Then
        subCargaControles
        cmbControl.mostrarItem iElem
    End If
Else
    idControl = cmbControl.itemID
    fraDisp.Visible = (cmbControl.itemID > 0)
    subGrid
End If


Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmbControl_Click"
subLog sERR

End Sub

Private Sub cmbControl_Click0()
cmbControl.ZOrder 0
End Sub

Private Sub cmbModo_Click0()
cmbModo.ZOrder 0
End Sub

Private Sub cmbPersona_Click0()
cmbPersona.ZOrder 0
End Sub

Private Sub cmbPuerto_Click0()
cmbPuerto.ZOrder 0
End Sub

Private Sub cmbTipo_Click()
cmbPuerto.itemID = 0
cmbPuerto.Limpiar
txtZkIp.Visible = False
cmbPuerto.Visible = False
If cmbTipo.itemID = 1 Then
    cmbPuerto.Visible = True
    lblPuerto.Caption = "Número de serie del Lector de Huellas"
    Dim objUU As DPFPReadersCollection
    Dim objUI As DPFPReaderDescription
    Dim i As Long
    Set objUU = New DPFPReadersCollection
    For Each objUI In objUU
        i = i + 1
        cmbPuerto.addElement objUI.SerialNumber, i
    Next
    chkEnrolaFun.Caption = "Enrola &Funcionarios"
    chkEnrolaVis.Caption = "Enrola &Visitantes"
    chkLogin.Caption = "Login"
ElseIf cmbTipo.itemID = 2 Then
    cmbPuerto.Visible = True
    lblPuerto.Caption = "Puerto serial del Lector 2D"
    chkEnrolaFun.Caption = "Registra &Funcionarios"
    chkEnrolaVis.Caption = "Registra &Visitantes"
    'chkLogin.Visible = False
    chkLogin.Caption = "Registrar Objetos"
    fnListaPuertosCOM
ElseIf cmbTipo.itemID = 3 Then
    txtZkIp.Visible = True
    lblPuerto.Caption = "Dirección IP del lector ZK"
End If
End Sub
Private Function fnListaPuertosCOM() As String
Dim i As Integer
For i = 1 To 100
    If COMAvailable(i) Then
       cmbPuerto.addElement CStr(i), CLng(i)
    End If
Next i
End Function

Private Sub cmbTipo_Click0()
cmbTipo.ZOrder 0
End Sub

Private Sub cmdAceptar_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    subLimpiar True
ElseIf KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If

End Sub

Private Sub Form_Load()
subCargaControles
cmbTipo.addElement "Lector de Huellas UareU", 1
cmbTipo.addElement "Lector 2D", 2
cmbTipo.addElement "Lector ZK IP", 3

cmbModo.addElement "Entrada", 1
cmbModo.addElement "Salida", 2
cmbModo.addElement "Entrada/Salida", 3

cmbPersona.addElement "Funcionarios", 1
cmbPersona.addElement "Visitantes", 2
cmbPersona.addElement "Ambos", 3


End Sub
Private Sub subCargaControles()
Dim objRs_ As New ADODB.Recordset
On Local Error GoTo errH
cmbControl.Limpiar
cmbControl.itemID = 0
sSql = "select id,nombre from tcontrol where terminal='" & sTerminal & "'"
With objRs_
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbControl.addElement "(Nuevo...)", -1
    While Not .EOF
        cmbControl.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Set objRs_ = Nothing
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subCargaControles"
subLog sERR
End Sub

Private Sub imgPuertasAsoc_Click()
frmPuertasAsoc.Show vbModal, Me
End Sub

Private Sub imgBorrarDisp_Click()
Dim bR As Byte
If idDisp > 0 Then
    bR = MsgBox("Desea eliminar el dispositivo Actual?", vbYesNo + vbQuestion)
    If bR = vbYes Then
        objCon.Execute "delete from tcontrol_disp where id=" & idDisp
        objCon.Execute "delete from tzk_asoc where idorigen=" & idDisp & " or iddestino=" & idDisp
        subLimpiar False
        subGrid
    End If
End If

End Sub

Private Sub imgEditaControl_Click()
Dim tmpID As Long
If cmbControl.itemID > 0 Then
    tmpID = cmbControl.itemID
    Load frmNuevoControl
    frmNuevoControl.lbl.Caption = "Modificando..."
    frmNuevoControl.idControl = cmbControl.itemID
    frmNuevoControl.subCargaControl
    frmNuevoControl.Show vbModal
    frmNuevoControl.idControl = 0
    Unload frmNuevoControl
    subCargaControles
    cmbControl.mostrarItem tmpID
End If
End Sub

Private Sub imgEliminarItem_Click()
Dim bR As Byte
If cmbControl.itemID > 0 Then
    bR = MsgBox("Desea eliminar el control Actual?", vbYesNo + vbQuestion)
    If bR = vbYes Then
        objCon.Execute "delete from tcontrol_disp where idcontrol=" & cmbControl.itemID
        objCon.Execute "delete from tcontrol where id=" & cmbControl.itemID
        objCon.Execute "delete from tpuertas_asoc where idcontrol=" & cmbControl.itemID & " or idcontrol_previo=" & cmbControl.itemID
        
        subLimpiar True
        subCargaControles
    End If
End If
End Sub

Private Sub imgGuardar_Click()
On Local Error GoTo errH
Dim objRs_ As New ADODB.Recordset
If cmbControl.itemID > 0 Then
    If Trim(txtNombre.Text) = vbNullString Then
        MsgBox "Debe inresar un nombre de dispositivo!", vbInformation
        txtNombre.SetFocus
        Exit Sub
    End If
    If cmbTipo.itemID = 0 Then
        MsgBox "Seleccione Tipo!", vbInformation
        cmbTipo.SetFocus
        Exit Sub
    End If
    If cmbTipo.itemID = 1 Or cmbTipo.itemID = 2 Then
        If cmbPuerto.itemID = 0 Then
            MsgBox "Seleccione " & lblPuerto.Caption & "!", vbInformation
            cmbPuerto.SetFocus
            Exit Sub
        End If
    ElseIf cmbTipo.itemID = 3 Then
        If Trim(txtZkIp.Text) = vbNullString Then
            MsgBox "Dirección IP no válida!", vbInformation
            txtZkIp.SetFocus
            Exit Sub
        End If
    End If
    If cmbModo.itemID = 0 Then
        MsgBox "Seleccione Modo!", vbInformation
        cmbModo.SetFocus
        Exit Sub
    End If
    If cmbPersona.itemID = 0 Then
        MsgBox "Seleccione Acceso a!", vbInformation
        cmbPersona.SetFocus
        Exit Sub
    End If
    sSql = "select * from tcontrol_disp where id=" & idDisp
    With objRs_
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenKeyset, adLockOptimistic
        If .EOF Then .AddNew
        !idControl = cmbControl.itemID
        !nombre = UCase(Trim(txtNombre.Text))
        !tipo = cmbTipo.itemID
        If cmbTipo.itemID = 1 Or cmbTipo.itemID = 2 Then
            !puerto = cmbPuerto.Text
        ElseIf cmbTipo.itemID = 3 Then
            !puerto = Trim(txtZkIp.Text)
        End If
        !modo = cmbModo.itemID
        !persona = cmbPersona.itemID
        !activo = IIf((chkActivo.Value = vbChecked), -1, 0)
        !enrola_vis = IIf((chkEnrolaVis.Value = vbChecked), -1, 0)
        !enrola_fun = IIf((chkEnrolaFun.Value = vbChecked), -1, 0)
        !login = IIf((chkLogin.Value = vbChecked), -1, 0)
        .UpDate
        'idDisp = !id
        .Close
    End With
    Set objRs_ = Nothing
    idDisp = 0
    subGrid
    subLimpiar False
End If
Exit Sub
errH:
If Err.Number = -2147467259 Then
    MsgBox "Ya existe un dispositivo con el nombre " & txtNombre.Text & "!", vbInformation
    objRs_.CancelUpdate
    txtNombre.SetFocus
Else
    sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subCargaControles"
    subLog sERR
End If
End Sub
Private Sub subGrid()
Dim objRs_ As New ADODB.Recordset
sSql = "select * from tcontrol_disp where idcontrol=" & cmbControl.itemID
Set objRs_ = objCon.Execute(sSql)
Set objGrid.DataSource = objRs_
objGrid.Columns("id").Visible = False
objGrid.Columns("idcontrol").Visible = False

End Sub

Private Sub imgZK_Click()
frmZK.Show vbModal, Me
End Sub

Private Sub objGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Local Error GoTo errH
If objGrid.Columns(0).Caption <> vbNullString Then
    idDisp = objGrid.Columns("id").Value
    subCargaDisp
End If
Exit Sub
errH:
End Sub
Private Sub subCargaDisp()
Dim objRs_ As New ADODB.Recordset
sSql = "select * from tcontrol_disp where id=" & idDisp
Set objRs_ = objCon.Execute(sSql)
txtNombre.Text = "" & objRs_!nombre
cmbTipo.mostrarItem objRs_!tipo
If cmbTipo.itemID = 1 Or cmbTipo.itemID = 2 Then
    cmbPuerto.mostrarItem 0, "" & objRs_!puerto
ElseIf cmbTipo.itemID = 3 Then
    txtZkIp.Text = "" & objRs_!puerto
End If
cmbModo.mostrarItem objRs_!modo
cmbPersona.mostrarItem objRs_!persona
chkActivo.Value = IIf(objRs_!activo, vbChecked, vbUnchecked)
chkEnrolaFun.Value = IIf(objRs_!enrola_fun, vbChecked, vbUnchecked)
chkEnrolaVis.Value = IIf(objRs_!enrola_vis, vbChecked, vbUnchecked)
chkLogin.Value = IIf(objRs_!login, vbChecked, vbUnchecked)
txtNombre.SetFocus
imgBorrarDisp.Visible = True
End Sub
Sub subLimpiar(Total As Boolean)
imgBorrarDisp.Visible = False
idDisp = 0
If Total Then cmbControl.itemID = 0
If Total Then fraDisp.Visible = False
txtNombre.Text = vbNullString
cmbTipo.itemID = 0
cmbPuerto.itemID = 0
cmbModo.itemID = 0
cmbPersona.itemID = 0
chkActivo.Value = vbUnchecked
chkEnrolaFun.Value = vbUnchecked
chkEnrolaVis.Value = vbUnchecked
chkLogin.Value = vbUnchecked
txtZkIp.Text = vbNullString
If Total Then cmbControl.SetFocus Else txtNombre.SetFocus
End Sub

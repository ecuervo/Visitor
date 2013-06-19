VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Begin VB.Form frmHorariosSET 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00909890&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de Horarios"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   Icon            =   "frmHorariosSET.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00909890&
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.ListBox lstFuncionarios 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7230
         ItemData        =   "frmHorariosSET.frx":0442
         Left            =   120
         List            =   "frmHorariosSET.frx":0444
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   840
         Width           =   6795
      End
      Begin A1AControles.A1ATextBox txtNombre 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   9476240
         passChar        =   ""
      End
      Begin A1AControles.A1AComboBox cmbHorarios 
         Height          =   315
         Left            =   7080
         TabIndex        =   4
         Top             =   1560
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         bkColor         =   9476240
         ColorFoco       =   16308367
      End
      Begin ARButtonCtrl.ARButton cmdAsignar 
         Height          =   375
         Left            =   7080
         TabIndex        =   6
         Tag             =   "12484943"
         Top             =   1920
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Asignar"
         ForeColor       =   16777215
         ForeColorOnMouse=   8987923
         BackColorOnMouse=   16777215
         BackColor       =   8987923
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin A1AControles.A1AComboBox cmbHorariosAdd 
         Height          =   315
         Left            =   7080
         TabIndex        =   8
         Top             =   2640
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         bkColor         =   9476240
         ColorFoco       =   16308367
      End
      Begin ARButtonCtrl.ARButton cmdAsignarAdd 
         Height          =   375
         Left            =   7080
         TabIndex        =   9
         Tag             =   "12484943"
         Top             =   3000
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Asignar"
         ForeColor       =   16777215
         ForeColorOnMouse=   8987923
         BackColorOnMouse=   16777215
         BackColor       =   8987923
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin ARButtonCtrl.ARButton cmdVer 
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Tag             =   "12484943"
         Top             =   1920
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   661
         Caption         =   "Ver empleados^"
         ForeColor       =   16777215
         ForeColorOnMouse=   8987923
         BackColorOnMouse=   16777215
         BackColor       =   8987923
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin ARButtonCtrl.ARButton cmdVer2 
         Height          =   375
         Left            =   8400
         TabIndex        =   12
         Tag             =   "12484943"
         Top             =   3000
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   661
         Caption         =   "Ver empleados^"
         ForeColor       =   16777215
         ForeColorOnMouse=   8987923
         BackColorOnMouse=   16777215
         BackColor       =   8987923
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin ARButtonCtrl.ARButton cmdDesmarcar 
         Height          =   315
         Left            =   5760
         TabIndex        =   13
         Tag             =   "12484943"
         Top             =   480
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Desmarcar"
         ForeColor       =   16777215
         ForeColorOnMouse=   8987923
         BackColorOnMouse=   16777215
         BackColor       =   8987923
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin ARButtonCtrl.ARButton cmdQuitar 
         Height          =   375
         Left            =   10200
         TabIndex        =   14
         Tag             =   "12484943"
         Top             =   3000
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   661
         Caption         =   "Quitar"
         ForeColor       =   -2147483631
         ForeColorOnMouse=   8987923
         BackColorOnMouse=   16777215
         BackColor       =   8987923
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
         Enabled         =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horarios adicionales:"
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
         Height          =   240
         Left            =   7080
         TabIndex        =   10
         Top             =   2400
         Width           =   2010
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   7080
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marque de la lista a que funcionarios se les asignará el horario seleccionado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   480
         Left            =   7200
         TabIndex        =   7
         Top             =   600
         Width           =   3930
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horario:"
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
         Height          =   240
         Left            =   7080
         TabIndex        =   5
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listado personal Activo"
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
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2220
      End
   End
End
Attribute VB_Name = "frmHorariosSET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCamb As Boolean

Private Sub cmdAsignarAdd_Click()
On Local Error GoTo errH
Dim bResp As Byte
Dim ix As Integer
Dim objR As New ADODB.Recordset
If lstFuncionarios.SelCount = 0 Then
    MsgBox "Seleccione funcionarios de la lista!", vbInformation
    Exit Sub
End If
If cmbHorariosAdd.itemID <= 0 Then
    MsgBox "Seleccione un horario adicional!", vbInformation
    cmbHorariosAdd.SetFocus
    Exit Sub
End If
bResp = MsgBox("Asignar el horario " & cmbHorarios.Text & " a los " & lstFuncionarios.SelCount & " Funcionarios seleccionados?", vbYesNo + vbQuestion)
If bResp = vbYes Then
    bCamb = False
    For ix = 0 To lstFuncionarios.ListCount - 1
        If lstFuncionarios.Selected(ix) Then
            sSql = "select idhorario from templeados where id=" & lstFuncionarios.ItemData(ix)
            Set objR = objCon.Execute(sSql)
            If cmbHorariosAdd.itemID <> objR!idHorario Then
                sSql = "if not exists(select id from thorarios_add where idempleado=" & lstFuncionarios.ItemData(ix) & " and idhorario=" & cmbHorariosAdd.itemID & ") "
                sSql = sSql & "insert into thorarios_add (idempleado,idhorario) values(" & lstFuncionarios.ItemData(ix) & "," & cmbHorariosAdd.itemID & ")"
                objCon.Execute sSql
                lstFuncionarios.Selected(ix) = False
            End If
        End If
    Next ix
    MsgBox "Asignación realizada satisfactoriamente.", vbInformation
    Set objR = Nothing
    bCamb = True
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdAsignarAdd_Click"
subLog sERR
End Sub

Private Sub cmdAsignar_Click()
On Local Error GoTo errH
Dim bResp As Byte
Dim ix As Integer
If lstFuncionarios.SelCount = 0 Then
    MsgBox "Seleccione funcionarios de la lista!", vbInformation
    Exit Sub
End If
If cmbHorarios.itemID <= 0 Then
    MsgBox "Seleccione un horario!", vbInformation
    cmbHorarios.SetFocus
    Exit Sub
End If
bResp = MsgBox("Asignar el horario " & cmbHorarios.Text & " a los " & lstFuncionarios.SelCount & " Funcionarios seleccionados?", vbYesNo + vbQuestion)
If bResp = vbYes Then
    bCamb = False
    For ix = 0 To lstFuncionarios.ListCount - 1
        If lstFuncionarios.Selected(ix) Then
            sSql = "update templeados set idhorario=" & cmbHorarios.itemID & " where id=" & lstFuncionarios.ItemData(ix)
            objCon.Execute (sSql)
            lstFuncionarios.Selected(ix) = False
        End If
    Next ix
    MsgBox "Asignación realizada satisfactoriamente.", vbInformation
    bCamb = True
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdAsignar_Click"
subLog sERR
End Sub

Private Sub cmdDesmarcar_Click()
On Local Error GoTo errH
Dim ix As Integer
For ix = 0 To lstFuncionarios.ListCount - 1
    bCamb = False
    lstFuncionarios.Selected(ix) = False
    lstFuncionarios.Refresh
    bCamb = True
Next ix
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdDesmarcar_Click"
subLog sERR
End Sub

Private Sub cmdQuitar_Click()
On Local Error GoTo errH
Dim bResp As Byte
Dim ix As Integer
bResp = MsgBox("Quitar el horario adicional asignado a " & lstFuncionarios.SelCount & " funcionarios seleccionados?", vbQuestion + vbYesNo)
If bResp = vbYes Then
    bCamb = False
    For ix = 0 To lstFuncionarios.ListCount - 1
        If lstFuncionarios.Selected(ix) Then
            sSql = "delete from thorarios_add where idempleado=" & lstFuncionarios.ItemData(ix) & " and idhorario=" & cmbHorariosAdd.itemID
            objCon.Execute sSql
            lstFuncionarios.Selected(ix) = False
        End If
    Next ix
    bCamb = True
    MsgBox "Se ha quitado el horario adicional.", vbInformation
    cmbHorariosAdd.itemID = 0
    cmbHorariosAdd.Text = vbNullString
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdQuitar_Click"
subLog sERR
End Sub

Private Sub cmdVer_Click()
On Local Error GoTo errH
Dim objR As New ADODB.Recordset
Dim ix As Integer
If cmbHorarios.itemID <> 0 Then
    sSql = "select id,idhorario from templeados"
    With objR
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenKeyset, adLockReadOnly
        bCamb = False
        For ix = 0 To lstFuncionarios.ListCount - 1
            .Find "id=" & lstFuncionarios.ItemData(ix), , adSearchForward, 1
            lstFuncionarios.Selected(ix) = (!idHorario = cmbHorarios.itemID)
        Next ix
        bCamb = True
        .Close
    End With
    lstFuncionarios.Refresh
Else
    MsgBox "Seleccione un horario!", vbInformation
    cmbHorarios.SetFocus
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdVer_Click"
subLog sERR
End Sub

Private Sub cmdVer2_Click()
On Local Error GoTo errH
Dim objR As New ADODB.Recordset
Dim ix As Integer
If cmbHorariosAdd.itemID <> 0 Then
    sSql = "select id,idempleado,idhorario from thorarios_add"
    With objR
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenKeyset, adLockReadOnly
        bCamb = False
        For ix = 0 To lstFuncionarios.ListCount - 1
            .Find "idempleado=" & lstFuncionarios.ItemData(ix), , adSearchForward, 1
            If Not .EOF Then
                lstFuncionarios.Selected(ix) = (!idHorario = cmbHorariosAdd.itemID)
            Else
                lstFuncionarios.Selected(ix) = False
            End If
        Next ix
        bCamb = True
        .Close
    End With
Else
    MsgBox "Seleccione un horario!", vbInformation
    cmbHorariosAdd.SetFocus
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdVer2_Click"
subLog sERR
End Sub

Private Sub Form_Load()
bCamb = True
subLista
End Sub

Private Sub lstFuncionarios_Click()
Dim objF As New ADODB.Recordset
If bCamb Then
    If lstFuncionarios.ListIndex <> -1 Then
        sSql = "select idhorario from templeados where id=" & lstFuncionarios.ItemData(lstFuncionarios.ListIndex)
        Set objF = objCon.Execute(sSql)
        If Val("" & objF!idHorario) > 0 Then
            cmbHorarios.mostrarItem CLng(Val("" & objF!idHorario))
        Else
            cmbHorarios.itemID = 0
            cmbHorarios.Text = vbNullString
        End If
        sSql = "select idhorario from thorarios_add where idempleado=" & lstFuncionarios.ItemData(lstFuncionarios.ListIndex)
        Set objF = objCon.Execute(sSql)
        cmdQuitar.Enabled = False
        If Not objF.EOF Then
            If Val("" & objF!idHorario) > 0 Then
                cmbHorariosAdd.mostrarItem CLng(Val("" & objF!idHorario))
                cmdQuitar.Enabled = True
                cmdQuitar.ForeColor = vbWhite
            End If
        Else
            cmbHorariosAdd.itemID = 0
            cmbHorariosAdd.Text = vbNullString
        End If
    End If
End If
End Sub

Private Sub txtNombre_txtCambio()
subLista
End Sub
Private Sub subLista()
On Local Error GoTo errH
If Trim(txtNombre.Text) <> vbNullString Then
    lstFuncionarios.Clear
    If modoBD = bdSQL Then
        sSql = "select id,'[' + isnull(codigo,'')+ '][' + isnull(documento,'') + '] ' + isnull(nombre,'') + ' ' + isnull(apellidos,'') as nombre from templeados where isnull(documento,'') + isnull(nombre,'') + ' ' + isnull(apellidos,'') like '%" & Trim(txtNombre.Text) & "%' and abs(activo)=1 and idcargo is not null order by nombre"
    ElseIf modoBD = bdACCESS Then
        sSql = "select id,'[' & codigo & '][' & documento & '] ' & nombre & ' ' & apellidos as nombre from templeados where documento & nombre & ' ' & apellidos like '%" & Trim(txtNombre.Text) & "%' and abs(activo)=1 and (not(idcargo)is null) order by nombre"
    End If
    
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenForwardOnly
        While Not .EOF
            lstFuncionarios.AddItem !nombre
            lstFuncionarios.ItemData(lstFuncionarios.NewIndex) = !id
            .MoveNext
        Wend
        .Close
    End With
Else
    lstFuncionarios.Clear
    If modoBD = bdSQL Then
        sSql = "select id,'[' + isnull(codigo,'')+ '][' + isnull(documento,'') + '] ' + isnull(nombre,'') + ' ' + isnull(apellidos,'') as nombre from templeados where abs(activo)=1 and idcargo is not null order by nombre"
    ElseIf modoBD = bdACCESS Then
        sSql = "select id,'[' & codigo & '][' & documento & '] ' & nombre & ' ' & apellidos as nombre from templeados where abs(activo)=1 and (not(idcargo)is null) order by nombre"
    End If
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenForwardOnly
        While Not .EOF
            lstFuncionarios.AddItem !nombre
            lstFuncionarios.ItemData(lstFuncionarios.NewIndex) = !id
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


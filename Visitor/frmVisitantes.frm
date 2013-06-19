VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.1#0"; "ARBUTTON.OCX"
Begin VB.Form frmVisitantes 
   BackColor       =   &H00A46B2E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visitantes"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   Icon            =   "frmVisitantes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8670
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00A46B2E&
      Height          =   5775
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   8415
      Begin VB.Timer tmrCam 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5880
         Top             =   4680
      End
      Begin VB.TextBox txtFechaF 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4470
         TabIndex        =   8
         Top             =   4005
         Width           =   1815
      End
      Begin VB.TextBox txtFechaI 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2550
         TabIndex        =   7
         Top             =   4005
         Width           =   1815
      End
      Begin VB.TextBox txtTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   210
         TabIndex        =   6
         Top             =   4020
         Width           =   2175
      End
      Begin VB.TextBox txtRh 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4890
         TabIndex        =   4
         Top             =   2340
         Width           =   1335
      End
      Begin VB.TextBox txtSexo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4890
         TabIndex        =   3
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   210
         TabIndex        =   1
         Top             =   1500
         Width           =   3735
      End
      Begin VB.ComboBox cmbDependencias 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         ItemData        =   "frmVisitantes.frx":2982
         Left            =   225
         List            =   "frmVisitantes.frx":2984
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3135
         Width           =   6015
      End
      Begin VB.TextBox txtApellidos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   225
         TabIndex        =   2
         Top             =   2340
         Width           =   4335
      End
      Begin VB.TextBox txtDoc1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   4335
      End
      Begin ARButtonCtrl.ARButton cmdGrabar 
         Height          =   495
         Left            =   6480
         TabIndex        =   9
         Tag             =   "12484943"
         Top             =   5040
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   873
         Caption         =   "&Guardar"
         ForeColor       =   16777215
         ForeColorOnMouse=   12484943
         BackColorOnMouse=   16777215
         BackColor       =   12484943
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin ARButtonCtrl.ARButton cmdBusca 
         Height          =   435
         Left            =   4080
         TabIndex        =   20
         Tag             =   "12484943"
         ToolTipText     =   "Buscar Funcionario"
         Top             =   1440
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   767
         Caption         =   "B"
         ForeColor       =   16777215
         ForeColorOnMouse=   12484943
         BackColorOnMouse=   16777215
         BackColor       =   12484943
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin VB.Image imgFoto 
         Appearance      =   0  'Flat
         Height          =   2340
         Left            =   6480
         Picture         =   "frmVisitantes.frx":2986
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1755
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   5
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   3960
         Width           =   1860
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4560
         TabIndex        =   19
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inicio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   4
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   3960
         Width           =   1860
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono/Extensión:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   2310
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   3
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3960
         Width           =   2340
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RH:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4800
         TabIndex        =   16
         Top             =   1920
         Width           =   450
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   2
         Left            =   4800
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4800
         TabIndex        =   15
         Top             =   1080
         Width           =   675
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   1
         Left            =   4800
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dependencia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   2760
         Width           =   1605
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   6
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   3900
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   7
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   4500
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   8
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   3120
         Width           =   6180
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   0
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   4500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Número de documento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   135
         TabIndex        =   13
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   135
         TabIndex        =   12
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Image imgHuella 
         Height          =   2100
         Left            =   6480
         Picture         =   "frmVisitantes.frx":3C45
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmVisitantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim idFuncionario As Long
Dim idDep As Long
Dim bFoto As Boolean

Private Sub cmbDependencias_Click()
On Local Error GoTo errH
Dim sNom As String
If cmbDependencias.ListIndex <> -1 Then
    idDep = cmbDependencias.ItemData(cmbDependencias.ListIndex)
    If idDep = 0 Then
        sNom = Trim(InputBox("Nombre de la dependencia:"))
        If sNom <> vbNullString Then
            sSql = "insert into tdependencias(nombre) values ('" & UCase(sNom) & "')"
            objCon.Execute sSql
            subDependencias
            cmbDependencias.Text = sNom
        End If
    End If
End If
Exit Sub
errH:
If Err.Number = -2147217873 Then
    cmbDependencias.Text = sNom
End If
End Sub

Private Sub cmdBusca_Click()
frmBuscar.Show vbModal
If Val(frmBuscar.Tag) <> 0 Then
    idFuncionario = Val(frmBuscar.Tag)
    Unload frmBuscar
    subDatos "id", CStr(idFuncionario)
End If
End Sub

Private Sub cmdGrabar_Click()
If Trim(txtDoc1.Text) = vbNullString Then
    MsgBox "Ingrese Número de documento!", vbInformation
    txtDoc1.Text = vbNullString
    txtDoc1.SetFocus
End If
sSql = "select * from tfuncionarios where id=" & idFuncionario
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If .EOF Then .AddNew
    !documento = Trim(txtDoc1.Text)
    !nombre = Trim(txtNombre.Text)
    !apellidos = Trim(txtApellidos.Text)
    !sexo = Trim(txtSexo.Text)
    !rh = Trim(txtRh.Text)
    !idDependencia = idDep
    !tel = Trim(txtTelefono.Text)
    If IsDate(Trim(txtFechaI.Text)) Then !fechai = fnFecha(Trim(txtFechaI.Text), False)
    If IsDate(Trim(txtFechaF.Text)) Then !fechaf = fnFecha(Trim(txtFechaF.Text), False)
    If bFoto Then
        SavePicture imgFoto.Picture, App.Path & "\tmpFoto"
        ConvertBMPtoJPG App.Path & "\tmpFoto", App.Path & "\tmpFoto" & ".jpg", True, 50, False
        fnGuardaFoto !foto, App.Path & "\tmpFoto.jpg"
    End If
    If bModificaHuella Then
        SavePicture imgHuella.Picture, App.Path & "\tmpHuella"
        ConvertBMPtoJPG App.Path & "\tmpHuella", App.Path & "\tmpHuella" & ".jpg", True, 50, False
        fnGuardaFoto !huella, App.Path & "\tmpHuella.jpg"
        !enrola = bHuellaMinuciasCAP
    End If
    .Update
    idFuncionario = !Id
    .Close
End With
subLimpiar
End Sub
Private Sub subLimpiar()
Dim cTrl As Control
For Each cTrl In Me
    If TypeName(cTrl) = "TextBox" Then
        cTrl.Text = vbNullString
    ElseIf TypeName(cTrl) = "ComboBox" Then
        cTrl.ListIndex = -1
    End If
Next
txtDoc1.SetFocus
idFuncionario = 0
idDep = 0
bFoto = False
imgFoto.Picture = LoadPicture(App.Path & "\imgfoto.jpg")
imgHuella.Picture = LoadPicture(App.Path & "\imghuella.jpg")

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
ElseIf KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    subLimpiar
End If
End Sub

Private Sub Form_Load()
subDependencias
End Sub
Sub subDependencias()
cmbDependencias.Clear
sSql = "select * from tdependencias order by nombre"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbDependencias.AddItem "(Nuevo...)": cmbDependencias.ItemData(cmbDependencias.NewIndex) = 0
    While Not objRst.EOF
        cmbDependencias.AddItem !nombre
        cmbDependencias.ItemData(cmbDependencias.NewIndex) = !Id
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub imgFoto_Click()
If bCam Then tmrCam.Enabled = Not tmrCam.Enabled
End Sub

Private Sub imgHuella_Click()
Set frmEnrola.objImagen = imgHuella
frmPrincipal.objUareU.StopCapture
frmEnrola.Show vbModal
frmPrincipal.objUareU.StartCapture
End Sub

Private Sub tmrCam_Timer()
imgFoto.Picture = frmPrincipal.objVideo.GrabFrame
bFoto = True
End Sub

Private Sub txtDoc1_Validate(Cancel As Boolean)
subDatos "documento", txtDoc1.Text
End Sub
Public Sub subDatos(sCampo As String, sValor As String)
sDoc = Trim(sValor)
If sDoc <> vbNullString Then
    sDoc = Replace(sDoc, ".", "")
    sDoc = Replace(sDoc, ",", "")
    sDoc = Replace(sDoc, "-", "")
    sSql = "select * from tfuncionarios where " & sCampo & "='" & sDoc & "'"
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenKeyset, adLockOptimistic
        If Not .EOF Then
            idFuncionario = !Id
            txtDoc1.Text = "" & !documento
            txtNombre.Text = "" & !nombre
            txtApellidos.Text = "" & !apellidos
            txtRh.Text = "" & !rh
            txtSexo.Text = "" & !sexo
            sSql = "select nombre from tdependencias where id=" & !idDependencia
            Set objRstA = objCon.Execute(sSql)
            If Not objRstA.EOF Then cmbDependencias.Text = objRstA!nombre
            txtTelefono.Text = "" & !tel
            txtFechaI.Text = "" & !fechai
            txtFechaF.Text = "" & !fechaf
            fnLeeFoto !foto, imgFoto
            fnLeeFoto !huella, imgHuella
        End If
        .Close
    End With
End If
End Sub
Private Sub txtFechaI_Validate(Cancel As Boolean)
If Trim(txtFechaI) <> vbNullString Then
    If Not IsDate(Trim(txtFechaI.Text)) Then
        MsgBox "Fecha no válida!", vbInformation
        txtFechaI.Text = vbNullString
        Cancel = True
    End If
End If
End Sub
Private Sub txtFechaF_Validate(Cancel As Boolean)
If Trim(txtFechaF) <> vbNullString Then
    If Not IsDate(Trim(txtFechaF.Text)) Then
        MsgBox "Fecha no válida!", vbInformation
        txtFechaF.Text = vbNullString
        Cancel = True
    End If
End If
End Sub

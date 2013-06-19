VERSION 5.00
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.0#0"; "A1AControles.ocx"
Begin VB.Form frmPerfiles 
   BackColor       =   &H0073EFF9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Perfiles de Seguridad"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "frmPerfiles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7590
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H0073EFF9&
      Height          =   6735
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame2 
         BackColor       =   &H0073EFF9&
         Caption         =   "Permitir / Denegar acceso a (Inventario)"
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
         Height          =   3375
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   7095
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Control Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   25
            Left            =   3480
            TabIndex        =   30
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Entrega 2do Piso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   24
            Left            =   3480
            TabIndex        =   29
            Top             =   2400
            Width           =   1935
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Proveedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   23
            Left            =   3480
            TabIndex        =   28
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Clientes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   22
            Left            =   5280
            TabIndex        =   27
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Almacenes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   21
            Left            =   3480
            TabIndex        =   26
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Configuración"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   20
            Left            =   5280
            TabIndex        =   25
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Reportes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   19
            Left            =   3480
            TabIndex        =   24
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Inventario Ventas Remisionado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   18
            Left            =   3480
            TabIndex        =   23
            Top             =   960
            Width           =   3495
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Inventario Ventas Facturación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   17
            Left            =   3480
            TabIndex        =   22
            Top             =   600
            Width           =   3495
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Inventario Ventas Entrada a Bodega"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   16
            Left            =   3480
            TabIndex        =   21
            Top             =   240
            Width           =   3495
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Producción 5. Plastificado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   15
            Left            =   120
            TabIndex        =   20
            Top             =   2760
            Width           =   3255
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Producción 4.Encuadernación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   19
            Top             =   2400
            Width           =   3255
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Producción 3.Plegado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   18
            Top             =   2040
            Width           =   3255
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Producción 2.Impresión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   17
            Top             =   1680
            Width           =   3255
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Producción 1. Papel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   3255
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Producción Orden de Producción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   3255
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Producción Materia Prima"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   2535
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Producto Terminado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2055
         End
      End
      Begin A1AControles.A1ATextBox txtNombre 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
      End
      Begin VB.Frame fraOpciones 
         BackColor       =   &H0073EFF9&
         Caption         =   "Permitir / Denegar acceso a"
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
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   7095
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Registrar &Visitantes Nuevos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   2775
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "&Minuta Digital"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   6
            Left            =   3120
            TabIndex        =   7
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "&Backup"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   5
            Left            =   1680
            TabIndex        =   6
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "&Herramientas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Re&portes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   3
            Left            =   4800
            TabIndex        =   4
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "&Registro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   2
            Left            =   3120
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "&Parámetros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   2
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkOpciones 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "&Funcionarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Image cmdCancelar 
         Height          =   555
         Left            =   120
         MouseIcon       =   "frmPerfiles.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "frmPerfiles.frx":074C
         Top             =   6000
         Width           =   1725
      End
      Begin VB.Image cmdAceptar 
         Height          =   555
         Left            =   5400
         MouseIcon       =   "frmPerfiles.frx":39DA
         MousePointer    =   99  'Custom
         Picture         =   "frmPerfiles.frx":3CE4
         Top             =   6000
         Width           =   1725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Perfil"
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
         Left            =   135
         TabIndex        =   10
         Top             =   240
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmPerfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public idPerfil As Integer
Private Sub cmbPermisos_Click0()
cmbPermisos.ZOrder 0
End Sub

Private Sub chkOpciones_Click(Index As Integer)
Dim i As Integer
If Index = 25 Then
    For i = 8 To 24
        chkOpciones(i).Value = chkOpciones(25).Value
    Next i
End If
End Sub

Private Sub cmdAceptar_Click()
Dim id As Integer, bCheck As Boolean
On Local Error GoTo errH
If Trim(txtNombre.Text) = vbNullString Then
    MsgBox "Ingrese un nombre para el perfil!", vbInformation
    txtNombre.SetFocus
    Exit Sub
End If
For id = 0 To chkOpciones.Count - 1
    If chkOpciones(id).Value = vbChecked Then
        bCheck = True
        Exit For
    End If
Next id
If Not bCheck Then
    MsgBox "Selecione al menos una casilla!", vbInformation
    Exit Sub
End If
sSql = "select * from tperfiles where id=" & idPerfil
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If .EOF Then .AddNew
    !nombre = UCase(Trim(txtNombre.Text))
    !funcionarios = IIf((chkOpciones(0).Value = vbChecked), -1, 0)
    !parámetros = IIf((chkOpciones(1).Value = vbChecked), -1, 0)
    !registro = IIf((chkOpciones(2).Value = vbChecked), -1, 0)
    !reportes = IIf((chkOpciones(3).Value = vbChecked), -1, 0)
    !herramientas = IIf((chkOpciones(4).Value = vbChecked), -1, 0)
    !backup1 = IIf((chkOpciones(5).Value = vbChecked), -1, 0)
    !minuta_digital = IIf((chkOpciones(6).Value = vbChecked), -1, 0)
    !reg_vis_nue = IIf((chkOpciones(7).Value = vbChecked), -1, 0)
    
    !pi8 = IIf((chkOpciones(8).Value = vbChecked), -1, 0)
    !pi9 = IIf((chkOpciones(9).Value = vbChecked), -1, 0)
    !pi10 = IIf((chkOpciones(10).Value = vbChecked), -1, 0)
    !pi11 = IIf((chkOpciones(11).Value = vbChecked), -1, 0)
    !pi12 = IIf((chkOpciones(12).Value = vbChecked), -1, 0)
    !pi13 = IIf((chkOpciones(13).Value = vbChecked), -1, 0)
    !pi14 = IIf((chkOpciones(14).Value = vbChecked), -1, 0)
    !pi15 = IIf((chkOpciones(15).Value = vbChecked), -1, 0)
    !pi16 = IIf((chkOpciones(16).Value = vbChecked), -1, 0)
    !pi17 = IIf((chkOpciones(17).Value = vbChecked), -1, 0)
    !pi18 = IIf((chkOpciones(18).Value = vbChecked), -1, 0)
    !pi19 = IIf((chkOpciones(19).Value = vbChecked), -1, 0)
    !pi20 = IIf((chkOpciones(20).Value = vbChecked), -1, 0)
    !pi21 = IIf((chkOpciones(21).Value = vbChecked), -1, 0)
    !pi22 = IIf((chkOpciones(22).Value = vbChecked), -1, 0)
    !pi23 = IIf((chkOpciones(23).Value = vbChecked), -1, 0)
    !pi24 = IIf((chkOpciones(24).Value = vbChecked), -1, 0)
    !pi25 = IIf((chkOpciones(25).Value = vbChecked), -1, 0)
    .UpDate
    idPerfil = !id
End With
Me.Hide
Exit Sub
errH:
If objRst.State = adStateOpen Then
    objRst.CancelUpdate
    objRst.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdAceptar_click"
subLog sERR

End Sub

Private Sub cmdCancelar_Click()
idPerfil = 0
Me.Hide
End Sub
Public Sub subCargaPerfil()
On Local Error GoTo errH
sSql = "select * from tperfiles where id=" & idPerfil
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    If Not .EOF Then
        txtNombre.Text = "" & !nombre
        chkOpciones(0).Value = IIf(!funcionarios, vbChecked, vbUnchecked)
        chkOpciones(1).Value = IIf(!parámetros, vbChecked, vbUnchecked)
        chkOpciones(2).Value = IIf(!registro, vbChecked, vbUnchecked)
        chkOpciones(3).Value = IIf(!reportes, vbChecked, vbUnchecked)
        chkOpciones(4).Value = IIf(!herramientas, vbChecked, vbUnchecked)
        chkOpciones(5).Value = IIf(!backup1, vbChecked, vbUnchecked)
        chkOpciones(6).Value = IIf(!minuta_digital, vbChecked, vbUnchecked)
        chkOpciones(7).Value = IIf(!reg_vis_nue, vbChecked, vbUnchecked)
        
        chkOpciones(8).Value = IIf(!pi8, vbChecked, vbUnchecked)
        chkOpciones(9).Value = IIf(!pi9, vbChecked, vbUnchecked)
        chkOpciones(10).Value = IIf(!pi10, vbChecked, vbUnchecked)
        chkOpciones(11).Value = IIf(!pi11, vbChecked, vbUnchecked)
        chkOpciones(12).Value = IIf(!pi12, vbChecked, vbUnchecked)
        chkOpciones(13).Value = IIf(!pi13, vbChecked, vbUnchecked)
        chkOpciones(14).Value = IIf(!pi14, vbChecked, vbUnchecked)
        chkOpciones(15).Value = IIf(!pi15, vbChecked, vbUnchecked)
        chkOpciones(16).Value = IIf(!pi16, vbChecked, vbUnchecked)
        chkOpciones(17).Value = IIf(!pi17, vbChecked, vbUnchecked)
        chkOpciones(18).Value = IIf(!pi18, vbChecked, vbUnchecked)
        chkOpciones(19).Value = IIf(!pi19, vbChecked, vbUnchecked)
        chkOpciones(20).Value = IIf(!pi20, vbChecked, vbUnchecked)
        chkOpciones(21).Value = IIf(!pi21, vbChecked, vbUnchecked)
        chkOpciones(22).Value = IIf(!pi22, vbChecked, vbUnchecked)
        chkOpciones(23).Value = IIf(!pi23, vbChecked, vbUnchecked)
        chkOpciones(24).Value = IIf(!pi24, vbChecked, vbUnchecked)
        chkOpciones(25).Value = IIf(!pi25, vbChecked, vbUnchecked)
    End If
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdAceptar_click"
subLog sERR
End Sub



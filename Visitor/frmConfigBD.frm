VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.1#0"; "ARBUTTON.OCX"
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#19.0#0"; "A1AControles.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfigBD 
   BackColor       =   &H00F8D88F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar Base de Datos"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3975
   ControlBox      =   0   'False
   Icon            =   "frmConfigBD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame F1 
      BackColor       =   &H00F8D88F&
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
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Frame fraSQL 
         BackColor       =   &H00F8D88F&
         Height          =   2535
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox txtContraseña 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1478
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   735
            Width           =   1815
         End
         Begin A1AControles.A1ATextBox txtServidor 
            Height          =   315
            Left            =   1200
            TabIndex        =   6
            Top             =   323
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
         End
         Begin VB.Shape Shape 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00A95900&
            Height          =   315
            Left            =   1395
            Shape           =   4  'Rounded Rectangle
            Top             =   720
            Width           =   1980
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contraseña: "
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
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Servidor:"
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
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.Frame fraAccess 
         BackColor       =   &H00F8D88F&
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CheckBox chkDefecto 
            Appearance      =   0  'Flat
            BackColor       =   &H00BE814F&
            Caption         =   "&Utilizar ruta por defecto"
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
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   3255
         End
         Begin MSComDlg.CommonDialog dlgRuta 
            Left            =   960
            Top             =   840
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ARButtonCtrl.ARButton cmdRuta 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   661
            Caption         =   "&Seleccionar Base de Datos"
            ForeColor       =   16777215
            ForeColorOnMouse=   12484943
            BackColorOnMouse=   16777215
            BackColor       =   12484943
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocus       =   2
         End
      End
      Begin A1AControles.A1AComboBox cmbMotor 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         bkColor         =   16308367
      End
      Begin VB.Image cmdCancelar 
         Height          =   555
         Left            =   120
         MouseIcon       =   "frmConfigBD.frx":70E2
         MousePointer    =   99  'Custom
         Picture         =   "frmConfigBD.frx":73EC
         Top             =   3480
         Width           =   1725
      End
      Begin VB.Image cmdAceptar 
         Height          =   555
         Left            =   1920
         MouseIcon       =   "frmConfigBD.frx":A67A
         MousePointer    =   99  'Custom
         Picture         =   "frmConfigBD.frx":A984
         Top             =   3480
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmConfigBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbMotor_Click()
If cmbMotor.itemID = 1 Then
    fraAccess.Visible = True
ElseIf cmbMotor.itemID = 2 Then
    fraSQL.Visible = True
End If
End Sub

Private Sub cmdAceptar_Click()
Dim sRuta As String
If cmbMotor.itemID = 0 Then
    MsgBox "Seleccione un elemento de la lista!", vbInformation
    cmbMotor.SetFocus
    Exit Sub
End If
If cmbMotor.itemID = 1 Then
    If cmdRuta.Tag = vbNullString Then
        If chkDefecto.Value = vbUnchecked Then
            MsgBox "Seleccione un archivo de base de datos o Marque la casilla Utilizar por defecto!", vbInformation
            Exit Sub
        Else
            sRuta = " "
        End If
    Else
        sRuta = cmdRuta.Tag
        If chkDefecto.Value = vbChecked Then
            sRuta = " "
        End If
    End If
    modoBD = bdACCESS
    fnEscribirIni "modobd", CStr(modoBD)
    fnEscribirIni "rutabd", sRuta
ElseIf cmbMotor.itemID = 2 Then
    If Trim(txtServidor.Text) = vbNullString Then
        MsgBox "Debe ingresar Servidor:!", vbInformation
        txtServidor.SetFocus
        Exit Sub
    Else
        sRuta = "Provider=SQLNCLI10.1;Password=" & txtContraseña.Text & ";Persist Security Info=True;User ID=sa;Initial Catalog=A1ABIOIDTAC;Data Source=" & txtServidor.Text
    End If
    modoBD = bdSQL
    fnEscribirIni "modobd", CStr(modoBD)
    fnEscribirIni "Conecta", sRuta
End If
Unload Me
End Sub

Private Sub cmdCancelar_Click()
End
End Sub

Private Sub cmdRuta_Click()
On Local Error GoTo errH
cmdRuta.Tag = vbNullString
dlgRuta.CancelError = True
dlgRuta.Filter = "Archivos de Base de datos|*.accdb;*.mdb"
dlgRuta.DialogTitle = "Seleccionar Base de Datos..."
dlgRuta.ShowOpen
cmdRuta.Tag = dlgRuta.FileName
Exit Sub
errH:
End Sub

Private Sub Form_Load()
cmbMotor.addElement "Microsoft Access", 1
cmbMotor.addElement "Microsoft SQL Server", 2
End Sub

VERSION 5.00
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFuncionarios 
   BackColor       =   &H0073EFF9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Funcionarios"
   ClientHeight    =   11055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11385
   Icon            =   "frmFuncionarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11055
   ScaleWidth      =   11385
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H0073EFF9&
      Height          =   10935
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   11175
      Begin VB.CheckBox chkNomina 
         Appearance      =   0  'Flat
         BackColor       =   &H0073EFF9&
         Caption         =   "&Reportar a nómina"
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
         Left            =   8280
         TabIndex        =   23
         Top             =   7080
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Frame fraHuelleros 
         BackColor       =   &H0073EFF9&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   9330
         TabIndex        =   62
         Top             =   3720
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Image imgZK 
            Height          =   540
            Left            =   0
            MouseIcon       =   "frmFuncionarios.frx":2982
            MousePointer    =   99  'Custom
            Picture         =   "frmFuncionarios.frx":2C8C
            Stretch         =   -1  'True
            ToolTipText     =   "Usar lector ZK"
            Top             =   840
            Width           =   1620
         End
         Begin VB.Image imgUareU 
            Height          =   540
            Left            =   0
            MouseIcon       =   "frmFuncionarios.frx":9E5C
            MousePointer    =   99  'Custom
            Picture         =   "frmFuncionarios.frx":A166
            Stretch         =   -1  'True
            ToolTipText     =   "Usar lector UareU"
            Top             =   120
            Width           =   1620
         End
      End
      Begin ARButtonCtrl.ARButton cmdEtoV 
         Height          =   495
         Left            =   9240
         TabIndex        =   60
         Tag             =   "12484943"
         ToolTipText     =   "Funcionario-->Visitante"
         Top             =   6120
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   873
         Caption         =   ""
         ForeColor       =   16777215
         ForeColorOnMouse=   8987923
         BackColorOnMouse=   16777215
         BackColor       =   8987923
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
         Style           =   1
         Picture         =   "frmFuncionarios.frx":D843
         PictureOn       =   "frmFuncionarios.frx":ED95
      End
      Begin MSComDlg.CommonDialog dlgFoto 
         Left            =   240
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0073EFF9&
         Caption         =   "Información Adicional"
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
         Height          =   3255
         Left            =   120
         TabIndex        =   47
         Top             =   7560
         Width           =   10935
         Begin VB.CheckBox chkActivo 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "&Activo"
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
            Left            =   5760
            TabIndex        =   58
            Top             =   1170
            Width           =   975
         End
         Begin A1AControles.A1AComboBox cmbEps 
            Height          =   315
            Left            =   120
            TabIndex        =   48
            Top             =   600
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            bkColor         =   7598073
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1AComboBox cmbArp 
            Height          =   315
            Left            =   3720
            TabIndex        =   50
            Top             =   600
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            bkColor         =   7598073
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1AComboBox cmbAfp 
            Height          =   315
            Left            =   7320
            TabIndex        =   52
            Top             =   600
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            bkColor         =   7598073
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1AComboBox cmbPerfiles 
            Height          =   315
            Left            =   120
            TabIndex        =   54
            Top             =   1200
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            bkColor         =   7598073
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1ATextBox txtFechaNac 
            Height          =   315
            Left            =   3720
            TabIndex        =   56
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   ""
            ColorFoco       =   16308367
         End
         Begin ARButtonCtrl.ARButton cmdActivos 
            Height          =   375
            Left            =   6840
            TabIndex        =   59
            Tag             =   "12484943"
            Top             =   1170
            Visible         =   0   'False
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   661
            Caption         =   "&Todos Activos"
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
         Begin ARButtonCtrl.ARButton cmdImprime 
            Height          =   375
            Left            =   8760
            TabIndex        =   67
            Tag             =   "12484943"
            ToolTipText     =   "Carnet del último registro..."
            Top             =   1170
            Visible         =   0   'False
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   661
            Caption         =   "&impresión carnet"
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
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Nacimiento"
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
            Left            =   3720
            TabIndex        =   57
            Top             =   960
            Width           =   2025
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   5475
            MouseIcon       =   "frmFuncionarios.frx":102E7
            MousePointer    =   99  'Custom
            Picture         =   "frmFuncionarios.frx":105F1
            ToolTipText     =   "Seleccionar fecha"
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Perfil de Seguridad:"
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
            TabIndex        =   55
            Top             =   960
            Width           =   1920
         End
         Begin VB.Image imgEditarPerfiles 
            Height          =   300
            Left            =   3405
            MouseIcon       =   "frmFuncionarios.frx":10840
            MousePointer    =   99  'Custom
            Picture         =   "frmFuncionarios.frx":10B4A
            ToolTipText     =   "Editar Nombre actual"
            Top             =   1200
            Width           =   300
         End
         Begin VB.Image imgEditaAFP 
            Height          =   300
            Left            =   10605
            MouseIcon       =   "frmFuncionarios.frx":10F7D
            MousePointer    =   99  'Custom
            Picture         =   "frmFuncionarios.frx":11287
            ToolTipText     =   "Editar Nombre actual"
            Top             =   600
            Width           =   300
         End
         Begin VB.Image imgEditaARP 
            Height          =   300
            Left            =   7005
            MouseIcon       =   "frmFuncionarios.frx":116BA
            MousePointer    =   99  'Custom
            Picture         =   "frmFuncionarios.frx":119C4
            ToolTipText     =   "Editar Nombre actual"
            Top             =   600
            Width           =   300
         End
         Begin VB.Image imgEditaEPS 
            Height          =   300
            Left            =   3405
            MouseIcon       =   "frmFuncionarios.frx":11DF7
            MousePointer    =   99  'Custom
            Picture         =   "frmFuncionarios.frx":12101
            ToolTipText     =   "Editar Nombre actual"
            Top             =   600
            Width           =   300
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AFP:"
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
            Left            =   7335
            TabIndex        =   53
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ARP:"
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
            Left            =   3735
            TabIndex        =   51
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EPS:"
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
            TabIndex        =   49
            Top             =   360
            Width           =   450
         End
      End
      Begin A1AControles.A1ATextBox txtFechaF 
         Height          =   315
         Left            =   4200
         TabIndex        =   21
         Top             =   7080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtFechaI 
         Height          =   315
         Left            =   2280
         TabIndex        =   20
         Top             =   7080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtExtension 
         Height          =   315
         Left            =   4920
         TabIndex        =   12
         Top             =   4320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtTelefono 
         Height          =   315
         Left            =   2280
         TabIndex        =   11
         Top             =   4320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1AComboBox cmbDepartamentos 
         Height          =   315
         Left            =   5040
         TabIndex        =   7
         Top             =   2400
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         bkColor         =   7598073
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1AComboBox cmbCompañias 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   2400
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         bkColor         =   7598073
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtSexo 
         Height          =   315
         Left            =   6600
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtNombre 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   1200
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtDoc1 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin VB.CheckBox chkLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H0073EFF9&
         Caption         =   "&Login"
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
         Left            =   480
         TabIndex        =   15
         Top             =   5280
         Width           =   855
      End
      Begin VB.Frame fraLogin 
         BackColor       =   &H0073EFF9&
         Enabled         =   0   'False
         Height          =   1335
         Left            =   360
         TabIndex        =   36
         Top             =   5400
         Width           =   8775
         Begin A1AControles.A1ATextBox txtConfirma 
            Height          =   315
            Left            =   5880
            TabIndex        =   18
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   "§"
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1ATextBox txtContraseña 
            Height          =   315
            Left            =   3000
            TabIndex        =   17
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   "§"
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1ATextBox txtUsuario 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   ""
            ColorFoco       =   16308367
         End
         Begin VB.Label Confirmar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirmar:"
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
            Left            =   5880
            TabIndex        =   39
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contraseña:"
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
            Left            =   3000
            TabIndex        =   38
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario:"
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
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   780
         End
      End
      Begin VB.Timer tmrCam 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8640
         Top             =   6240
      End
      Begin IMAGXPR5LibCtl.ImagXpress imgFoto 
         Height          =   2340
         Left            =   9240
         TabIndex        =   35
         Top             =   960
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   4128
         ErrStr          =   "U9EROCBXRIS-GC305XPXEP"
         ErrCode         =   1061571217
         ErrInfo         =   -850013016
         Persistence     =   -1  'True
         _cx             =   99746432
         _cy             =   1
         Picture         =   "frmFuncionarios.frx":12534
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "A1A Group"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         AutoSize        =   2
         ScrollBarLargeChangeH=   10
         ScrollBarSmallChangeH=   1
         DrawFillColor   =   255
         SaveJPGSubSampling=   2
         OLEDropMode     =   0
         CompressInMemory=   2
      End
      Begin A1AControles.A1ATextBox txtApellidos 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   1800
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtRh 
         Height          =   315
         Left            =   7890
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtOficina 
         Height          =   315
         Left            =   360
         TabIndex        =   10
         Top             =   4320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtMovil 
         Height          =   315
         Left            =   6570
         TabIndex        =   13
         Top             =   4320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtEmail 
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   4920
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1AComboBox cmbCargos 
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   3000
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         bkColor         =   7598073
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1AComboBox cmbHorarios 
         Height          =   315
         Left            =   5040
         TabIndex        =   9
         Top             =   3000
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         bkColor         =   7598073
         ColorFoco       =   16308367
      End
      Begin A1AControles.A1ATextBox txtTarjetaNum 
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   7080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin ARButtonCtrl.ARButton cmdTitular 
         Height          =   300
         Left            =   5160
         TabIndex        =   63
         Tag             =   "12484943"
         ToolTipText     =   "Asociar Titular"
         Top             =   600
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   "T"
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
         Alignment       =   1
         ShowFocus       =   2
      End
      Begin A1AControles.A1ATextBox txtDocTitular 
         Height          =   315
         Left            =   3000
         TabIndex        =   64
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
         Enabled         =   0   'False
      End
      Begin A1AControles.A1AComboBox cmbParentesco 
         Height          =   315
         Left            =   5640
         TabIndex        =   1
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         bkColor         =   7598073
         ColorFoco       =   16308367
      End
      Begin ARButtonCtrl.ARButton cmdHorarioPer 
         Height          =   375
         Left            =   6120
         TabIndex        =   69
         Tag             =   "12484943"
         ToolTipText     =   "Modificar horarios del personal"
         Top             =   3540
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         Caption         =   "Personalizar horarios"
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
      Begin ARButtonCtrl.ARButton cmdHorarioSet 
         Height          =   375
         Left            =   5040
         TabIndex        =   68
         Tag             =   "12484943"
         Top             =   3540
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   661
         Caption         =   "Aplicar"
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
      Begin A1AControles.A1ATextBox txtCódigo 
         Height          =   315
         Left            =   6240
         TabIndex        =   22
         Top             =   7080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   7598073
         passChar        =   ""
         ColorFoco       =   16308367
      End
      Begin ARButtonCtrl.ARButton cmdGrabar 
         Height          =   495
         Left            =   9225
         TabIndex        =   24
         Tag             =   "12484943"
         Top             =   5520
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   873
         Caption         =   "&Guardar"
         ForeColor       =   16777215
         ForeColorOnMouse=   8987923
         BackColorOnMouse=   16777215
         BackColor       =   8987923
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
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   6240
         TabIndex        =   71
         Top             =   6840
         Width           =   675
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aplica el Horario seleccionado a todo el personal del Departamento seleccionado"
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
         Height          =   480
         Left            =   720
         TabIndex        =   70
         Top             =   3480
         Width           =   4530
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   4800
         Picture         =   "frmFuncionarios.frx":3DAAA
         ToolTipText     =   "Editar Nombre actual"
         Top             =   3577
         Width           =   300
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   4095
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parentesco:"
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
         Left            =   5655
         TabIndex        =   66
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Titular"
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
         Left            =   3015
         TabIndex        =   65
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarjeta N°:"
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
         Left            =   360
         TabIndex        =   61
         Top             =   6840
         Width           =   1005
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   5865
         MouseIcon       =   "frmFuncionarios.frx":3DF9C
         MousePointer    =   99  'Custom
         Picture         =   "frmFuncionarios.frx":3E2A6
         ToolTipText     =   "Seleccionar fecha"
         Top             =   7110
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3915
         MouseIcon       =   "frmFuncionarios.frx":3E4F5
         MousePointer    =   99  'Custom
         Picture         =   "frmFuncionarios.frx":3E7FF
         ToolTipText     =   "Seleccionar fecha"
         Top             =   7080
         Width           =   240
      End
      Begin VB.Image cmdEditaHorarios 
         Height          =   300
         Left            =   8805
         MouseIcon       =   "frmFuncionarios.frx":3EA4E
         MousePointer    =   99  'Custom
         Picture         =   "frmFuncionarios.frx":3ED58
         ToolTipText     =   "Editar Nombre actual"
         Top             =   3000
         Width           =   300
      End
      Begin VB.Image imgEditaCargo 
         Height          =   300
         Left            =   4605
         MouseIcon       =   "frmFuncionarios.frx":3F18B
         MousePointer    =   99  'Custom
         Picture         =   "frmFuncionarios.frx":3F495
         ToolTipText     =   "Editar Nombre actual"
         Top             =   3000
         Width           =   300
      End
      Begin VB.Image imgEditaDepartamento 
         Height          =   300
         Left            =   8805
         MouseIcon       =   "frmFuncionarios.frx":3F8C8
         MousePointer    =   99  'Custom
         Picture         =   "frmFuncionarios.frx":3FBD2
         ToolTipText     =   "Editar Nombre actual"
         Top             =   2400
         Width           =   300
      End
      Begin VB.Image imgEditaCompañia 
         Height          =   300
         Left            =   4605
         MouseIcon       =   "frmFuncionarios.frx":40005
         MousePointer    =   99  'Custom
         Picture         =   "frmFuncionarios.frx":4030F
         ToolTipText     =   "Editar Nombre actual"
         Top             =   2400
         Width           =   300
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
         ForeColor       =   &H00A95900&
         Height          =   240
         Left            =   5040
         TabIndex        =   46
         Top             =   2760
         Width           =   765
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo:"
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
         Left            =   840
         TabIndex        =   45
         Top             =   2760
         Width           =   630
      End
      Begin VB.Image imgBuscar 
         Height          =   300
         Left            =   8805
         MouseIcon       =   "frmFuncionarios.frx":40742
         MousePointer    =   99  'Custom
         Picture         =   "frmFuncionarios.frx":40A4C
         ToolTipText     =   "Buscar..."
         Top             =   1215
         Width           =   300
      End
      Begin VB.Image Image3 
         Height          =   765
         Left            =   120
         Picture         =   "frmFuncionarios.frx":40E51
         Top             =   2640
         Width           =   525
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   120
         Picture         =   "frmFuncionarios.frx":42800
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   120
         Picture         =   "frmFuncionarios.frx":43DAE
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
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
         Left            =   360
         TabIndex        =   44
         Top             =   4680
         Width           =   600
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Móvil:"
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
         Left            =   6570
         TabIndex        =   43
         Top             =   4080
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extensión:"
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
         Left            =   4920
         TabIndex        =   42
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oficina:"
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
         Left            =   360
         TabIndex        =   41
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento:"
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
         Left            =   5040
         TabIndex        =   40
         Top             =   2160
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin:"
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
         Left            =   4200
         TabIndex        =   34
         Top             =   6840
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inicio:"
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
         Left            =   2280
         TabIndex        =   33
         Top             =   6840
         Width           =   1530
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefonos:"
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
         Left            =   2280
         TabIndex        =   32
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RH:"
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
         Left            =   7890
         TabIndex        =   31
         Top             =   1560
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo:"
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
         Left            =   6600
         TabIndex        =   30
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañía:"
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
         Left            =   855
         TabIndex        =   29
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID No de documento"
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
         Left            =   855
         TabIndex        =   28
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00A95900&
         Height          =   240
         Left            =   855
         TabIndex        =   27
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos:"
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
         Left            =   855
         TabIndex        =   26
         Top             =   1560
         Width           =   945
      End
      Begin VB.Image imgHuella 
         Height          =   2100
         Left            =   9225
         Picture         =   "frmFuncionarios.frx":45507
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmFuncionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public idEmpleado As Long
Dim idCarnet As Long
Dim idTitular As Long
Dim idDep As Long
Dim bFoto As Boolean
Dim sDOC As String

Private Sub chkLogin_Click()
If chkLogin.Value = vbChecked Then
    fraLogin.Enabled = True
Else
    fraLogin.Enabled = False
End If
End Sub

Private Sub cmbAfp_Click()
On Local Error GoTo errH:
Dim iElem As Long, sElem As String
If cmbAfp.itemID = -1 Then
    frmNuevo.Show vbModal
    sElem = frmNuevo.Tag
    Unload frmNuevo
    If sElem <> vbNullString Then
        sSql = "select * from tafp where id=0"
        With objRstA
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            .AddNew
            !nombre = sElem
            .UpDate
            iElem = !id
            .Close
        End With
        subAFP
        cmbAfp.mostrarItem iElem
    End If
End If
Exit Sub
errH:
If objRstA.State = adStateOpen Then
    objRstA.CancelUpdate
    objRstA.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdAfp_click"
subLog sERR

End Sub

Private Sub cmbAfp_Click0()
cmbAfp.ZOrder 0
End Sub

Private Sub cmbArp_Click()
On Local Error GoTo errH:
Dim iElem As Long, sElem As String
If cmbArp.itemID = -1 Then
    frmNuevo.Show vbModal
    sElem = frmNuevo.Tag
    Unload frmNuevo
    If sElem <> vbNullString Then
        sSql = "select * from tarp where id=0"
        With objRstA
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            .AddNew
            !nombre = sElem
            .UpDate
            iElem = !id
            .Close
        End With
        subARP
        cmbArp.mostrarItem iElem
    End If
End If
Exit Sub
errH:
If objRstA.State = adStateOpen Then
    objRstA.CancelUpdate
    objRstA.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdArp_click"
subLog sERR
End Sub

Private Sub cmbArp_Click0()
cmbArp.ZOrder 0
End Sub

Private Sub cmbCargos_Click()
On Local Error GoTo errH:
Dim iElem As Long, sElem As String
If cmbCargos.itemID = -1 Then
    frmNuevo.Show vbModal
    sElem = frmNuevo.Tag
    Unload frmNuevo
    If sElem <> vbNullString Then
        sSql = "select * from tcargos where id=0"
        With objRstA
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            .AddNew
            !nombre = sElem
            .UpDate
            iElem = !id
            .Close
        End With
        subCargos
        cmbCargos.mostrarItem iElem
    End If
End If
Exit Sub
errH:
If objRstA.State = adStateOpen Then
    objRstA.CancelUpdate
    objRst.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdCargos_click"
subLog sERR

End Sub

Private Sub cmbCargos_Click0()
cmbCargos.ZOrder 0
End Sub

Private Sub cmbCompañias_Click()
On Local Error GoTo errH:
Dim iElem As Long, sElem As String
If cmbCompañias.itemID = -1 Then
    frmNuevo.Show vbModal
    sElem = frmNuevo.Tag
    Unload frmNuevo
    If sElem <> vbNullString Then
        sSql = "select * from tcompañias where id=0"
        With objRstA
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            .AddNew
            !nombre = sElem
            .UpDate
            iElem = !id
            .Close
        End With
        subCompañias cmbCompañias, True
        cmbCompañias.mostrarItem iElem
    End If
Else
    subDepartamentos
End If
Exit Sub
errH:
If objRstA.State = adStateOpen Then
    objRstA.CancelUpdate
    objRstA.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdCompañias_click"
subLog sERR
End Sub

Private Sub cmbCompañias_Click0()
cmbCompañias.ZOrder 0
End Sub

Private Sub cmbDepartamentos_Click()
On Local Error GoTo errH:
Dim sElem As String
If cmbDepartamentos.itemID = -1 Then
    idDep = 0
    idDepartamento = 0
    frmDepartamentos.Show vbModal
    sElem = frmDepartamentos.Tag
    frmDepartamentos.Tag = vbNullString
    Unload frmDepartamentos
    If sElem <> vbNullString Then
        subDepartamentos
        cmbDepartamentos.mostrarItem Val(sElem)
        sElem = vbNullString
    End If
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmbHorarios_click"
subLog sERR
End Sub

Private Sub cmbDepartamentos_Click0()
cmbDepartamentos.ZOrder 0
End Sub

Private Sub cmbEps_Click()
On Local Error GoTo errH:
Dim iElem As Long, sElem As String
If cmbEps.itemID = -1 Then
    frmNuevo.Show vbModal
    sElem = frmNuevo.Tag
    Unload frmNuevo
    If sElem <> vbNullString Then
        sSql = "select * from teps where id=0"
        With objRstA
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            .AddNew
            !nombre = sElem
            .UpDate
            iElem = !id
            .Close
        End With
        subEPS
        cmbEps.mostrarItem iElem
    End If
End If
Exit Sub
errH:
If objRstA.State = adStateOpen Then
    objRstA.CancelUpdate
    objRstA.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdEps_click"
subLog sERR

End Sub

Private Sub cmbEps_Click0()
cmbEps.ZOrder 0
End Sub

Private Sub cmbHorarios_Click()
On Local Error GoTo errH:
Dim sElem As String
If cmbHorarios.itemID = -1 Then
    frmHorarios.Show vbModal
    sElem = frmHorarios.Tag
    Unload frmHorarios
    If sElem <> vbNullString Then
        subHorarios cmbHorarios, True
        cmbHorarios.mostrarItem Val(sElem)
    End If
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmbHorarios_click"
subLog sERR

End Sub

Private Sub cmbHorarios_Click0()
cmbHorarios.ZOrder 0
End Sub

Private Sub cmbParentesco_Click()
On Local Error GoTo errH:
Dim iElem As Long, sElem As String
If cmbParentesco.itemID = -1 Then
    frmNuevo.Show vbModal
    sElem = frmNuevo.Tag
    Unload frmNuevo
    If sElem <> vbNullString Then
        sSql = "select * from tparentesco where id=0"
        With objRstA
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            .AddNew
            !nombre = sElem
            .UpDate
            iElem = !id
            .Close
        End With
        subParentesco
        cmbParentesco.mostrarItem iElem
    End If
End If
Exit Sub
errH:
If objRstA.State = adStateOpen Then
    objRstA.CancelUpdate
    objRstA.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmbParentesco_Click"
subLog sERR

End Sub

Private Sub cmbParentesco_Click0()
cmbParentesco.ZOrder 0
End Sub

Private Sub cmbPerfiles_Click()
On Local Error GoTo errH:
Dim iElem As Long
If cmbPerfiles.itemID = -1 Then
    frmPerfiles.Show vbModal
    iElem = frmPerfiles.idPerfil
    frmPerfiles.idPerfil = 0
    Unload frmPerfiles
    If iElem > 0 Then
        subPerfiles
        cmbPerfiles.mostrarItem iElem
    End If
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdPerfiles_click"
subLog sERR
End Sub

Private Sub cmbPerfiles_Click0()
cmbPerfiles.ZOrder 0
End Sub

Private Sub cmdActivos_Click()
Dim bResp As Byte
bResp = MsgBox("Seguro que desea establecer el estado ACTIVO a todos los Funcionarios?", vbYesNo + vbCritical)
If bResp = vbYes Then
    sSql = "update templeados set activo=-1"
    objCon.Execute sSql
End If
End Sub

Private Sub cmdEditaHorarios_Click()
Dim tH As Long
If cmbHorarios.itemID > 0 Then
    tH = cmbHorarios.itemID
    Load frmHorarios
    frmHorarios.cmbHorarios.mostrarItem cmbHorarios.itemID
    frmHorarios.Show vbModal
    Unload frmHorarios
    subHorarios cmbHorarios, True
    cmbHorarios.mostrarItem tH
End If
End Sub

Private Sub cmdGrabar_Click()
On Local Error GoTo errH
If Trim(txtDoc1.Text) = vbNullString Then
    MsgBox "Ingrese Número de documento!", vbInformation
    txtDoc1.Text = vbNullString
    txtDoc1.SetFocus
    Exit Sub
End If
'If cmbHorarios.itemID = 0 Then
'    MsgBox "Debe asignar un horario!", vbInformation
'    cmbHorarios.SetFocus
'    Exit Sub
'End If
If Not IsDate(Trim(txtFechaI.Text)) Then
    MsgBox "La fecha de inicio es obligatoria!", vbInformation
    txtFechaI.SetFocus
    Exit Sub
End If
If Not IsDate(Trim(txtFechaF.Text)) Then
    MsgBox "La fecha de finalización es obligatoria!", vbInformation
    txtFechaF.SetFocus
    Exit Sub
End If
If CDate(Trim(txtFechaF.Text)) < CDate(Trim(txtFechaI.Text)) Then
    MsgBox "Las fechas de inicio y finalización son inconsistentes!", vbInformation
    Exit Sub
End If
If cmbPerfiles.itemID = 0 Then
    MsgBox "Debe asignar un Perfil de Seguridad!", vbInformation
    cmbPerfiles.SetFocus
    Exit Sub
End If
If Trim(txtFechaNac.Text) <> vbNullString Then
    If Not IsDate(Trim(txtFechaNac.Text)) Then
        MsgBox "La fecha de nacimiento no es válida!", vbInformation
        txtFechaNac.SetFocus
        Exit Sub
    End If
End If
sSql = "select * from templeados where id=" & idEmpleado
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If .EOF Then .AddNew
    !documento = Trim(txtDoc1.Text)
    !nombre = fnMayúscula(Trim(txtNombre.Text))
    !apellidos = fnMayúscula(Trim(txtApellidos.Text))
    !sexo = UCase(Trim(txtSexo.Text))
    !rh = UCase(Trim(txtRh.Text))
    !idcompañia = cmbCompañias.itemID
    !idDepartamento = cmbDepartamentos.itemID
    !idcargo = cmbCargos.itemID
    !idHorario = cmbHorarios.itemID
    !oficina = Trim(txtOficina.Text)
    !tel = Trim(txtTelefono.Text)
    !extension = Trim(txtExtension.Text)
    !movil = Trim(txtMovil.Text)
    !email = Trim(txtEmail.Text)
    If IsDate(Trim(txtFechaI.Text)) Then !fechaI = fnFecha(Trim(txtFechaI.Text), False)
    If IsDate(Trim(txtFechaF.Text)) Then !fechaf = fnFecha(Trim(txtFechaF.Text), False)
    If bFoto And bModificaFoto Then
        If objGAATools.fnExisteArchivo(App.Path & "\tmpFoto") Then Kill App.Path & "\tmpFoto"
        DoEvents
        'SavePicture imgFoto1.Picture, App.Path & "\tmpFoto"
        If dlgFoto.FileName = vbNullString Then
            imgFoto.SaveFileType = FT_BMP
            imgFoto.SaveFileName = App.Path & "\tmpFoto"
            imgFoto.SaveFile
            ConvertBMPtoJPG App.Path & "\tmpFoto", App.Path & "\tmpFoto" & ".jpg", True, 50, False
            fnGuardaFoto !foto, App.Path & "\tmpFoto.jpg"
        Else
            fnGuardaFoto !foto, dlgFoto.FileName
        End If
    End If
    If bHuella And bModificaHuella Then
        If objGAATools.fnExisteArchivo(App.Path & "\tmpHuella") Then Kill App.Path & "\tmpHuella"
        DoEvents
        SavePicture imgHuella.Picture, App.Path & "\tmpHuella"
        ConvertBMPtoJPG App.Path & "\tmpHuella", App.Path & "\tmpHuella" & ".jpg", True, 50, False
        fnGuardaFoto !huella, App.Path & "\tmpHuella.jpg"
        !enrola = bHuellaMinuciasCAP
    End If
    !ideps = cmbEps.itemID
    !idarp = cmbArp.itemID
    !idafp = cmbAfp.itemID
    !idPerfil = cmbPerfiles.itemID
    If IsDate(Trim(txtFechaNac.Text)) Then !fecha_nac = fnFecha(Trim(txtFechaNac.Text), False)
    !activo = IIf((chkActivo.Value = vbChecked), -1, 0)
    !tarjeta_num = Trim(txtTarjetaNum.Text)
    !idTitular = idTitular
    !idparentesco = cmbParentesco.itemID
    !nomina = IIf((chkNomina.Value = vbChecked), -1, 0)
    !codigo = Trim(txtCódigo.Text)
    .UpDate
    idEmpleado = !id
    idCarnet = !id
    .Close
End With
If chkLogin.Value = vbChecked Then
    sSql = "select id,usuario,contraseña from templeados where usuario='" & txtUsuario.Text & "'"
    Set objRst = objCon.Execute(sSql)
    If Not objRst.EOF Then
        If objRst!id <> idEmpleado Then
            MsgBox "El Usuario " & txtUsuario.Text & " ya existe!. Debe digitar otro diferente.", vbInformation
            txtUsuario.Text = vbNullString
            txtUsuario.SetFocus
            Exit Sub
        End If
    End If
    sSql = "select usuario,contraseña from templeados where id=" & idEmpleado
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenKeyset, adLockOptimistic
        If .EOF Then .AddNew
        If txtConfirma.Text <> txtConfirma.Text Then
            MsgBox "La confirmación de la contraseña no coincide!", vbInformation
            txtConfirma.SetFocus
        Else
            !usuario = txtUsuario.Text
            !contraseña = txtContraseña.Text
        End If
        .UpDate
        .Close
    End With
End If
If Trim(txtTarjetaNum.Text) <> vbNullString Then
    'If (Not oZKs) = -1 Then
    If Not objGAATools.fnArrVacioCls(oZKs) Then
        For idxZK = 1 To UBound(oZKs)
            If oZKs(idxZK).bConectado Then
                If frmPrincipal.objZK(idxZK).SetStrCardNumber(Trim(txtTarjetaNum.Text)) Then
                    frmPrincipal.objZK(idxZK).GetLastError lZkErr
                    If lZkErr = 1 Then
                        zkID = idEmpleado
                        zkUSR = "fun" & zkID
                        If frmPrincipal.objZK(idxZK).SetUserInfo(1, zkID, zkUSR, zkUSR, 0, True) Then
                            frmPrincipal.objZK(idxZK).GetLastError lZkErr
                            
                        End If
                    End If
                End If
            End If
        Next idxZK
    End If
End If
idEmpleado = 0
frmPrincipal.subCargaPerfil
subLimpiar
cmdImprime.Visible = True
Exit Sub
errH:
idEmpleado = 0
If objRst.State = adStateOpen Then
    objRst.CancelUpdate
    objRst.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdGrabar_click"
subLog sERR
End Sub
Private Sub subLimpiar()
On Error GoTo errH
Dim cTrl As Control
For Each cTrl In Me
    If TypeName(cTrl) = "A1ATextBox" Then
        cTrl.Text = vbNullString
    ElseIf TypeName(cTrl) = "A1AComboBox" Then
        cTrl.itemID = 0
        cTrl.Limpiar
    End If
Next
idEmpleado = 0
sDOC = vbNullString
idDep = 0
idDepartamento = 0
bFoto = False
imgFoto.Picture = LoadPicture(App.Path & "\imgfoto.jpg")
imgHuella.Picture = LoadPicture(App.Path & "\imghuella.jpg")
chkLogin.Value = vbUnchecked
subCompañias cmbCompañias, True
subCargos
subHorarios cmbHorarios, True
subEPS
subARP
subAFP
subPerfiles
subParentesco
txtDoc1.SetFocus
chkActivo.Value = vbUnchecked
fraHuelleros.Visible = False
idTitular = 0
chkNomina.Value = vbChecked
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subLimpiar"
subLog sERR
End Sub

Private Sub cmdEtoV_Click()
Dim bResp As Byte
On Local Error GoTo errH
    If idEmpleado > 0 Then
        bResp = MsgBox("Esta persona dejará de ser Funcionario y será registrado como Visitante." & vbCr & "¿Desea continuar?", vbYesNo + vbQuestion)
        If bResp = vbYes Then
            sSql = "insert into tvisitantes_huella(documento,idtipodoc,idtratamiento,nombre,apellidos,sexo,rh,email,telefono,foto,huella,enrola) "
            sSql = sSql & "select documento,1,1,nombre,apellidos,sexo,rh,email,movil,foto,huella,enrola "
            sSql = sSql & "from templeados where id=" & idEmpleado
            objCon.Execute (sSql)
            
            
            sSql = "delete from templeados where id=" & idEmpleado
            objCon.Execute sSql
            
            sSql = "update tacceso set idhuellero_entra=id where idtpersona=" & idEmpleado & " and idtipoper=1"
            objCon.Execute sSql
            
            MsgBox txtNombre.Text & " " & txtApellidos.Text & " ahora está registrado como visitante.", vbInformation
           
            subLimpiar
        End If
    End If
Exit Sub
errH:
MsgBox "Error " & Err.Number & ". " & Err.Description

End Sub

Private Sub cmdHorarioMas_Click()

End Sub

Private Sub cmdHorarioPer_Click()
Load frmHorariosSET
subHorarios frmHorariosSET.cmbHorarios, False
subHorarios frmHorariosSET.cmbHorariosAdd, False
frmHorariosSET.Show , Me
End Sub

Private Sub cmdHorarioSet_Click()
Dim bResp As Byte
Dim objRHor As New ADODB.Recordset
If cmbDepartamentos.itemID <= 0 Then
    MsgBox "Debe seleccionar un Departamento!", vbInformation
    cmbDepartamentos.SetFocus
    Exit Sub
End If
If cmbHorarios.itemID <= 0 Then
    MsgBox "Debe seleccionar un Horario!", vbInformation
    cmbHorarios.SetFocus
    Exit Sub
End If
sSql = "select count(id) as cnt from templeados where abs(activo)=1 and iddepartamento=" & cmbDepartamentos.itemID
Set objRHor = objCon.Execute(sSql)
If objRHor!cnt = 0 Then
    MsgBox "No hay personal en este Departamento!", vbInformation
Else
    bResp = MsgBox("Se asignará el horario " & cmbHorarios.Text & " a " & objRHor!cnt & " personas " & vbCr & _
    "del departamento " & cmbDepartamentos.Text & ". ¿Desea continuar?", vbYesNo + vbQuestion)
    If bResp = vbYes Then
        sSql = "update templeados set idhorario=" & cmbHorarios.itemID & " where abs(activo)=1 and iddepartamento=" & cmbDepartamentos.itemID
        objCon.Execute sSql
        MsgBox "Asignación de horario aplicada satisfactoriamente!", vbInformation
    End If
End If
End Sub

Private Sub cmdImprime_Click()
If idCarnet > 0 Then
    Load frmHerramientas
    frmHerramientas.subCargar idCarnet
    frmHerramientas.Show
End If
End Sub

Private Sub cmdTitular_Click()
If idTitular > 0 Then
    Load frmBuscar
    frmBuscar.subBuscaUno idTitular
End If
frmBuscar.Show vbModal
If Val(frmBuscar.Tag) <> 0 Then
    idTitular = Val(frmBuscar.Tag)
    Unload frmBuscar
    If idTitular = idEmpleado Then
        MsgBox "El titular no puede ser la misma persona!", vbInformation
        idTitular = 0
    Else
        subDocTitular
    End If
End If

End Sub
Private Sub subDocTitular()
Dim objRs_ As New ADODB.Recordset
If Val(idTitular) > 0 Then
    sSql = "select documento from templeados where id=" & idTitular
    Set objRs_ = objCon.Execute(sSql)
    If Not objRs_.EOF Then
        txtDocTitular.Text = "" & objRs_!documento
    End If
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
ElseIf KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    idCarnet = 0
    cmdImprime.Visible = False
    subLimpiar
End If
End Sub

Public Sub subCompañias(ByRef oCombo As A1AComboBox, bNuevo As Boolean)
On Local Error GoTo errH
oCombo.Limpiar
sSql = "select * from tcompañias order by nombre"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    If bNuevo Then oCombo.addElement "(Nuevo...)", -1
    While Not objRstA.EOF
        oCombo.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subDependencias"
subLog sERR
End Sub
Sub subCargos()
On Local Error GoTo errH
cmbCargos.Limpiar
sSql = "select * from tcargos order by nombre"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbCargos.addElement "(Nuevo...)", -1
    While Not objRstA.EOF
        cmbCargos.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subCargos"
subLog sERR
End Sub
Sub subHorarios(ByRef oCombo As A1AComboBox, bNuevo As Boolean)
On Local Error GoTo errH
oCombo.Limpiar
sSql = "select * from thorarios order by nombre"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    If bNuevo Then oCombo.addElement "(Nuevo...)", -1
    oCombo.addElement "(Ninguno)", 0
    While Not objRstA.EOF
        oCombo.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subhorarios"
subLog sERR
End Sub
Sub subEPS()
On Local Error GoTo errH
cmbEps.Limpiar
sSql = "select * from teps order by nombre"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbEps.addElement "(Nuevo...)", -1
    While Not objRstA.EOF
        cmbEps.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subEps"
subLog sERR
End Sub
Sub subARP()
On Local Error GoTo errH
cmbArp.Limpiar
sSql = "select * from tarp order by nombre"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbArp.addElement "(Nuevo...)", -1
    While Not objRstA.EOF
        cmbArp.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subArp"
subLog sERR
End Sub
Sub subAFP()
On Local Error GoTo errH
cmbAfp.Limpiar
sSql = "select * from tafp order by nombre"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbAfp.addElement "(Nuevo...)", -1
    While Not objRstA.EOF
        cmbAfp.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subAfp"
subLog sERR
End Sub
Sub subParentesco()
On Local Error GoTo errH
cmbParentesco.Limpiar
sSql = "select * from tparentesco order by nombre"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbParentesco.addElement "(Nuevo...)", -1
    While Not objRstA.EOF
        cmbParentesco.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subParentesco"
subLog sERR
End Sub

Sub subPerfiles()
On Local Error GoTo errH
cmbPerfiles.Limpiar
sSql = "select id,nombre from tperfiles order by id"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbPerfiles.addElement "(Nuevo...)", -1
    While Not objRstA.EOF
        cmbPerfiles.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subPerfiles"
subLog sERR
End Sub
Sub subDepartamentos()
On Local Error GoTo errH
idDep = 0
idDepartamento = 0
cmbDepartamentos.Limpiar
cmbDepartamentos.ZOrder 0
sSql = "select * from tDepartamentos where idcompañia=" & cmbCompañias.itemID & " order by nombre"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbDepartamentos.addElement "(Nuevo...)", -1
    While Not objRstA.EOF
        cmbDepartamentos.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subDependencias"
subLog sERR
End Sub

Private Sub Form_Load()
subCompañias cmbCompañias, True
subCargos
subHorarios cmbHorarios, True
subEPS
subARP
subAFP
subPerfiles
subParentesco
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmPrincipal.subLimpiar
End Sub

Private Sub imgEditaAFP_Click()
On Local Error GoTo errH
Dim tmpID As Long
If cmbAfp.itemID > 0 Then
    Load frmMsg
    frmMsg.lblTitulo.Caption = "Modificar nombre:"
    frmMsg.txtNombre.Text = cmbAfp.Text
    frmMsg.Show vbModal
    If frmMsg.bM Then
        sSql = "update tafp set nombre='" & Trim(frmMsg.txtNombre.Text) & "' where id=" & cmbAfp.itemID
        objCon.Execute sSql
        tmpID = cmbAfp.itemID
        subAFP
        cmbAfp.mostrarItem tmpID
    End If
End If
Exit Sub
errH:
If Err.Number = -2147467259 Then
    MsgBox "El valor ingresado ya está en la lista!", vbInformation
    cmbAfp.itemID = 0
End If
End Sub

Private Sub imgEditaARP_Click()
On Local Error GoTo errH
Dim tmpID As Long
If cmbArp.itemID > 0 Then
    Load frmMsg
    frmMsg.lblTitulo.Caption = "Modificar nombre:"
    frmMsg.txtNombre.Text = cmbArp.Text
    frmMsg.Show vbModal
    If frmMsg.bM Then
        sSql = "update tarp set nombre='" & Trim(frmMsg.txtNombre.Text) & "' where id=" & cmbArp.itemID
        objCon.Execute sSql
        tmpID = cmbArp.itemID
        subARP
        cmbArp.mostrarItem tmpID
    End If
End If
Exit Sub
errH:
If Err.Number = -2147467259 Then
    MsgBox "El valor ingresado ya está en la lista!", vbInformation
    cmbArp.itemID = 0
End If
End Sub

Private Sub imgEditaCargo_Click()
On Local Error GoTo errH
Dim tmpID As Long
If cmbCargos.itemID > 0 Then
    Load frmMsg
    frmMsg.lblTitulo.Caption = "Modificar nombre:"
    frmMsg.txtNombre.Text = cmbCargos.Text
    frmMsg.Show vbModal
    If frmMsg.bM Then
        sSql = "update tcargos set nombre='" & Trim(frmMsg.txtNombre.Text) & "' where id=" & cmbCargos.itemID
        objCon.Execute sSql
        tmpID = cmbCargos.itemID
        subCargos
        cmbCargos.mostrarItem tmpID
    End If
End If
Exit Sub
errH:
If Err.Number = -2147467259 Then
    MsgBox "El valor ingresado ya está en la lista!", vbInformation
    cmbCargos.itemID = 0
End If

End Sub

Private Sub imgEditaCompañia_Click()
On Local Error GoTo errH
Dim tmpID As Long
If cmbCompañias.itemID > 0 Then
    Load frmMsg
    frmMsg.lblTitulo.Caption = "Modificar nombre:"
    frmMsg.txtNombre.Text = cmbCompañias.Text
    frmMsg.Show vbModal
    If frmMsg.bM Then
        sSql = "update tcompañias set nombre='" & Trim(frmMsg.txtNombre.Text) & "' where id=" & cmbCompañias.itemID
        objCon.Execute sSql
        tmpID = cmbCompañias.itemID
        subCompañias cmbCompañias, True
        cmbCompañias.mostrarItem tmpID
    End If
End If
Exit Sub
errH:
If Err.Number = -2147467259 Then
    MsgBox "El valor ingresado ya está en la lista!", vbInformation
    cmbCompañias.itemID = 0
End If
End Sub

Private Sub imgBuscar_Click()
frmBuscar.Show vbModal
If Val(frmBuscar.Tag) <> 0 Then
    idEmpleado = Val(frmBuscar.Tag)
    Unload frmBuscar
    subDatos "id", CStr(idEmpleado)
End If
End Sub

Private Sub imgEditaDepartamento_Click()
Dim tD As Long
If cmbDepartamentos.itemID > 0 Then
    tD = cmbDepartamentos.itemID
    Load frmDepartamentos
    frmDepartamentos.cmbDepartamentos.mostrarItem cmbDepartamentos.itemID
    frmDepartamentos.Show vbModal
    Unload frmDepartamentos
    subDepartamentos
    cmbDepartamentos.mostrarItem tD
End If
End Sub

Private Sub imgEditaEPS_Click()
On Local Error GoTo errH
Dim tmpID As Long
If cmbEps.itemID > 0 Then
    Load frmMsg
    frmMsg.lblTitulo.Caption = "Modificar nombre:"
    frmMsg.txtNombre.Text = cmbEps.Text
    frmMsg.Show vbModal
    If frmMsg.bM Then
        sSql = "update teps set nombre='" & Trim(frmMsg.txtNombre.Text) & "' where id=" & cmbEps.itemID
        objCon.Execute sSql
        tmpID = cmbEps.itemID
        subEPS
        cmbEps.mostrarItem tmpID
    End If
End If
Exit Sub
errH:
If Err.Number = -2147467259 Then
    MsgBox "El valor ingresado ya está en la lista!", vbInformation
    cmbEps.itemID = 0
End If
End Sub

Private Sub imgEditarPerfiles_Click()
On Local Error GoTo errH
Dim tmpID As Long
If cmbPerfiles.itemID > 0 Then
    If cmbPerfiles.itemID = 1 Then
        MsgBox "El perfil seleccionado no se puede modificar!"
    Else
        tmpID = cmbPerfiles.itemID
        Load frmPerfiles
        frmPerfiles.idPerfil = cmbPerfiles.itemID
        frmPerfiles.subCargaPerfil
        frmPerfiles.Show vbModal
        Unload frmPerfiles
        subPerfiles
        cmbPerfiles.mostrarItem tmpID
    End If
End If
Exit Sub
errH:
If Err.Number = -2147467259 Then
    MsgBox "El valor ingresado ya está en la lista!", vbInformation
    cmbEps.itemID = 0
End If

End Sub

Private Sub imgFecha_Click(Index As Integer)
frmCalendario.Show vbModal
If frmCalendario.Tag <> vbNullString Then
    If Index = 0 Then txtFechaI.Text = frmCalendario.Tag
    If Index = 1 Then txtFechaF.Text = frmCalendario.Tag
    If Index = 2 Then txtFechaNac.Text = frmCalendario.Tag
End If
Unload frmCalendario
End Sub

Private Sub imgFoto_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo errH:
If Button = vbLeftButton Then
    If bCam Then tmrCam.Enabled = Not tmrCam.Enabled
Else
    dlgFoto.CancelError = True
    dlgFoto.Filter = "Archivos de imagen|*.jpg;*.bmp;*.gif"
    dlgFoto.DialogTitle = "Seleccionar foto"
    dlgFoto.ShowOpen
    imgFoto.Picture = LoadPicture(dlgFoto.FileName)
    bFoto = True: bModificaFoto = True
End If
Exit Sub
errH:
If Err.Number <> 32755 Then
    sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_imgFoto_MouseUp"
    subLog sERR
End If
End Sub

Private Sub imgHuella_Click()
fraHuelleros.Visible = True
End Sub

Private Sub imgTitular_Click()

End Sub

Private Sub imgUareU_Click()
If bHuellasU Then
    fraHuelleros.Visible = False
    bHuellaOrigen = 2
    Set frmEnrola.objImagen = imgHuella
    '''frmPrincipal.objUareU.StopCapture
    frmEnrola.Show vbModal
    '''frmPrincipal.subUareU
End If
End Sub

Private Sub imgZK_Click()
If idEmpleado = 0 Then
    MsgBox "Por favor guarde los datos del funcionario, cargue los datos, y luego acceda a esta función nuevamente.", vbInformation
    fraHuelleros.Visible = False
Else
''''    bHuellaOrigen = 2
''''    Set frmEnrolaZK.objImagen = imgHuella
''''    '''frmPrincipal.objUareU.StopCapture
''''    frmEnrolaZK.Show vbModal
''''    '''frmPrincipal.subUareU
    frmZkDedos.Show vbModal
End If
End Sub

Private Sub tmrCam_Timer()
If bFoto Then
    bModificaFoto = True
Else
    bFoto = True
    bModificaFoto = True
End If
imgFoto.Picture = frmPrincipal.imgFoto.Picture
End Sub

Private Sub txtApellidos_Validate(Cancel As Boolean)
txtApellidos.Text = fnMayúscula(txtApellidos.Text)
End Sub

Private Sub txtDoc1_Validate(Cancel As Boolean)
idEmpleado = 0
subDatos "documento", txtDoc1.Text
End Sub
Public Sub subDatos(sCampo As String, sValor As String)
On Local Error GoTo errH
sDOC = Trim(sValor)
If sDOC <> vbNullString Then
    sDOC = Replace(sDOC, ".", "")
    sDOC = Replace(sDOC, ",", "")
    sDOC = Replace(sDOC, "-", "")
    If sCampo = "id" Then
        sSql = "select * from templeados where " & sCampo & "=" & sDOC
    Else
        sSql = "select * from templeados where " & sCampo & "='" & sDOC & "'"
    End If
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenKeyset, adLockOptimistic
        If Not .EOF Then
            idEmpleado = !id
            idCarnet = !id
            txtDoc1.Text = "" & !documento
            txtNombre.Text = "" & !nombre
            txtApellidos.Text = "" & !apellidos
            txtRh.Text = "" & !rh
            txtSexo.Text = "" & !sexo
            cmbCompañias.mostrarItem Val("" & !idcompañia)
            cmbDepartamentos.mostrarItem Val("" & !idDepartamento)
            cmbCargos.mostrarItem Val("" & !idcargo)
            cmbHorarios.mostrarItem Val("" & !idHorario)
            txtOficina.Text = "" & !oficina
            txtTelefono.Text = "" & !tel
            txtExtension.Text = "" & !extension
            txtMovil.Text = "" & !movil
            txtEmail.Text = "" & !email
            
            txtFechaI.Text = "" & !fechaI
            txtFechaF.Text = "" & !fechaf
            If Not IsNull(!foto) Then
                bFoto = True
                fnLeeFoto !foto, imgFoto
            End If
            If Not IsNull(!huella) Then
                bHuella = True
                fnLeeFoto !huella, imgHuella
            End If
            txtUsuario.Text = "" & objRst!usuario
            txtContraseña.Text = "" & objRst!contraseña
            txtConfirma.Text = "" & objRst!contraseña
            cmbEps.mostrarItem Val("" & !ideps)
            cmbArp.mostrarItem Val("" & !idarp)
            cmbAfp.mostrarItem Val("" & !idafp)
            cmbPerfiles.mostrarItem Val("" & !idPerfil)
            txtFechaNac.Text = "" & !fecha_nac
            If Not IsNull(!activo) Then
                chkActivo.Value = IIf(!activo, vbChecked, vbUnchecked)
            End If
            txtTarjetaNum.Text = "" & !tarjeta_num
            idTitular = Val("" & !idTitular)
            cmbParentesco.mostrarItem Val("" & !idparentesco)
            cmdImprime.Visible = True
            If Not IsNull(!nomina) Then
                chkNomina.Value = IIf(!nomina, vbChecked, vbUnchecked)
            End If
            txtCódigo.Text = "" & !codigo
        End If
        .Close
    End With
    subDocTitular
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subDatos"
subLog sERR
End Sub

Private Sub txtEmail_Validate(Cancel As Boolean)
If Trim(txtEmail.Text) <> vbNullString Then
    txtEmail.Text = LCase(txtEmail.Text)
    If Not fnEmail(txtEmail.Text) Then
        MsgBox "La dirección de correo electrónico no es válida!", vbInformation
        Cancel = True
    End If
End If
End Sub

Private Sub txtFechaI_Validate(Cancel As Boolean)
If Trim(txtFechaI.Text) <> vbNullString Then
    If Not IsDate(Trim(txtFechaI.Text)) Then
        MsgBox "Fecha no válida!", vbInformation
        txtFechaI.Text = vbNullString
        Cancel = True
    End If
End If
End Sub
Private Sub txtFechaF_Validate(Cancel As Boolean)
If Trim(txtFechaF.Text) <> vbNullString Then
    If Not IsDate(Trim(txtFechaF.Text)) Then
        MsgBox "Fecha no válida!", vbInformation
        txtFechaF.Text = vbNullString
        Cancel = True
    End If
End If
End Sub

Private Sub txtNombre_Validate(Cancel As Boolean)
txtNombre.Text = fnMayúscula(txtNombre.Text)
End Sub

Private Sub txtRh_Validate(Cancel As Boolean)
txtRh.Text = UCase(txtRh.Text)
End Sub

Private Sub txtSexo_Validate(Cancel As Boolean)
txtSexo.Text = UCase(txtSexo.Text)
End Sub

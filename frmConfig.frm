VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Begin VB.Form frmConfig 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00909890&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9135
   ControlBox      =   0   'False
   Icon            =   "frmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   9135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
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
      Height          =   5535
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   4335
      Begin A1AControles.A1ATextBox txtPuertoDatos 
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   9476240
         passChar        =   ""
         ColorFoco       =   7598073
      End
      Begin A1AControles.A1ATextBox txtRun 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   3960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   9476240
         passChar        =   ""
      End
      Begin A1AControles.A1AComboBox cmbCamaras 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         bkColor         =   9476240
         ColorFoco       =   7598073
      End
      Begin A1AControles.A1AComboBox cmbCamObjetos 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   2760
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         bkColor         =   9476240
         ColorFoco       =   7598073
      End
      Begin ARButtonCtrl.ARButton cmdRun 
         Height          =   315
         Left            =   3840
         TabIndex        =   8
         Top             =   3960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         Caption         =   ">"
         ForeColor       =   16777215
         ForeColorOnFocus=   16777215
         BackColorOnMouse=   16777215
         BackColor       =   255
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
      Begin A1AControles.A1AComboBox cmbImpresoras 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         bkColor         =   9476240
         ColorFoco       =   7598073
      End
      Begin A1AControles.A1AComboBox cmbImpresorasR 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         bkColor         =   9476240
         ColorFoco       =   7598073
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto recepción datos:"
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
         Left            =   240
         TabIndex        =   31
         Top             =   4590
         Width           =   2490
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora para Reportes:"
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
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   2685
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora para stickers:"
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
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ejecutar Instruccion SQL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Solo para usuarios avanzados!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   3240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cámara objetos:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cámara fotografía:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4335
      Begin VB.Frame Frame3 
         BackColor       =   &H00909890&
         Caption         =   "Frame3"
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
         Height          =   1095
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   975
         Begin VB.Image imgEST 
            Height          =   360
            Index           =   1
            Left            =   120
            MouseIcon       =   "frmConfig.frx":70E2
            MousePointer    =   99  'Custom
            Picture         =   "frmConfig.frx":73EC
            ToolTipText     =   "Buscar..."
            Top             =   600
            Width           =   360
         End
         Begin VB.Image imgEST 
            Height          =   360
            Index           =   0
            Left            =   120
            MouseIcon       =   "frmConfig.frx":7AEE
            MousePointer    =   99  'Custom
            Picture         =   "frmConfig.frx":7DF8
            ToolTipText     =   "Buscar..."
            Top             =   240
            Width           =   360
         End
      End
      Begin MSComCtl2.DTPicker horaSalida 
         Height          =   375
         Left            =   1440
         TabIndex        =   29
         Top             =   5040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   83099650
         CurrentDate     =   41114
      End
      Begin VB.Image imgAntiPass 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":84FA
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":8804
         Top             =   6840
         Width           =   360
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anti passback"
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
         Left            =   720
         TabIndex        =   37
         Top             =   6900
         Width           =   1485
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Validar eventos ZK"
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
         Left            =   720
         TabIndex        =   36
         Top             =   6540
         Width           =   1980
      End
      Begin VB.Image imgZk_ev 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":8F06
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":9210
         Top             =   6480
         Width           =   360
      End
      Begin VB.Image imgNoEsLabor 
         Height          =   360
         Left            =   720
         MouseIcon       =   "frmConfig.frx":9912
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":9C1C
         Top             =   3960
         Width           =   360
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No E/S en horas laborales"
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
         Left            =   1200
         TabIndex        =   35
         Top             =   4020
         Width           =   2760
      End
      Begin VB.Image imgIntegridad 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":A31E
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":A628
         Top             =   5760
         Width           =   360
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Integridad Puertas Dependientes"
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
         Left            =   720
         TabIndex        =   34
         Top             =   5820
         Width           =   3435
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E/S por la misma puerta"
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
         Left            =   1200
         TabIndex        =   33
         Top             =   6180
         Width           =   2505
      End
      Begin VB.Image imgPuertaES 
         Height          =   360
         Left            =   720
         MouseIcon       =   "frmConfig.frx":AD2A
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":B034
         Top             =   6120
         Width           =   360
      End
      Begin VB.Image imgVoz 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":B736
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":BA40
         Top             =   5400
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Habilitar voz"
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
         Left            =   720
         TabIndex        =   32
         Top             =   5460
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
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
         Left            =   720
         TabIndex        =   30
         Top             =   5100
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Habilitar hora auto salida"
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
         Left            =   720
         TabIndex        =   28
         Top             =   4740
         Width           =   2640
      End
      Begin VB.Image imgHoraAutoSalida 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":C142
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":C44C
         Top             =   4680
         Width           =   360
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso manual visitantes"
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
         Left            =   720
         TabIndex        =   27
         Top             =   4380
         Width           =   2655
      End
      Begin VB.Image imgIngresoManualVis 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":CB4E
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":CE58
         Top             =   4320
         Width           =   360
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restricción por horario"
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
         Left            =   720
         TabIndex        =   24
         Top             =   3660
         Width           =   2385
      End
      Begin VB.Image imgRestrictHorario 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":D55A
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":D864
         Top             =   3600
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salida automática visitantes"
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
         Left            =   720
         TabIndex        =   23
         Top             =   3300
         Width           =   2940
      End
      Begin VB.Image imgAutoSalida 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":DF66
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":E270
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargar datos visita anterior"
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
         Left            =   720
         TabIndex        =   22
         Top             =   2940
         Width           =   2835
      End
      Begin VB.Image imgDatosVisAnterior 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":E972
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":EC7C
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar mensajes de error"
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
         Left            =   720
         TabIndex        =   21
         Top             =   2580
         Width           =   2745
      End
      Begin VB.Image imgErrores 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":F37E
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":F688
         Top             =   2520
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enrolar visitantes"
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
         Left            =   720
         TabIndex        =   20
         Top             =   2220
         Width           =   1815
      End
      Begin VB.Image imgEnrolaVis 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":FD8A
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":10094
         Top             =   2160
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresión de sticker"
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
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   2115
      End
      Begin VB.Image imgFre 
         Height          =   360
         Left            =   720
         MouseIcon       =   "frmConfig.frx":10796
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":10AA0
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frecuentes"
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
         Left            =   1200
         TabIndex        =   18
         Top             =   1860
         Width           =   1170
      End
      Begin VB.Image imgAut 
         Height          =   360
         Left            =   720
         MouseIcon       =   "frmConfig.frx":111A2
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":114AC
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Automáticos"
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
         Left            =   1200
         TabIndex        =   17
         Top             =   1500
         Width           =   1290
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visitantes"
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
         Left            =   720
         TabIndex        =   16
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Image imgVis 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":11BAE
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":11EB8
         Top             =   1080
         Width           =   360
      End
      Begin VB.Image imgFun 
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmConfig.frx":125BA
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":128C4
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionarios"
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
         Left            =   720
         TabIndex        =   15
         Top             =   780
         Width           =   1350
      End
   End
   Begin ARButtonCtrl.ARButton cmdAccesos 
      Height          =   495
      Left            =   4680
      TabIndex        =   11
      Top             =   5880
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      Caption         =   "Configurar Control de Acceso"
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
   Begin VB.Image cmdAceptar 
      Height          =   555
      Left            =   7320
      MouseIcon       =   "frmConfig.frx":12FC6
      MousePointer    =   99  'Custom
      Picture         =   "frmConfig.frx":132D0
      Top             =   6480
      Width           =   1725
   End
   Begin VB.Image cmdCancelar 
      Height          =   555
      Left            =   4680
      MouseIcon       =   "frmConfig.frx":1655E
      MousePointer    =   99  'Custom
      Picture         =   "frmConfig.frx":16868
      Top             =   6480
      Width           =   1725
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim bMostrarErroresL As Boolean

Dim idImpresoraL As Integer
Dim sImpresoraL As String
Dim idImpresoraRL As Integer
Dim sImpresoraRL As String
Dim bEnrolaVISL As Boolean
Dim bStickerFL As Boolean
Dim bStickerVL As Boolean
Dim bStickerAL As Boolean
Dim bStickerFrL As Boolean
Dim idCamaraL As Byte
Dim idCamaraOL As Byte
Dim bDatosVisAnteriorL As Boolean
Dim bAutoSalidaL As Boolean
Dim bRestrictHorarioL As Boolean
Dim bIngresoManualVisL As Boolean
Dim sHoraAutoSalidaL As String
Dim bHoraAL As Boolean
Dim iPuerto_DatosL As Integer
Dim bVozL As Boolean
Dim bPuertaESL As Boolean
Dim bIntegridadL As Boolean
Dim bNoESLaborL As Boolean
Dim bZk_evL As Boolean
Dim bAntiPassL As Boolean



'Dim idDISP As Long
''

''Private Sub chkRegistradora_Click()
''fraRegistradora.Visible = (chkRegistradora.Value = vbChecked)
''If cmbModo.itemID = 1 Then
''    lblE.Visible = True
''    txtPuertoE.Visible = True
''
''    lblS.Visible = False
''    txtPuertoS.Visible = False
''ElseIf cmbModo.itemID = 2 Then
''    lblE.Visible = False
''    txtPuertoE.Visible = False
''
''    lblS.Visible = True
''    txtPuertoS.Visible = True
''ElseIf cmbModo.itemID = 3 Then
''    lblE.Visible = True
''    txtPuertoE.Visible = True
''    lblS.Visible = True
''    txtPuertoS.Visible = True
''End If
''End Sub

Private Sub cmbCamaras_Click()
idCamaraL = cmbCamaras.itemID - 1
End Sub

Private Sub cmbCamaras_Click0()
cmbCamaras.ZOrder 0
End Sub

Private Sub cmbCamObjetos_Click()
idCamaraOL = cmbCamaras.itemID - 1
End Sub

Private Sub cmbCamObjetos_Click0()
cmbCamObjetos.ZOrder 0
End Sub



Private Sub cmbImpresoras_Click()

If cmbImpresoras.itemID <> -1 Then
    sImpresoraL = cmbImpresoras.Text
    idImpresoraL = cmbImpresoras.itemID - 1
End If
End Sub

Private Sub cmbImpresoras_Click0()
cmbImpresoras.ZOrder 0
End Sub

Private Sub cmbImpresorasR_Click()

If cmbImpresorasR.itemID <> -1 Then
    idImpresoraRL = cmbImpresorasR.itemID - 1
    sImpresoraRL = cmbImpresorasR.Text
End If
End Sub

Private Sub cmbImpresorasR_Click0()
cmbImpresorasR.ZOrder 0
End Sub
Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdAceptar_Click()
On Local Error GoTo errH:
bMostrarErrores = bMostrarErroresL

idImpresora = idImpresoraL
sImpresora = sImpresoraL
idImpresoraR = idImpresoraRL
sImpresoraR = sImpresoraRL

bEnrolaVis = bEnrolaVISL
bStickerF = bStickerFL
bStickerV = bStickerVL
bStickerA = bStickerAL
bStickerFr = bStickerFrL
bVoz = bVozL
bPuertaES = bPuertaESL
bIntegridad = bIntegridadL
idCamara = idCamaraL
idCamaraO = idCamaraOL
bDatosVisAnterior = bDatosVisAnteriorL
bAutoSalida = bAutoSalidaL
bRestrictHorario = bRestrictHorarioL
bNoESLabor = bNoESLaborL
bZk_ev = bZk_evL
bAntiPass = bAntiPassL
bIngresoManualVis = bIngresoManualVisL
sHoraAutoSalida = sHoraAutoSalidaL
iPuerto_Datos = iPuerto_DatosL
subConfig True
Unload Me
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_Load"
subLog sERR
End Sub

Private Sub cmdAccesos_Click()
Dim zCn As Integer
On Local Error GoTo errH
'If (Not oZKs) = -1 Then
If Not objGAATools.fnArrVacioCls(oZKs) Then
    zCn = UBound(oZKs)
    If zCn > 0 Then
        frmPrincipal.subDesconectarZK
    End If
End If
Sigue:
frmConfigControl.Show vbModal
'If bSin = False Then
frmPrincipal.subCargaZK
frmPrincipal.subRegistradoras
Exit Sub
errH:
If Err.Number = 9 Then
    GoTo Sigue
End If
End Sub

Private Sub cmdRun_Click()
Dim bResp As Byte
On Local Error GoTo errH
If txtRun.Text <> vbNullString Then
    bResp = MsgBox("Está seguro?", vbYesNo + vbExclamation)
    If bResp Then
        objCon.Execute txtRun.Text
        DoEvents
        MsgBox "Hecho!", vbExclamation
    End If
End If
Exit Sub
errH:
MsgBox "Error " & Err.Number & ". " & Err.Description, vbCritical

End Sub

Private Sub Form_Load()
Dim ix As Byte, iPrn As Long
On Local Error GoTo errH:
bStickerFL = bStickerF: subSwitch imgFun, bStickerF
bStickerVL = bStickerV: subSwitch imgVis, bStickerV
bStickerAL = bStickerA: subSwitch imgAut, bStickerA
bStickerFrL = bStickerFr: subSwitch imgFre, bStickerFr
bEnrolaVISL = bEnrolaVis: subSwitch imgEnrolaVis, bEnrolaVis
bMostrarErroresL = bMostrarErrores: subSwitch imgErrores, bMostrarErrores
bVozL = bVoz: subSwitch imgVoz, bVoz
bPuertaESL = bPuertaES: subSwitch imgPuertaES, bPuertaES
bIntegridadL = bIntegridad: subSwitch imgIntegridad, bIntegridad
bNoESLaborL = bNoESLabor: subSwitch imgNoEsLabor, bNoESLabor
bZk_evL = bZk_ev: subSwitch imgZk_ev, bZk_ev
bAntiPassL = bAntiPass: subSwitch imgAntiPass, bAntiPass
For iPrn = 1 To Printers.Count
    cmbImpresoras.addElement Printers.Item(iPrn - 1).DeviceName, iPrn
    If iPrn = idImpresora + 1 Then
        cmbImpresoras.mostrarItem iPrn
    End If
Next iPrn
For iPrn = 1 To Printers.Count
    cmbImpresorasR.addElement Printers.Item(iPrn - 1).DeviceName, iPrn
    If iPrn = idImpresoraR + 1 Then
        cmbImpresorasR.mostrarItem iPrn
    End If
Next iPrn
idCamaraL = idCamara
idCamaraOL = idCamaraO

bDatosVisAnteriorL = bDatosVisAnterior: subSwitch imgDatosVisAnterior, bDatosVisAnterior
bAutoSalidaL = bAutoSalida: subSwitch imgAutoSalida, bAutoSalida
bRestrictHorarioL = bRestrictHorario: subSwitch imgRestrictHorario, bRestrictHorario
bIngresoManualVisL = bIngresoManualVis: subSwitch imgIngresoManualVis, bIngresoManualVis

bHoraAL = Not (sHoraAutoSalida = vbNullString)
If bHoraAL Then
    horaSalida.Enabled = True
    horaSalida.Value = CDate(sHoraAutoSalida)
Else
    horaSalida.Enabled = False
End If
iPuerto_DatosL = iPuerto_Datos
txtPuertoDatos.Text = iPuerto_DatosL
subSwitch imgHoraAutoSalida, bHoraAL

If frmPrincipal.objVideo.GetVideoDeviceCount > 0 Then
    For ix = 0 To frmPrincipal.objVideo.GetVideoDeviceCount - 1
       cmbCamaras.addElement frmPrincipal.objVideo.GetVideoDeviceName(ix), CLng(ix) + 1
       cmbCamObjetos.addElement frmPrincipal.objVideo.GetVideoDeviceName(ix), CLng(ix) + 1
    Next ix
    cmbCamaras.mostrarItem CLng(idCamaraL + 1)
    cmbCamObjetos.mostrarItem CLng(idCamaraOL + 1)
End If
On Error Resume Next
cmbImpresoras.Text = sImpresora
'''

Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_Load"
subLog sERR
End Sub
Private Sub subSwitch(ByRef img As Image, ByRef var As Boolean)
If var Then
    img.Picture = imgEST(1).Picture
Else
    img.Picture = imgEST(0).Picture
End If
End Sub
Private Sub horaSalida_Change()
sHoraAutoSalidaL = Format(horaSalida, "HH:MM:SS")
End Sub

Private Sub imgAntiPass_Click()
bAntiPassL = Not bAntiPassL
subSwitch imgAntiPass, bAntiPassL

End Sub

Private Sub imgAutoSalida_Click()
bAutoSalidaL = Not bAutoSalidaL
subSwitch imgAutoSalida, bAutoSalidaL
End Sub

Private Sub imgDatosVisAnterior_Click()
bDatosVisAnteriorL = Not bDatosVisAnteriorL
subSwitch imgDatosVisAnterior, bDatosVisAnteriorL
End Sub

Private Sub imgErrores_Click()
bMostrarErroresL = Not bMostrarErroresL
subSwitch imgErrores, bMostrarErroresL
End Sub

Private Sub imgAut_Click()
bStickerAL = Not bStickerAL
subSwitch imgAut, bStickerAL

End Sub

Private Sub imgEnrolaVis_Click()
bEnrolaVISL = Not bEnrolaVISL
subSwitch imgEnrolaVis, bEnrolaVISL
End Sub

Private Sub imgFre_Click()
bStickerFrL = Not bStickerFrL
subSwitch imgFre, bStickerFrL

End Sub

Private Sub imgFun_Click()
bStickerFL = Not bStickerFL
subSwitch imgFun, bStickerFL
End Sub

Private Sub imgHoraAutoSalida_Click()
bHoraAL = Not bHoraAL
If bHoraAL Then
    horaSalida.Enabled = True
    sHoraAutoSalidaL = Format(horaSalida, "HH:MM:SS")
Else
    horaSalida.Enabled = False
    sHoraAutoSalidaL = vbNullString
End If
subSwitch imgHoraAutoSalida, bHoraAL
End Sub

Private Sub imgIngresoManualVis_Click()
bIngresoManualVisL = Not bIngresoManualVisL
subSwitch imgIngresoManualVis, bIngresoManualVisL
End Sub

Private Sub imgIntegridad_Click()
bIntegridadL = Not bIntegridadL
subSwitch imgIntegridad, bIntegridadL

End Sub

Private Sub imgNoEsLabor_Click()
bNoESLaborL = Not bNoESLaborL
subSwitch imgNoEsLabor, bNoESLaborL

End Sub

Private Sub imgPuertaES_Click()
bPuertaESL = Not bPuertaESL
subSwitch imgPuertaES, bPuertaESL

End Sub

Private Sub imgRestrictHorario_Click()
bRestrictHorarioL = Not bRestrictHorarioL
subSwitch imgRestrictHorario, bRestrictHorarioL
End Sub

Private Sub imgVis_Click()
bStickerVL = Not bStickerVL
subSwitch imgVis, bStickerVL
End Sub


Private Sub imgVoz_Click()
bVozL = Not bVozL
subSwitch imgVoz, bVozL
End Sub

Private Sub imgZk_ev_Click()
bZk_evL = Not bZk_evL
subSwitch imgZk_ev, bZk_evL

End Sub

Private Sub txtPuertoDatos_txtCambio()
iPuerto_DatosL = Val(Trim(txtPuertoDatos.Text))
End Sub

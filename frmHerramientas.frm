VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Begin VB.Form frmHerramientas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F8D88F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   10905
   Icon            =   "frmHerramientas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10905
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
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
      Height          =   6495
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   10695
      Begin VB.Frame fraCarnets 
         BackColor       =   &H00F8D88F&
         Caption         =   "Impresión de Carnets"
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
         Height          =   5895
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   10455
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "Teléfonos"
            Enabled         =   0   'False
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
            Height          =   375
            Left            =   2520
            TabIndex        =   36
            Top             =   3480
            Width           =   1935
         End
         Begin VB.CheckBox Check10 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "AFP"
            Enabled         =   0   'False
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
            Height          =   375
            Left            =   2520
            TabIndex        =   35
            Top             =   4920
            Width           =   1935
         End
         Begin VB.CheckBox Check9 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "ARP"
            Enabled         =   0   'False
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
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   4920
            Width           =   1935
         End
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "EPS"
            Enabled         =   0   'False
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
            Height          =   375
            Left            =   2520
            TabIndex        =   33
            Top             =   4560
            Width           =   1935
         End
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "Móvil"
            Enabled         =   0   'False
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
            Height          =   375
            Left            =   2520
            TabIndex        =   32
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "Extensión"
            Enabled         =   0   'False
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
            Height          =   375
            Left            =   2520
            TabIndex        =   31
            Top             =   3840
            Width           =   1935
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "Oficina"
            Enabled         =   0   'False
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
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   4560
            Width           =   1935
         End
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "Cargo"
            Enabled         =   0   'False
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
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "Departamento"
            Enabled         =   0   'False
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
            Height          =   375
            Left            =   240
            TabIndex        =   28
            Top             =   3840
            Width           =   1935
         End
         Begin VB.CheckBox chkCia1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "Compañía"
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
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   3480
            Width           =   1935
         End
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
            Height          =   2190
            Left            =   120
            TabIndex        =   0
            Top             =   840
            Width           =   4395
         End
         Begin VB.FileListBox lstArchivos 
            Height          =   285
            Left            =   4320
            TabIndex        =   12
            Top             =   0
            Visible         =   0   'False
            Width           =   2295
         End
         Begin A1AControles.A1AComboBox cmbDiseño 
            Height          =   315
            Left            =   5880
            TabIndex        =   1
            Top             =   360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            bkColor         =   16308367
         End
         Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer objVisor 
            Height          =   5055
            Left            =   4680
            TabIndex        =   13
            Top             =   720
            Width           =   5595
            lastProp        =   600
            _cx             =   9869
            _cy             =   8916
            DisplayGroupTree=   0   'False
            DisplayToolbar  =   -1  'True
            EnableGroupTree =   0   'False
            EnableNavigationControls=   -1  'True
            EnableStopButton=   0   'False
            EnablePrintButton=   -1  'True
            EnableZoomControl=   -1  'True
            EnableCloseButton=   0   'False
            EnableProgressControl=   0   'False
            EnableSearchControl=   0   'False
            EnableRefreshButton=   0   'False
            EnableDrillDown =   0   'False
            EnableAnimationControl=   0   'False
            EnableSelectExpertButton=   0   'False
            EnableToolbar   =   -1  'True
            DisplayBorder   =   -1  'True
            DisplayTabs     =   0   'False
            DisplayBackgroundEdge=   -1  'True
            SelectionFormula=   ""
            EnablePopupMenu =   0   'False
            EnableExportButton=   0   'False
            EnableSearchExpertButton=   0   'False
            EnableHelpButton=   0   'False
            LaunchHTTPHyperlinksInNewBrowser=   -1  'True
            EnableLogonPrompts=   -1  'True
         End
         Begin ARButtonCtrl.ARButton cmdBuscar 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Tag             =   "12484943"
            ToolTipText     =   "Carnet del último registro..."
            Top             =   360
            Visible         =   0   'False
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   661
            Caption         =   "Buscar funcionario"
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Datos para incluir en el carnet"
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
            TabIndex        =   27
            Top             =   3120
            Width           =   2880
         End
         Begin VB.Image imgPagina 
            Height          =   480
            Left            =   9000
            MouseIcon       =   "frmHerramientas.frx":70E2
            MousePointer    =   99  'Custom
            Picture         =   "frmHerramientas.frx":73EC
            ToolTipText     =   "Alternar Cara"
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diseño:"
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
            Left            =   4800
            TabIndex        =   11
            Top             =   405
            Width           =   705
         End
      End
      Begin VB.Frame fraNovedades 
         BackColor       =   &H00F8D88F&
         Caption         =   "Novedades Empleados"
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
         Height          =   5895
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   8895
         Begin A1AControles.A1AComboBox cmbTipoNovedad 
            Height          =   315
            Left            =   720
            TabIndex        =   2
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            bkColor         =   16308367
         End
         Begin MSComCtl2.DTPicker horaI 
            Height          =   375
            Left            =   2640
            TabIndex        =   5
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   17235970
            CurrentDate     =   40892
         End
         Begin A1AControles.A1ATextBox txtFechaI 
            Height          =   315
            Left            =   720
            TabIndex        =   4
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
         End
         Begin A1AControles.A1ATextBox txtFechaF 
            Height          =   315
            Left            =   4440
            TabIndex        =   6
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
         End
         Begin MSComCtl2.DTPicker horaF 
            Height          =   375
            Left            =   6360
            TabIndex        =   7
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   17235970
            CurrentDate     =   40892
         End
         Begin A1AControles.A1ATextBox txtAnotacion 
            Height          =   315
            Left            =   720
            TabIndex        =   8
            Top             =   1920
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
         End
         Begin A1AControles.A1AComboBox cmbEmpleados 
            Height          =   315
            Left            =   3240
            TabIndex        =   3
            Top             =   720
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   556
            bkColor         =   16308367
         End
         Begin MSDataGridLib.DataGrid objGrid 
            Height          =   2295
            Left            =   120
            TabIndex        =   24
            Top             =   3240
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
            ForeColor       =   11098368
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empleado"
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
            Left            =   3240
            TabIndex        =   23
            Top             =   480
            Width           =   960
         End
         Begin VB.Image cmdAceptar 
            Height          =   555
            Left            =   6210
            MouseIcon       =   "frmHerramientas.frx":7957
            MousePointer    =   99  'Custom
            Picture         =   "frmHerramientas.frx":7C61
            Top             =   2400
            Width           =   1725
         End
         Begin VB.Image cmdCancelar 
            Height          =   555
            Left            =   4080
            MouseIcon       =   "frmHerramientas.frx":AEEF
            MousePointer    =   99  'Custom
            Picture         =   "frmHerramientas.frx":B1F9
            Top             =   2400
            Width           =   1725
         End
         Begin VB.Label Label6 
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
            Left            =   720
            TabIndex        =   22
            Top             =   1680
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Novedad:"
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
            Left            =   720
            TabIndex        =   21
            Top             =   480
            Width           =   1680
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora:"
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
            Left            =   6360
            TabIndex        =   20
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora:"
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
            Left            =   2640
            TabIndex        =   19
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label Label2 
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
            Left            =   4440
            TabIndex        =   18
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   6075
            MouseIcon       =   "frmHerramientas.frx":E487
            MousePointer    =   99  'Custom
            Picture         =   "frmHerramientas.frx":E791
            ToolTipText     =   "Seleccionar fecha"
            Top             =   1320
            Width           =   240
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
            Left            =   720
            TabIndex        =   17
            Top             =   1080
            Width           =   1530
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   2355
            MouseIcon       =   "frmHerramientas.frx":E9E0
            MousePointer    =   99  'Custom
            Picture         =   "frmHerramientas.frx":ECEA
            ToolTipText     =   "Seleccionar fecha"
            Top             =   1320
            Width           =   240
         End
      End
      Begin VB.Label lblNovedades 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00A95900&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOVEDADES"
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
         Left            =   2040
         MouseIcon       =   "frmHerramientas.frx":EF39
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblCarnets 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00A95900&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CARNETS"
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
         Left            =   120
         MouseIcon       =   "frmHerramientas.frx":F243
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   120
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmHerramientas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCrystal As New CRAXDRT.Application
Dim objRpt As New CRAXDRT.Report
Dim objFondo As CRAXDRT.OLEObject
Dim objFondo1 As CRAXDRT.OLEObject
Dim idCarnet As Long

Dim idNovedad As Variant

Private Sub chkCia1_Click()
subVerR
End Sub

Private Sub cmbDiseño_Click()
On Local Error GoTo errH
If cmbDiseño.itemID > 0 Then
    Set objRpt = objCrystal.OpenReport(App.Path & "\Carnets\" & cmbDiseño.Text)
    lstFuncionarios_Click
    While objVisor.IsBusy
        DoEvents
    Wend
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmbDiseño_Click"
subLog sERR
End Sub

Private Sub cmbDiseño_Click0()
cmbDiseño.ZOrder 0
End Sub

Private Sub subVerR()
Dim o As CRAXDRT.TextObject
sSql = "select * from vcarnet where id=" & idCarnet
Set objRst = objCon.Execute(sSql)
With objRpt
    .DiscardSavedData
    .Database.SetDataSource objRst
    .ReadRecords
    If Left(cmbDiseño.Text, 6) <> "Carnet" Then
        .Sections("DetailSection1").ReportObjects("cia1").Suppress = Not (chkCia1.Value = vbChecked)
    End If
    
    DoEvents
End With
With objVisor
    .ReportSource = objRpt
    .EnableExportButton = True
    .DisplayGroupTree = False
    .ViewReport
    .Zoom 2 '1 PageWidth,2 Whole Page
End With

End Sub
Private Sub subFuncionarios()
On Local Error GoTo errH
lstFuncionarios.Clear
cmbEmpleados.Limpiar
sSql = "select id,nombre,apellidos from templeados where id>1 order by nombre,apellidos"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not objRst.EOF
        lstFuncionarios.AddItem Trim("" & !nombre & " " & !apellidos)
        lstFuncionarios.ItemData(lstFuncionarios.NewIndex) = !id
        
        cmbEmpleados.addElement Trim("" & !nombre & " " & !apellidos), !id
        .MoveNext
    Wend
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subFuncionarios"
subLog sERR
End Sub

Private Sub cmbEmpleados_Click0()
cmbEmpleados.ZOrder 0
End Sub

Private Sub cmbTipoNovedad_Click()
On Local Error GoTo errH:
Dim iElem As Long, sElem As String
If cmbTipoNovedad.itemID = -1 Then
    frmNuevo.Show vbModal
    sElem = frmNuevo.Tag
    Unload frmNuevo
    If sElem <> vbNullString Then
        sSql = "select * from ttipo_novedad where id=0"
        With objRstA
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            .AddNew
            !nombre = sElem
            .UpDate
            iElem = !id
            .Close
        End With
        subTipoNovedades
        cmbTipoNovedad.mostrarItem iElem
    End If
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

Private Sub cmbTipoNovedad_Click0()
cmbTipoNovedad.ZOrder 0
End Sub

Private Sub cmdAceptar_Click()
On Local Error GoTo errH
If cmbTipoNovedad.itemID = 0 Then
    MsgBox "Seleccione un Tipo de Novedad!", vbInformation
    cmbTipoNovedad.SetFocus
    Exit Sub
End If
If cmbEmpleados.itemID = 0 Then
    MsgBox "Seleccione un Empleado!", vbInformation
    cmbEmpleados.SetFocus
    Exit Sub
End If
If CDate(txtFechaI.Text & " " & horaI.Hour & ":" & horaI.Minute & ":" & horaI.Second) >= CDate(txtFechaF.Text & " " & horaF.Hour & ":" & horaF.Minute & ":" & horaF.Second) Then
    MsgBox "La fechas u horas inconsistentes!", vbInformation
    txtFechaI.SetFocus
    Exit Sub
End If
sSql = "select * from tnovedades where id=" & Val(idNovedad)
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If .EOF Then .AddNew
    !idtipo = cmbTipoNovedad.itemID
    !idEmpleado = cmbEmpleados.itemID
    !fechaI = fnFecha(txtFechaI.Text & " " & horaI.Hour & ":" & horaI.Minute & ":" & horaI.Second, True)
    !fechaf = fnFecha(txtFechaF.Text & " " & horaF.Hour & ":" & horaF.Minute & ":" & horaF.Second, True)
    !anotacion = Trim(txtAnotacion.Text)
    .UpDate
    .Close
End With
subLimpiar
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_cmdAceptar_Click"
subLog sERR
End Sub
Private Sub subLimpiar()
idNovedad = vbNullString
cmbTipoNovedad.itemID = 0
iniNovedad
cmbEmpleados.itemID = 0
txtAnotacion.Text = vbNullString
cmbTipoNovedad.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    subLimpiar
ElseIf KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
iniCarnet
iniNovedad
End Sub
Private Sub subListar()
sSql = "select * from vnovedades order by id desc"
Set objRst = objCon.Execute(sSql)
Set objGrid.DataSource = objRst
objGrid.Columns("id").Visible = False
End Sub
Private Sub iniNovedad()
txtFechaI.Text = Date
txtFechaF.Text = Date
horaI.Hour = Hour(Time)
horaI.Minute = Minute(Time)
horaI.Second = Second(Time)

horaF.Hour = Hour(Time)
horaF.Minute = Minute(Time)
horaF.Second = Second(Time)
subTipoNovedades
subListar
End Sub
Sub subTipoNovedades()
On Local Error GoTo errH
cmbTipoNovedad.Limpiar
sSql = "select * from ttipo_novedad order by nombre"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    cmbTipoNovedad.addElement "(Nuevo...)", -1
    While Not objRstA.EOF
        cmbTipoNovedad.addElement !nombre, !id
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subTipoNovedades"
subLog sERR
End Sub
Private Sub imgFecha_Click(Index As Integer)
frmCalendario.Show vbModal
If frmCalendario.Tag <> vbNullString Then
    If Index = 0 Then txtFechaI.Text = frmCalendario.Tag
    If Index = 1 Then txtFechaF.Text = frmCalendario.Tag
End If
Unload frmCalendario

End Sub

Private Sub imgPagina_Click()
If objVisor.GetCurrentPageNumber = 1 Then objVisor.ShowNthPage 2 Else objVisor.ShowNthPage 1
End Sub
Private Sub lblCarnets_Click()
fraCarnets.Visible = True
fraNovedades.Visible = False
End Sub
Sub iniCarnet()
Dim ix As Integer
objCrystal.SetLicenseKeycode "AV860-01CS00G-U7000NC"
idCarnet = 0
If Dir(App.Path & "\Carnets\") = vbNullString Then
    MsgBox "No se han encontrado diseños de carnets!", vbInformation
Else
    lstArchivos.Path = App.Path & "\Carnets\"
    lstArchivos.Pattern = "*.rpt"
    For ix = 0 To lstArchivos.ListCount - 1
        cmbDiseño.addElement lstArchivos.List(ix), ix + 1
    Next ix
    'subVerR
    subFuncionarios
    fraCarnets.Visible = True
End If

End Sub
Private Sub lblNovedades_Click()
fraCarnets.Visible = False
fraNovedades.Visible = True
End Sub

Private Sub lstFuncionarios_Click()
If lstFuncionarios.ListIndex <> -1 Then
    If cmbDiseño.itemID > 0 Then
        idCarnet = lstFuncionarios.ItemData(lstFuncionarios.ListIndex)
        sSql = "select * from templeados where id=" & idCarnet
        With objRst
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            objPDF.setSecLevel_nCols -1, 3
            sPDF = "A1A" & Trim(!documento) & "®" & UCase(Trim(!nombre)) & "®" & UCase(Trim(!apellidos)) & "®®®®tc"
            Debug.Print sPDF
            Set objPDFI = objPDF.fnEncode(sPDF)
            If Dir(App.Path & "\tmp_pdf") <> vbNullString Then Kill App.Path & "\tmp_pdf"
            DoEvents
            SavePicture objPDFI, App.Path & "\tmp_pdf"
            DoEvents
            fnGuardaFoto !pdf, App.Path & "\tmp_pdf"
            .UpDate
            .Close
        End With
        subVerR
    End If
End If
End Sub

Public Sub subCargar(id As Long)
Dim i As Integer
For i = 0 To lstFuncionarios.ListCount - 1
    If lstFuncionarios.ItemData(i) = id Then
        lstFuncionarios.ListIndex = i
    End If
Next i
End Sub

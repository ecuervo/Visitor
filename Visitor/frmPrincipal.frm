VERSION 5.00
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{912FB004-DD9A-11D3-BD8D-DAAFCB8D9378}#1.0#0"; "videocapx.ocx"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmPrincipal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   13335
   ControlBox      =   0   'False
   Icon            =   "frmPrincipal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   13335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra1 
      BackColor       =   &H00F8D88F&
      BorderStyle     =   0  'None
      Height          =   10320
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   13335
      Begin zkemkeeperCtl.CZKEM objZK 
         Height          =   615
         Index           =   0
         Left            =   4080
         OleObjectBlob   =   "frmPrincipal.frx":494E
         TabIndex        =   72
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Timer tmrPosHuella 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   3120
         Top             =   1200
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   615
         Left            =   480
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   3495
         Begin VB.Timer tmrZk 
            Enabled         =   0   'False
            Interval        =   2000
            Left            =   3120
            Top             =   120
         End
         Begin MSWinsockLib.Winsock wSerDatos 
            Left            =   2280
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.Timer tmrEspera 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   1440
            Top             =   120
         End
         Begin MSComDlg.CommonDialog objDlg 
            Left            =   1680
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Timer tmrHora 
            Interval        =   1000
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer tmrMovimiento 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   480
            Top             =   0
         End
         Begin VB.Timer tmrCamPreview 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   960
            Top             =   0
         End
         Begin VIDEOCAPXLib.VideoCapX objVideo1 
            Height          =   300
            Left            =   1440
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   465
            _Version        =   131072
            _ExtentX        =   820
            _ExtentY        =   529
            _StockProps     =   1
            BackColor       =   10790052
            PreviewScale    =   -1  'True
            PreviewAudio    =   0   'False
         End
         Begin IMAGXPR5LibCtl.ImagXpress ImgFoto 
            Height          =   285
            Left            =   1920
            TabIndex        =   34
            Top             =   0
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            ErrStr          =   "U9EROCBXRIS-GC305XPXEP"
            ErrCode         =   237570808
            ErrInfo         =   1689331687
            Persistence     =   -1  'True
            _cx             =   79824544
            _cy             =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings"
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
      End
      Begin VB.Frame fra2 
         BackColor       =   &H00F8D88F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   8415
         Left            =   -120
         TabIndex        =   23
         Top             =   1920
         Width           =   13575
         Begin VB.PictureBox picPuertas 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   3255
            Left            =   10440
            ScaleHeight     =   3225
            ScaleWidth      =   2385
            TabIndex        =   81
            Top             =   3600
            Visible         =   0   'False
            Width           =   2415
            Begin VB.ListBox lstPuertas 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00892513&
               Height          =   2430
               ItemData        =   "frmPrincipal.frx":4972
               Left            =   45
               List            =   "frmPrincipal.frx":4974
               TabIndex        =   82
               Top             =   360
               Width           =   2295
            End
            Begin ARButtonCtrl.ARButton cmdPuertaE 
               Height          =   300
               Left            =   45
               TabIndex        =   83
               Tag             =   "12484943"
               ToolTipText     =   "Registrar Objetos"
               Top             =   2880
               Visible         =   0   'False
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   529
               Caption         =   "Entra"
               ForeColor       =   16777215
               ForeColorOnMouse=   8987923
               BackColorOnMouse=   16777215
               BackColor       =   8987923
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ShowFocus       =   2
            End
            Begin ARButtonCtrl.ARButton cmdPuertaS 
               Height          =   300
               Left            =   1320
               TabIndex        =   84
               Tag             =   "12484943"
               ToolTipText     =   "Registrar Objetos"
               Top             =   2860
               Visible         =   0   'False
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   529
               Caption         =   "Sale"
               ForeColor       =   16777215
               ForeColorOnMouse=   8987923
               BackColorOnMouse=   16777215
               BackColor       =   8987923
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ShowFocus       =   2
            End
            Begin VB.Image imgCerrrarPuertas 
               Height          =   315
               Left            =   2040
               MouseIcon       =   "frmPrincipal.frx":4976
               MousePointer    =   99  'Custom
               Picture         =   "frmPrincipal.frx":4C80
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   3135
            Left            =   2400
            TabIndex        =   74
            Top             =   4440
            Visible         =   0   'False
            Width           =   975
            Begin ARButtonCtrl.ARButton cmdLibera 
               Height          =   375
               Index           =   1
               Left            =   360
               TabIndex        =   75
               Tag             =   "12484943"
               Top             =   795
               Visible         =   0   'False
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               Caption         =   ""
               ForeColor       =   16777215
               ForeColorOnMouse=   12484943
               BackColorOnMouse=   16777215
               BackColor       =   16576215
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
               Picture         =   "frmPrincipal.frx":5202
            End
            Begin ARButtonCtrl.ARButton cmdLibera 
               Height          =   375
               Index           =   2
               Left            =   360
               TabIndex        =   76
               Tag             =   "12484943"
               Top             =   1215
               Visible         =   0   'False
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               Caption         =   ""
               ForeColor       =   16777215
               ForeColorOnMouse=   12484943
               BackColorOnMouse=   16777215
               BackColor       =   16576215
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
               Picture         =   "frmPrincipal.frx":5704
            End
            Begin ARButtonCtrl.ARButton cmdLibera 
               Height          =   375
               Index           =   3
               Left            =   360
               TabIndex        =   77
               Tag             =   "12484943"
               Top             =   1650
               Visible         =   0   'False
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               Caption         =   ""
               ForeColor       =   16777215
               ForeColorOnMouse=   12484943
               BackColorOnMouse=   16777215
               BackColor       =   16576215
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
               Picture         =   "frmPrincipal.frx":5C06
            End
            Begin ARButtonCtrl.ARButton cmdLibera 
               Height          =   375
               Index           =   4
               Left            =   360
               TabIndex        =   78
               Tag             =   "12484943"
               Top             =   2070
               Visible         =   0   'False
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               Caption         =   ""
               ForeColor       =   16777215
               ForeColorOnMouse=   12484943
               BackColorOnMouse=   16777215
               BackColor       =   16576215
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
               Picture         =   "frmPrincipal.frx":6108
            End
            Begin ARButtonCtrl.ARButton cmdLibera 
               Height          =   375
               Index           =   5
               Left            =   360
               TabIndex        =   79
               Tag             =   "12484943"
               Top             =   2505
               Visible         =   0   'False
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               Caption         =   ""
               ForeColor       =   16777215
               ForeColorOnMouse=   12484943
               BackColorOnMouse=   16777215
               BackColor       =   16576215
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
               Picture         =   "frmPrincipal.frx":660A
            End
            Begin ARButtonCtrl.ARButton cmdLibera 
               Height          =   375
               Index           =   0
               Left            =   360
               TabIndex        =   80
               Tag             =   "12484943"
               Top             =   360
               Visible         =   0   'False
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               Caption         =   ""
               ForeColor       =   16777215
               ForeColorOnMouse=   12484943
               BackColorOnMouse=   16777215
               BackColor       =   16576215
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
               Picture         =   "frmPrincipal.frx":6B0C
            End
         End
         Begin VB.Frame fraDatosOcultos 
            Caption         =   "Frame2"
            Height          =   1095
            Left            =   3240
            TabIndex        =   70
            Top             =   1320
            Visible         =   0   'False
            Width           =   1695
            Begin A1AControles.A1ATextBox txtFechaNace 
               Height          =   315
               Left            =   120
               TabIndex        =   71
               Top             =   360
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               Text            =   ""
               bkColor         =   16777215
               passChar        =   ""
            End
         End
         Begin VB.CheckBox chkExtra 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "Extra"
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
            Height          =   255
            Left            =   5760
            TabIndex        =   9
            Top             =   1800
            Width           =   830
         End
         Begin VB.PictureBox picFrecuentes 
            BackColor       =   &H00892513&
            Height          =   4215
            Left            =   -5880
            ScaleHeight     =   4155
            ScaleWidth      =   6405
            TabIndex        =   63
            Top             =   960
            Visible         =   0   'False
            Width           =   6460
            Begin VB.ListBox lstFrecuentes 
               Appearance      =   0  'Flat
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
               Height          =   3810
               ItemData        =   "frmPrincipal.frx":700E
               Left            =   120
               List            =   "frmPrincipal.frx":7010
               Style           =   1  'Checkbox
               TabIndex        =   64
               Top             =   120
               Width           =   5775
            End
            Begin VB.Image imgCerrarFrecuentes 
               Height          =   315
               Left            =   6000
               MouseIcon       =   "frmPrincipal.frx":7012
               MousePointer    =   99  'Custom
               Picture         =   "frmPrincipal.frx":731C
               Top             =   120
               Width           =   315
            End
         End
         Begin VB.PictureBox picGrid 
            BackColor       =   &H00BE814F&
            Height          =   4335
            Left            =   -5640
            ScaleHeight     =   4275
            ScaleWidth      =   6405
            TabIndex        =   60
            Top             =   660
            Visible         =   0   'False
            Width           =   6460
            Begin MSDataGridLib.DataGrid objGrid 
               Height          =   4095
               Left            =   60
               TabIndex        =   61
               Top             =   60
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   7223
               _Version        =   393216
               AllowUpdate     =   0   'False
               BackColor       =   8987923
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
         End
         Begin VB.PictureBox picAnotaciones 
            BackColor       =   &H00BE814F&
            Height          =   4335
            Left            =   -6105
            ScaleHeight     =   4275
            ScaleWidth      =   7485
            TabIndex        =   68
            Top             =   120
            Visible         =   0   'False
            Width           =   7545
            Begin MSDataGridLib.DataGrid objGridAnotaciones 
               Height          =   4095
               Left            =   60
               TabIndex        =   69
               Top             =   60
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   7223
               _Version        =   393216
               AllowUpdate     =   0   'False
               BackColor       =   8987923
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
            Begin VB.Image imgQuitaAnotacion 
               Height          =   300
               Left            =   7087
               MouseIcon       =   "frmPrincipal.frx":789E
               MousePointer    =   99  'Custom
               Picture         =   "frmPrincipal.frx":7BA8
               ToolTipText     =   "Quitar anotación"
               Top             =   600
               Width           =   300
            End
            Begin VB.Image imgCerrarAnotaciones 
               Height          =   315
               Left            =   7080
               MouseIcon       =   "frmPrincipal.frx":809A
               MousePointer    =   99  'Custom
               Picture         =   "frmPrincipal.frx":83A4
               Top             =   120
               Width           =   315
            End
         End
         Begin VIDEOCAPXLib.VideoCapX objVideo 
            Height          =   2340
            Left            =   120
            TabIndex        =   65
            Top             =   5760
            Visible         =   0   'False
            Width           =   1755
            _Version        =   131072
            _ExtentX        =   3096
            _ExtentY        =   4128
            _StockProps     =   1
         End
         Begin VB.CheckBox chkFrecuente 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "Frecuente"
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
            Height          =   255
            Left            =   8160
            TabIndex        =   62
            ToolTipText     =   "Registra Automáticamente el ingreso de visitantes"
            Top             =   4800
            Width           =   1380
         End
         Begin IMAGXPR5LibCtl.ImagXpress imgFotos 
            Height          =   1515
            Index           =   2
            Left            =   8520
            TabIndex        =   56
            Top             =   3120
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   2672
            ErrStr          =   "U9EROCBXRIS-GC305XPXEP"
            ErrCode         =   237778165
            ErrInfo         =   -1307435588
            Persistence     =   -1  'True
            _cx             =   79824240
            _cy             =   1
            Picture         =   "frmPrincipal.frx":8926
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings"
               Size            =   15.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   11098368
            AutoSize        =   2
            BorderType      =   3
            ScrollBarLargeChangeH=   10
            ScrollBarSmallChangeH=   1
            DrawFillColor   =   255
            SaveJPGSubSampling=   2
            OLEDropMode     =   0
            CompressInMemory=   2
         End
         Begin A1AControles.A1AComboBox cmbTipoID 
            Height          =   315
            Left            =   3600
            TabIndex        =   1
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            bkColor         =   16308367
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1AComboBox cmbTratamiento 
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            bkColor         =   16308367
            ColorFoco       =   7598073
         End
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
            ItemData        =   "frmPrincipal.frx":33E9C
            Left            =   7560
            List            =   "frmPrincipal.frx":33E9E
            TabIndex        =   26
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CheckBox chkNotsms 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Mensaje SMS MMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Left            =   1440
            TabIndex        =   21
            Top             =   7800
            Width           =   3855
         End
         Begin VB.CheckBox chkNotCorreo 
            Appearance      =   0  'Flat
            BackColor       =   &H0073EFF9&
            Caption         =   "Correo electrónico"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   375
            Left            =   1440
            TabIndex        =   20
            Top             =   7440
            Width           =   1935
         End
         Begin A1AControles.A1ATextBox txtExtension 
            Height          =   315
            Left            =   1440
            TabIndex        =   18
            Top             =   6720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   ""
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1ATextBox txtOficina 
            Height          =   315
            Left            =   3720
            TabIndex        =   17
            Top             =   6195
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   ""
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1ATextBox txtUbicacion 
            Height          =   315
            Left            =   1440
            TabIndex        =   16
            Top             =   6195
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   ""
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1ATextBox txtLocalizacion 
            Height          =   315
            Left            =   1440
            TabIndex        =   15
            Top             =   5670
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   ""
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1ATextBox txtDepartamento 
            Height          =   315
            Left            =   1440
            TabIndex        =   14
            Top             =   5130
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   ""
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1ATextBox txtCompañia 
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Top             =   4605
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   ""
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1ATextBox txtEmpleado 
            Height          =   315
            Left            =   1440
            TabIndex        =   12
            Top             =   4080
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   ""
            ColorFoco       =   16308367
         End
         Begin VB.CheckBox chkAuto 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8D88F&
            Caption         =   "&Automático"
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
            Height          =   255
            Left            =   6600
            TabIndex        =   37
            ToolTipText     =   "Registra Automáticamente el ingreso de visitantes"
            Top             =   4800
            Width           =   1500
         End
         Begin ARButtonCtrl.ARButton cmdObjetos 
            Height          =   495
            Left            =   1560
            TabIndex        =   27
            Tag             =   "12484943"
            ToolTipText     =   "Registrar Objetos"
            Top             =   2760
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   873
            Caption         =   "Objetos reportados Activos"
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
         End
         Begin IMAGXPR5LibCtl.ImagXpress imgFoto1 
            Height          =   2340
            Left            =   6600
            TabIndex        =   31
            Top             =   0
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   4128
            ErrStr          =   "U9EROCBXRIS-GC305XPXEP"
            ErrCode         =   237778165
            ErrInfo         =   -1307435588
            Persistence     =   -1  'True
            _cx             =   79823968
            _cy             =   1
            Picture         =   "frmPrincipal.frx":33EA0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings"
               Size            =   15.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   11098368
            AutoSize        =   2
            BorderType      =   3
            ScrollBarLargeChangeH=   10
            ScrollBarSmallChangeH=   1
            DrawFillColor   =   255
            SaveJPGSubSampling=   2
            OLEDropMode     =   0
            CompressInMemory=   2
         End
         Begin VB.CheckBox chkSensor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8400
            TabIndex        =   36
            ToolTipText     =   "Habilitar la cámara como sensor de movimiento"
            Top             =   0
            Width           =   185
         End
         Begin A1AControles.A1ATextBox txtChat 
            Height          =   315
            Left            =   3480
            TabIndex        =   19
            Top             =   6720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   7598073
            passChar        =   ""
            ColorFoco       =   16308367
         End
         Begin A1AControles.A1ATextBox txtNombre 
            Height          =   315
            Left            =   2640
            TabIndex        =   4
            Top             =   840
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1ATextBox txtApellidos 
            Height          =   315
            Left            =   2640
            TabIndex        =   5
            Top             =   1200
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1ATextBox txtOrganizacion 
            Height          =   315
            Left            =   1320
            TabIndex        =   8
            Top             =   1800
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1ATextBox txtEmail 
            Height          =   315
            Left            =   3120
            TabIndex        =   11
            Top             =   2400
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1ATextBox txtTelefono 
            Height          =   315
            Left            =   1320
            TabIndex        =   10
            Top             =   2400
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
            ColorFoco       =   7598073
         End
         Begin IMAGXPR5LibCtl.ImagXpress imgFotos 
            Height          =   1515
            Index           =   0
            Left            =   8520
            TabIndex        =   54
            Top             =   0
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   2672
            ErrStr          =   "U9EROCBXRIS-GC305XPXEP"
            ErrCode         =   237778165
            ErrInfo         =   -1307435588
            Persistence     =   -1  'True
            _cx             =   79823696
            _cy             =   1
            Picture         =   "frmPrincipal.frx":5F416
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings"
               Size            =   15.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   11098368
            AutoSize        =   2
            BorderType      =   3
            ScrollBarLargeChangeH=   10
            ScrollBarSmallChangeH=   1
            DrawFillColor   =   255
            SaveJPGSubSampling=   2
            OLEDropMode     =   0
            CompressInMemory=   2
         End
         Begin IMAGXPR5LibCtl.ImagXpress imgFotos 
            Height          =   1515
            Index           =   1
            Left            =   8520
            TabIndex        =   55
            Top             =   1560
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   2672
            ErrStr          =   "U9EROCBXRIS-GC305XPXEP"
            ErrCode         =   237778165
            ErrInfo         =   -1307435588
            Persistence     =   -1  'True
            _cx             =   79823424
            _cy             =   1
            Picture         =   "frmPrincipal.frx":8A98C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings"
               Size            =   15.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   11098368
            AutoSize        =   2
            BorderType      =   3
            ScrollBarLargeChangeH=   10
            ScrollBarSmallChangeH=   1
            DrawFillColor   =   255
            SaveJPGSubSampling=   2
            OLEDropMode     =   0
            CompressInMemory=   2
         End
         Begin A1AControles.A1AComboBox cmbSexo 
            Height          =   315
            Left            =   5760
            TabIndex        =   6
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            bkColor         =   16308367
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1AComboBox cmbRH 
            Height          =   315
            Left            =   5760
            TabIndex        =   7
            Top             =   1200
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            bkColor         =   16308367
            ColorFoco       =   7598073
         End
         Begin ARButtonCtrl.ARButton cmdVtoE 
            Height          =   495
            Left            =   5160
            TabIndex        =   67
            Tag             =   "12484943"
            ToolTipText     =   "Visitante-->Funcionario"
            Top             =   2760
            Width           =   1380
            _ExtentX        =   2434
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
            Picture         =   "frmPrincipal.frx":B5F02
            PictureOn       =   "frmPrincipal.frx":B76C4
         End
         Begin A1AControles.A1ATextBox txtDoc1 
            Height          =   315
            Left            =   1320
            TabIndex        =   0
            Top             =   240
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
            ColorFoco       =   7598073
         End
         Begin A1AControles.A1ATextBox txtTarjetaNum 
            Height          =   315
            Left            =   4680
            TabIndex        =   2
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            Text            =   ""
            bkColor         =   16308367
            passChar        =   ""
            ColorFoco       =   7598073
         End
         Begin VB.Image imgMonitor 
            Height          =   585
            Left            =   9000
            MouseIcon       =   "frmPrincipal.frx":B8F76
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":B9280
            ToolTipText     =   "Ver/Ocultar ventana de eventos"
            Top             =   7560
            Width           =   690
         End
         Begin VB.Image imgPuertas 
            Height          =   2505
            Left            =   12840
            MouseIcon       =   "frmPrincipal.frx":BA816
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":BAB20
            ToolTipText     =   "Control de puertas"
            Top             =   3880
            Width           =   435
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
            Left            =   4680
            TabIndex        =   73
            Top             =   0
            Width           =   1005
         End
         Begin VB.Image imgAlerta2 
            Height          =   960
            Left            =   6600
            MouseIcon       =   "frmPrincipal.frx":BE4CA
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":BE7D4
            Top             =   7320
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Image imgAlerta1 
            Height          =   960
            Left            =   5640
            MouseIcon       =   "frmPrincipal.frx":C1816
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":C1B20
            Top             =   7320
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Image imgAlerta 
            Height          =   960
            Left            =   5580
            MouseIcon       =   "frmPrincipal.frx":C4B62
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":C4E6C
            ToolTipText     =   "Anotaciones"
            Top             =   4440
            Width           =   960
         End
         Begin VB.Image imgAnotacion 
            Height          =   960
            Left            =   5580
            MouseIcon       =   "frmPrincipal.frx":C7EAE
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":C81B8
            ToolTipText     =   "Agregar anotación negativa"
            Top             =   3360
            Width           =   960
         End
         Begin VB.Image imgEntra 
            Height          =   1320
            Left            =   10440
            MouseIcon       =   "frmPrincipal.frx":CB1FA
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":CB504
            Top             =   3840
            Width           =   2520
         End
         Begin VB.Image imgMasivo 
            Height          =   930
            Left            =   9910
            MouseIcon       =   "frmPrincipal.frx":CF9C0
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":CFCCA
            ToolTipText     =   "Registrar salida masiva"
            Top             =   4035
            Width           =   450
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo/RH"
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
            Left            =   5640
            TabIndex        =   66
            Top             =   600
            Width           =   825
         End
         Begin VB.Image imgFrecuentes 
            Height          =   240
            Left            =   9600
            MouseIcon       =   "frmPrincipal.frx":D1354
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":D165E
            ToolTipText     =   "Cambiar Visitantes Frecuentes"
            Top             =   4800
            Width           =   240
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   7320
            X2              =   7320
            Y1              =   4680
            Y2              =   4920
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   8880
            X2              =   8880
            Y1              =   4680
            Y2              =   4920
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   9720
            X2              =   7320
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   9720
            X2              =   9720
            Y1              =   4440
            Y2              =   4680
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   10440
            X2              =   9720
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Image imgMinuta 
            Height          =   435
            Left            =   9998
            MouseIcon       =   "frmPrincipal.frx":D1AA0
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":D1DAA
            Top             =   7920
            Width           =   3195
         End
         Begin VB.Image imgFlecha 
            Height          =   720
            Left            =   9920
            MouseIcon       =   "frmPrincipal.frx":D666C
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":D6976
            ToolTipText     =   "Ver visitantes en espera"
            Top             =   5460
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Image imgObjeto 
            Height          =   1815
            Left            =   6750
            Stretch         =   -1  'True
            Top             =   5280
            Width           =   2775
         End
         Begin VB.Image imgFeclas 
            Height          =   375
            Index           =   1
            Left            =   9600
            MouseIcon       =   "frmPrincipal.frx":D7D38
            MousePointer    =   99  'Custom
            ToolTipText     =   "Funcionarios"
            Top             =   7080
            Width           =   255
         End
         Begin VB.Image imgFeclas 
            Height          =   375
            Index           =   0
            Left            =   6435
            MouseIcon       =   "frmPrincipal.frx":D8042
            MousePointer    =   99  'Custom
            ToolTipText     =   "Funcionarios"
            Top             =   7080
            Width           =   255
         End
         Begin VB.Image imgBuscar 
            Height          =   300
            Left            =   3240
            MouseIcon       =   "frmPrincipal.frx":D834C
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":D8656
            Top             =   255
            Width           =   300
         End
         Begin VB.Image imgCarnet 
            Height          =   885
            Left            =   7455
            MouseIcon       =   "frmPrincipal.frx":D8A5B
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":D8D65
            ToolTipText     =   "Previsualizar Sticker"
            Top             =   7440
            Width           =   1365
         End
         Begin VB.Image imgHuella 
            Height          =   2100
            Left            =   6600
            Picture         =   "frmPrincipal.frx":DCD43
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   1785
         End
         Begin VB.Image Image2 
            Height          =   2280
            Left            =   6480
            Picture         =   "frmPrincipal.frx":DE2CA
            Top             =   5160
            Width           =   3315
         End
         Begin VB.Label lblVisT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   12120
            TabIndex        =   58
            Top             =   2760
            Width           =   1035
         End
         Begin VB.Label lblVisMes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   10680
            TabIndex        =   57
            Top             =   2760
            Width           =   1035
         End
         Begin VB.Label lblEmpleados 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   10560
            TabIndex        =   53
            Top             =   6720
            Width           =   1035
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A1A Photo Instant Chat"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   240
            Left            =   3240
            TabIndex        =   52
            Top             =   6495
            Width           =   2055
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellidos"
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
            Left            =   1650
            TabIndex        =   51
            Top             =   1200
            Width           =   885
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tratamiento"
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
            Left            =   1320
            TabIndex        =   50
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
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
            Left            =   3600
            TabIndex        =   49
            Top             =   0
            Width           =   420
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Extensión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   240
            Left            =   1440
            TabIndex        =   48
            Top             =   6495
            Width           =   870
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Oficina"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   240
            Left            =   3720
            TabIndex        =   47
            Top             =   5970
            Width           =   600
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   240
            Left            =   1440
            TabIndex        =   46
            Top             =   5970
            Width           =   855
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Localización"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   240
            Left            =   1440
            TabIndex        =   45
            Top             =   5445
            Width           =   1080
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   240
            Left            =   1440
            TabIndex        =   44
            Top             =   4920
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compañía"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   240
            Left            =   1440
            TabIndex        =   43
            Top             =   4395
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empleado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   240
            Left            =   1440
            TabIndex        =   42
            Top             =   3885
            Width           =   870
         End
         Begin VB.Image Image1 
            Height          =   4845
            Left            =   240
            Picture         =   "frmPrincipal.frx":E42FF
            Top             =   3480
            Width           =   5235
         End
         Begin VB.Image imgEspera 
            Height          =   1320
            Left            =   10440
            MouseIcon       =   "frmPrincipal.frx":F1270
            MousePointer    =   99  'Custom
            Picture         =   "frmPrincipal.frx":F157A
            Top             =   5160
            Width           =   2520
         End
         Begin VB.Label lblTurno 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "275"
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
            Height          =   285
            Left            =   10920
            TabIndex        =   41
            Top             =   3440
            Width           =   420
         End
         Begin VB.Label lblVisitantes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   10680
            TabIndex        =   40
            Top             =   2160
            Width           =   1035
         End
         Begin VB.Label lblDia 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "27"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   12000
            TabIndex        =   39
            Top             =   280
            Width           =   315
         End
         Begin VB.Label lblFecha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   10680
            TabIndex        =   38
            Top             =   880
            Width           =   2595
         End
         Begin VB.Image imgAcceso 
            Height          =   7860
            Left            =   9840
            Picture         =   "frmPrincipal.frx":F5ABC
            Top             =   0
            Width           =   3510
         End
         Begin VB.Image imgIconos 
            Height          =   570
            Index           =   6
            Left            =   720
            Picture         =   "frmPrincipal.frx":FC692
            Top             =   2760
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono:"
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
            Left            =   1320
            TabIndex        =   35
            Top             =   2160
            Width           =   900
         End
         Begin VB.Image imgIconos 
            Height          =   420
            Index           =   5
            Left            =   480
            Picture         =   "frmPrincipal.frx":FCBD1
            Top             =   2400
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Correo electrónico:"
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
            Left            =   3120
            TabIndex        =   30
            Top             =   2160
            Width           =   1830
         End
         Begin VB.Image imgIconos 
            Height          =   705
            Index           =   2
            Left            =   360
            Picture         =   "frmPrincipal.frx":FD08D
            Top             =   1680
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombres"
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
            TabIndex        =   29
            Top             =   600
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID N° DOC"
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
            Left            =   1320
            TabIndex        =   28
            Top             =   0
            Width           =   945
         End
         Begin VB.Image imgIconos 
            Height          =   750
            Index           =   0
            Left            =   240
            Picture         =   "frmPrincipal.frx":FD630
            Top             =   0
            Width           =   1035
         End
         Begin VB.Image imgIconos 
            Height          =   930
            Index           =   1
            Left            =   360
            Picture         =   "frmPrincipal.frx":FDE75
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Organización:"
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
            Left            =   1320
            TabIndex        =   25
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dependencia:"
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
            Left            =   1440
            TabIndex        =   24
            Top             =   5520
            Width           =   1320
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 16"
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
         TabIndex        =   59
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000EBC6C&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         Height          =   255
         Left            =   5900
         Shape           =   4  'Rounded Rectangle
         Top             =   1485
         Width           =   1175
      End
      Begin VB.Image imgHerramientas 
         Height          =   375
         Left            =   6120
         MouseIcon       =   "frmPrincipal.frx":FE545
         MousePointer    =   99  'Custom
         ToolTipText     =   "Otras Utilidades"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgBackup 
         Height          =   375
         Left            =   7560
         MouseIcon       =   "frmPrincipal.frx":FE84F
         MousePointer    =   99  'Custom
         ToolTipText     =   "Copia de seguridad"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgLogo 
         Height          =   870
         Left            =   7080
         MouseIcon       =   "frmPrincipal.frx":FEB59
         MousePointer    =   99  'Custom
         Picture         =   "frmPrincipal.frx":FEE63
         Stretch         =   -1  'True
         ToolTipText     =   "417 x 58 - 100% x 14%"
         Top             =   960
         Width           =   6255
      End
      Begin VB.Image imgRegistro 
         Height          =   375
         Left            =   3000
         MouseIcon       =   "frmPrincipal.frx":1058EB
         MousePointer    =   99  'Custom
         ToolTipText     =   "Licencia"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgWeb 
         Height          =   375
         Left            =   8760
         MouseIcon       =   "frmPrincipal.frx":105BF5
         MousePointer    =   99  'Custom
         ToolTipText     =   "Funcionarios"
         Top             =   0
         Width           =   2895
      End
      Begin VB.Image imgIdeas 
         Height          =   375
         Left            =   5400
         MouseIcon       =   "frmPrincipal.frx":105EFF
         MousePointer    =   99  'Custom
         ToolTipText     =   "Funcionarios"
         Top             =   0
         Width           =   2895
      End
      Begin VB.Image imgReportes 
         Height          =   375
         Left            =   4560
         MouseIcon       =   "frmPrincipal.frx":106209
         MousePointer    =   99  'Custom
         ToolTipText     =   "Ir a Reportes"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgFuncionarios 
         Height          =   375
         Left            =   0
         MouseIcon       =   "frmPrincipal.frx":106513
         MousePointer    =   99  'Custom
         ToolTipText     =   "Ir a Funcionarios"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgConfig 
         Height          =   375
         Left            =   1560
         MouseIcon       =   "frmPrincipal.frx":10681D
         MousePointer    =   99  'Custom
         ToolTipText     =   "Establecer parámetros"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgMinimizar 
         Height          =   375
         Left            =   12240
         MouseIcon       =   "frmPrincipal.frx":106B27
         MousePointer    =   99  'Custom
         ToolTipText     =   "Minimizar Aplicación"
         Top             =   0
         Width           =   255
      End
      Begin VB.Image imgCerrar 
         Height          =   375
         Left            =   12960
         MouseIcon       =   "frmPrincipal.frx":106E31
         MousePointer    =   99  'Custom
         ToolTipText     =   "Cerrar la Aplicación"
         Top             =   0
         Width           =   255
      End
      Begin VB.Image imgApp1 
         Height          =   1860
         Left            =   0
         Picture         =   "frmPrincipal.frx":10713B
         Top             =   0
         Width           =   13335
      End
   End
   Begin VB.Menu mnuSistema 
      Caption         =   "&Sistema"
      Visible         =   0   'False
      Begin VB.Menu mnuConfig 
         Caption         =   "&Configurar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuIr 
      Caption         =   "&Ir a.."
      Visible         =   0   'False
      Begin VB.Menu mnuFuncionarios 
         Caption         =   "&Funcionarios"
      End
      Begin VB.Menu mnuConsultas 
         Caption         =   "&Consultas"
      End
      Begin VB.Menu mnuReportes 
         Caption         =   "&Reportes"
      End
   End
   Begin VB.Menu mnuInterrogante 
      Caption         =   "?"
      Visible         =   0   'False
      Begin VB.Menu mnuLicencia 
         Caption         =   "&Licencia"
      End
   End
   Begin VB.Menu mnuControl 
      Caption         =   "mnuControl"
      Visible         =   0   'False
      Begin VB.Menu mnuLibera 
         Caption         =   "LIBERAR PARA &ENTRADA"
         Index           =   0
      End
      Begin VB.Menu mnuLibera 
         Caption         =   "LIBERAR PARA &SALIDA"
         Index           =   1
      End
      Begin VB.Menu divi 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiberaC 
         Caption         =   "&Cancelar"
      End
   End
   Begin VB.Menu mnuZKs 
      Caption         =   "mnuZKs"
      Visible         =   0   'False
      Begin VB.Menu mnuLiberaZK 
         Caption         =   "&LIBERAR"
      End
      Begin VB.Menu divi1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZKsC 
         Caption         =   "&Cancelar"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sApp As String
Enum enumTipoP
    tpNONE
    tpFUNC
    tpVISI
End Enum
Dim TipoPER As enumTipoP
Dim bEntra As Boolean

Dim txtDestino As Control
Dim bBuscar As Boolean

'
Dim idOrganizacion As Integer
Dim idTurno As Variant
Dim idAutoriza As Integer
Dim idEmpleado As Integer
Dim bFoto As Boolean
Dim iCicloFoto As Byte
'Dim bActivo As Boolean

Dim sNOM As String, sAPE As String
Dim sTIPO_PDF As String
Dim bCapturando As Boolean

Dim bREG As Boolean
Dim idControl As Integer
Dim WithEvents obj2D1 As A1A2D.clsCC
Attribute obj2D1.VB_VarHelpID = -1
Dim WithEvents obj2D2 As A1A2D.clsCC
Attribute obj2D2.VB_VarHelpID = -1
Dim WithEvents obj2D3 As A1A2D.clsCC
Attribute obj2D3.VB_VarHelpID = -1
Dim WithEvents obj2D4 As A1A2D.clsCC
Attribute obj2D4.VB_VarHelpID = -1
Dim WithEvents obj2D5 As A1A2D.clsCC
Attribute obj2D5.VB_VarHelpID = -1
Dim WithEvents obj2D6 As A1A2D.clsCC
Attribute obj2D6.VB_VarHelpID = -1
Dim WithEvents obj2D7 As A1A2D.clsCC
Attribute obj2D7.VB_VarHelpID = -1
Dim WithEvents obj2D8 As A1A2D.clsCC
Attribute obj2D8.VB_VarHelpID = -1
Dim WithEvents obj2D9 As A1A2D.clsCC
Attribute obj2D9.VB_VarHelpID = -1
Dim WithEvents obj2D10 As A1A2D.clsCC
Attribute obj2D10.VB_VarHelpID = -1

Public WithEvents objUareU As DPFPCapture
Attribute objUareU.VB_VarHelpID = -1

Public objPhidget As PhidgetInterfaceKit

''''
Dim bPerFun As Boolean
Dim bPerPar As Boolean
Dim bPerReg As Boolean
Dim bPerRep As Boolean
Dim bPerHer As Boolean
Dim bPerBac As Boolean
Dim bPerMin As Boolean
Dim bPerVis As Boolean

Dim bEsperaTemp As Boolean
Dim DatosDD_ As A1A2D.Datos2D
Dim DatosDD As A1A2D.Datos2D
Dim colSexo(2) As String
Dim colRH(8) As String
Dim bEnrola() As Byte
Dim bAutorizado As Boolean

Public sDocTmp As String
Dim bCuentaZK As Byte
Public bEnrolandoZK As Boolean
Dim idAnotacion As Variant

Dim sHuellaZK As String
Dim iPuertoEManual As Integer
Dim iPuertoSManual As Integer
Private Sub chkAuto_Click()
If chkAuto.Value = vbChecked Then
    bAuto = True
Else
    bAuto = False
End If
subConfig True
End Sub

Private Sub chkSensor_Click()
If bCam Then
    If chkSensor.Value = vbChecked Then
        tmrMovimiento.Enabled = True
        bSensor = True
    Else
        tmrMovimiento.Enabled = False
        bSensor = False
    End If
    subConfig True
End If
End Sub

Private Sub cmbRH_Click0()
cmbRH.ZOrder 0
End Sub

Private Sub cmbSexo_Click0()
cmbSexo.ZOrder 0
End Sub

Private Sub cmbTipoID_Click0()
cmbTipoID.ZOrder 0
End Sub

Private Sub cmbTratamiento_Click0()
cmbTratamiento.ZOrder 0
End Sub

Private Sub cmdLibera_Click(Index As Integer)
idControl = Val(cmdLibera(Index).Tag)
If objPhidget.IsAttached Then
    Me.PopupMenu mnuControl
Else
    mnuLiberaZK.Caption = cmdLibera(Index).toolTipText
    Me.PopupMenu mnuZKs
End If
End Sub

Private Sub cmdPuertaE_Click()
subRELEVO CLng(iPuertoEManual)
End Sub

Private Sub cmdPuertaS_Click()
subRELEVO CLng(iPuertoSManual)
End Sub

Private Sub cmdVtoE_Click()
Dim bResp As Byte
On Local Error GoTo errH
If Not objVISIManual Is Nothing Then
    If objVISIManual.idVISI > 0 Then
        If objVISIManual.bEntraVISI Then
            bResp = MsgBox("Esta persona dejará de ser Visitante y será registrado como Funcionario." & vbCr & "¿Desea continuar?", vbYesNo + vbQuestion)
            If bResp = vbYes Then
                sSql = "insert into templeados(documento,nombre,apellidos,sexo,rh,email,movil,foto,huella,enrola) "
                sSql = sSql & "select documento,nombre,apellidos,sexo,rh,email,telefono,foto,huella,enrola "
                sSql = sSql & "from tvisitantes_huella where id=" & objVISIManual.idVISI
                objCon.Execute (sSql)
                
                
                sSql = "delete from tvisitantes_huella where id=" & objVISIManual.idVISI
                objCon.Execute sSql
                
                sSql = "update tacceso set idhuellero_sale=id where idtpersona=" & objVISIManual.idVISI & " and idtipoper=2"
                objCon.Execute sSql
                
                MsgBox txtNombre.Text & " " & txtApellidos.Text & " ahora está registrado como funcionario." & vbCr & _
                "Deberá ir al módulo de funcionarios para completar la información.", vbInformation
               
                Set objVISIManual = Nothing
                subLimpiar
            End If
        Else
            MsgBox "Registre primero la salida y luego si realice el procedimiento!", vbInformation
        End If
    End If
End If
Exit Sub
errH:
MsgBox "Error " & Err.Number & ". " & Err.Description
End Sub
Private Sub Form_Unload(Cancel As Integer)
objPhidget.Close
DoEvents
End Sub

Private Sub imgAlerta_Click()
Dim objRs_ As New ADODB.Recordset
If imgAlerta.Tag = "1" Then
    sSql = "select id,Fecha_Hora,Anotacion from tanotaciones where documento='" & sDocTmp & "' order by id"
    Set objRs_ = objCon.Execute(sSql)
    Set objGridAnotaciones.DataSource = objRs_
    objGridAnotaciones.Columns("id").Visible = False
    objGridAnotaciones.Columns("Anotacion").Width = 6000
    picAnotaciones.ZOrder 0
    picAnotaciones.Visible = True
End If
End Sub

Private Sub imgAnotacion_Click()
If sDocTmp <> vbNullString Then
    Load frmAnotacion
    frmAnotacion.Show vbModal
End If
End Sub

Private Sub imgBackup_Click()
subValidaPermiso "frmBackup"
End Sub

Private Sub imgBuscar_Click()
frmBuscar1.Show vbModal
If frmBuscar1.Tag <> vbNullString Then
    txtDoc1.Text = frmBuscar1.Tag
    Unload frmBuscar1
    txtDoc1_Validate False
Else
    Unload frmBuscar1
End If
End Sub

Private Sub imgCerrarAnotaciones_Click()
picAnotaciones.Visible = False
End Sub

Private Sub imgCerrarFrecuentes_Click()
picFrecuentes.Visible = False
End Sub

Private Sub imgCerrrarPuertas_Click()
picPuertas.Visible = False
End Sub

Private Sub imgEspera_Click()
If Not objVISIManual Is Nothing Then
    If objVISIManual.bEspera Then
        Exit Sub
    Else
        bEsperaTemp = True
        imgEntra_Click
        subEnEspera
    End If
End If
End Sub

Private Sub imgFeclas_Click(Index As Integer)
subObjetos Index + 1
End Sub

Private Sub imgFlecha_Click()
picGrid.ZOrder 0
picGrid.Visible = Not picGrid.Visible
End Sub

Private Sub imgHerramientas_Click()
subValidaPermiso "frmHerramientas"
End Sub

Private Sub imgLogo_Click()
On Local Error GoTo errH
objDlg.CancelError = True
objDlg.Filter = "Archivos de Imagen|*.jpg;*.bmp;*.gif"
objDlg.ShowOpen
FileCopy objDlg.FileName, App.Path & "\SuLogo"
imgLogo.Picture = LoadPicture(App.Path & "\SuLogo")
Exit Sub
errH:
If Err.Number <> 0 Then
    sERR = "Error " & Err.Number & ". " & Err.Description & Me.name & "_imgEntra_Click"
    subLog sERR
End If
End Sub

Private Sub imgCarnet_Click()
On Local Error Resume Next
Load frmCarnet
frmCarnet.imgLogo.Picture = imgLogo.Picture
frmCarnet.imgFoto1.Picture = imgFoto1.Picture
frmCarnet.lblNombre.Caption = txtNombre.Text
frmCarnet.lblApellidos.Caption = txtApellidos.Text
frmCarnet.lblEmpresa.Caption = txtOrganizacion.Text
frmCarnet.lblDependencia.Caption = txtDepartamento.Text
frmCarnet.lblFecha.Caption = Date & "-" & Time
frmCarnet.imgObjetos.Visible = objVISIManual.bObjetos
frmCarnet.Show vbModal
End Sub

Private Sub imgMasivo_Click()
Dim bResp As Byte
bResp = MsgBox("Registrar salida a todos los Visitantes?", vbYesNo + vbQuestion)
If bResp = vbYes Then
    If modoBD = bdACCESS Then
        sSql = "update tacceso set sale='" & fnFecha(Now, True) & "',idhuellero_sale=-1,idlogin_s=" & idLogin & " where idtipoper=2 and sale is null"
    ElseIf modoBD = bdSQL Then
        sSql = "update tacceso set sale=getdate(),idhuellero_sale=-1,idlogin_s=" & idLogin & " where idtipoper=2 and sale is null"
    End If
    objCon.Execute sSql
    DoEvents
    MsgBox "Hecho!", vbInformation
End If
End Sub

Private Sub imgMinuta_Click()
subValidaPermiso "frmMinuta"
End Sub

Private Sub imgMonitor_Click()
frmMonitor.Visible = Not frmMonitor.Visible
End Sub

Private Sub imgPuertas_Click()
picPuertas.Visible = Not picPuertas.Visible
End Sub

Private Sub imgQuitaAnotacion_Click()
If Val(idAnotacion) > 0 Then
    If MsgBox("Quitar la anotación actual?", vbYesNo + vbQuestion) = vbYes Then
        sSql = "delete from tanotaciones where id=" & Val(idAnotacion)
        objCon.Execute sSql
        idAnotacion = vbNullString
        imgCerrarAnotaciones_Click
    End If
End If
End Sub

Private Sub imgRegistro_Click()
subValidaPermiso "frmLic"
End Sub

Private Sub imgReportes_Click()
subValidaPermiso "frmReportes"
End Sub
Private Sub imgEntra_Click()
Dim sArr() As String, stNom As String
On Local Error GoTo errH:
AccesoTipo = accMANUAL
If objVISIManual Is Nothing Then Exit Sub
If objVISIManual.bEntraVISI Then
    If Trim(txtNombre.Text) = vbNullString Then
        MsgBox "Ingrese Nombre!", vbInformation
        txtNombre.SetFocus
        Exit Sub
    End If
    If Trim(txtApellidos.Text) = vbNullString Then
        MsgBox "Ingrese Apellidos!", vbInformation
        txtApellidos.SetFocus
        Exit Sub
    End If
    idOrganizacion = Val(txtOrganizacion.Tag)
    If idOrganizacion = 0 Then
        If Trim(txtOrganizacion.Text) <> vbNullString Then
            sSql = "select * from torganizaciones where id=" & idOrganizacion
            With objRst
                If .State = adStateOpen Then .Close
                .Open sSql, objCon, adOpenKeyset, adLockOptimistic
                .AddNew
                !nombre = UCase(Trim(txtOrganizacion.Text))
                .UpDate
                txtOrganizacion.Tag = !id
                idOrganizacion = !id
                .Close
            End With
        End If
    End If
    idEmpleado = Val(txtEmpleado.Tag)
    If idEmpleado = 0 Then
        MsgBox "Ingrese Funcionario!", vbInformation
        txtEmpleado.SetFocus
        Exit Sub
    End If
    If idAutoriza = 0 Then idAutoriza = idEmpleado
    If bFoto Then
        If bModificaFoto Then
            If Dir(App.Path & "\tmpFoto") <> vbNullString Then Kill App.Path & "\tmpFoto"
            DoEvents
            imgFoto1.SaveFileType = FT_BMP
            imgFoto1.SaveFileName = App.Path & "\tmpFoto"
            imgFoto1.SaveFile
            ConvertBMPtoJPG App.Path & "\tmpFoto", App.Path & "\tmpFoto" & ".jpg", True, 50, False
        End If
    End If
    If bHuellasU Then
        If bHuella = False Then
            imgHuella_Click
        End If
        If bHuella Then
            If bModificaHuella Then
                If Dir(App.Path & "\tmpHuella") <> vbNullString Then Kill App.Path & "\tmpHuella"
                DoEvents
                SavePicture imgHuella.Picture, App.Path & "\tmpHuella"
                ConvertBMPtoJPG App.Path & "\tmpHuella", App.Path & "\tmpHuella" & ".jpg", True, 50, False
            End If
        End If
    End If
    
    sSql = "select * from tvisitantes_huella where id=" & Val(objVISIManual.idVISI)
    With objRst
        If .State = adStateOpen Then .Close
        '''OJO CAMBIO NUEVO
        .CursorLocation = adUseClient
        .Open sSql, objCon, adOpenKeyset, adLockOptimistic
        If .EOF Then
            .AddNew
            !idLogin = idLogin
        Else
            If IsNull(!idLogin) Then !idLogin = idLogin
        End If
        !documento = Trim(txtDoc1.Text)
        !idtipodoc = cmbTipoID.itemID
        !tarjeta = Trim(txtTarjetaNum.Text)
        !idtratamiento = cmbTratamiento.itemID
        !nombre = fnMayúscula(Trim(txtNombre.Text))
        objVISIManual.sNOM = !nombre
        !apellidos = fnMayúscula(Trim(txtApellidos.Text))
        objVISIManual.sAPE = !apellidos
        !sexo = cmbSexo.Text
        objVISIManual.sSEXO = !sexo
        !rh = cmbRH.Text
        !email = Trim(txtEmail.Text)
        !telefono = Trim(txtTelefono.Text)
        If idOrganizacion > 0 Then !idtorganizacion = idOrganizacion
        If bFoto And bModificaFoto Then fnGuardaFoto !foto, App.Path & "\tmpFoto.jpg"
        DoEvents
        If bHuella And bModificaHuella Then
            fnGuardaFoto !huella, App.Path & "\tmpHuella.jpg"
            DoEvents
            If bEnrolaVis Then
                !enrola = bHuellaMinuciasCAP
            Else
                Dim Tmp() As Byte
                !enrola = Null
            End If
        End If
        '''''
        'If Val(objVISIManual.idVISI) = 0 Or IsNull(!pdf) Then
        
            objPDF.setSecLevel_nCols -1, 3
            sPDF = "A1A" & txtDoc1.Text & "®" & UCase(Trim(txtNombre.Text)) & "®" & UCase(Trim(txtApellidos.Text)) & "®" & idTurno & "®VIAIPI®" & idTurno & "®ts"
            Debug.Print sPDF
            Set objPDFI = objPDF.fnEncode(sPDF)
            If Dir(App.Path & "\tmp_pdf") <> vbNullString Then Kill App.Path & "\tmp_pdf"
            DoEvents
            SavePicture objPDFI, App.Path & "\tmp_pdf"
            DoEvents
            fnGuardaFoto !pdf, App.Path & "\tmp_pdf", False
            
        'End If
        !frecuente = IIf((chkFrecuente.Value = vbChecked), -1, 0)
        !extra = IIf((chkExtra.Value = vbChecked), -1, 0)
        If Trim(txtFechaNace.Text) <> vbNullString Then !fechanac = fnFecha(CDate(txtFechaNace.Text), False)
        ''''
        .UpDate
        objVISIManual.idVISI = !id
        .Close
    End With
'    objVISIManual.idAccesoVISI = !id
'    If objVISIManual.bObjetos Then
'        sSql = "insert into tobjetos(idacceso,descripcion,serial,estado,foto)"
'        sSql = sSql & " select " & objVISIManual.idAccesoVISI & ",tmp_objetos.descripcion,tmp_objetos.serial,1,tmp_objetos.foto from tmp_objetos"
'        objCon.Execute (sSql)
'    End If
    
    
    If objVISIManual.bObjetos Then
        sSql = "update tmp_objetos set idvisitante=" & objVISIManual.idVISI & " where idvisitante is null"
        objCon.Execute (sSql)
    End If
            
    If bIngresoManualVis = False Then
        With objRst
            If .State = adStateOpen Then .Close
            sSql = "select * from tautoriza_acceso_vis where documento='" & objVISIManual.sDOC & "'"
            .Open sSql, objCon, adOpenKeyset, adLockOptimistic
            If .EOF Then .AddNew
            !fecha = fnFecha(Date, False)
            !documento = Trim(txtDoc1.Text)
    
            'If (Not bHuellaMinuciasCAP) = -1 Then
            If objGAATools.fnArrVacioByte(bHuellaMinuciasCAP) Then
                'If (Not bEnrola) = -1 Then
                If objGAATools.fnArrVacioByte(bEnrola) Then
                    !enrola = Null
                Else
                    !enrola = bEnrola
                End If
            Else
                !enrola = bHuellaMinuciasCAP
            End If
            !idEmpleado = idEmpleado
            .UpDate
            .Close
        End With
        
        If objVISIManual.bEntraVISI Then
            If chkAuto.Value = vbUnchecked Then subSticker
        End If
        Set objVISIManual = Nothing
    Else
        subAccesoVISIManual
    End If
    subLimpiar
Else
    subAccesoVISIManual
    subLimpiar
End If
Exit Sub
errH:
If objRst.State = adStateOpen Then
    objRst.CancelUpdate
    objRst.Close
End If
sERR = "Error " & Err.Number & ". " & Err.Description & Me.name & "_imgEntra_Click"
subLog sERR
'Resume
End Sub

Private Sub subSticker()
On Local Error GoTo errH
Dim bSticker As Boolean
Dim iCount As Integer
Dim miPrn As Printer
If idImpresora <> -1 Then
'    If TipoPER = 1 And bStickerF Then bSticker = True
'    If TipoPER = 2 And bStickerV Then bSticker = True
    If bStickerV Then
'''        sSql = "select COUNT(id) as cnt from tobjetos where idacceso=" & objVISIManual.idAccesoVISI
'''        Set objRst = objCon.Execute(sSql)
'''        iCount = objRst!cnt
            Set Printer = Printers(idImpresora)
            Load frmSticker1
            With frmSticker1
                .imgLogo.Picture = imgLogo.Picture
                .lblDoc.Caption = objVISIManual.sDOC
                .lblNombre.Caption = objVISIManual.sNOM
                .lblApellidos.Caption = objVISIManual.sAPE
                .lblEmpresa.Caption = txtOrganizacion.Text
                .lblDependencia.Caption = txtDepartamento.Text
                .lblEntra.Caption = Now
                .ImgFoto.Picture = imgFoto1.Picture
                '.imgObjetos.Visible = (iCount > 0)
                .imgObjetos.Visible = objVISIManual.bObjetos
                .lblVis.Caption = lblTurno.Caption
                If objGAATools.fnExisteArchivo(App.Path & "\tmp_pdf") Then
                    .img2D.Picture = LoadPicture(App.Path & "\tmp_pdf")
                End If
                .subFormato
                '.Show
            End With
        
    End If
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-subSticker"
subLog sERR
End Sub
Private Sub subStickerAuto()
On Local Error GoTo errH
Dim bSticker As Boolean
Dim iCount As Integer
Dim miPrn As Printer
If idImpresora <> -1 Then
'    If TipoPER = 1 And bStickerF Then bSticker = True
'    If TipoPER = 2 And bStickerV Then bSticker = True
    If bStickerV Then
'''        sSql = "select COUNT(id) as cnt from tobjetos where idacceso=" & objVISIManual.idAccesoVISI
'''        Set objRst = objCon.Execute(sSql)
'''        iCount = objRst!cnt
            Set Printer = Printers(idImpresora)
            Load frmSticker1
            With frmSticker1
                .imgLogo.Picture = imgLogo.Picture
                .lblDoc.Caption = objVISI.sDOC
                .lblNombre.Caption = objVISI.sNOM
                .lblApellidos.Caption = objVISI.sAPE
                .lblEmpresa.Caption = txtOrganizacion.Text
                .lblDependencia.Caption = txtDepartamento.Text
                .lblEntra.Caption = Now
                .ImgFoto.Picture = imgFoto1.Picture
                '.imgObjetos.Visible = (iCount > 0)
                .imgObjetos.Visible = objVISI.bObjetos
                .lblVis.Caption = lblTurno.Caption
                If objGAATools.fnExisteArchivo(App.Path & "\tmp_pdf") Then
                    .img2D.Picture = LoadPicture(App.Path & "\tmp_pdf")
                End If
                .subFormato
                '.Show
            End With
        
    End If
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-subSticker"
subLog sERR
End Sub

Private Sub subContar()
On Local Error GoTo errH
If modoBD = bdACCESS Then
    sSql = "select count(idtpersona) as f from vcuenta where fec=#" & fnFecha(Date, False) & "# and idtipoper =1"
    Set objRst = objCon.Execute(sSql)
    lblEmpleados.Caption = "" & objRst!f
    
    sSql = "select count(idtpersona) as v from vcuenta where fec=#" & fnFecha(Date, False) & "# and idtipoper =2"
    Set objRst = objCon.Execute(sSql)
    lblVisitantes.Caption = "" & objRst!v
    
    sSql = "select count(idtpersona) as vm from vcuenta where month(fec)=" & Month(Date) & " and idtipoper =2"
    Set objRst = objCon.Execute(sSql)
    lblVisMes.Caption = "" & objRst!vm
    
    sSql = "select count(idtpersona) as vt from vcuenta where idtipoper =2"
    Set objRst = objCon.Execute(sSql)
    lblVisT.Caption = "" & objRst!vt
ElseIf modoBD = bdSQL Then
    sSql = "select count(idtpersona) as f from vcuenta where fec='" & fnFecha(Date, False) & "' and idtipoper =1"
    Set objRst = objCon.Execute(sSql)
    lblEmpleados.Caption = "" & objRst!f
    
    sSql = "select count(idtpersona) as v from vcuenta where fec='" & fnFecha(Date, False) & "' and idtipoper =2"
    Set objRst = objCon.Execute(sSql)
    lblVisitantes.Caption = "" & objRst!v
    
    sSql = "select count(idtpersona) as vm from vcuenta where datepart(month,fec)=" & Month(Date) & " and idtipoper =2"
    Set objRst = objCon.Execute(sSql)
    lblVisMes.Caption = "" & objRst!vm
    
    sSql = "select count(idtpersona) as vt from vcuenta where idtipoper =2"
    Set objRst = objCon.Execute(sSql)
    lblVisT.Caption = "" & objRst!vt
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-subContar"
subLog sERR
End Sub
Private Sub cmdObjetos_Click()
If Not objVISIManual Is Nothing Or Not objVISI Is Nothing Then
    Load frmObjetos
    frmObjetos.Show vbModal
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    If ActiveControl.name = "txtDoc1" Then
        txtDoc1_Validate False
    Else
        If lstEmerge.Visible Then lstEmerge.ListIndex = 0
        SendKeys "{TAB}"
    End If
ElseIf KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    subLimpiar
End If
End Sub

Private Sub Form_Load()
10    On Local Error GoTo errH
      Dim fechaI As Date
      Dim iPrn As Byte
      Dim iDias1 As Integer
      Dim iTipoLIC As Integer
20    sApp = App.Title ' & " Ver. " & App.Major & "." & App.Minor & "." & App.Revision
30    If App.PrevInstance Then
40        MsgBox sApp & " ya está en ejecución!", vbInformation
50        End
60    Else
          '''
70        sTerminal = fnNombrePC
80        If Not fnConectaCFG Then
90            MsgBox "Error cargando configurarión!", vbCritical
100           End
110       Else
120           iTipoLIC = ValidateSoft(31, "Visitor15")
              'ACA DEBE VALIDAR EL ESTADO DEL SOFT
              ' Retorna 1  cuando el sof esta Activo
              ' Retorna 2  cuando el sof esta pirata
              ' Retorna 3  cuando el sof esta DEM0
              ' Retorna 4  cuando el sof esta demo vencido
              If bSinAccess = False Then
130               sSql = "select * from tconfig"
140               Set objRst = objConCFG.Execute(sSql)
150               If objRst.EOF Then
160                   If objGAATools.fnExisteArchivo(App.Path & "\A1A APP Config.exe") Then
170                       Shell App.Path & "\A1A APP Config.exe", vbNormalFocus
180                       DoEvents
190                   Else
200                       MsgBox "No se encuentra aplicativo de Configuración!", vbCritical
210                   End If
220                   End
230               End If
    
240               modoBD = Val(objRst!modo)
250               sBD = "" & objRst!bd
260               sBDe = "" & objRst!bde
                End If
      '''        If modoBD = bdNONE Then
      '''            frmConfigBD.Show vbModal
      '''        End If
270           If Not fnConecta Then
280               Screen.MousePointer = vbNormal
290               MsgBox "Error de acceso a la base de datos!", vbCritical
300               subLog "Error de acceso a la base de datos!"
310               End
320           End If
330           iDias = 31 - iDias
340           If iTipoLIC <> 1 Then
350               sSql = "select * from tlicencia where terminal='" & sTerminal & "'"
360               Set objRst = objCon.Execute(sSql)
          '        objRst.Close
          '        objRst.Open sSql, objCon, adOpenKeyset, adLockOptimistic
          '        objRst.Delete
          '        objRst.Update
          '        objRst.Close
          
370               If Not objRst.EOF Then
380                   bEOF = False
390                   iDias1 = 30 - DateDiff("d", objRst!fechaI, Date)
400               Else
410                   bEOF = True
420               End If
430               If Not bEOF Then
440                   If iDias <> iDias1 Then
450                       MsgBox "La información de licencia de este producto ha sido alterada!", vbCritical
460                       End
470                   End If
480               End If
490           End If
500           If iTipoLIC = 1 Then
510           ElseIf iTipoLIC = 2 Then
520               MsgBox "La información de licencia de este producto ha sido alterada!", vbCritical
530               End
540           ElseIf iTipoLIC = 3 Then
550               Load frmLic
560               mnuLicencia.Enabled = True
570               Unload frmLic
580           ElseIf iTipoLIC = 4 Then
590               MsgBox "El tiempo de prueba de este producto ha terminado!", vbInformation
600               Load frmLic
610               Unload frmLic
620           End If
              '''
630           bMostrarErrores = True
640           subConfig False
650           subUareU
660           subVoces
670           frmLogin.Show vbModal
680           Screen.MousePointer = vbHourglass
690           Me.Caption = sApp
700           subCargaPerfil
710           If bEntrena = vbYes Then Me.Caption = Me.Caption & " (Modo Entrenamiento)"
720           If Dir(App.Path & "\SuLogo") <> vbNullString Then
730               imgLogo.Picture = LoadPicture(App.Path & "\SuLogo")
740           End If
750           lblFecha.Caption = FormatDateTime(Date, vbShortDate)
760           If idImpresora = -1 Then
770               If MsgBox("No hay una ipresora configurada para Stickers." _
                  & vbCr & "¿Desea seleccionar una ahora?", vbQuestion + vbYesNo) = vbYes Then
780                   Screen.MousePointer = vbNormal
790                   subValidaPermiso "frmConfig"
800               Else
810                   MsgBox "No se imprimirán Stickers!", vbInformation
820               End If
830           End If
              
840           picAnotaciones.Move 2288, 120
850           picFrecuentes.Move 3120, 720
860           picGrid.Move 3360, 3540
              
870           cmbSexo.addElement "M", 1
880           cmbSexo.addElement "F", 2
890           colSexo(1) = "M"
900           colSexo(2) = "F"
              
910           colRH(1) = "A+"
920           colRH(2) = "A-"
930           colRH(3) = "B+"
940           colRH(4) = "B-"
950           colRH(5) = "O+"
960           colRH(6) = "O-"
970           colRH(7) = "AB+"
980           colRH(8) = "AB-"
              
990           cmbRH.addElement "A+", 1
1000          cmbRH.addElement "A-", 2
1010          cmbRH.addElement "B+", 3
1020          cmbRH.addElement "B-", 4
1030          cmbRH.addElement "O+", 5
1040          cmbRH.addElement "O-", 6
1050          cmbRH.addElement "AB+", 7
1060          cmbRH.addElement "AB-", 8
              
1070          subMuestraCAM
1080          subCarga2D
          '''    If iPuerto <> 0 Then
          '''        Set obj2D = New A1A2D.clsCC
          '''        obj2D.Inicializa iPuerto, "@1@"
          '''    End If
          '        subListarDependencias
1090          subContar
1100          bBuscar = True
1110          If Not bAuto Then
1120              chkAuto.Value = vbUnchecked
1130          Else
1140              chkAuto.Value = vbChecked
1150          End If
1160          chkAuto_Click
1170          subTurno
1180          subListarTipoID
1190          subListarTratamiento
              
      '''        sSql = "delete from tmp_objetos where idvisitante is null"
      '''        objCon.Execute sSql
          
1200          imgFoto1_Click
              
1210          subEnEspera

1230          subPhidget
              
1240          If modoBD = bdSQL Then
1250              sSql = "delete from tautoriza_acceso_vis where convert(date,fecha)<>convert(date,GETDATE())"
1260          ElseIf modoBD = bdACCESS Then
1270              sSql = "delete from tautoriza_acceso_vis where fecha<>date()"
1280          End If
1290          objCon.Execute sSql
      '        wServidor.LocalPort = 1515
      '        wServidor.Listen
1300          subCargaZK
1220          subRegistradoras
1310          If iPuerto_Datos > 0 Then
1320              wSerDatos.LocalPort = iPuerto_Datos
1330              wSerDatos.Listen
1340          End If
1350          Screen.MousePointer = vbNormal
1360      End If
1370  End If
1380  Exit Sub
errH:
1390  Screen.MousePointer = vbNormal
1400  sERR = "Error linea " & Erl & ":" & Err.Description & Err.Number & ". " & Err.Description & "-" & Me.name & "_Load"
1410  subLog sERR
End Sub
Public Sub subCargaZK()
Screen.MousePointer = vbHourglass
If modoBD = bdSQL Then
    sSql = "select cd.puerto,cd.nombre,cd.enrola_fun from tcontrol c "
    sSql = sSql & "join tcontrol_disp cd on c.id = cd.idcontrol "
    sSql = sSql & " where abs(c.activa)=1 and abs(cd.activo)=1 and cd.tipo=3 and c.terminal='" & sTerminal & "'"
ElseIf modoBD = bdACCESS Then
    sSql = "SELECT tcontrol_disp.puerto,tcontrol_disp.nombre "
    sSql = sSql & "FROM tcontrol INNER JOIN tcontrol_disp ON tcontrol.id = tcontrol_disp.idcontrol "
    sSql = sSql & "WHERE (((Abs([tcontrol].[activa]))=1) AND ((Abs([tcontrol_disp].[activo]))=1) AND ((tcontrol_disp.tipo)=3) AND ((tcontrol.terminal)='" & sTerminal & "'));"
End If
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        ReDim Preserve oZKs(objZK.Count)
        Set oZKs(objZK.Count) = New clsZK
        Load objZK(objZK.Count)
        objZK(objZK.Count - 1).MachineNumber = objZK.Count - 1
        oZKs(objZK.Count - 1).iIndex = objZK.Count - 1
        oZKs(objZK.Count - 1).sIP = "" & !puerto
        oZKs(objZK.Count - 1).sNombre = "" & !nombre
        oZKs(objZK.Count - 1).bEnrola = IIf(IsNull(!enrola_fun), False, !enrola_fun)
        If objZK(objZK.Count - 1).Connect_Net("" & !puerto, 4370) Then
            oZKs(objZK.Count - 1).bConectado = True
            objZK(objZK.Count - 1).Tag = "" & !puerto
            objZK(objZK.Count - 1).CancelOperation
'            DoEvents
'            objZK(objZK.Count - 1).SetDeviceInfo objZK(objZK.Count - 1).MachineNumber, 5, 0
'            objZK(objZK.Count - 1).GetLastError lZkErr
            DoEvents
            objZK(objZK.Count - 1).StartIdentify
        Else
            oZKs(objZK.Count - 1).bConectado = False
            subLog "Error abriendo lextor ZK " & !puerto
        End If
        .MoveNext
    Wend
End With
Screen.MousePointer = vbNormal
End Sub
Public Sub subCargaPerfil()
On Local Error GoTo errH
sSql = "select id,nombre from tperfiles"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If .EOF Then
        .AddNew
        !nombre = "CONTROL TOTAL"
        .UpDate
        'sSql = "update templeados set activo=-1"
        'objCon.Execute sSql
    End If
    .Close
End With

sSql = "select idperfil from templeados where id=" & idLogin
Set objRst = objCon.Execute(sSql)
idPerf = Val("" & objRst!idPerfil)

bPerFun = False
bPerPar = False
bPerReg = False
bPerRep = False
bPerHer = False
bPerBac = False
bPerMin = False
bPerVis = False

If idPerf = 0 Then
    MsgBox "Usuario sin perfil asignado! - idLoggin=" & idLogin & " - IdPerf=" & idPerf, vbCritical
    End
ElseIf idPerf = 1 Then
    bPerFun = True
    bPerPar = True
    bPerReg = True
    bPerRep = True
    bPerHer = True
    bPerBac = True
    bPerMin = True
    bPerVis = True
Else
    sSql = "select * from tperfiles where id=" & idPerf
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenForwardOnly
        If .EOF Then
            MsgBox "Usuario sin perfil asignado! - idLoggin=" & idLogin & " - IdPerf=" & idPerf, vbCritical
            End
        Else
            bPerFun = !funcionarios
            bPerPar = !parámetros
            bPerReg = !registro
            bPerRep = !reportes
            bPerHer = !herramientas
            bPerBac = !backup1
            bPerMin = !minuta_digital
            bPerVis = !reg_vis_nue
        End If
    End With
    If Not bPerFun And Not bPerPar And Not bPerReg And Not bPerRep And Not bPerHer And Not bPerBac And Not bPerMin And Not bPerVis Then
        MsgBox "Usuario sin perfil asignado! - idLoggin=" & idLogin & " - IdPerf=" & idPerf, vbCritical
        End
    End If
    
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subCargaPerfil"
subLog sERR
End Sub
Private Sub subPhidget()
Dim ms As Boolean
Set objPhidget = New PhidgetInterfaceKit
objPhidget.Open
subEsperar 1
If objPhidget.IsAttached = False Then
    iPhidgetPuertos = 0
    ms = bMostrarErrores
    bMostrarErrores = False
    subLog "Tarjeta Phidget no encontrada!"
    bMostrarErrores = ms
Else
    iPhidgetPuertos = objPhidget.OutputCount
End If

End Sub
Sub subListarTipoID()
On Local Error GoTo errH
sSql = "select * from ttipodoc"
cmbTipoID.Limpiar
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        cmbTipoID.addElement !nombre, !id
        If !defecto Then cmbTipoID.porDefectoElUltimo
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subListarTipoID"
subLog sERR
End Sub
Sub subListarTratamiento()
On Local Error GoTo errH
sSql = "select * from ttratamiento"
cmbTratamiento.Limpiar
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        cmbTratamiento.addElement !nombre, !id
        If !defecto Then cmbTratamiento.porDefectoElUltimo
        .MoveNext
    Wend
    .Close
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subListarTratamiento"
subLog sERR
End Sub
Private Sub subVoces()
On Local Error GoTo errH
bHabla = False
Set objVoces = objHabla.GetVoices
For idx = 0 To objVoces.Count - 1
    If objVoces(idx).GetDescription = "Carlos" Or objVoces(idx).GetDescription = "Soledad" Then
    'If objVoces(idx).GetDescription = "Soledad" Then
        Set objHabla.Voice = objVoces(idx)
        bHabla = True
        Exit For
    End If
Next
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subVoces"
subLog sERR
End Sub
Public Sub subUareU()
On Local Error GoTo errH
Set objUareUs = New DPFPReadersCollection
If objUareUs.Count = 0 Then
    bHuellasU = False
Else
    bHuellasU = True
    Set objVerifica = New DPFPVerification
    Set objCreaFea = New DPFPFeatureExtraction
    For Each objUInf In objUareUs
        Set objUareU = New DPFPCapture
        objUareU.Priority = CapturePriorityHigh
        objUareU.StartCapture
        objUareU.ReaderSerialNumber = objUInf.SerialNumber
    Next
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subUareU"
subLog sERR
End Sub
Public Sub subRegistradoras()
On Local Error GoTo errH
sSql = "select id,nombre from tcontrol where abs(activa)=1 and terminal='" & sTerminal & "' order by id"
lstPuertas.Clear
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        lstPuertas.AddItem "" & !nombre
        lstPuertas.ItemData(lstPuertas.NewIndex) = !id
        .MoveNext
    Wend
End With
Exit Sub
errH:
If Err.Number = 9 Then
    
Else
    sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subRegistradoras"
    subLog sERR
End If
End Sub

Private Sub subRegistradoras_OLD()
Dim zCn As Integer
On Local Error GoTo errH
Dim i As Byte
If objPhidget.IsAttached Then
    sSql = "select id,nombre from tcontrol where abs(activa)=1 order by id"
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenForwardOnly
        While Not .EOF
            cmdLibera(i).Tag = !id
            cmdLibera(i).Visible = True
            cmdLibera(i).toolTipText = !nombre
            i = i + 1
            If i = 6 Then .MoveLast
            .MoveNext
        Wend
    End With
Else
    'If (Not oZKs) = -1& Then
    If Not objGAATools.fnArrVacioCls(oZKs) Then
        zCn = UBound(oZKs)
        If zCn > 0 Then
            For i = 1 To zCn
                If oZKs(i).bConectado Then
                    cmdLibera(i - 1).Tag = oZKs(i).iIndex
                    cmdLibera(i - 1).Visible = True
                    cmdLibera(i - 1).toolTipText = oZKs(i).sNombre
                End If
                If i = 6 Then Exit For
            Next i
        End If
    End If
End If
Exit Sub
errH:
If Err.Number = 9 Then
    
Else
    sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subRegistradoras"
    subLog sERR
End If
End Sub
Private Sub subCarga2D()
On Local Error GoTo errH
Dim i As Byte, bError As Boolean
bError = False
If modoBD = bdSQL Then
    sSql = "select cd.puerto from tcontrol c "
    sSql = sSql & "join tcontrol_disp cd on c.id = cd.idcontrol "
    sSql = sSql & " where abs(c.activa)=1 and abs(cd.activo)=1 and cd.tipo=2 and c.terminal='" & sTerminal & "'"
ElseIf modoBD = bdACCESS Then
    sSql = "SELECT tcontrol_disp.puerto "
    sSql = sSql & "FROM tcontrol INNER JOIN tcontrol_disp ON tcontrol.id = tcontrol_disp.idcontrol "
    sSql = sSql & "WHERE (((Abs([tcontrol].[activa]))=1) AND ((Abs([tcontrol_disp].[activo]))=1) AND ((tcontrol_disp.tipo)=2) AND ((tcontrol.terminal)='" & sTerminal & "'));"
End If
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        i = i + 1
        Select Case i
            Case 1
                Set obj2D1 = New A1A2D.clsCC
                If Not obj2D1.Inicializa(CInt("" & !puerto), "@1@") Then bError = True
            Case 2
                Set obj2D2 = New A1A2D.clsCC
                If Not obj2D2.Inicializa(CInt("" & !puerto), "@1@") Then bError = True
            Case 3
                Set obj2D3 = New A1A2D.clsCC
                If Not obj2D3.Inicializa(CInt("" & !puerto), "@1@") Then bError = True
            Case 4
                Set obj2D4 = New A1A2D.clsCC
                If Not obj2D4.Inicializa(CInt("" & !puerto), "@1@") Then bError = True
            Case 5
                Set obj2D5 = New A1A2D.clsCC
                If Not obj2D5.Inicializa(CInt("" & !puerto), "@1@") Then bError = True
            Case 6
                Set obj2D6 = New A1A2D.clsCC
                If Not obj2D6.Inicializa(CInt("" & !puerto), "@1@") Then bError = True
            Case 7
                Set obj2D7 = New A1A2D.clsCC
                If Not obj2D7.Inicializa(CInt("" & !puerto), "@1@") Then bError = True
            Case 8
                Set obj2D8 = New A1A2D.clsCC
                If Not obj2D8.Inicializa(CInt("" & !puerto), "@1@") Then bError = True
            Case 9
                Set obj2D9 = New A1A2D.clsCC
                If Not obj2D9.Inicializa(CInt("" & !puerto), "@1@") Then bError = True
            Case 10
                Set obj2D10 = New A1A2D.clsCC
                If Not obj2D10.Inicializa(CInt("" & !puerto), "@1@") Then bError = True
        End Select
        If i = 10 Then .MoveLast
        .MoveNext
    Wend
    .Close
End With
If bError Then
    subLog "Error al abrir puerto Lector 2D!"
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subRegistradoras"
End Sub

Private Sub subTurno()
On Local Error GoTo errH
If modoBD = bdACCESS Then
    sSql = "select max(turno) as turno1 from tacceso where CDate(Format([entra],'dd-mm-yyyy'))=#" & fnFecha(Date, False) & "#"
ElseIf modoBD = bdSQL Then
    sSql = "select max(turno) as turno1 from tacceso where convert(date,entra)='" & fnFecha(Date, False) & "'"
End If
Set objRst = objCon.Execute(sSql)
idTurno = Val("" & objRst!turno1) + 1
lblTurno.Caption = Val(idTurno)
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subTurno"
End Sub
Private Sub subMuestraCAM()
On Local Error GoTo errH
Dim idxCam As Integer
If objVideo.GetVideoDeviceCount = 0 Then
    bCam = False
    tmrMovimiento.Enabled = False
    tmrCamPreview.Enabled = False
    MsgBox "No se ha encontrado ninguna cámara conectada en el sistema!", vbInformation
Else
    If idCamara = 0 Then
        If objVideo.GetVideoDeviceCount > 0 Then
            GoTo setCam
        Else
            bCam = False
            tmrMovimiento.Enabled = False
            tmrCamPreview.Enabled = False
            MsgBox "No hay nunguna cámara configurada!", vbInformation
        End If
    Else
        If objVideo.GetVideoDeviceCount > 0 Then
setCam:
            If Left(objVideo.GetVideoDeviceName(idCamara), 4) = "713x" Then
                objVideo.VideoDeviceIndex = 1
            Else
                objVideo.VideoDeviceIndex = idCamara
            End If
            If objVideo.GetVideoDeviceName(objVideo.VideoDeviceIndex) <> "Error: 3" Then
                idCamara = objVideo.VideoDeviceIndex
                subConfig True
                objVideo.AudioDeviceIndex = -1
                objVideo.CaptureAudio = False
                objVideo.UseVideoFilter = vcxBoth
                objVideo.SetVideoFormat 320, 240
                '''objVideo.SetCrop 70, 0, 180, 240
                objVideo.Connected = True
                objVideo.Preview = True
                '''objVideo.SetTextOverlay 0, "A1A", 0, 0, "Arial", 10, vbRed, -1
                bCam = True
            Else
                bCam = False
                tmrMovimiento.Enabled = False
                MsgBox "No se ha encontrado ninguna cámara conectada en el sistema!", vbInformation
            End If
            If bCam Then
                If Not bSensor Then
                    chkSensor.Value = vbUnchecked
                Else
                    chkSensor.Value = vbChecked
                End If
                chkSensor_Click
                tmrCamPreview.Enabled = True
            Else
                tmrCamPreview.Enabled = False
            End If
        Else
            bCam = False
            tmrMovimiento.Enabled = False
            tmrCamPreview.Enabled = False
            MsgBox "No se ha encontrado ninguna cámara conectada en el sistema!", vbInformation
        End If
    End If
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subMuestraCAM"
subLog sERR
End Sub

Private Sub Form_Resize()
'imgFondo.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
'fra1.Move  (Me.ScaleWidth / 2) - (fra1.Width / 2), (Me.ScaleHeight / 2) - (fra1.Height / 2)
fra1.Move (Me.ScaleWidth / 2) - (fra1.Width / 2), 0
End Sub

Private Sub imgConfig_Click()
subValidaPermiso "frmConfig"
End Sub

Private Sub imgCerrar_Click()
If bHuellasU Then
    objUareU.StopCapture
    DoEvents
End If
subDesconectarZK
DoEvents
End
End Sub

Private Sub imgFoto_Click()
If bCam Then subTomarFoto
End Sub

Private Sub imgFoto1_Click()
If bCam Then
    If tmrCamPreview.Enabled = False Then tmrCamPreview.Enabled = True
    If bCam Then subTomarFoto
End If
End Sub

Private Sub imgFotos_Click(Index As Integer)
imgFoto1.Picture = imgFotos(Index).Picture
End Sub

Private Sub imgFuncionarios_Click()
subValidaPermiso "frmFuncionarios"
End Sub
Private Sub subValidaPermiso(sModulo As String)
On Local Error GoTo errH
Select Case sModulo
    Case "frmFuncionarios": If bPerFun Then frmFuncionarios.Show Else MsgBox "No tiene permiso para acceder a este módulo", vbInformation
    Case "frmConfig"
        If bPerPar Then
            subCerrar2d
'            subDesconectarZK
            Load frmConfig
            frmConfig.Show vbModal
            subCarga2D
'            subCargaZK
        Else
            MsgBox "No tiene permiso para acceder a este módulo", vbInformation
        End If
    Case "frmLic"
        If bPerReg Then
            On Local Error GoTo errH
            frmLic.Show vbModal
            Unload frmLic
        Else
            MsgBox "No tiene permiso para acceder a este módulo", vbInformation
        End If
    Case "frmReportes": If bPerRep Then frmReportes.Show Else MsgBox "No tiene permiso para acceder a este módulo", vbInformation
    Case "frmHerramientas": If bPerHer Then frmHerramientas.Show Else MsgBox "No tiene permiso para acceder a este módulo", vbInformation
    Case "frmBackup": If bPerBac Then frmBackup.Show vbModal Else MsgBox "No tiene permiso para acceder a este módulo", vbInformation
    Case "frmMinuta"
        If bPerMin Then
            If Not frmMinuta.Visible Then
                Load frmMinuta
                'frmMinuta.Tag = frmMinuta.ScaleHeight
                frmMinuta.Height = 0
                frmMinuta.bModo = 1
                frmMinuta.Move frmPrincipal.Left + ((frmPrincipal.ScaleWidth / 2) - (frmMinuta.ScaleWidth / 2))
                frmMinuta.tmrScroll.Enabled = True
            Else
                frmMinuta.SetFocus
            End If
        Else
            MsgBox "No tiene permiso para acceder a este módulo", vbInformation
        End If
End Select
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subValidaPermiso"
subLog sERR
End Sub
Public Sub subDesconectarZK()
Dim i As Integer
For i = 1 To objZK.Count - 1
    If oZKs(i).bConectado Then objZK(i).Disconnect
    DoEvents
    Unload objZK(i)
    DoEvents
Next i
End Sub
Private Sub subCerrar2d()
If Not obj2D1 Is Nothing Then obj2D1.cerrarPuerto
If Not obj2D2 Is Nothing Then obj2D2.cerrarPuerto
If Not obj2D3 Is Nothing Then obj2D3.cerrarPuerto
If Not obj2D4 Is Nothing Then obj2D4.cerrarPuerto
If Not obj2D5 Is Nothing Then obj2D5.cerrarPuerto
If Not obj2D6 Is Nothing Then obj2D6.cerrarPuerto
If Not obj2D7 Is Nothing Then obj2D7.cerrarPuerto
If Not obj2D8 Is Nothing Then obj2D8.cerrarPuerto
If Not obj2D9 Is Nothing Then obj2D9.cerrarPuerto
If Not obj2D10 Is Nothing Then obj2D10.cerrarPuerto
End Sub
Private Sub imgHuella_Click()
If bHuellasU Then
    bHuellaOrigen = 1
    Set frmEnrola.objImagen = imgHuella
    'frmPrincipal.objUareU.StopCapture
    frmEnrola.Show vbModal
    'frmPrincipal.objUareU.StartCapture
    'subUareU
End If
End Sub

Private Sub imgMinimizar_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub imgFrecuentes_Click()
On Local Error GoTo errH
sSql = "select id,nombre,apellidos,frecuente from tvisitantes_huella where abs(frecuente)=1 order by nombre,apellidos"
lstFrecuentes.Clear
lstFrecuentes.Tag = 0
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        lstFrecuentes.AddItem !nombre & " " & !apellidos
        lstFrecuentes.ItemData(lstFrecuentes.NewIndex) = !id
        lstFrecuentes.Selected(lstFrecuentes.NewIndex) = !frecuente
        .MoveNext
    Wend
    .Close
End With
lstFrecuentes.Tag = 1
picFrecuentes.ZOrder 0
picFrecuentes.Visible = True
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_imgFrecuentes_Click"
subLog sERR
End Sub

Private Sub imgWeb_Click()
ShellExecute Me.hWnd, "Open", "http://www.a1agroup.com", "", "", 1
End Sub

Private Sub lstEmerge_Click()
Dim sDest As String, sTm As String, iStm As Integer
On Local Error GoTo errH

    bBuscar = False
    txtDestino.Tag = lstEmerge.ItemData(lstEmerge.ListIndex)
    sTm = lstEmerge.Text
    iStm = InStr(1, sTm, "-")
    If iStm <> 0 Then sTm = Trim(Mid(sTm, 1, iStm - 1))
    
    txtDestino.Text = sTm
    
    lstEmerge.Visible = False
    'SendKeys "{TAB}"
    If lstEmerge.Tag = "txtEmpleado" Then
        'sSql = "select nombre from tDepartamentos where id=" & Val(txtDestino.Tag)
        sSql = "select d.nombre as depto,d.localizacion,d.ubicacion,c.nombre as cia,e.oficina,e.extension,e.usuario from ((templeados e left join tdepartamentos d on e.iddepartamento=d.id)"
        sSql = sSql & " left join tcompañias as c on e.idcompañia=c.id)where e.id=" & Val(txtDestino.Tag)
        Set objRst = objCon.Execute(sSql)
        If Not objRst.EOF Then
            txtDepartamento.Text = "" & objRst!depto
            txtCompañia.Text = "" & objRst!cia
            txtLocalizacion.Text = "" & objRst!localizacion
            txtUbicacion.Text = "" & objRst!ubicacion
            txtOficina.Text = "" & objRst!oficina
            txtExtension.Text = "" & objRst!extension
            txtChat.Text = "" & objRst!usuario
        End If
    ElseIf lstEmerge.Tag = "txtDepartamento" Then
        idDepartamento = Val(txtDestino.Tag)
        txtEmpleado.Text = vbNullString
        txtEmpleado.Tag = vbNullString
        txtCompañia.Text = vbNullString
        txtLocalizacion.Text = vbNullString
        txtUbicacion.Text = vbNullString
        txtOficina.Text = vbNullString
        txtExtension.Text = vbNullString
        txtChat.Text = vbNullString
    End If
    bBuscar = True

Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subLimpiar"
subLog sERR
End Sub

Private Sub lstFrecuentes_ItemCheck(Item As Integer)
On Local Error GoTo errH
If lstFrecuentes.Tag = 1 Then
    If lstFrecuentes.Selected(Item) = True Then
        sSql = "update tvisitantes_huella set frecuente=-1 where id=" & lstFrecuentes.ItemData(Item)
    Else
        sSql = "update tvisitantes_huella set frecuente=0 where id=" & lstFrecuentes.ItemData(Item)
    End If
    objCon.Execute sSql
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_lstFrecuentes_ItemCheck"
subLog sERR

End Sub

Private Sub lstPuertas_Click()
Dim objC As New ADODB.Recordset
Dim idC As Integer
Dim zCn As Integer, i As Integer
If lstPuertas.ListIndex <> -1 Then
    cmdPuertaE.Visible = False
    cmdPuertaS.Visible = False
    iPuertoEManual = 0
    iPuertoSManual = 0
    idC = lstPuertas.ItemData(lstPuertas.ListIndex)
    If objPhidget.IsAttached Then
        sSql = "select puerto_e,puerto_s from tcontrol where id=" & idC
        Set objC = objCon.Execute(sSql)
        If Not objC.EOF Then
            iPuertoEManual = objC!puerto_e
            iPuertoSManual = objC!puerto_s
            cmdPuertaE.Visible = True
            cmdPuertaS.Visible = True
        End If
    Else
        If Not objGAATools.fnArrVacioCls(oZKs) Then
            zCn = UBound(oZKs)
            If zCn > 0 Then
                sSql = "select puerto,modo from tcontrol_disp where abs(activo)=1 and tipo=3 and modo in (1,2) and idcontrol=" & idC
                Set objC = objCon.Execute(sSql)
                While Not objC.EOF
                    For i = 1 To zCn
                        If oZKs(i).sIP = objC!puerto Then
                            If oZKs(i).bConectado Then
                                If objC!modo = 1 Then
                                    cmdPuertaE.Visible = True
                                    iPuertoEManual = objZK(i).Index
                                ElseIf objC!modo = 2 Then
                                    cmdPuertaS.Visible = True
                                    iPuertoSManual = objZK(i).Index
                                End If
                                Exit For
                            End If
                        End If
                    Next i
                    objC.MoveNext
                Wend
            End If
        End If
    End If
End If
End Sub

Private Sub mnuConfig_Click()
'Load frmConfig
'frmConfig.Show vbModal
End Sub

Private Sub mnuConsultas_Click()
'frmConsultas.Show
End Sub

Private Sub mnuFuncionarios_Click()
'frmFuncionarios.Show
End Sub

Private Sub mnuLibera_Click(Index As Integer)
sSql = "select puerto_e,puerto_s from tcontrol where id=" & idControl
Set objRst = objCon.Execute(sSql)
If objRst.EOF Then
    Exit Sub
Else
    bPulsoE = Val("" & objRst!puerto_e)
    bPulsoS = Val("" & objRst!puerto_s)
    
    If Index = 0 Then
        bPuertoREG = bPulsoE
    ElseIf Index = 1 Then
        bPuertoREG = bPulsoS
    End If
    subRELEVO bPuertoREG
    idControl = 0
    iDispMODO = 0
    bPuertoREG = 0
    bPulsoE = 0
    bPulsoS = 0
End If
End Sub

Private Sub mnuLiberaC_Click()
idControl = 0
End Sub

Private Sub mnuLiberaZK_Click()
subRELEVO CLng(idControl)
End Sub

Private Sub mnuLicencia_Click()
'On Local Error GoTo errH
'Load frmLic
'Unload frmLic
'Exit Sub
'errH:
End Sub

Private Sub mnuReportes_Click()
'frmReportes.Show
End Sub

Private Sub mnuSalir_Click()
Unload Me
End Sub

Private Sub mnuZKsC_Click()
idControl = 0
End Sub

Private Sub obj2D1_Lectura(Datos2D As A1A2D.Datos2D)
subLectura2D Datos2D
End Sub

Private Sub obj2D10_Lectura(Datos2D As A1A2D.Datos2D)
subLectura2D Datos2D
End Sub

Private Sub obj2D2_Lectura(Datos2D As A1A2D.Datos2D)
subLectura2D Datos2D
End Sub

Private Sub obj2D3_Lectura(Datos2D As A1A2D.Datos2D)
subLectura2D Datos2D
End Sub

Private Sub obj2D4_Lectura(Datos2D As A1A2D.Datos2D)
subLectura2D Datos2D
End Sub

Private Sub obj2D5_Lectura(Datos2D As A1A2D.Datos2D)
subLectura2D Datos2D
End Sub

Private Sub obj2D6_Lectura(Datos2D As A1A2D.Datos2D)
subLectura2D Datos2D
End Sub

Private Sub obj2D7_Lectura(Datos2D As A1A2D.Datos2D)
subLectura2D Datos2D
End Sub

Private Sub obj2D8_Lectura(Datos2D As A1A2D.Datos2D)
subLectura2D Datos2D
End Sub

Private Sub obj2D9_Lectura(Datos2D As A1A2D.Datos2D)
subLectura2D Datos2D
End Sub

Private Sub objGrid_Click()
Dim sD As String
sD = objGrid.Columns("Documento").Text
If sD <> vbNullString Then
    picGrid.Visible = False
    txtDoc1.Text = sD
    txtDoc1_Validate False
End If
End Sub

Private Sub objGridAnotaciones_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Local Error Resume Next
idAnotacion = Val(objGridAnotaciones.Columns("id").Value)
End Sub

Private Sub objUareU_OnComplete(ByVal ReaderSerNum As String, ByVal pSample As Object)
Dim RespV As DPFPVerificationResult
Dim Resp As DPFPCaptureFeedbackEnum
'
Dim bProceso As Boolean
'
If modoBD = bdSQL Then
    sSql = "select d.id,d.modo,d.persona,d.activo,c.id as idc,c.puerto_e,c.puerto_s,c.activa"
    sSql = sSql & ",d.enrola_fun,d.enrola_vis,d.login"
    sSql = sSql & " from tcontrol_disp as d join tcontrol c on d.idcontrol=c.id"
    sSql = sSql & " where d.puerto='" & ReaderSerNum & "' and c.terminal='" & sTerminal & "'"
ElseIf modoBD = bdACCESS Then
    sSql = "SELECT tcontrol_disp.id, tcontrol_disp.modo, tcontrol_disp.persona, tcontrol_disp.activo, tcontrol.id AS idc, tcontrol.puerto_e, tcontrol.puerto_s, tcontrol.activa, "
    sSql = sSql & "tcontrol_disp.enrola_fun, tcontrol_disp.enrola_vis, tcontrol_disp.login "
    sSql = sSql & "FROM tcontrol_disp INNER JOIN tcontrol ON tcontrol_disp.idcontrol = tcontrol.id "
    sSql = sSql & "WHERE (((tcontrol_disp.puerto)='" & ReaderSerNum & "') AND ((tcontrol.terminal)='" & sTerminal & "'));"
End If
Set objRst = objCon.Execute(sSql)
If objRst.EOF Then
    fnHablar "Este huellero no está configurado."
    Exit Sub
Else
    If objRst!activo = False Then
        fnHablar "Este huellero está desactivado."
        Exit Sub
    Else
        If Not IsNull(objRst!idC) Then
            If Not objRst!activa Then
                bREG = False
                fnHablar "Portería desactivada."
                bPulsoE = 0
                bPulsoS = 0
                Exit Sub
            Else
                bREG = True
                idPuerta = objRst!idC
                idDisp = objRst!id
                iDispMODO = Val("" & objRst!modo)
                iDispPER = Val("" & objRst!persona)
                
                bPulsoE = Val("" & objRst!puerto_e)
                bPulsoS = Val("" & objRst!puerto_s)
                bDispEnrolaFun = objRst!enrola_fun
                bDispEnrolaVis = objRst!enrola_vis
                bDispLogin = objRst!login
                
            End If
        ElseIf Not objRst!activa Then
            bREG = False
            fnHablar "Portería desactivada."
            bPulsoE = 0
            bPulsoS = 0
            Exit Sub
        End If
    End If
End If
If Screen.ActiveForm.name = "frmEnrola" Then
    Resp = objCreaFea.CreateFeatureSet(pSample, DataPurposeEnrollment)
Else
    Resp = objCreaFea.CreateFeatureSet(pSample, DataPurposeVerification)
End If
If Resp = CaptureFeedbackGood Then
    sndPlaySound App.Path & "\huella.wav", 1
    Set objPlantilla = New DPFPTemplate
    '''Valida Funcionarios
    If Screen.ActiveForm.name = "frmEnrola" Then
        If bHuellaOrigen = 2 Then
            If Not bDispEnrolaFun Then
                fnHablar "Este huellero no está configurado para enrolar funcionarios."
                Exit Sub
            End If
        End If
        If bHuellaOrigen = 1 Then
            If Not bDispEnrolaVis Then
                fnHablar "Este huellero no está configurado para enrolar visitantes."
                Exit Sub
            End If
        End If
        
        frmEnrola.imgHuella.Picture = objConv.ConvertToPicture(pSample)

        bProceso = True
        If frmEnrola.bHuellaMon = False Then
            frmEnrola.bHuellaMon = True
            frmEnrola.sSerial = ReaderSerNum
        Else
            If ReaderSerNum <> frmEnrola.sSerial Then Exit Sub
        End If
        If bHuellaOrigen = 1 And bEnrolaVis = False Then bProceso = False
        'sndPlaySound App.Path & "\huella1.wav", 1
        If Not bProceso Then
            frmEnrola.lblPaso(3).ForeColor = Val(frmEnrola.Tag)
            frmEnrola.lblPaso(2).ForeColor = Val(frmEnrola.Tag)
            frmEnrola.lblPaso(1).ForeColor = Val(frmEnrola.Tag)
            frmEnrola.lblPaso(0).ForeColor = Val(frmEnrola.Tag)
            frmEnrola.objImagen.Picture = frmEnrola.imgHuella.Picture
            frmEnrola.imgHuella.Picture = LoadPicture(App.Path & "\huella_ok.bmp")
            DoEvents
            If bHuella Then
                bModificaHuella = True
            Else
                bHuella = True
                bModificaHuella = True
            End If
            frmEnrola.tmrCerrar.Enabled = True
            Exit Sub
        Else
            objCreaPlantilla.AddFeatures objCreaFea.FeatureSet
            'Debug.Print objCreaPlantilla.TemplateStatus
            Select Case objCreaPlantilla.TemplateStatus
                Case TemplateStatusUnknown, TemplateStatusCreationFailed
                    frmEnrola.lblPaso(3).ForeColor = Val(frmEnrola.Tag)
                    frmEnrola.lblPaso(2).ForeColor = Val(frmEnrola.Tag)
                    frmEnrola.lblPaso(1).ForeColor = Val(frmEnrola.Tag)
                    frmEnrola.lblPaso(0).ForeColor = Val(frmEnrola.Tag)
                    
                    frmEnrola.circ(3).BorderColor = Val(frmEnrola.Tag)
                    frmEnrola.circ(2).BorderColor = Val(frmEnrola.Tag)
                    frmEnrola.circ(1).BorderColor = Val(frmEnrola.Tag)
                    frmEnrola.circ(0).BorderColor = Val(frmEnrola.Tag)
                    frmEnrola.imgHuella.Picture = LoadPicture(App.Path & "\huella_error.bmp")
                    'Debug.Print "Faltan " & objCreaPlantilla.FeaturesNeeded
                Case TemplateStatusInsufficient
                    frmEnrola.lblPaso(objCreaPlantilla.FeaturesNeeded).ForeColor = vbWhite
                    frmEnrola.circ(objCreaPlantilla.FeaturesNeeded).BorderColor = vbWhite
                    Exit Sub
                Case TemplateStatusTemplateReady
                    frmEnrola.lblPaso(objCreaPlantilla.FeaturesNeeded).ForeColor = vbWhite
                    frmEnrola.circ(objCreaPlantilla.FeaturesNeeded).BorderColor = vbWhite
                    Set objPlantilla = objCreaPlantilla.Template
                    bHuellaMinuciasCAP = objPlantilla.Serialize
                    frmEnrola.objImagen.Picture = frmEnrola.imgHuella.Picture
                    frmEnrola.imgHuella.Picture = LoadPicture(App.Path & "\huella_ok.bmp")
                    DoEvents
                    If bHuella Then
                        bModificaHuella = True
                    Else
                        bHuella = True
                        bModificaHuella = True
                    End If
                    frmEnrola.tmrCerrar.Enabled = True
                    '''
'                    WriteBinary
                    '''
                    Exit Sub
            End Select
        End If
    End If
    If iDispPER = 1 Or iDispPER = 3 Then
        With objRst
            If .State = adStateOpen Then .Close
            sSql = "select id,documento,enrola from templeados"
            .Open sSql, objCon, adOpenForwardOnly
            While Not .EOF
                If Not IsNull(!enrola) Then
                    bHuellaMinucias = !enrola
                    ''
'                    ReadBinary
                    ''
                    objPlantilla.Deserialize bHuellaMinucias
                    Set RespV = objVerifica.Verify(objCreaFea.FeatureSet, objPlantilla)
                    If RespV.Verified = True Then
                        If Screen.ActiveForm.name = "frmLogin" Then
                            If Not bDispLogin Then
                                fnHablar "Este huellero está inactivo para Login."
                                Exit Sub
                            Else
                                AccesoTipo = accLOGIN
                                tmrPosHuella.Tag = "" & objRst!documento
                                tmrPosHuella.Enabled = True
                                Exit Sub
                            End If
                        Else
                            AccesoTipo = accHUELLA
                            tmrPosHuella.Tag = "" & objRst!documento
                            tmrPosHuella.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
                .MoveNext
            Wend
            .Close
        End With
    End If
    If Screen.ActiveForm.name = "frmPrincipal" Then
        If iDispPER = 2 Or iDispPER = 3 Then
            With objRst
                If .State = adStateOpen Then .Close
                If bIngresoManualVis = False Then
                    sSql = "select * from tautoriza_acceso_vis"
                Else
                    sSql = "select documento,enrola,0 as idempleado from tvisitantes_huella"
                End If
                .Open sSql, objCon, adOpenForwardOnly
                While Not .EOF
                    If Not IsNull(!enrola) Then
                        bHuellaMinucias = !enrola
                        objPlantilla.Deserialize bHuellaMinucias
                        Set RespV = objVerifica.Verify(objCreaFea.FeatureSet, objPlantilla)
                        If RespV.Verified = True Then
                            AccesoTipo = accHUELLA
                            tmrPosHuella.Tag = "" & objRst!documento
                            idEmpleado = Val("" & objRst!idEmpleado)
                            If bIngresoManualVis = False Then bAutorizado = True
                            tmrPosHuella.Enabled = True
                            Exit Sub
                        End If
                    End If
                    .MoveNext
                Wend
                .Close
            End With
            With objRst
                If .State = adStateOpen Then .Close
                'sSql = "select id,documento,enrola from tvisitantes_huella"
                '''!!!sSql = "select * from v_entrando order by cuenta desc"
                sSql = "select * from v_saliendo"
                .Open sSql, objCon, adOpenForwardOnly
                While Not .EOF
                    If Not IsNull(!enrola) Then
                        bHuellaMinucias = !enrola
                        objPlantilla.Deserialize bHuellaMinucias
                        Set RespV = objVerifica.Verify(objCreaFea.FeatureSet, objPlantilla)
                        If RespV.Verified = True Then
                            AccesoTipo = accHUELLA
                            tmrPosHuella.Tag = "" & objRst!documento
                            tmrPosHuella.Enabled = True
                            Exit Sub
                        End If
                    End If
                    .MoveNext
                Wend
                .Close
            End With
        End If
        fnHablar "Huella no registrada."
    End If
End If
End Sub

Private Sub objVideo_Click()
If bCam Then subTomarFoto
End Sub

Private Sub objZK_OnAttTransaction(Index As Integer, ByVal EnrollNumber As Long, ByVal IsInValid As Long, ByVal AttState As Long, ByVal VerifyMethod As Long, ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long)
'Debug.Print "Si"
End Sub

Private Sub objZK_OnEnrollFinger(Index As Integer, ByVal EnrollNumber As Long, ByVal FingerIndex As Long, ByVal ActionResult As Long, ByVal TemplateLength As Long)
Dim sMot As String
If ActionResult <> 6 Then
    frmEnrolaZK.lblPaso(3 - bCuentaZK).ForeColor = vbWhite
    frmEnrolaZK.circ(3 - bCuentaZK).BorderColor = vbWhite
    Select Case ActionResult
        Case 0
            frmEnrolaZK.imgHuella.Picture = LoadPicture(App.Path & "\huella_ok.bmp")
            frmEnrolaZK.tmrCerrar.Enabled = True
            'objZK(Index).RefreshData 0
            bEnrolandoZK = False
            subGuardaTMP Index
        Case Else
            If ActionResult = 3 Then
                sMot = "Falló almacenamiento de datos"
            ElseIf ActionResult = 4 Then
                sMot = "Falló enrolamiento"
            ElseIf ActionResult = 5 Then
                sMot = "Huella repetida"
            ElseIf ActionResult = 6 Then
                sMot = "Operación cancelada"
            End If
            bEnrolandoZK = False
            objZK(Index).StartIdentify
            frmEnrolaZK.imgHuella.Picture = LoadPicture(App.Path & "\huella_error.bmp")
            frmEnrolaZK.tmrCerrar.Enabled = True
            MsgBox "Ha fallado el enrolamiento!. Por favor vuelva a intentarlo." & _
            vbCr & "Motivo: " & ActionResult & ". " & sMot, vbCritical
    End Select
Else
    bEnrolandoZK = False
    objZK(Index).StartIdentify
    frmEnrolaZK.bCancela = False
    Unload frmEnrolaZK
End If
End Sub
Private Sub subGuardaTMP(idx As Integer)
Dim lFlag As Long
Dim hLen As Long
Dim sDedo As String
If idHuellaZK = 0 Then
    sDedo = "zktmp"
Else
    sDedo = "zktmp" & idHuellaZK
End If
If objZK(idx).ReadAllUserID(0) Then
    sHuellaZK = vbNullString
    If objZK(idx).GetUserTmpExStr(1, frmFuncionarios.idEmpleado, idHuellaZK, lFlag, sHuellaZK, hLen) Then
        objCon.Execute "update templeados set " & sDedo & "='" & sHuellaZK & "' where id=" & frmFuncionarios.idEmpleado
'        objZK(idx).CancelOperation
'        objZK(idx).StartIdentify
        subGuardaZK idx, 2
    End If
End If

End Sub
Public Sub subGuardaZK(idxZK As Integer, idOper As Byte)
'idOper: 1 enrola,2 Replica,3 Guarda
Dim Zs As Integer
If oZKs(idxZK).bConectado Then
    bCuentaZK = 0
    If idOper = 1 Then
        objZK(idxZK).CancelOperation
        DoEvents
        If objZK(idxZK).SetUserInfo(1, zkID, zkUSR, zkUSR, 0, True) Then
            If Not objZK(idxZK).DelUserTmp(objZK(idxZK).MachineNumber, zkID, idHuellaZK) Then
                objZK(idxZK).GetLastError lZkErr
            End If
            If objZK(idxZK).StartEnroll(zkID, idHuellaZK) Then
                bEnrolandoZK = True
            End If
        Else
            MsgBox "No se pudo registrar el ID del funcionario!", vbInformation
            bEnrolandoZK = False
            objZK(idxZK).StartIdentify
        End If
    ElseIf idOper = 2 Then
        objZK(idxZK).StartIdentify
        If modoBD = bdSQL Then
            sSql = "select tdd.puerto from tcontrol_disp td join tzk_asoc ta on td.id=ta.idorigen "
            sSql = sSql & "join tcontrol_disp tdd on tdd.id=ta.iddestino where td.puerto='" & oZKs(idxZK).sIP & "'"
        ElseIf modoBD = bdACCESS Then
            sSql = "select puerto from v_zkasoc where p1='" & oZKs(idxZK).sIP & "'"
        End If
        With objRstA
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenForwardOnly
            While Not .EOF
                'If (Not oZKs) = -1 Then
                If Not objGAATools.fnArrVacioCls(oZKs) Then
                    For Zs = 1 To UBound(oZKs)
                        If oZKs(Zs).sIP = "" & !puerto Then
                            If oZKs(Zs).bConectado Then
                                subGuardaZK oZKs(Zs).iIndex, 3
                            End If
                        End If
                    Next Zs
                End If
                .MoveNext
            Wend
        End With
    ElseIf idOper = 3 Then
        objZK(idxZK).CancelOperation
        If objZK(idxZK).SetUserInfo(1, zkID, zkUSR, zkUSR, 0, True) Then
            If Not objZK(idxZK).DelUserTmp(objZK(idxZK).MachineNumber, zkID, idHuellaZK) Then
                objZK(idxZK).GetLastError lZkErr
            End If
            If Not objZK(idxZK).SetUserTmpExStr(1, zkID, idHuellaZK, 0, sHuellaZK) Then
                objZK(idxZK).GetLastError lZkErr
            End If
        Else
            MsgBox "No se pudo registrar el ID del funcionario!", vbInformation
        End If
        bEnrolandoZK = False
        objZK(idxZK).StartIdentify
        DoEvents
    End If
    'frmPrincipal.bEnrolandoZK = True
End If
End Sub
Private Sub objZK_OnFinger(Index As Integer)
'Debug.Print objZK(Index).MachineNumber
Dim bProceso As Boolean
If bEnrolandoZK Then
    If modoBD = bdSQL Then
        sSql = "select d.id,d.modo,d.persona,d.activo,c.id as idc,c.puerto_e,c.puerto_s,c.activa"
        sSql = sSql & ",d.enrola_fun,d.enrola_vis,d.login"
        sSql = sSql & " from tcontrol_disp as d join tcontrol c on d.idcontrol=c.id"
        sSql = sSql & " where d.puerto='" & objZK(Index).Tag & "' and c.terminal='" & sTerminal & "'"
    ElseIf modoBD = bdACCESS Then
        sSql = "SELECT tcontrol_disp.id, tcontrol_disp.modo, tcontrol_disp.persona, tcontrol_disp.activo, tcontrol.id AS idc, tcontrol.puerto_e, tcontrol.puerto_s, tcontrol.activa, "
        sSql = sSql & "tcontrol_disp.enrola_fun, tcontrol_disp.enrola_vis, tcontrol_disp.login "
        sSql = sSql & "FROM tcontrol_disp INNER JOIN tcontrol ON tcontrol_disp.idcontrol = tcontrol.id "
        sSql = sSql & "WHERE (((tcontrol_disp.puerto)='" & objZK(Index).Tag & "') AND ((tcontrol.terminal)='" & sTerminal & "'));"
    End If
    Set objRst = objCon.Execute(sSql)
    If objRst.EOF Then
        fnHablar "Este dispositivo no está configurado."
        Exit Sub
    Else
        If objRst!activo = False Then
            fnHablar "Este dispositivo está desactivado."
            Exit Sub
        Else
            If Not IsNull(objRst!idC) Then
                If Not objRst!activa Then
                    bREG = False
                    fnHablar "Portería desactivada."
                    bPulsoE = 0
                    bPulsoS = 0
                    Exit Sub
                Else
                    bREG = True
                    idDisp = objRst!id
                    iDispMODO = Val("" & objRst!modo)
                    iDispPER = Val("" & objRst!persona)
                    
                    bPulsoE = Val("" & objRst!puerto_e)
                    bPulsoS = Val("" & objRst!puerto_s)
                    bDispEnrolaFun = objRst!enrola_fun
                    bDispEnrolaVis = objRst!enrola_vis
                    bDispLogin = objRst!login
                    
                End If
            ElseIf Not objRst!activa Then
                bREG = False
                fnHablar "Portería desactivada."
                bPulsoE = 0
                bPulsoS = 0
                Exit Sub
            End If
        End If
    End If

    If Screen.ActiveForm.name = "frmEnrolaZK" Then
        If bHuellaOrigen = 2 Then
            If Not bDispEnrolaFun Then
                fnHablar "Este dispositivo no está configurado para enrolar funcionarios."
                objZK(Index).CancelOperation
                Exit Sub
            End If
        End If
        If bHuellaOrigen = 1 Then
            If Not bDispEnrolaVis Then
                fnHablar "Este huellero no está configurado para enrolar visitantes."
                Exit Sub
            End If
        End If
        
        'frmEnrola.imgHuella.Picture = objConv.ConvertToPicture(pSample)

        bProceso = True
        If frmEnrola.bHuellaMon = False Then
            frmEnrola.bHuellaMon = True
            frmEnrola.sSerial = objZK(Index).Tag
        Else
            If objZK(Index).Tag <> frmEnrola.sSerial Then Exit Sub
        End If
        If bHuellaOrigen = 1 And bEnrolaVis = False Then bProceso = False
        'sndPlaySound App.Path & "\huella1.wav", 1
        If Not bProceso Then
            frmEnrola.lblPaso(3).ForeColor = Val(frmEnrola.Tag)
            frmEnrola.lblPaso(2).ForeColor = Val(frmEnrola.Tag)
            frmEnrola.lblPaso(1).ForeColor = Val(frmEnrola.Tag)
            frmEnrola.lblPaso(0).ForeColor = Val(frmEnrola.Tag)
            frmEnrola.objImagen.Picture = frmEnrola.imgHuella.Picture
            frmEnrola.imgHuella.Picture = LoadPicture(App.Path & "\huella_ok.bmp")
            DoEvents
            If bHuella Then
                bModificaHuella = True
            Else
                bHuella = True
                bModificaHuella = True
            End If
            frmEnrola.tmrCerrar.Enabled = True
            Exit Sub
        Else
            bCuentaZK = bCuentaZK + 1
            If bCuentaZK < 3 Then
                frmEnrolaZK.lblPaso(3 - bCuentaZK).ForeColor = vbWhite
                frmEnrolaZK.circ(3 - bCuentaZK).BorderColor = vbWhite
            Else
                frmEnrolaZK.lblPaso(3 - bCuentaZK).ForeColor = vbWhite
                frmEnrolaZK.circ(3 - bCuentaZK).BorderColor = vbWhite
'                frmEnrolaZK.imgHuella.Picture = LoadPicture(App.Path & "\huella_ok.bmp")
'                frmEnrolaZK.tmrCerrar.Enabled = True
            End If
        End If
    End If
End If
End Sub

Private Sub objZK_OnHIDNum(Index As Integer, ByVal CardNumber As Long)
If bZk_ev Then
    subLecturaZK Index, 0, CardNumber
End If
End Sub
Private Sub subLecturaZK(Index As Integer, idEmp As Long, CardNumber As Long)
idxZK = Index
If modoBD = bdSQL Then
    sSql = "select d.id,d.modo,d.persona,d.activo,c.id as idc,c.puerto_e,c.puerto_s,c.activa"
    sSql = sSql & ",d.enrola_fun,d.enrola_vis,d.login"
    sSql = sSql & " from tcontrol_disp as d join tcontrol c on d.idcontrol=c.id"
    sSql = sSql & " where d.puerto='" & objZK(Index).Tag & "' and c.terminal='" & sTerminal & "'"
ElseIf modoBD = bdACCESS Then
    sSql = "SELECT tcontrol_disp.id, tcontrol_disp.modo, tcontrol_disp.persona, tcontrol_disp.activo, tcontrol.id AS idc, tcontrol.puerto_e, tcontrol.puerto_s, tcontrol.activa, "
    sSql = sSql & "tcontrol_disp.enrola_fun, tcontrol_disp.enrola_vis, tcontrol_disp.login "
    sSql = sSql & "FROM tcontrol_disp INNER JOIN tcontrol ON tcontrol_disp.idcontrol = tcontrol.id "
    sSql = sSql & "WHERE (((tcontrol_disp.puerto)='" & objZK(Index).Tag & "') AND ((tcontrol.terminal)='" & sTerminal & "'));"
End If
Set objRst = objCon.Execute(sSql)
If objRst.EOF Then
    fnHablar "Este dispositivo no está configurado."
    Exit Sub
Else
    If objRst!activo = False Then
        fnHablar "Este dispositivo está desactivado."
        Exit Sub
    Else
        If Not IsNull(objRst!idC) Then
            If Not objRst!activa Then
                bREG = False
                fnHablar "Portería desactivada."
                bPulsoE = 0
                bPulsoS = 0
                Exit Sub
            Else
                bREG = True
                idPuerta = objRst!idC
                idDisp = objRst!id
                iDispMODO = Val("" & objRst!modo)
                iDispPER = Val("" & objRst!persona)
                
                bPulsoE = Val("" & objRst!puerto_e)
                bPulsoS = Val("" & objRst!puerto_s)
                bDispEnrolaFun = objRst!enrola_fun
                bDispEnrolaVis = objRst!enrola_vis
                bDispLogin = objRst!login
                
            End If
        ElseIf Not objRst!activa Then
            bREG = False
            fnHablar "Portería desactivada."
            bPulsoE = 0
            bPulsoS = 0
            Exit Sub
        End If
    End If
End If
If iDispPER = 1 Or iDispPER = 3 Then 'Solo empleados
    If CardNumber > 0 Then
        sSql = "select documento from templeados where tarjeta_num='" & CardNumber & "'"
        Set objRst = objCon.Execute(sSql)
        If Not objRst.EOF Then
            AccesoTipo = accTARJETA
            subAccesoTipo "" & objRst!documento
        End If
    ElseIf idEmp > 0 Then
        sSql = "select documento from templeados where id=" & idEmp
        Set objRst = objCon.Execute(sSql)
        If Not objRst.EOF Then
            AccesoTipo = accHUELLA_ZK
            subAccesoTipo "" & objRst!documento
        End If
    End If
End If

End Sub
Private Sub objZK_OnVerify(Index As Integer, ByVal UserID As Long)
Dim i As Long
If bZk_ev Then
    If UserID > -1 Then
        i = Timer
        While Not objFUNC Is Nothing
            DoEvents
            If Timer - i > 1 Then
                Set objFUNC = Nothing
            End If
        Wend
        subLecturaZK Index, UserID, 0
    End If
End If
End Sub

Private Sub tmrCamPreview_Timer()
On Error Resume Next
ImgFoto.Picture = objVideo.GrabFrame
ImgFoto.Rotate 90
'imgFoto.DrawText 198, 5, "45", vbRed
If bCapturando Then
    imgFoto1.Picture = ImgFoto.Picture
End If

End Sub

Private Sub tmrEspera_Timer()
imgFlecha.Visible = Not imgFlecha.Visible
End Sub

Private Sub tmrHora_Timer()
lblFecha.Caption = FormatDateTime(Date + Time, vbGeneralDate)
lblDia.Caption = Right("00" & Day(Date), 2)

If Format(Time, "HH:MM:SS") = sHoraAutoSalida Then
    'MsgBox "YES"
    If modoBD = bdACCESS Then
        sSql = "update tacceso set sale='" & fnFecha(Now, True) & "',idhuellero_sale=-1,idlogin_s=" & idLogin & " where idtipoper=2 and sale is null"
    ElseIf modoBD = bdSQL Then
        sSql = "update tacceso set sale=getdate(),idhuellero_sale=-1,idlogin_s=" & idLogin & " where idtipoper=2 and sale is null"
    End If
    objCon.Execute sSql
    DoEvents
End If
End Sub

Private Sub tmrMovimiento_Timer()
Dim curID As Integer
If objVideo.DetectMotion > 5 Then subTomarFoto
End Sub
Private Sub subTomarFoto()
Static curID As Byte
Static bFIN As Boolean
If bFoto Then
    bModificaFoto = True
Else
    bFoto = True
    bModificaFoto = True
End If
imgFoto1.Picture = ImgFoto.Picture
bCapturando = Not bCapturando
'txtNombre.SetFocus
If chkSensor.Value = vbChecked Then
'    curID = imgFotos.Count - 1
'    If iCicloFoto > 0 Then
'        If bFIN = False Then
'            Load imgFotos(curID + 1)
'            curID = imgFotos.Count - 1
'            If imgFotos.Count Mod 2 = 0 Then
'                imgFotos(curID).Move imgFotos(curID - 1).Left + imgFotos(curID).Width, imgFotos(curID - 1).Top
'            Else
'                imgFotos(curID).Move imgFotos(curID - 2).Left, imgFotos(curID - 2).Top + imgFotos(curID - 2).Height
'            End If
'        Else
'            curID = iCicloFoto
'        End If
'    Else
'        curID = iCicloFoto
'    End If
    imgFotos(curID).Visible = True
    imgFotos(curID).Picture = ImgFoto.Picture
    curID = curID + 1
    If curID = 3 Then
        bFIN = True
        curID = 0
    End If
End If
End Sub

Private Sub txtApellidos_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tmrPosHuella_Timer()
tmrPosHuella.Enabled = False
subAccesoTipo tmrPosHuella.Tag
End Sub

Private Sub txtApellidos_Validate(Cancel As Boolean)
txtApellidos.Text = fnMayúscula(txtApellidos.Text)
End Sub

Private Sub txtAutoriza_Change()
'subLista txtAutoriza
End Sub

Private Sub txtAutoriza_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAutoriza_LostFocus()
lstEmerge.Visible = False
End Sub

Private Sub txtDoc1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDoc2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtDepartamento_LostFocus()
lstEmerge.Visible = False
End Sub

Private Sub txtDepartamento_txtCambio()
subLista txtDepartamento
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

Private Sub txtOrganizacion_Change()

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
            Case "txtOrganizacion"
                sSql = "select * from torganizaciones where nombre like '%" & txtOrganizacion.Text & "%' order by nombre"
                'lstEmerge.Tag = "tempresa"
            Case "txtAutoriza"
                'If modoBD = bdACCESS Then
                '    sSql = "select id,nombre + ' ' + apellidos as nombre from templeados  where nombre + ' ' + apellidos like '%" & txtCompañia.Text & "%' order by nombre"
                'ElseIf modoBD = bdSQL Then
                '    sSql = "select id,isnull(nombre,'') + ' ' + isnull(apellidos,'') as nombre from templeados where isnull(nombre,'') + ' ' + isnull(apellidos,'') like '%" & txtCompañia.Text & "%' order by nombre"
                'End If
            Case "txtEmpleado"
                If modoBD = bdACCESS Then
                    If idDepartamento = 0 Then
                        sSql = "select id,nombre + ' ' + apellidos + '-' + tel as nombre from templeados where nombre + ' ' + apellidos like '%" & txtEmpleado.Text & "%' and id>1 and abs(activo)=1 order by nombre"
                    Else
                        sSql = "select id,nombre + ' ' + apellidos + '-' + tel as nombre from templeados where nombre + ' ' + apellidos like '%" & txtCompañia.Text & "%' and idDepartamento=" & idDepartamento & " order by nombre"
                    End If
                ElseIf modoBD = bdSQL Then
                    If idDepartamento = 0 Then
                        sSql = "select id,isnull(nombre,'') + ' ' + isnull(apellidos,'') + '-' + isnull(tel,'') as nombre from templeados where isnull(nombre,'') + ' ' + isnull(apellidos,'') like '%" & txtEmpleado.Text & "%' and id>1 and abs(activo)=1 order by nombre"
                    Else
                        sSql = "select id,isnull(nombre,'') + ' ' + isnull(apellidos,'') + '-' + isnull(tel,'') as nombre from templeados where isnull(nombre,'') + ' ' + isnull(apellidos,'') like '%" & txtCompañia.Text & "%' and idDepartamento=" & idDepartamento & " order by nombre"
                    End If
                End If

                'lstEmerge.Tag = "tpersona"
            Case "txtDepartamento"
                'sSql = "select * from tDepartamentos where nombre like '%" & txtDepartamento.Text & "%' order by nombre"
                sSql = "select d.id,c.nombre + ' - ' + d.nombre nombre from tcompañias c join tdepartamentos d on c.id=d.idcompañia where d.nombre like '%" & txtDepartamento.Text & "%' order by nombre"
        End Select
        Set txtDestino = txt
        With objRst
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenForwardOnly
            While Not .EOF
                lstEmerge.AddItem !nombre
                lstEmerge.ItemData(lstEmerge.NewIndex) = !id
                .MoveNext
            Wend
'            If lstEmerge.ListCount = 1 Then
'                'lstEmerge.ListIndex = 0
'
'            Else
            If lstEmerge.ListCount >= 1 Then
                lstEmerge.Move txt.Left, txt.Top + txt.Height, txt.Width
                lstEmerge.Visible = True
            End If
            .Close
        End With
    Else
        lstEmerge.Visible = False
        txtEmpleado.Text = vbNullString
        txtEmpleado.Tag = vbNullString
        txtCompañia.Text = vbNullString
        txtDepartamento.Text = vbNullString
        txtLocalizacion.Text = vbNullString
        txtUbicacion.Text = vbNullString
        txtOficina.Text = vbNullString
        txtExtension.Text = vbNullString
        txtChat.Text = vbNullString
    End If
End If
End Sub

Private Sub txtDoc1_Validate(Cancel As Boolean)
Dim sDOC As String
sDOC = Trim(txtDoc1.Text)
If sDOC <> vbNullString Then
    sDOC = Replace(sDOC, ".", "")
    sDOC = Replace(sDOC, ",", "")
    sDOC = Replace(sDOC, "-", "")
    AccesoTipo = accMANUAL
    idDisp = 0
    sDocTmp = sDOC
    subAccesoTipo sDOC
End If
End Sub
Private Sub subAccesoTipo(sDOC As String)
On Local Error GoTo errH:
Dim objRs_ As New ADODB.Recordset
tmrPosHuella.Tag = vbNullString
If AccesoTipo = accLOGIN Then
    sSql = "select id,usuario,contraseña,foto from templeados where documento='" & sDOC & "'"
    Set objRs_ = objCon.Execute(sSql)
    If "" & objRs_!usuario <> vbNullString Then
        idLogin = objRs_!id
        fnLeeFoto objRs_!foto, frmLogin.ImgFoto
        frmLogin.ImgFoto.Visible = True
        fnHablar "Acceso autorizado para el usuario " & objRs_!usuario & "."
        subEsperar 1
        frmLogin.Hide
        Unload frmLogin
        Exit Sub
    Else
        fnHablar "Funcionario sin login asociado!."
        Exit Sub
    End If
Else
    sDocTmp = sDOC
    TipoPER = tpNONE
    sSql = "select id from templeados where documento='" & sDOC & "'"
    Set objRs_ = objCon.Execute(sSql)
    If objRs_.EOF Then
        subAnotaciones 'Solo para visitantes
        If AccesoTipo = acc2D Then
            Dim sFec As String
            sFec = DatosDD.ccDiaNace & "/" & DatosDD.ccMesNace & "/" & DatosDD.ccAñosNace
            If IsDate(sFec) Then
                txtFechaNace.Text = fnFecha(CDate(sFec), False)
            End If
        End If
        sSql = "select id from tvisitantes_huella where documento='" & sDOC & "'"
        Set objRs_ = objCon.Execute(sSql)
        If objRs_.EOF Then
            TipoPER = tpNONE
            Set objVISIManual = New clsDatosVISI
            objVISIManual.bEntraVISI = True
        Else
            TipoPER = tpVISI
            If TipoPER = iDispPER Or iDispPER = 0 Or iDispPER = 3 Then
                Set objVISI = New clsDatosVISI
                objVISI.idVISI = Val("" & objRs_!id)
                objVISI.iDispMODO = iDispMODO
                iDispMODO = 0
                objVISI.idDisp = idDisp
                idDisp = 0
                objVISI.AccesoTipo = AccesoTipo
                AccesoTipo = accNONE
                objVISI.bREG = bREG
                bREG = False
                sSql = "select * from tautoriza_acceso_vis where documento='" & sDOC & "'"
                Set objRst = objCon.Execute(sSql)
                If Not objRst.EOF Then
                    bAutorizado = True
                    idEmpleado = Val("" & objRst!idEmpleado)
                End If
            Else
                fnHablar "Por esta portería solo pueden ingresar Funcionarios."
                Exit Sub
            End If
        End If
    Else
        TipoPER = tpFUNC
        If TipoPER = iDispPER Or iDispPER = 0 Or iDispPER = 3 Then
            Set objFUNC = New clsDatosFUNC
            objFUNC.idFUNC = Val("" & objRs_!id)
            objFUNC.sDOC = txtDoc1.Text
            'txtDoc1.Text = vbNullString
            objFUNC.iDispMODO = iDispMODO
            iDispMODO = 0
            objFUNC.idDisp = idDisp
            idDisp = 0
            objFUNC.idPuerta = idPuerta
            idPuerta = 0
            objFUNC.AccesoTipo = AccesoTipo
            AccesoTipo = accNONE
            objFUNC.bREG = bREG
        Else
            fnHablar "Por esta portería solo pueden ingresar Visitantes."
            Exit Sub
        End If
    End If
    Set objRs_ = Nothing
    Select Case TipoPER
        Case tpNONE
            If AccesoTipo = accMANUAL Then
                cmbTipoID.SetFocus
            ElseIf AccesoTipo = acc2D Then
                If bDispEnrolaVis Then
                    txtDoc1.Text = DatosDD.ccNumero
                    If DatosDD.ccTipo = Cédula_2D_V1 Or DatosDD.ccTipo = Cédula_2D_V2 Or DatosDD.ccTipo = PASE_2011 Or DatosDD.ccTipo = TPROPIEDAD_2011 Then
                        cmbTipoID.mostrarItem 1
                    ElseIf DatosDD.ccTipo = TI_2D Then
                        cmbTipoID.mostrarItem 2
                    End If
                    If DatosDD.ccSexo = "F" Then cmbTratamiento.mostrarItem 2 Else cmbTratamiento.mostrarItem 1
                    txtNombre.Text = fnMayúscula(Trim(DatosDD.ccNombre1 & " " & DatosDD.ccNombre2))
                    txtApellidos.Text = fnMayúscula(Trim(DatosDD.ccApellido1 & " " & DatosDD.ccApellido2))
                    If txtNombre.Text <> vbNullString Then
                        txtOrganizacion.SetFocus
                    Else
                        cmbTipoID.SetFocus
                    End If
                    cmbSexo.mostrarItem fnBuscaSEXO(DatosDD.ccSexo)
                    cmbRH.mostrarItem fnBuscaRH(DatosDD.ccRH)
                    
                Else
                    fnHablar "Este lector no está configurado para registrar visitantes."
                    AccesoTipo = accNONE
                    bREG = False
                    bDispEnrolaFun = False
                    bDispEnrolaVis = False
                End If
                
            End If
        Case tpFUNC
            subDatosFUNC
        Case tpVISI
            bAutoStmp = bAutoSalida
            subDatosVISI
    End Select
    '
    'Select Case AccesoTIPO
    '    Case accMANUAL
    '
    '    Case acc2D
    '
    '    Case accHUELLA
    '
    'End Select
End If
Exit Sub
errH:
Set objFUNC = Nothing
subLog "Error " & Err.Number & ". " & Err.Description & " - SubAccesoTipo"
End Sub
Private Sub subAnotaciones()
Dim objRs_ As New ADODB.Recordset
sSql = "select id from tanotaciones where documento='" & sDocTmp & "'"
Set objRs_ = objCon.Execute(sSql)
If objRs_.EOF Then
    imgAlerta.Picture = imgAlerta2.Picture
    imgAlerta.Tag = "0"
Else
    imgAlerta.Picture = imgAlerta1.Picture
    imgAlerta.Tag = "1"
End If
Set objRs_ = Nothing

End Sub
Private Function fnBuscaSEXO(sSEXO As String) As Integer
Dim i As Integer
For i = 1 To 2
    If colSexo(i) = sSEXO Then
        fnBuscaSEXO = i
        Exit For
    End If
Next
End Function

Private Function fnBuscaRH(sRH As String) As Integer
Dim i As Integer
For i = 1 To 8
    If colRH(i) = sRH Then
        fnBuscaRH = i
        Exit For
    End If
Next
End Function

Sub subDatosFUNC()
On Local Error GoTo errH:
Dim i As Byte, cnt As Byte
Dim bHoy As Byte
Dim bPermiso As Boolean

Dim sHoraA1 As String
Dim sHoraA2 As String
Dim sHora1 As String
Dim sHora2 As String
Dim sMinu As Long
Dim bNoche As Boolean
Dim sMod As String
Dim sArr() As String, stNom As String
Dim bHLibre As Boolean
If objFUNC.AccesoTipo = accMANUAL Then txtDoc1.Text = vbNullString
sSql = "select id,nombre,apellidos,sexo,fechai,fechaf,idhorario,activo from templeados where id=" & Val(objFUNC.idFUNC)
Set objRst = objCon.Execute(sSql)
objFUNC.sNOM = "" & objRst!nombre
objFUNC.sAPE = "" & objRst!apellidos
objFUNC.sSEXO = "" & objRst!sexo
If objFUNC.iDispMODO = 1 Then
    sMod = "Entra"
ElseIf objFUNC.iDispMODO = 2 Then
    sMod = "Sale"
ElseIf objFUNC.iDispMODO = 3 Then
    sMod = "E/S"
End If
idLog = Val(idLog) + 1
    sArr = Split(objFUNC.sNOM, " ")
   If UBound(sArr) >= 0 Then
       stNom = sArr(0)
   End If
   sArr = Split(objFUNC.sAPE, " ")
   If UBound(sArr) >= 0 Then
       stNom = stNom & " " & sArr(0)
   End If

subMonitor idLog & ". " & sMod & " [" & objFUNC.idFUNC & " " & stNom & "]", False, False
If Not objRst!activo Then
    fnHablar "Funcionario inactivo!"
    subMonitor "{INACTIVO}", False, True
    Set objFUNC = Nothing
    Exit Sub
End If
If Not IsNull(objRst!fechaI) Then
    If objRst!fechaI > Date Then
        fnHablar "No está autorizado para ingresar!"
        subMonitor "{FECHAI}", False, True
        Set objFUNC = Nothing
        Exit Sub
    End If
End If
If Not IsNull(objRst!fechaf) Then
    If objRst!fechaf < Date Then
        fnHablar "No está autorizado para ingresar!"
        subMonitor "{FECHAF}", False, True
        Set objFUNC = Nothing
        Exit Sub
    End If
End If
If Val("" & objRst!idHorario) = 0 Then
    fnHablar "No tiene un horario asignado, acceso denegado!"
    subMonitor "{SIN_HORARIO}", False, True
    Set objFUNC = Nothing
    Exit Sub
End If
sSql = "select * from thorarios where id=" & objRst!idHorario
Set objRst = objCon.Execute(sSql)
sMinu = Val("" & objRst!minutos)
If IsNull(objRst!libre) Then
    bHLibre = False
Else
    bHLibre = objRst!libre
End If
If IsNull(objRst!entra2) Then
    If objRst!sale1 < objRst!entra1 Then
        'Nocturno Jornada Unica
        bNoche = True
        sHora1 = Date & " " & objRst!entra1
        sHora2 = Date + 1 & " " & objRst!sale1
    Else
        'Diurno Jornada Unica
        sHora1 = Date & " " & objRst!entra1
        sHora2 = Date & " " & objRst!sale1
    End If
ElseIf objRst!sale2 > objRst!entra1 Then
    'Diurno Jornada Normal (Doble)
    sHora1 = Date & " " & objRst!entra1
    sHora2 = Date & " " & objRst!sale2
Else
    'Nocturno Jornada Normal(DOble)
    bNoche = True
    sHora1 = Date & " " & objRst!entra1
    sHora2 = Date + 1 & " " & objRst!sale2
End If
objFUNC.sHorarioE = sHora1
objFUNC.sHorarioS = sHora2


If bNoche Then
    objFUNC.bHorario_noche = 1
Else
    objFUNC.bHorario_noche = 0
End If
bHoy = Weekday(Date)
sHoraA1 = DateAdd("n", Val("" & sMinu) * (-1), CDate(Date & " " & Time))
sHoraA2 = DateAdd("n", Val("" & sMinu), CDate(Date & " " & Time))
If bHLibre = False Then
    If bRestrictHorario Then
        If bNoESLabor Then
            sSql = vbNullString
            If objFUNC.iDispMODO = 1 Then
                sSql = "select 0 from vcan_entrar where id=" & Val(objFUNC.idFUNC) & " and getdate() between ae1 and ae2"
    ''            sSql = sSql & " or ae0<getdate()"
            ElseIf objFUNC.iDispMODO = 2 Then
                sSql = "select 0 from vcan_salir where id=" & Val(objFUNC.idFUNC) & " and (getdate() between as1 and as2 or getdate()>as3)"
            End If
            If sSql = vbNullString Then
                subMonitor "{HORARIO}", False, True
                Set objFUNC = Nothing
                Exit Sub
            Else
                Set objRstO = objCon.Execute(sSql)
                If objRstO.EOF Then
                    fnHablar "Acceso denegado."
                    subMonitor "{HORARIO}", False, True
                    Set objFUNC = Nothing
                    Exit Sub
                End If
            End If
        Else
            If (CDate(sHoraA1) < CDate(sHora1)) Or (sHoraA2 > CDate(sHora2)) Then
                'nodeberia entrar
                If objRst.Fields("d" & bHoy) Or bNoche Then
                    
                Else
                    sSql = "select * from tnovedades where idempleado=" & objFUNC.idFUNC
                    Set objRst = objCon.Execute(sSql)
                    If Not objRst.EOF Then
                        If objRst!idtipo = 3 Then 'Autorizacion
                            If Now >= objRst!fechaI And Now <= objRst!fechaf Then
                                fnHablar "Permiso asignado."
                            Else
                                fnHablar "No tiene autorización para ingresar a esta hora!"
                                subMonitor "{SIN_AUTORIZACION}", False, True
                                Set objFUNC = Nothing
                                Exit Sub
                            End If
                        Else
                            fnHablar "No tiene autorización para ingresar a esta hora!"
                            subMonitor "{SIN_AUTORIZACION}", False, True
                            Set objFUNC = Nothing
                            Exit Sub
                        End If
                    Else
                        fnHablar "No tiene autorización para ingresar a esta hora!"
                        subMonitor "{SIN_AUTORIZACION}", False, True
                        Set objFUNC = Nothing
                        Exit Sub
                    End If
                End If
            Else
                If Not (objRst.Fields("d" & bHoy) Or bNoche) Then
                    sSql = "select * from tnovedades where idempleado=" & objFUNC.idFUNC
                    Set objRst = objCon.Execute(sSql)
                    If Not objRst.EOF Then
                        If objRst!idtipo = 3 Then 'Autorizacion
                            If Now >= objRst!fechaI And Now <= objRst!fechaf Then
                                fnHablar "Permiso asignado."
                            Else
                                fnHablar "No tiene autorización para ingresar a esta hora!"
                                subMonitor "{SIN_AUTORIZACION}", False, True
                                Set objFUNC = Nothing
                                Exit Sub
                            End If
                        Else
                            fnHablar "No tiene autorización para ingresar a esta hora!"
                            subMonitor "{SIN_AUTORIZACION}", False, True
                            Set objFUNC = Nothing
                            Exit Sub
                        End If
                    Else
                        fnHablar "No tiene autorización para ingresar a esta hora!"
                        subMonitor "{SIN_AUTORIZACION}", False, True
                        Set objFUNC = Nothing
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End If
subAccesoFUNC
Exit Sub
errH:
Set objFUNC = Nothing
subLog "Error " & Err.Number & " " & Err.Description & "-subDatosFUNC"
End Sub

Sub subDatosFUNC_OLD()
Dim i As Byte, cnt As Byte
Dim bHoy As Byte
Dim bPermiso As Boolean

Dim sHoraA1 As String
Dim sHoraA2 As String
Dim sHora1 As String
Dim sHora2 As String
Dim sMinu As Byte
Dim bNoche As Boolean
If objFUNC.AccesoTipo = accMANUAL Then txtDoc1.Text = vbNullString
sSql = "select id,nombre,apellidos,sexo,fechai,fechaf,idhorario,activo from templeados where id=" & Val(objFUNC.idFUNC)
Set objRst = objCon.Execute(sSql)
objFUNC.sNOM = "" & objRst!nombre
objFUNC.sAPE = "" & objRst!apellidos
objFUNC.sSEXO = "" & objRst!sexo
If Not objRst!activo Then
    fnHablar "Funcionario inactivo!"
    Set objFUNC = Nothing
    Exit Sub
End If
If Not IsNull(objRst!fechaI) Then
    If objRst!fechaI > Date Then
        fnHablar "No está autorizado para ingresar!"
        Set objFUNC = Nothing
        Exit Sub
    End If
End If
If Not IsNull(objRst!fechaf) Then
    If objRst!fechaf < Date Then
        fnHablar "No está autorizado para ingresar!"
        Set objFUNC = Nothing
        Exit Sub
    End If
End If
If Val("" & objRst!idHorario) = 0 Then
    fnHablar "No tiene un horario asignado, acceso denegado!"
    Set objFUNC = Nothing
    Exit Sub
ElseIf bRestrictHorario Then
    sSql = "select * from thorarios where id=" & objRst!idHorario
    Set objRst = objCon.Execute(sSql)
    sMinu = Val("" & objRst!minutos)
    If IsNull(objRst!entra2) Then
        If objRst!sale1 < objRst!entra1 Then
            'Nocturno Jornada Unica
            bNoche = True
            sHora1 = Date & " " & objRst!entra1
            sHora2 = Date + 1 & " " & objRst!sale1
        Else
            'Diurno Jornada Unica
            sHora1 = Date & " " & objRst!entra1
            sHora2 = Date & " " & objRst!sale1
        End If
    ElseIf objRst!sale2 > objRst!entra1 Then
        'Diurno Jornada Normal (Doble)
        sHora1 = Date & " " & objRst!entra1
        sHora2 = Date & " " & objRst!sale2
    Else
        'Nocturno Jornada Normal(DOble)
        bNoche = True
        sHora1 = Date & " " & objRst!entra1
        sHora2 = Date + 1 & " " & objRst!sale2
    End If
    bHoy = Weekday(Date)
    sHoraA1 = DateAdd("n", Val("" & sMinu) * (-1), Time)
    sHoraA1 = Date & " " & sHoraA1
    sHoraA2 = DateAdd("n", Val("" & sMinu), Time)
    If bNoche Then
        sHoraA2 = Date + 1 & " " & sHoraA2
    Else
        sHoraA2 = Date & " " & sHoraA2
    End If
    If (CDate(sHoraA1) < CDate(sHora1)) Or (sHoraA2 > CDate(sHora2)) Then
        'entra si es diurno normal,diurno unica,nocturno normal
        If objRst.Fields("d" & bHoy) Or bNoche Then
            
        Else
            sSql = "select * from tnovedades where idempleado=" & objFUNC.idFUNC
            Set objRst = objCon.Execute(sSql)
            If Not objRst.EOF Then
                If objRst!idtipo = 3 Then 'Autorizacion
                    If Now >= objRst!fechaI And Now <= objRst!fechaf Then
                        fnHablar "Permiso asignado."
                    Else
                        fnHablar "No tiene autorización para ingresar a esta hora!"
                        Set objFUNC = Nothing
                        Exit Sub
                    End If
                Else
                    fnHablar "No tiene autorización para ingresar a esta hora!"
                    Set objFUNC = Nothing
                    Exit Sub
                End If
            Else
                fnHablar "No tiene autorización para ingresar a esta hora!"
                Set objFUNC = Nothing
                Exit Sub
            End If
        End If
    Else
        If objRst.Fields("d" & bHoy) Then
        
        Else
            sSql = "select * from tnovedades where idempleado=" & objFUNC.idFUNC
            Set objRst = objCon.Execute(sSql)
            If Not objRst.EOF Then
                If objRst!idtipo = 3 Then 'Autorizacion
                    If Now >= objRst!fechaI And Now <= objRst!fechaf Then
                        fnHablar "Permiso asignado."
                    Else
                        fnHablar "No tiene autorización para ingresar a esta hora!"
                        Set objFUNC = Nothing
                        Exit Sub
                    End If
                Else
                    fnHablar "No tiene autorización para ingresar a esta hora!"
                    Set objFUNC = Nothing
                    Exit Sub
                End If
            Else
                fnHablar "No tiene autorización para ingresar a esta hora!"
                Set objFUNC = Nothing
                Exit Sub
            End If
        End If
    End If
End If
subAccesoFUNC
End Sub
Private Sub subAccesoFUNC()
10    On Local Error GoTo errH
      Dim objRs_ As New ADODB.Recordset
      Dim objP As New ADODB.Recordset
      Dim sArr() As String, stNom As String, sHora As String
      Dim idPuerta_ant As Integer
      Dim sModoAnt As String
      Dim iLecDep As Byte
20    sArr = Split(Time, ":")
30    sHora = Val(sArr(0)) & " y " & Val(sArr(1)) & " " & Mid(sArr(2), 4)
    
    sSql = "select max(id) as ult from tacceso where idtpersona=" & Val(objFUNC.idFUNC) & " and idtipoper=" & tpFUNC
    '''''
''''    sSql = "select max(id) as ult from v_espuerta where idtpersona=" & Val(objFUNC.idFUNC) & " and idtipoper=" & tpFUNC & " and (idpe=" & objFUNC.idPuerta & " or idps=" & objFUNC.idPuerta & ")"

    
    Set objRs_ = objCon.Execute(sSql)
    objFUNC.idAccesoFUNC = Val("" & objRs_!ult)
    
    sSql = "select a.sale,cd.idcontrol from tacceso as a left join tcontrol_disp as cd on a.idhuellero_entra=cd.id where a.id=" & Val(objFUNC.idAccesoFUNC)
    Set objRs_ = objCon.Execute(sSql)
    If objFUNC.iDispMODO <> 3 Then
        If bAntiPass Then
            objFUNC.bEntraFUNC = Not IsNull(objRs_!sale)
        Else
            objFUNC.bEntraFUNC = (objFUNC.iDispMODO = 1)
        End If
    Else
        objFUNC.bEntraFUNC = Not IsNull(objRs_!sale)
    End If
    If bIntegridad Then
        sSql = "select count(id) as cuenta from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta
        Set objP = objCon.Execute(sSql)
        iLecDep = objP!cuenta
        sSql = "select top 1 * from vacceso_integridad where idfunc=" & objFUNC.idFUNC & " order by fh desc"
        Set objP = objCon.Execute(sSql)
        If objP.EOF Then
            If objFUNC.iDispMODO <> 1 Then
                fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
                Set objFUNC = Nothing
                Exit Sub
            Else
                If iLecDep = 1 Then
                    sSql = "select idcontrol_previo,modo from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta
                    Set objP = objCon.Execute(sSql)
                    idPuerta_ant = objP!idcontrol_previo
                    If objP!modo = "S" Then
                        If idPuerta_ant < objFUNC.idPuerta Then
                            fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
                            Set objFUNC = Nothing
                            Exit Sub
                        End If
                    Else
                    
                    End If
                ElseIf iLecDep = 2 Then
                    fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
                    Set objFUNC = Nothing
                    Exit Sub
                End If
            End If
        Else
            idPuerta_ant = objP!idControl
            sModoAnt = objP!modo
            If objFUNC.idPuerta <> idPuerta_ant Then
                If iLecDep = 1 Then
                    sSql = "select idcontrol_previo,modo from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta & " and idcontrol_previo=" & idPuerta_ant
                    If objFUNC.iDispMODO = 1 Then
                        sSql = sSql & " and modo='E'"
                    ElseIf objFUNC.iDispMODO = 2 Then
                        sSql = sSql & " and modo='S'"
                    End If
                    Set objP = objCon.Execute(sSql)
                    If objP.EOF Then
                        fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
                        Set objFUNC = Nothing
                        Exit Sub
                    Else
                        If objFUNC.iDispMODO = 1 Then
                            objFUNC.bEntraFUNC = True
                        ElseIf objFUNC.iDispMODO = 2 Then
                            objFUNC.bEntraFUNC = False
                        End If
                    End If
                ElseIf iLecDep = 2 Then
                    sSql = "select idcontrol_previo,modo from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta & " and idcontrol_previo=" & idPuerta_ant
                    If objFUNC.iDispMODO = 1 Then
                        sSql = sSql & " and modo='E'"
                    ElseIf objFUNC.iDispMODO = 2 Then
                        sSql = sSql & " and modo='S'"
                    End If
                    Set objP = objCon.Execute(sSql)
                    If objP.EOF Then
                        fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
                        Set objFUNC = Nothing
                        Exit Sub
                    Else
                        If Not (idPuerta_ant > objFUNC.idPuerta And sModoAnt = "S") And Not (idPuerta_ant < objFUNC.idPuerta And sModoAnt = "E") Then
                            fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
                            Set objFUNC = Nothing
                            Exit Sub
                        Else
                            If objFUNC.iDispMODO = 1 Then
                                objFUNC.bEntraFUNC = True
                            ElseIf objFUNC.iDispMODO = 2 Then
                                objFUNC.bEntraFUNC = False
                            End If
                        End If
                    End If
                End If
            Else
                If sModoAnt = "S" And objFUNC.iDispMODO = 1 Then
                    objFUNC.bEntraFUNC = True
                ElseIf sModoAnt = "E" And objFUNC.iDispMODO = 2 Then
                    objFUNC.bEntraFUNC = False
                Else
                    fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
                    Set objFUNC = Nothing
                    Exit Sub
                End If
            End If
        End If
        
        If bPuertaES Then
            If objFUNC.bEntraFUNC = False Then
                If Val("" & objRs_!idControl) <> 0 Then
                    If Val("" & objRs_!idControl) <> objFUNC.idPuerta Then
                        fnHablar "Debe salir por la misma puerta que ingresó!."
                        Set objFUNC = Nothing
                        Exit Sub
                    End If
                Else
                    If objFUNC.idDisp <> 0 Then
                        fnHablar "Al igual que la entrada, debe registrar su salida de forma manual."
                        Set objFUNC = Nothing
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

100   sArr = Split(objFUNC.sNOM, " ")
110   If UBound(sArr) >= 0 Then
120       stNom = sArr(0)
130   End If
140   sArr = Split(objFUNC.sAPE, " ")
150   If UBound(sArr) >= 0 Then
160       stNom = stNom & " " & sArr(0)
170   End If

        If bIntegridad And objFUNC.iDispMODO = 2 Then
            sSql = "select top 1 * from tacceso where idtipoper=1 and idtpersona=" & objFUNC.idFUNC & " and sale is null order by id"
            Debug.Print objFUNC.bEntraFUNC
        Else
            sSql = "select * from tacceso where id=" & Val(objFUNC.idAccesoFUNC)
        End If
190   With objRst
200       If .State = adStateOpen Then .Close
210       .Open sSql, objCon, adOpenKeyset, adLockOptimistic
220       If objFUNC.bEntraFUNC Then
              'MsgBox iDispMODO
230           If objFUNC.AccesoTipo = accHUELLA_ZK Then
240               objFUNC.bRegPUERTO = idxZK
250           Else
260               objFUNC.bRegPUERTO = bPulsoE
270           End If
280           If objFUNC.iDispMODO = 0 Or objFUNC.iDispMODO = 1 Or objFUNC.iDispMODO = 3 Then 'E,ES
290               .AddNew
300               !idTipoPer = tpFUNC
310               !idtpersona = Val(objFUNC.idFUNC)
320               !entra = fnFecha(Now, True)
330               !idhuellero_entra = objFUNC.idDisp
340               !idlogin_e = idLogin
350               !terminal_e = sTerminal
352                If objFUNC.sHorarioE <> vbNullString Then
353                    !horario_e = fnFecha(CDate(objFUNC.sHorarioE), True)
354                    !horario_noche = objFUNC.bHorario_noche
355                End If
380               fnHablar stNom & " " & sHora
390           ElseIf objFUNC.iDispMODO = 2 Then
400               objFUNC.bREG = False
410               If objFUNC.AccesoTipo = acc2D Then
420                   fnHablar "Este lector solo registra la salida."
                    subMonitor "{SIN_ENTRADA}", False, True
430               ElseIf objFUNC.AccesoTipo = accHUELLA Then
440                   fnHablar "Este huellero solo registra la salida."
                    subMonitor "{SIN_ENTRADA}", False, True
450               ElseIf objFUNC.AccesoTipo = accTARJETA Or objFUNC.AccesoTipo = accHUELLA_ZK Then
460                   fnHablar "Este dispositivo solo registra la salida."
                    subMonitor "{SIN_ENTRADA}", False, True
470               End If
480               Set objFUNC = Nothing
490               Exit Sub
500           End If
510       Else
520           If objFUNC.AccesoTipo = accHUELLA_ZK Then
530               objFUNC.bRegPUERTO = idxZK
540           Else
550               objFUNC.bRegPUERTO = bPulsoS
560           End If
570           If objFUNC.iDispMODO = 0 Or objFUNC.iDispMODO = 2 Or objFUNC.iDispMODO = 3 Then 'S,ES
580               !sale = fnFecha(Now, True)
590               !idhuellero_sale = objFUNC.idDisp
600               !idlogin_s = idLogin
610               !terminal_s = sTerminal
                If objFUNC.sHorarioS <> vbNullString Then
                    !horario_s = fnFecha(CDate(objFUNC.sHorarioS), True)
                    !horario_noche = objFUNC.bHorario_noche
                End If
640               fnHablar sHora & " " & stNom
650           ElseIf objFUNC.iDispMODO = 1 Then
660               objFUNC.bREG = False
670               If objFUNC.AccesoTipo = acc2D Then
680                   fnHablar "Este lector solo registra la entrada."
                    subMonitor "{SIN_SALIDA}", False, True
690               ElseIf objFUNC.AccesoTipo = accHUELLA Then
700                   fnHablar "Este huellero solo registra la entrada."
                    subMonitor "{SIN_SALIDA}", False, True
710               ElseIf objFUNC.AccesoTipo = accTARJETA Or objFUNC.AccesoTipo = accHUELLA_ZK Then
720                   fnHablar "Este dispositivo solo registra la entrada."
                    subMonitor "{SIN_SALIDA}", False, True
730               End If
'740               subLimpiar
                    Set objFUNC = Nothing
750               Exit Sub
760           End If
770       End If
780       .UpDate
        If objFUNC.bEntraFUNC Then
            subMonitor "{REG_ENTRA_OK}", False, False
        Else
            subMonitor "{REG_SALE_OK}", False, False
        End If
790       objFUNC.idAccesoFUNC = !id
800       .Close
810   End With
820   If objFUNC.bREG Then subRELEVO objFUNC.bRegPUERTO
830   Set objFUNC = Nothing
subMonitor "{END}", False, True
840   Exit Sub
errH:
Set objFUNC = Nothing
subLog "Error " & Err.Number & " " & Err.Description & "-subAccesoFUNC Linea No. " & Erl
End Sub
'Private Sub subAccesoFUNC_OLD2()
'10    On Local Error GoTo errH
'      Dim objRs_ As New ADODB.Recordset
'      Dim objP As New ADODB.Recordset
'      Dim sArr() As String, stNom As String, sHora As String
'      Dim idPuerta_ant As Integer
'      Dim sModoAnt As String
'      Dim iLecDep As Byte
'20    sArr = Split(Time, ":")
'30    sHora = Val(sArr(0)) & " y " & Val(sArr(1)) & " " & Mid(sArr(2), 4)
'
'    sSql = "select max(id) as ult from tacceso where idtpersona=" & Val(objFUNC.idFUNC) & " and idtipoper=" & tpFUNC
'    Set objRs_ = objCon.Execute(sSql)
'    objFUNC.idAccesoFUNC = Val("" & objRs_!ult)
'
'    sSql = "select a.sale,cd.idcontrol from tacceso as a left join tcontrol_disp as cd on a.idhuellero_entra=cd.id where a.id=" & Val(objFUNC.idAccesoFUNC)
'    Set objRs_ = objCon.Execute(sSql)
'    objFUNC.bEntraFUNC = Not IsNull(objRs_!sale)
'    If bIntegridad Then
'        sSql = "select count(id) as cuenta from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta
'        Set objP = objCon.Execute(sSql)
'        iLecDep = objP!cuenta
'        sSql = "select top 1 * from vacceso_integridad where idfunc=" & objFUNC.idFUNC & " order by fh desc"
'        Set objP = objCon.Execute(sSql)
'        If objP.EOF Then
'            If iLecDep = 1 Then
'
'                sSql = "select idcontrol_previo,modo from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta
'                Set objP = objCon.Execute(sSql)
'                If objP!modo <> "S" And objFUNC.iDispMODO <> 1 Then
'                    fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
'                    Set objFUNC = Nothing
'                    Exit Sub
'                Else
'                    If objFUNC.iDispMODO = 1 Then
'                        objFUNC.bEntraFUNC = True
'                    Else
'                        fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
'                        Set objFUNC = Nothing
'                        Exit Sub
'                    End If
'                End If
'            ElseIf iLecDep = 2 Then
'                fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
'                Set objFUNC = Nothing
'                Exit Sub
'            End If
'        Else
'            idPuerta_ant = objP!idControl
'            sModoAnt = objP!modo
'            If objFUNC.idPuerta <> idPuerta_ant Then
'                If iLecDep = 1 Then
'                    sSql = "select idcontrol_previo,modo from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta & " and idcontrol_previo=" & idPuerta_ant
'                    Set objP = objCon.Execute(sSql)
'                    If objP.EOF Then
'                        fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
'                        Set objFUNC = Nothing
'                        Exit Sub
'                    Else
'                        If objP!modo <> sModoAnt Then
'                            fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
'                            Set objFUNC = Nothing
'                            Exit Sub
'                        Else
'                            objFUNC.bEntraFUNC = False
'                        End If
'                    End If
'                ElseIf iLecDep = 2 Then
'                    sSql = "select idcontrol_previo,modo from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta & " and idcontrol_previo=" & idPuerta_ant
'                    'sSql = sSql & " and modo<>'" & objP!modo & "'"
'                    Set objP = objCon.Execute(sSql)
'                    If objP!modo <> sModoAnt Then
'                        fnHablar "No cumple el protocolo de integridad!. Acceso denegado."
'                        Set objFUNC = Nothing
'                        Exit Sub
'                    Else
'                        If objFUNC.iDispMODO = 1 Then
'                            objFUNC.bEntraFUNC = True
'                        ElseIf objFUNC.iDispMODO = 2 Then
'                            objFUNC.bEntraFUNC = False
'                        End If
'                    End If
'                End If
'            Else
'                If sModoAnt = "S" And objFUNC.iDispMODO = 1 Then
'                    objFUNC.bEntraFUNC = True
'                ElseIf sModoAnt = "E" And objFUNC.iDispMODO = 2 Then
'                    objFUNC.bEntraFUNC = False
'                Else
'                    Set objFUNC = Nothing
'                    fnHablar "Este dispositivo solo registra la salida."
'                    Exit Sub
'                End If
'            End If
'        End If
'
'        If bPuertaES Then
'            If objFUNC.bEntraFUNC = False Then
'                If Val("" & objRs_!idControl) <> 0 Then
'                    If Val("" & objRs_!idControl) <> objFUNC.idPuerta Then
'                        fnHablar "Debe salir por la misma puerta que ingresó!."
'                        Set objFUNC = Nothing
'                        Exit Sub
'                    End If
'                Else
'                    If objFUNC.idDisp <> 0 Then
'                        fnHablar "Al igual que la entrada, debe registrar su salida de forma manual."
'                        Set objFUNC = Nothing
'                        Exit Sub
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'100   sArr = Split(objFUNC.sNOM, " ")
'110   If UBound(sArr) >= 0 Then
'120       stNom = sArr(0)
'130   End If
'140   sArr = Split(objFUNC.sAPE, " ")
'150   If UBound(sArr) >= 0 Then
'160       stNom = stNom & " " & sArr(0)
'170   End If
'
'180   sSql = "select * from tacceso where id=" & Val(objFUNC.idAccesoFUNC)
'190   With objRst
'200       If .State = adStateOpen Then .Close
'210       .Open sSql, objCon, adOpenKeyset, adLockOptimistic
'220       If objFUNC.bEntraFUNC Then
'              'MsgBox iDispMODO
'230           If objFUNC.AccesoTipo = accHUELLA_ZK Then
'240               objFUNC.bRegPUERTO = idxZK
'250           Else
'260               objFUNC.bRegPUERTO = bPulsoE
'270           End If
'280           If objFUNC.iDispMODO = 0 Or objFUNC.iDispMODO = 1 Or objFUNC.iDispMODO = 3 Then 'E,ES
'290               .AddNew
'300               !idTipoPer = tpFUNC
'310               !idtpersona = Val(objFUNC.idFUNC)
'320               !entra = fnFecha(Now, True)
'330               !idhuellero_entra = objFUNC.idDisp
'340               !idlogin_e = idLogin
'350               !terminal_e = sTerminal
'                If objFUNC.sHorarioE <> vbNullString Then
'                    !horario_e = fnFecha(CDate(objFUNC.sHorarioE), True)
'                    !horario_noche = objFUNC.bHorario_noche
'                End If
'380               fnHablar stNom & " " & sHora
'390           ElseIf objFUNC.iDispMODO = 2 Then
'400               objFUNC.bREG = False
'410               If objFUNC.AccesoTipo = acc2D Then
'420                   fnHablar "Este lector solo registra la salida."
'430               ElseIf objFUNC.AccesoTipo = accHUELLA Then
'440                   fnHablar "Este huellero solo registra la salida."
'450               ElseIf objFUNC.AccesoTipo = accTARJETA Or objFUNC.AccesoTipo = accHUELLA_ZK Then
'460                   fnHablar "Este dispositivo solo registra la salida."
'470               End If
'480               Set objFUNC = Nothing
'490               Exit Sub
'500           End If
'510       Else
'520           If objFUNC.AccesoTipo = accHUELLA_ZK Then
'530               objFUNC.bRegPUERTO = idxZK
'540           Else
'550               objFUNC.bRegPUERTO = bPulsoS
'560           End If
'570           If objFUNC.iDispMODO = 0 Or objFUNC.iDispMODO = 2 Or objFUNC.iDispMODO = 3 Then 'S,ES
'580               !sale = fnFecha(Now, True)
'590               !idhuellero_sale = objFUNC.idDisp
'600               !idlogin_s = idLogin
'610               !terminal_s = sTerminal
'                If objFUNC.sHorarioS <> vbNullString Then
'                    !horario_s = fnFecha(CDate(objFUNC.sHorarioS), True)
'                    !horario_noche = objFUNC.bHorario_noche
'                End If
'640               fnHablar sHora & " " & stNom
'650           ElseIf objFUNC.iDispMODO = 1 Then
'660               objFUNC.bREG = False
'670               If objFUNC.AccesoTipo = acc2D Then
'680                   fnHablar "Este lector solo registra la entrada."
'690               ElseIf objFUNC.AccesoTipo = accHUELLA Then
'700                   fnHablar "Este huellero solo registra la entrada."
'710               ElseIf objFUNC.AccesoTipo = accTARJETA Or objFUNC.AccesoTipo = accHUELLA_ZK Then
'720                   fnHablar "Este dispositivo solo registra la entrada."
'730               End If
'740               subLimpiar
'750               Exit Sub
'760           End If
'770       End If
'780       .UpDate
'790       objFUNC.idAccesoFUNC = !id
'800       .Close
'810   End With
'820   If objFUNC.bREG Then subRELEVO objFUNC.bRegPUERTO
'830   Set objFUNC = Nothing
'840   Exit Sub
'errH:
'850   MsgBox "Error " & Err.Number & " " & Err.Description & "-subAccesoFUNC Linea No. " & Erl
'End Sub
'
''''Private Sub subAccesoFUNC_OLD()
''''10    On Local Error GoTo errH
''''      Dim objRs_ As New ADODB.Recordset
''''      Dim objP As New ADODB.Recordset
''''      Dim sArr() As String, stNom As String, sHora As String
''''      Dim idPuerta_ant As Integer
''''      Dim iLecDep As Byte
''''20    sArr = Split(Time, ":")
''''30    sHora = Val(sArr(0)) & " y " & Val(sArr(1)) & " " & Mid(sArr(2), 4)
''''
''''    sSql = "select max(id) as ult from tacceso where idtpersona=" & Val(objFUNC.idFUNC) & " and idtipoper=" & tpFUNC
''''    Set objRs_ = objCon.Execute(sSql)
''''    objFUNC.idAccesoFUNC = Val("" & objRs_!ult)
''''
''''    sSql = "select a.sale,cd.idcontrol from tacceso as a left join tcontrol_disp as cd on a.idhuellero_entra=cd.id where a.id=" & Val(objFUNC.idAccesoFUNC)
''''    Set objRs_ = objCon.Execute(sSql)
''''    objFUNC.bEntraFUNC = Not IsNull(objRs_!sale)
''''    If bIntegridad Then
''''        sSql = "select idcontrol from tcontrol_disp where id=(select case isnull(sale,0) when 0 then idhuellero_entra else idhuellero_sale end iddisp from tacceso where id=" & objFUNC.idAccesoFUNC & ")"
''''        Set objP = objCon.Execute(sSql)
''''        If Not objP.EOF Then idPuerta_ant = objP!idControl
''''        If idPuerta_ant <> objFUNC.idPuerta Then
''''            sSql = "select count(id) as cuenta from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta
''''            Set objP = objCon.Execute(sSql)
''''            iLecDep = objP!cuenta
''''            If iLecDep = 0 Then
''''
''''            ElseIf iLecDep = 1 Then
''''                sSql = "select idcontrol_previo,modo from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta
''''                If idPuerta_ant <> 0 Then
''''                    sSql = sSql & " and idcontrol_previo=" & idPuerta_ant
''''                End If
''''                Set objP = objCon.Execute(sSql)
''''                If objP.EOF Then
''''                    fnHablar "No cumple el protocolo de integridad de puertas!. Acceso denegado."
''''                    Set objFUNC = Nothing
''''                    Exit Sub
''''                Else
''''                    If objP!modo = "S" Then
''''                        If objFUNC.iDispMODO = 1 Or objFUNC.iDispMODO = 3 Then
''''                            If objFUNC.bEntraFUNC Then
''''                                If idPuerta_ant <> 0 Then objFUNC.bEntraFUNC = False
''''                            Else
''''                                fnHablar "No cumple el protocolo de integridad de puertas!. Acceso denegado."
''''                                Set objFUNC = Nothing
''''                                Exit Sub
''''                            End If
''''                        End If
''''                    ElseIf objP!modo = "E" Then
''''                        If objFUNC.iDispMODO = 1 Or objFUNC.iDispMODO = 3 Then
''''    '                        If objFUNC.bEntraFUNC Then
''''                                If idPuerta_ant <> 0 Then
''''                                    objFUNC.bEntraFUNC = True
''''                                Else
''''                                    fnHablar "No cumple el protocolo de integridad de puertas!. Acceso denegado."
''''                                    Set objFUNC = Nothing
''''                                    Exit Sub
''''                                End If
''''    '                        Else
''''    '                            fnHablar "No cumple el protocolo de integridad de puertas!. Acceso denegado."
''''    '                            Set objFUNC = Nothing
''''    '                            Exit Sub
''''    '                        End If
''''                        End If
''''                    End If
''''                End If
''''            ElseIf iLecDep = 2 Then
''''                sSql = "select idcontrol_previo,modo from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta & " and idcontrol_previo=" & idPuerta_ant
''''                Set objP = objCon.Execute(sSql)
''''                If objP.EOF Then
''''                    fnHablar "No cumple el protocolo de integridad de puertas!. Acceso denegado."
''''                    Set objFUNC = Nothing
''''                    Exit Sub
''''                Else
''''                    If Not (objP!modo = "E") Then
''''                        If objFUNC.bEntraFUNC Then
''''                            objFUNC.bEntraFUNC = False
''''                        Else
''''                            fnHablar "No cumple el protocolo de integridad de puertas!. Acceso denegado."
''''                            Set objFUNC = Nothing
''''                            Exit Sub
''''                        End If
''''                    Else
''''                        If objFUNC.bEntraFUNC Then
''''
''''                        Else
''''                            objFUNC.bEntraFUNC = True
''''                            If Not (objP!modo = "S" And objFUNC.bEntraFUNC = False) Then
''''                                fnHablar "No cumple el protocolo de integridad de puertas!. Acceso denegado."
''''                                Set objFUNC = Nothing
''''                                Exit Sub
''''                            End If
''''                        End If
''''                    End If
''''                End If
''''            Else
''''
''''            End If
''''
''''
''''
''''
'''''                sSql = "select idcontrol_previo,modo from tpuertas_asoc where idcontrol=" & objFUNC.idPuerta
'''''                With objP
'''''                    If .State = adStateOpen Then .Close
'''''                    .Open sSql, objCon, adOpenForwardOnly
'''''
'''''                    If objP.EOF Then
'''''                        If Not (objFUNC.bEntraFUNC And idPuerta_ant = 0 And (objFUNC.iDispMODO = 1 Or objFUNC.iDispMODO = 3)) Then
'''''                            fnHablar "No cumple el protocolo de integridad de puertas!. Acceso denegado."
'''''                            Set objFUNC = Nothing
'''''                            Exit Sub
'''''                        End If
'''''                    ElseIf Not (.RecordCount = 1 And objFUNC.bEntraFUNC And !modo = "S" And (objFUNC.iDispMODO = 1 Or objFUNC.iDispMODO = 3)) Then
'''''                        If Not (objP!modo = "E" And objFUNC.bEntraFUNC) Then
'''''                            If .RecordCount > 1 Then
'''''                                While Not .EditMode
'''''
'''''                                Wend
'''''                            End If
'''''                            fnHablar "No cumple el protocolo de integridad de puertas!. Acceso denegado."
'''''                            Set objFUNC = Nothing
'''''                            Exit Sub
'''''                        ElseIf Not (objP!modo = "S" And objFUNC.bEntraFUNC = False) Then
'''''                            fnHablar "No cumple el protocolo de integridad de puertas!. Acceso denegado."
'''''                            Set objFUNC = Nothing
'''''                            Exit Sub
'''''                        End If
'''''                    End If
'''''                    .Close
'''''                End With
''''
''''        End If
''''        If bPuertaES Then
''''            If objFUNC.bEntraFUNC = False Then
''''                If Val("" & objRs_!idControl) <> 0 Then
''''                    If Val("" & objRs_!idControl) <> objFUNC.idPuerta Then
''''                        fnHablar "Debe salir por la misma puerta que ingresó!."
''''                        Set objFUNC = Nothing
''''                        Exit Sub
''''                    End If
''''                Else
''''                    If objFUNC.idDisp <> 0 Then
''''                        fnHablar "Al igual que la entrada, debe registrar su salida de forma manual."
''''                        Set objFUNC = Nothing
''''                        Exit Sub
''''                    End If
''''                End If
''''            End If
''''        End If
''''    End If
''''
''''
''''100   sArr = Split(objFUNC.sNOM, " ")
''''110   If UBound(sArr) >= 0 Then
''''120       stNom = sArr(0)
''''130   End If
''''140   sArr = Split(objFUNC.sAPE, " ")
''''150   If UBound(sArr) >= 0 Then
''''160       stNom = stNom & " " & sArr(0)
''''170   End If
''''
''''180   sSql = "select * from tacceso where id=" & Val(objFUNC.idAccesoFUNC)
''''190   With objRst
''''200       If .State = adStateOpen Then .Close
''''210       .Open sSql, objCon, adOpenKeyset, adLockOptimistic
''''220       If objFUNC.bEntraFUNC Then
''''              'MsgBox iDispMODO
''''230           If objFUNC.AccesoTipo = accHUELLA_ZK Then
''''240               objFUNC.bRegPUERTO = idxZK
''''250           Else
''''260               objFUNC.bRegPUERTO = bPulsoE
''''270           End If
''''280           If objFUNC.iDispMODO = 0 Or objFUNC.iDispMODO = 1 Or objFUNC.iDispMODO = 3 Then 'E,ES
''''290               .AddNew
''''300               !idTipoPer = tpFUNC
''''310               !idtpersona = Val(objFUNC.idFUNC)
''''320               !entra = fnFecha(Now, True)
''''330               !idhuellero_entra = objFUNC.idDisp
''''340               !idlogin_e = idLogin
''''350               !terminal_e = sTerminal
''''                If objFUNC.sHorarioE <> vbNullString Then
''''                    !horario_e = fnFecha(CDate(objFUNC.sHorarioE), True)
''''                    !horario_noche = objFUNC.bHorario_noche
''''                End If
''''380               fnHablar stNom & " " & sHora
''''390           ElseIf objFUNC.iDispMODO = 2 Then
''''400               objFUNC.bREG = False
''''410               If objFUNC.AccesoTipo = acc2D Then
''''420                   fnHablar "Este lector solo registra la salida."
''''430               ElseIf objFUNC.AccesoTipo = accHUELLA Then
''''440                   fnHablar "Este huellero solo registra la salida."
''''450               ElseIf objFUNC.AccesoTipo = accTARJETA Or objFUNC.AccesoTipo = accHUELLA_ZK Then
''''460                   fnHablar "Este dispositivo solo registra la salida."
''''470               End If
''''480               Set objFUNC = Nothing
''''490               Exit Sub
''''500           End If
''''510       Else
''''520           If objFUNC.AccesoTipo = accHUELLA_ZK Then
''''530               objFUNC.bRegPUERTO = idxZK
''''540           Else
''''550               objFUNC.bRegPUERTO = bPulsoS
''''560           End If
''''570           If objFUNC.iDispMODO = 0 Or objFUNC.iDispMODO = 2 Or objFUNC.iDispMODO = 3 Then 'S,ES
''''580               !sale = fnFecha(Now, True)
''''590               !idhuellero_sale = objFUNC.idDisp
''''600               !idlogin_s = idLogin
''''610               !terminal_s = sTerminal
''''                If objFUNC.sHorarioS <> vbNullString Then
''''                    !horario_s = fnFecha(CDate(objFUNC.sHorarioS), True)
''''                    !horario_noche = objFUNC.bHorario_noche
''''                End If
''''640               fnHablar sHora & " " & stNom
''''650           ElseIf objFUNC.iDispMODO = 1 Then
''''660               objFUNC.bREG = False
''''670               If objFUNC.AccesoTipo = acc2D Then
''''680                   fnHablar "Este lector solo registra la entrada."
''''690               ElseIf objFUNC.AccesoTipo = accHUELLA Then
''''700                   fnHablar "Este huellero solo registra la entrada."
''''710               ElseIf objFUNC.AccesoTipo = accTARJETA Or objFUNC.AccesoTipo = accHUELLA_ZK Then
''''720                   fnHablar "Este dispositivo solo registra la entrada."
''''730               End If
''''740               subLimpiar
''''750               Exit Sub
''''760           End If
''''770       End If
''''780       .UpDate
''''790       objFUNC.idAccesoFUNC = !id
''''800       .Close
''''810   End With
''''820   If objFUNC.bREG Then subRELEVO objFUNC.bRegPUERTO
''''830   Set objFUNC = Nothing
''''840   Exit Sub
''''errH:
''''850   MsgBox "Error " & Err.Number & " " & Err.Description & "-subAccesoFUNC Linea No. " & Erl
''''End Sub
''''
Private Sub subAccesoVISI()
Dim objRs_ As New ADODB.Recordset
Dim sArr() As String, stNom As String

sSql = "select max(id) as ult from tacceso where idtpersona=" & Val(objVISI.idVISI) & " and idtipoper=" & tpVISI
Set objRs_ = objCon.Execute(sSql)
objVISI.idAccesoVISI = Val("" & objRs_!ult)
sSql = "select sale,idempleado,idautoriza from tacceso where id=" & Val(objVISI.idAccesoVISI)
Set objRs_ = objCon.Execute(sSql)
objVISI.bEntraVISI = Not IsNull(objRs_!sale)
If objVISI.bEntraVISI = True Then
    If objVISI.bFrecuente = False Then
        If bAutoStmp Then
            bAutoStmp = False
            subDatosVISI
            Exit Sub
        End If
    End If
End If
If bAutorizado = False Then
    If Not objRs_.EOF Then
        idEmpleado = objRs_!idEmpleado
        idAutoriza = objRs_!idAutoriza
    End If
End If
bAutorizado = False
sArr = Split(objVISI.sNOM, " ")
If UBound(sArr) >= 0 Then
    stNom = sArr(0)
End If
sArr = Split(objVISI.sAPE, " ")
If UBound(sArr) >= 0 Then
    stNom = stNom & " " & sArr(0)
End If
sSql = "select * from tacceso where id=" & Val(objVISI.idAccesoVISI)
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If objVISI.bEntraVISI Then
        'MsgBox iDispMODO
        If objVISI.iDispMODO = 0 Or objVISI.iDispMODO = 1 Or objVISI.iDispMODO = 3 Then 'E,ES
            objVISI.bRegPUERTO = bPulsoE
            .AddNew
            !idTipoPer = TipoPER
            !idtpersona = Val(objVISI.idVISI)
            !entra = fnFecha(Now, True)
            !idhuellero_entra = objVISI.idDisp
            !idEmpleado = idEmpleado
            !idAutoriza = idAutoriza
            !turno = Val(lblTurno.Caption)
            !idlogin_e = idLogin
            !terminal_e = sTerminal
            If objVISI.sSEXO <> vbNullString Then
                fnHablar "Bienvenid" & IIf((objVISI.sSEXO = "M"), "o ", "a ") & stNom & "."
            Else
                fnHablar "Bienvenido " & stNom & "."
            End If
        ElseIf objVISI.iDispMODO = 2 Then
            bREG = False
            If objVISI.AccesoTipo = acc2D Then
                fnHablar "Este lector solo registra la salida."
            ElseIf objVISI.AccesoTipo = accHUELLA Then
                fnHablar "Este huellero solo registra la salida."
            ElseIf objVISI.AccesoTipo = accTARJETA Then
                fnHablar "Este lector solo registra la salida."
                If wSerDatos.State = sckConnected Then
                    wSerDatos.SendData "cancel|"
                    DoEvents
                End If
            End If
            Set objVISI = Nothing
            subLimpiar
            Exit Sub
        End If
    Else
        If objVISI.iDispMODO = 0 Or objVISI.iDispMODO = 2 Or objVISI.iDispMODO = 3 Then 'S,ES
            objVISI.bRegPUERTO = bPulsoS
            sSql = "delete from tmp_objetos where idvisitante=" & objVISI.idVISI
            objCon.Execute sSql
            sSql = "select * from tobjetos where idacceso=" & objVISI.idAccesoVISI & " and abs(estado)=1"
            Set objRstA = objCon.Execute(sSql)
            If Not objRstA.EOF Then
                objVISI.bObjetos = True
                cmdObjetos_Click
            End If
            
            !sale = fnFecha(Now, True)
            !idhuellero_sale = objVISI.idDisp
            !idlogin_s = idLogin
            !terminal_s = sTerminal
            fnHablar "Hasta pronto " & stNom & "."
        ElseIf objVISI.iDispMODO = 1 Then
            If objVISI.AccesoTipo = acc2D Then
                fnHablar "Este lector solo registra la entrada."
            ElseIf objVISI.AccesoTipo = accHUELLA Then
                fnHablar "Este huellero solo registra la entrada."
            ElseIf objVISI.AccesoTipo = accTARJETA Then
                fnHablar "Este lector solo registra la entrada."
                If wSerDatos.State = sckConnected Then
                    wSerDatos.SendData "cancel|"
                    DoEvents
                End If
            End If
            Set objVISI = Nothing
            subLimpiar
            Exit Sub
        End If
    End If
    .UpDate
    objVISI.idAccesoVISI = !id
    .Close
End With
'If wServidor.State = sckConnected Then
'    wServidor.SendData CStr(objVISI.idAccesoVISI)
'End If
If objVISI.bEntraVISI = False Then
    sSql = "update tvisitantes_huella set tarjeta='' where id=" & objVISI.idVISI
    objCon.Execute sSql
End If
If objVISI.bEntraVISI Then
    sSql = "select * from tmp_objetos where idvisitante=" & Val(objVISI.idVISI)
    Set objRst = objCon.Execute(sSql)
    If Not objRst.EOF Then
        sSql = "insert into tobjetos(idacceso,descripcion,serial,estado,foto)"
        sSql = sSql & " select " & Val(objVISI.idAccesoVISI) & ",tmp_objetos.descripcion,tmp_objetos.serial,1,tmp_objetos.foto from tmp_objetos"
        sSql = sSql & " where idvisitante=" & objVISI.idVISI
        objCon.Execute (sSql)
        
        sSql = "delete from tmp_objetos where idvisitante=" & Val(objVISI.idVISI)
        objCon.Execute (sSql)
    End If
End If
If objVISI.bEntraVISI Then
    If bStickerFr And objVISI.bFrecuente Then
        subStickerAuto
    End If
    If bStickerA And chkAuto.Value = vbChecked Then
        subStickerAuto
    End If
End If
    
If objVISI.AccesoTipo = accMANUAL Then
    Set objVISI = Nothing
    subLimpiar
Else
    If objVISI.bREG Then
        If objVISI.AccesoTipo = accTARJETA Then
            If objVISI.bEntraVISI Then
                wSerDatos.SendData "tarest|1"
            Else
                wSerDatos.SendData "tarest|2"
            End If
            DoEvents
        Else
            subRELEVO objVISI.bRegPUERTO
            DoEvents
        End If
    End If
    Set objVISI = Nothing
    subLimpiar
End If
End Sub
Private Sub subAccesoVISIManual()
Dim objRs_ As New ADODB.Recordset
Dim sArr() As String, stNom As String
sArr = Split(objVISIManual.sNOM, " ")
If UBound(sArr) >= 0 Then
    stNom = sArr(0)
End If
sArr = Split(objVISIManual.sAPE, " ")
If UBound(sArr) >= 0 Then
    stNom = stNom & " " & sArr(0)
End If
sSql = "select * from tacceso where id=" & Val(objVISIManual.idAccesoVISI)
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If objVISIManual.bEntraVISI Then
        If objVISIManual.iDispMODO = 0 Or objVISIManual.iDispMODO = 1 Or objVISIManual.iDispMODO = 3 Then 'E,ES
            objVISIManual.bRegPUERTO = bPulsoE
            If objVISIManual.bEspera = False Then .AddNew
            !idTipoPer = tpVISI
            !idtpersona = Val(objVISIManual.idVISI)
            !entra = fnFecha(Now, True)
            !idhuellero_entra = objVISIManual.idDisp
            !idEmpleado = idEmpleado
            !idAutoriza = idAutoriza
            !turno = idTurno
            !espera = IIf(bEsperaTemp, -1, 0)
            !idlogin_e = idLogin
            !terminal_e = sTerminal
            If bEsperaTemp Then
                fnHablar "En espera."
            Else
                If objVISIManual.sSEXO <> vbNullString Then
                    fnHablar "Bienvenid" & IIf((objVISIManual.sSEXO = "M"), "o ", "a ") & stNom & "."
                Else
                    fnHablar "Bienvenido " & stNom & "."
                End If
            End If
            .UpDate
            objVISIManual.idAccesoVISI = !id
            If objVISIManual.bObjetos Then
                sSql = "insert into tobjetos(idacceso,descripcion,serial,estado,foto)"
                sSql = sSql & " select " & objVISIManual.idAccesoVISI & ",tmp_objetos.descripcion,tmp_objetos.serial,1,tmp_objetos.foto from tmp_objetos"
                objCon.Execute (sSql)
                
'                sSql = "delete from tmp_objetos where idvisitante=" & objVISI.idVISI
'                objCon.Execute sSql
            End If
        ElseIf iDispMODO = 2 Then
            fnHablar "Este huellero solo registra la salida."
            Set objVISIManual = Nothing
            Exit Sub
        End If
    Else
        If objVISIManual.iDispMODO = 0 Or objVISIManual.iDispMODO = 2 Or objVISIManual.iDispMODO = 3 Then 'S,ES
            objVISIManual.bRegPUERTO = bPulsoS
            sSql = "select * from tobjetos where idacceso=" & objVISIManual.idAccesoVISI
            Set objRst = objCon.Execute(sSql)
            If Not objRst.EOF Then
                objVISIManual.bObjetos = True
                cmdObjetos_Click
            End If
            
            !sale = fnFecha(Now, True)
            If IsNull(!entra) Then !entra = !sale
            !idhuellero_sale = objVISIManual.idDisp
            !idlogin_s = idLogin
            !terminal_s = sTerminal
            fnHablar "Hasta pronto " & stNom & "."
        ElseIf objVISIManual.iDispMODO = 1 Then
            fnHablar "Este huellero solo registra la entrada."
            Set objVISIManual = Nothing
            Exit Sub
        End If
    End If
    .UpDate
    objVISIManual.idAccesoVISI = !id
    .Close
End With
If objVISIManual.bEntraVISI Then subSticker
If objVISIManual.bREG Then subRELEVO objVISIManual.bRegPUERTO
Set objVISIManual = Nothing
subLimpiar
End Sub
Sub subDatosVISI()
sSql = "select documento,nombre,apellidos,sexo,rh,foto,enrola,frecuente from tvisitantes_huella where id=" & Val(objVISI.idVISI)
Set objRst = objCon.Execute(sSql)
If Not objRst.EOF Then
    If objVISI.AccesoTipo = accMANUAL Then cmbTipoID.SetFocus
    objVISI.sDOC = "" & objRst!documento
    'If objVISI.AccesoTipo = accTARJETA Then txtDoc1.Text = objVISI.sDOC
    objVISI.sNOM = "" & objRst!nombre
    objVISI.sAPE = "" & objRst!apellidos
    objVISI.sSEXO = "" & objRst!sexo
    objVISI.sRH = "" & objRst!rh
    If bAutorizado Then
        objVISI.bFrecuente = True
        'bAutorizado = False
        sSql = "delete from tautoriza_acceso_vis where documento='" & objVISI.sDOC & "'"
        objCon.Execute sSql
    Else
        If Not IsNull(objRst!frecuente) Then objVISI.bFrecuente = objRst!frecuente Else objVISI.bFrecuente = False
    End If
    If objVISI.bFrecuente Or bAutoStmp Or bAuto Then
        sSql = "select max(id) as ult from tacceso where idtpersona=" & Val(objVISI.idVISI) & " and idtipoper=" & tpVISI
        Set objRstA = objCon.Execute(sSql)
        objVISI.idAccesoVISI = Val("" & objRstA!ult)
        
        sSql = "select entra,sale,espera,idempleado from tacceso where id=" & Val(objVISI.idAccesoVISI)
        Set objRst = objCon.Execute(sSql)
        If Not objRst.EOF Then
            objVISI.bEntraVISI = Not IsNull(objRst!sale)
        Else
            objVISI.bEntraVISI = True
        End If
        If objVISI.bEntraVISI Then
            sSql = "select * from vvisita_datos where id=" & objVISI.idAccesoVISI
            Set objRst = objCon.Execute(sSql)
            If Not objRst.EOF Then
                bBuscar = False
                txtEmpleado.Text = "" & objRst!empleado
                txtEmpleado.Tag = Val("" & objRst!idEmpleado)
                txtCompañia.Text = "" & objRst!compañia
                txtDepartamento.Text = "" & objRst!departamento
                txtLocalizacion.Text = "" & objRst!localizacion
                txtUbicacion.Text = "" & objRst!ubicacion
                txtOficina.Text = "" & objRst!oficina
                txtExtension.Text = "" & objRst!extension
                txtChat.Text = "" & objRst!usuario
                bBuscar = True
            End If
        End If
        subAccesoVISI
    Else
        Set objVISIManual = New clsDatosVISI
        objVISIManual.AccesoTipo = objVISI.AccesoTipo
        objVISIManual.idVISI = objVISI.idVISI
        objVISIManual.sDOC = objVISI.sDOC
        objVISIManual.sNOM = objVISI.sNOM
        objVISIManual.sAPE = objVISI.sAPE
        objVISIManual.sSEXO = objVISI.sSEXO
        objVISIManual.sRH = objVISI.sRH
        objVISIManual.bFrecuente = objVISI.bFrecuente
        objVISIManual.bREG = objVISI.bREG
        objVISIManual.bRegPUERTO = objVISI.bRegPUERTO
        If Not IsNull(objRst!enrola) Then bEnrola = objRst!enrola
        'objVISIManual.iDispMODO = objVISI.iDispMODO No es visitante frecuente Entra Manualmente
        'objVISIManual.idDisp = objVISI.idDisp No es visitante frecuente No utiliza dispositivo de entrada
        Set objVISI = Nothing
        If objVISIManual.AccesoTipo = accHUELLA Or objVISIManual.AccesoTipo = acc2D Then txtDoc1.Text = objVISIManual.sDOC
        sSql = "select idtipodoc,idtratamiento,email,telefono,idtorganizacion,foto,huella,extra,fechanac from tvisitantes_huella where id=" & objVISIManual.idVISI
        With objRst
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockReadOnly
            cmbTipoID.mostrarItem Val("" & !idtipodoc)
            cmbTratamiento.mostrarItem Val("" & !idtratamiento)
            txtNombre.Text = objVISIManual.sNOM
            txtApellidos.Text = objVISIManual.sAPE
            
            cmbSexo.mostrarItem fnBuscaSEXO(objVISIManual.sSEXO)
            cmbRH.mostrarItem fnBuscaRH(objVISIManual.sRH)
            
            sSql = "select nombre from torganizaciones where id=" & Val("" & !idtorganizacion)
            Set objRstA = objCon.Execute(sSql)
            If Not objRstA.EOF Then
                bBuscar = False
                txtOrganizacion.Text = "" & objRstA!nombre
                txtOrganizacion.Tag = !idtorganizacion
                bBuscar = True
            End If
            
            txtTelefono.Text = "" & !telefono
            txtEmail.Text = "" & !email
            If txtFechaNace.Text = vbNullString Then
                If Not IsNull(!fechanac) Then txtFechaNace.Text = fnFecha(CDate("" & !fechanac), False)
            End If
            If Not IsNull(!foto) Then
                bFoto = True
                tmrCamPreview.Enabled = False
                fnLeeFoto !foto, imgFoto1
            End If
            If Not IsNull(!huella) Then
                bHuella = True
                fnLeeFoto !huella, imgHuella
            End If
        
            chkFrecuente.Value = IIf(objVISIManual.bFrecuente, vbChecked, vbUnchecked)
            If IsNull(!extra) Then chkExtra.Value = vbUnchecked Else chkExtra.Value = IIf(!extra, vbChecked, vbUnchecked)
            
        
            sSql = "select max(id) as ult from tacceso where idtpersona=" & Val(objVISIManual.idVISI) & " and idtipoper=" & tpVISI
            Set objRstA = objCon.Execute(sSql)
            objVISIManual.idAccesoVISI = Val("" & objRstA!ult)
'''            If objVISIManual.idAccesoVISI = 0 Then
'''                sSql = "insert into tacceso(idtipoper,idtpersona) values (2," & objVISIManual.idVISI & ")"
'''                objCon.Execute sSql
'''                DoEvents
'''                sSql = "select max(id) as ult from tacceso where idtpersona=" & Val(objVISIManual.idVISI) & " and idtipoper=" & tpVISI
'''                Set objRstA = objCon.Execute(sSql)
'''                objVISIManual.idAccesoVISI = Val("" & objRstA!ult)
'''            End If
            sSql = "select entra,sale,espera,idempleado from tacceso where id=" & Val(objVISIManual.idAccesoVISI)
            Set objRst = objCon.Execute(sSql)
            If Not objRst.EOF Then
                objVISIManual.bEntraVISI = Not IsNull(objRst!sale)
                If FormatDateTime(objRst!entra, vbShortDate) = Date Then
                    If IsNull(objRst!espera) Then objVISIManual.bEspera = False Else objVISIManual.bEspera = objRst!espera
                End If
            Else
                objVISIManual.bEntraVISI = True
                objVISIManual.bEspera = False
            End If
        
            If objVISIManual.bEspera Then
                imgEspera.Visible = False
                objVISIManual.bEntraVISI = True
                txtEmpleado.Tag = Val("" & objRst!idEmpleado)
                sSql = "select * from tobjetos where idacceso=" & objVISIManual.idAccesoVISI
            Else
                imgEspera.Visible = objVISIManual.bEntraVISI
            End If
            If objVISIManual.bEntraVISI Then
                imgEntra.Picture = LoadPicture(App.Path & "\Ingreso.jpg")
                If bDatosVisAnterior Then
                    sSql = "select * from vvisita_datos where id=" & objVISIManual.idAccesoVISI
                    Set objRst = objCon.Execute(sSql)
                    If Not objRst.EOF Then
                        bBuscar = False
                        txtEmpleado.Text = "" & objRst!empleado
                        txtEmpleado.Tag = Val("" & objRst!idEmpleado)
                        txtCompañia.Text = "" & objRst!compañia
                        txtDepartamento.Text = "" & objRst!departamento
                        txtLocalizacion.Text = "" & objRst!localizacion
                        txtUbicacion.Text = "" & objRst!ubicacion
                        txtOficina.Text = "" & objRst!oficina
                        txtExtension.Text = "" & objRst!extension
                        txtChat.Text = "" & objRst!usuario
                        txtNombre.SetFocus
                        bBuscar = True
                    End If
                End If
            Else
                imgEntra.Picture = LoadPicture(App.Path & "\Salida.jpg")
            End If
            
            If objVISIManual.bEspera Or Not objVISIManual.bEntraVISI Then
                sSql = "select * from vvisita_datos where id=" & objVISIManual.idAccesoVISI
                Set objRst = objCon.Execute(sSql)
                If Not objRst.EOF Then
                    bBuscar = False
                    txtEmpleado.Text = "" & objRst!empleado
                    txtCompañia.Text = "" & objRst!compañia
                    txtDepartamento.Text = "" & objRst!departamento
                    txtLocalizacion.Text = "" & objRst!localizacion
                    txtUbicacion.Text = "" & objRst!ubicacion
                    txtOficina.Text = "" & objRst!oficina
                    txtExtension.Text = "" & objRst!extension
                    txtChat.Text = "" & objRst!usuario
                    txtNombre.SetFocus
                    bBuscar = True
                End If
                If Not objVISIManual.bEntraVISI Then
                    sSql = "select id,foto from tobjetos where idacceso=" & objVISIManual.idAccesoVISI & " and foto is not null"
                    With objRstO
                        If .State = adStateOpen Then .Close
                        .Open sSql, objCon, adOpenKeyset, adLockOptimistic
                        If Not objRstO.EOF Then subObjetos 0
                    End With
                End If
            Else
                txtEmpleado.SetFocus
            End If
        End With
    End If
End If
End Sub
Private Sub subObjetos(bDir As Byte)
If bDir = 1 Then
    If Not objRstO.BOF Then
        objRstO.MovePrevious
        If objRstO.BOF Then objRstO.MoveFirst
    Else
        objRstO.MoveFirst
    End If
ElseIf bDir = 2 Then
    If Not objRstO.EOF Then
        objRstO.MoveNext
        If objRstO.EOF Then objRstO.MoveLast
    Else
        objRstO.MoveLast
    End If
End If
fnLeeFoto objRstO!foto, imgObjeto
End Sub
Private Sub txtOrganizacion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEmpleado_txtCambio()
subLista txtEmpleado
End Sub



Private Sub txtOrganizacion_LostFocus()
lstEmerge.Visible = False
End Sub

Private Sub txtFicha_Change()
'lblTurno.Caption = txtFicha.Text
End Sub

Private Sub txtFicha_Validate(Cancel As Boolean)
''Dim sDoc As String, sCampo As String
''sDoc = Trim(txtFicha.Text)
''sFicha = vbNullString
''sPlaca = vbNullString
''If sDoc <> vbNullString Then
''    sDoc = Replace(sDoc, ".", "")
''    sDoc = Replace(sDoc, ",", "")
''    sDoc = Replace(sDoc, "-", "")
''    sFicha = sDoc
''    sCampo = "ficha"
''    subDatos sCampo, sDoc
''End If
End Sub



Private Sub txtEmpleado_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEmpleado_LostFocus()
lstEmerge.Visible = False
End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Sub subLimpiar()
On Local Error GoTo errH
Dim i As Byte
bEsperaTemp = False

TipoPER = 0


idEmpleado = 0
idDepartamento = 0
idAutoriza = 0
sFecha1 = vbNullString
sFecha2 = vbNullString
sTel = vbNullString

sAc(0) = vbNullString
sAc(1) = vbNullString
sAc(2) = vbNullString


txtDoc1.Text = vbNullString
cmbTipoID.itemID = 0
txtTarjetaNum.Text = vbNullString
cmbTratamiento.itemID = 0
txtNombre.Text = vbNullString
txtApellidos.Text = vbNullString
cmbSexo.itemID = 0
cmbRH.itemID = 0
txtOrganizacion.Text = vbNullString: txtOrganizacion.Tag = vbNullString
txtTelefono.Text = vbNullString
txtEmail.Text = vbNullString
txtFechaNace.Text = vbNullString

txtEmpleado.Text = vbNullString: txtEmpleado.Tag = vbNullString
txtCompañia.Text = vbNullString
txtDepartamento.Text = vbNullString
txtLocalizacion.Text = vbNullString
txtUbicacion.Text = vbNullString
txtOficina.Text = vbNullString
txtExtension.Text = vbNullString
txtChat.Text = vbNullString
chkFrecuente.Value = vbUnchecked
subListarTipoID
subListarTratamiento


fra2.Caption = vbNullString


bFoto = False

Set imgFoto1.Picture = Nothing
For i = 0 To imgFotos.Count - 1
    Set imgFotos(i).Picture = LoadPicture(App.Path & "\imgfoto.jpg")
    'imgFotos(I).Visible = False
Next i
iCicloFoto = 0
'---
txtOrganizacion.Text = vbNullString

txtEmpleado.Text = vbNullString
txtDoc1.SetFocus
'Set cmdEntra.Picture = LoadPicture(App.Path & "\ingreso.jpg")
subTurno

imgFoto1.Picture = LoadPicture(App.Path & "\imgFoto.jpg")
imgHuella.Picture = LoadPicture(App.Path & "\imgHuella.jpg")
imgEntra.Picture = LoadPicture(App.Path & "\Ingreso.jpg")
imgEspera.Visible = True

'''sSql = "delete from tmp_objetos where idvisitante is null"
'''objCon.Execute sSql

'cmdEntra.BackColor = Val(cmdEntra.Tag)
cmdObjetos.BackColor = Val(cmdObjetos.Tag)

sNOM = vbNullString
sAPE = vbNullString
bHuella = False
bModificaHuella = False
bModificaFoto = False
sTIPO_PDF = vbNullString
'subListarDependencias

imgObjeto.Picture = LoadPicture()
subContar
subEnEspera
If bCam Then tmrCamPreview.Enabled = True
iDispMODO = 0
idDisp = 0

AccesoTipo = 0
bPuertoREG = 0

bPulsoE = 0
bPulsoS = 0
sDocTmp = vbNullString
Set imgAlerta.Picture = LoadPicture()
chkExtra.Value = vbUnchecked

Erase bHuellaMinuciasCAP
Erase bHuellaMinucias
Erase bEnrola
DatosDD = DatosDD_
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-subLimpiar"
subLog sERR
End Sub
Private Sub subEnEspera()
picGrid.Visible = False
imgFlecha.Visible = False
sSql = "select Documento,Visitante,Empleado from ven_espera order by id"
With objRstA
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    If Not .EOF Then
        tmrEspera.Enabled = True
        Set objGrid.DataSource = objRstA
        objGrid.Columns("Documento").Visible = False
        objGrid.Columns("Visitante").Width = 3000
        objGrid.Columns("Empleado").Width = 3000
    Else
        tmrEspera.Enabled = False
    End If
End With
End Sub
Private Sub txtNovedad_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombre_Validate(Cancel As Boolean)
txtNombre.Text = fnMayúscula(txtNombre.Text)
End Sub

Private Sub txtOrganizacion_txtCambio()
subLista txtOrganizacion
End Sub

Private Sub txtOrganizacion_Validate(Cancel As Boolean)
txtOrganizacion.Text = UCase(txtOrganizacion.Text)
End Sub
Private Sub subRELEVO(pP As Long)
10    On Local Error GoTo errH
20    If Not objFUNC Is Nothing Then
30        If objFUNC.AccesoTipo = accHUELLA_ZK Then
'40            If pP = 1 Then
50                objZK(idxZK).ACUnlock objZK(idxZK).MachineNumber, 10
                subMonitor "{PULSO}{" & objFUNC.idFUNC & "}", False, False
'60            End If
70            idxZK = -1
80        Else
90            If objPhidget.IsAttached Then
100               objPhidget.OutputState(pP) = True
110               subEsperar 1
120               objPhidget.OutputState(pP) = False
130           End If
140       End If
Phidget:
150   ElseIf objPhidget.IsAttached Then
160       objPhidget.OutputState(pP) = True
170       subEsperar 1
180       objPhidget.OutputState(pP) = False
        Else
            If Not objVISI Is Nothing Then
                If objVISI.AccesoTipo = accHUELLA_ZK Then
                    objZK(pP).ACUnlock objZK(pP).MachineNumber, 10
                End If
            Else
                objZK(pP).ACUnlock objZK(pP).MachineNumber, 10
            End If
190   End If
200   Exit Sub
errH:
210   subLog "Error linea " & Erl & ":" & Err.Number & ". " & Err.Description & "_subRELEVO"
220
End Sub
Public Sub subLectura2D(ByRef Datos2D As A1A2D.Datos2D)
Dim tDoc As String
Dim objVisTmp As clsDatosVISI
AccesoTipo = acc2D
If modoBD = bdSQL Then
    sSql = "select d.id,d.modo,d.persona,d.activo,c.id as idc,c.puerto_e,c.puerto_s,c.activa,d.enrola_fun,d.enrola_vis,login"
    sSql = sSql & " from tcontrol_disp as d join tcontrol c on d.idcontrol=c.id"
    sSql = sSql & " where d.puerto='" & Datos2D.ccPuerto & "' and c.terminal='" & sTerminal & "'"
ElseIf modoBD = bdACCESS Then
    sSql = "SELECT tcontrol_disp.id, tcontrol_disp.modo, tcontrol_disp.persona, tcontrol_disp.activo, tcontrol.id AS idc, tcontrol.puerto_e, tcontrol.puerto_s, tcontrol.activa, "
    sSql = sSql & "tcontrol_disp.enrola_fun, tcontrol_disp.enrola_vis, tcontrol_disp.login "
    sSql = sSql & "FROM tcontrol_disp INNER JOIN tcontrol ON tcontrol_disp.idcontrol = tcontrol.id "
    sSql = sSql & "WHERE (((tcontrol_disp.puerto)='" & Datos2D.ccPuerto & "') AND ((tcontrol.terminal)='" & sTerminal & "'));"
End If
Set objRst = objCon.Execute(sSql)
If objRst.EOF Then
    fnHablar "Este lector no está configurado."
    Exit Sub
Else
    If objRst!activo = False Then
        fnHablar "Este lector está desactivado."
        Exit Sub
    Else
        If Not IsNull(objRst!idC) Then
            If Not objRst!activa Then
                bREG = False
                fnHablar "Portería desactivada."
                bPulsoE = 0
                bPulsoS = 0
                Exit Sub
            Else
                bREG = True
                idPuerta = objRst!idC
                idDisp = objRst!id
                iDispMODO = Val("" & objRst!modo)
                iDispPER = Val("" & objRst!persona)
                
                bPulsoE = Val("" & objRst!puerto_e)
                bPulsoS = Val("" & objRst!puerto_s)
                bDispEnrolaFun = objRst!enrola_fun
                bDispEnrolaVis = objRst!enrola_vis
                bDispLogin = objRst!login
            End If
        ElseIf Not objRst!activa Then
            bREG = False
            fnHablar "Portería desactivada."
            bPulsoE = 0
            bPulsoS = 0
            Exit Sub
        End If
    End If
End If
    
If Screen.ActiveForm.name = "frmFuncionarios" Then
        If Datos2D.ccTipo = Cédula_2D_V1 Or Datos2D.ccTipo = Cédula_2D_V2 Or Datos2D.ccTipo = A1A Or Datos2D.ccTipo = TI_2D Or Datos2D.ccTipo = PASE_2011 Or Datos2D.ccTipo = TPROPIEDAD_2011 Then
            If Datos2D.ccTipo = Cédula_2D_V1 Or Datos2D.ccTipo = Cédula_2D_V2 Or Datos2D.ccTipo = PASE_2011 Or Datos2D.ccTipo = TPROPIEDAD_2011 Then
                If bDispEnrolaFun Then
                    frmFuncionarios.txtDoc1.Text = Datos2D.ccNumero
                    frmFuncionarios.subDatos "documento", Datos2D.ccNumero
                    If frmFuncionarios.txtNombre.Text = vbNullString Then
                        frmFuncionarios.txtNombre.Text = Trim(Datos2D.ccNombre1 & " " & Datos2D.ccNombre2)
                        frmFuncionarios.txtApellidos.Text = Trim(Datos2D.ccApellido1 & " " & Datos2D.ccApellido2)
                        frmFuncionarios.txtRh.Text = Datos2D.ccRH
                        frmFuncionarios.txtSexo.Text = Datos2D.ccSexo
                        frmFuncionarios.txtFechaNac.Text = Datos2D.ccDiaNace & "/" & Datos2D.ccMesNace & "/" & Datos2D.ccAñosNace
                        'frmFuncionarios.cmbDependencias.SetFocus
                    End If
                Else
                    fnHablar "Este lector no está configurado para registrar funcionarios."
                End If
            
            End If
        End If
ElseIf Screen.ActiveForm.name = "frmObjetos" Then
    If bDispLogin Then 'Objetos
        If objVISI Is Nothing Then
            Set objVisTmp = objVISIManual
        Else
            Set objVisTmp = objVISI
        End If
        If Not objVisTmp Is Nothing Then
            If objVisTmp.bEntraVISI Then
                frmObjetos.txtSerial.Text = UCase(Datos2D.ccNumero)
                frmObjetos.txtDesc.SetFocus
                'frmObjetos.subAgregar
            Else
                sSql = "update tobjetos set estado=0 where idacceso=" & Val(objVisTmp.idAccesoVISI) & " and serial='" & UCase(Datos2D.ccNumero) & "'"
                objCon.Execute sSql
                sSql = "select Descripcion,Serial,Estado from tobjetos where idacceso=" & Val(objVisTmp.idAccesoVISI) & " and abs(estado)=1"
                frmObjetos.subCargarGrid sSql
                If frmObjetos.objRs_.EOF Then
                    objVisTmp.bObjetos = False
                    Unload frmObjetos
                End If
            End If
        End If
    End If
Else
    If Datos2D.ccTipo = Código_1D Or Datos2D.ccTipo = Cédula_2D_V1 Or Datos2D.ccTipo = Cédula_2D_V2 Or Datos2D.ccTipo = A1A Or Datos2D.ccTipo = TI_2D Or Datos2D.ccTipo = PASE_2011 Or Datos2D.ccTipo = TPROPIEDAD_2011 Then
        If Datos2D.ccTipo = A1A Then
            sTIPO_PDF = UBound(Datos2D.ccA1A)
            sTIPO_PDF = Datos2D.ccA1A(Val(sTIPO_PDF))
            If Len(sTIPO_PDF) = 2 Then
                sTIPO_PDF = Mid(sTIPO_PDF, 2)
            End If
            tDoc = Datos2D.ccA1A(0)
        Else
            tDoc = Datos2D.ccNumero
        End If
        DatosDD = Datos2D
        subAccesoTipo tDoc
    End If
End If
'Datos2D.ccNumero = vbNullString
'Datos2D.ccNombre1 = vbNullString
'Datos2D.ccNombre2 = vbNullString
'Datos2D.ccApellido1 = vbNullString
'Datos2D.ccApellido2 = vbNullString
'Datos2D.ccSexo = vbNullString
'Datos2D.ccTipo = 0

End Sub
Private Function fnConectaCFG() As Boolean
Dim i As Integer
On Local Error GoTo errH
fnConectaCFG = False
If Not objGAATools.fnExisteArchivo(App.Path & "\conectar.txt") Then
    objConCFG.CursorLocation = adUseClient
    objConCFG.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\A1AAppConfig.accdb;Jet OLEDB:Database Password=A1AAppconfig.2012;"
    fnConectaCFG = True
    bSinAccess = False
Else
    i = FreeFile
    Open App.Path & "\conectar.txt" For Input As #i
    Line Input #i, sBD
    Close #i
    If sBD <> vbNullString Then
        bSinAccess = True
        fnConectaCFG = True
        modoBD = bdSQL
    Else
        bSinAccess = False
    End If
End If
Exit Function
errH:
subLog "Error " & Err.Number & ". " & Err.Description & "_fnConectaCFG"
fnConectaCFG = False
End Function

Private Sub wSerDatos_Close()
wSerDatos.Close
DoEvents
wSerDatos.Listen
DoEvents
End Sub

Private Sub wSerDatos_ConnectionRequest(ByVal requestID As Long)
wSerDatos.Close
DoEvents
wSerDatos.Accept requestID
DoEvents
End Sub

Private Sub wSerDatos_DataArrival(ByVal bytesTotal As Long)
Dim sRec As String
Dim sArr() As String
sRec = ""
wSerDatos.GetData sRec
sArr = Split(sRec, "|")
If UBound(sArr()) > 0 Then
    Select Case sArr(0)
        Case "evtar"
            'txtTarjetaNum.Text = sArr(1)
            'Validar estado persona
            'Enviar estado
            DoEvents
            bREG = True
            iDispMODO = Val(sArr(2))
            wSerDatos.SendData "ok|0"
            subDatosTar sArr(1)
    End Select
End If
End Sub
Sub subDatosTar(sTar As String)
Dim objRs_ As New ADODB.Recordset
'Visitantes
'sSql = "select documento from tvisitantes_huella where tarjeta='" & sTar & "'"
'Set objRs_ = objCon.Execute(sSql)
'If Not objRs_.EOF Then
'    AccesoTipo = accTARJETA
'    subAccesoTipo objRs_!documento
'End If

'Empleados
sSql = "select documento from templeados where tarjeta_num='" & sTar & "'"
Set objRs_ = objCon.Execute(sSql)
If Not objRs_.EOF Then
    AccesoTipo = accTARJETA
    subAccesoTipo objRs_!documento
End If


End Sub
'Private Sub wServidor_Close()
'wServidor.Close
'wServidor.Listen
'End Sub

'Private Sub wServidor_ConnectionRequest(ByVal requestID As Long)
'wServidor.Close
'wServidor.Accept requestID
'End Sub

Private Sub wSerDatos_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wSerDatos.Close
DoEvents
wSerDatos.Listen
DoEvents
subLog Me.name & "_wSerDatos_Error_" & Description
End Sub

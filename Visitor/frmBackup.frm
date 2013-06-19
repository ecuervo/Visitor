VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.1#0"; "ARBUTTON.OCX"
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#19.0#0"; "A1AControles.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmBackup 
   BackColor       =   &H00F8D88F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00F8D88F&
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      Begin MSComDlg.CommonDialog dlgRuta 
         Left            =   1440
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin A1AControles.A1ATextBox txtRuta 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   16308367
         passChar        =   ""
      End
      Begin ARButtonCtrl.ARButton cmdAceptar 
         Default         =   -1  'True
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   767
         Caption         =   "&Aceptar"
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
      Begin ARButtonCtrl.ARButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   4560
         TabIndex        =   4
         Top             =   1080
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   767
         Caption         =   "&Cancelar"
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
      Begin ARButtonCtrl.ARButton cmdVer 
         Height          =   315
         Left            =   5640
         TabIndex        =   5
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Caption         =   "..."
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
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta"
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
         TabIndex        =   2
         Top             =   240
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objRuta As New vbalFolderBrowse6.cBrowseForFolder
Dim sRuta As String

Private Sub cmdAceptar_Click()
Dim sBKP As String
On Local Error GoTo errH
Screen.MousePointer = vbHourglass
If objCon.State = adStateOpen Then objCon.Close
DoEvents
sBKP = "_BKP_" & Right("00" & Day(Date), 2) & Right("00" & Month(Date), 2) & Year(Date)
sBKP = sBKP & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2)
sBKP = sBKP & ".accdb"

FileCopy App.Path & "\A1ABIOIDTAC.accdb", sRuta & "A1ABIOIDTAC" & sBKP
DoEvents
Screen.MousePointer = vbNormal
MsgBox "Copia realizada con éxito!", vbInformation
If Not fnConecta Then
    MsgBox "Error al reintentar abrir la base de datos. Reinicie la aplicación."
End If
Unload Me
Exit Sub
errH:
Screen.MousePointer = vbNormal
MsgBox "Error " & Err.Number & ". " & Err.Description
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdVer_Click()
objRuta.Title = "Guardar copia de seguridad en..."
objRuta.UseNewUI = True
sRuta = objRuta.BrowseForFolder
txtRuta.Text = sRuta
End Sub

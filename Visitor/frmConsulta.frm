VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.1#0"; "ARBUTTON.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmConsultas 
   BackColor       =   &H00A46B2E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consultas"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10710
   ControlBox      =   0   'False
   Icon            =   "frmConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   10710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00A46B2E&
      Caption         =   "Historial de Accesos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10455
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
         TabIndex        =   5
         Top             =   1380
         Width           =   7095
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
         Left            =   210
         TabIndex        =   0
         Top             =   780
         Width           =   2895
      End
      Begin ARButtonCtrl.ARButton cmdConsultas 
         Default         =   -1  'True
         Height          =   435
         Left            =   3360
         TabIndex        =   1
         Top             =   720
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   767
         Caption         =   "&Consultar"
         ForeColor       =   16777215
         ForeColorOnMouse=   12484943
         BackColorOnMouse=   16777215
         BackColor       =   12484943
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
         Arrow           =   1
      End
      Begin MSDataGridLib.DataGrid objGrid 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   10775342
         ForeColor       =   12632256
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
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   6
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   7260
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   3060
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
      Begin VB.Image imgCerrar 
         Height          =   360
         Left            =   10080
         Picture         =   "frmConsulta.frx":70E2
         Top             =   0
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultas_Click()
Dim sD As String, idPer As Variant
sD = Trim(txtDoc1.Text)
If sD <> vbNullString Then
    sD = Replace(sD, ".", "")
    sD = Replace(sD, ",", "")
    sD = Replace(sD, "-", "")
    
    sSql = "select id,nombre,apellidos from tvisitantes_huella where documento='" & sD & "'"
    Set objRst = objCon.Execute(sSql)
    If Not objRst.EOF Then
        txtNombre.Text = "" & objRst!nombre & " " & objRst!apellidos
        idPer = objRst!Id
        sSql = "select a.entra,a.sale,f.nombre + ' ' + f.apellidos Visita_a"
        sSql = sSql & " from tacceso a left join templeados f on a.idEmpleado=f.id"
        sSql = sSql & " where a.idtpersona = " & idPer
        With objRst
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenKeyset, adLockReadOnly
            If Not .EOF Then
                Set objGrid.DataSource = objRst
            End If
            
        End With
    Else
        MsgBox "No encontrado!", vbInformation
        txtDoc1.SetFocus
    End If
End If
End Sub

Private Sub imgCerrar_Click()
Unload Me
End Sub

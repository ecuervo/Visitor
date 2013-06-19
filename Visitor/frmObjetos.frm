VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmObjetos 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00909890&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Objetos"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00909890&
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtSerial 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Top             =   660
         Width           =   4935
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1500
         Width           =   4935
      End
      Begin ARButtonCtrl.ARButton cmdAceptar 
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Top             =   4320
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
         Height          =   435
         Left            =   5040
         TabIndex        =   7
         Top             =   4320
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
      Begin MSDataGridLib.DataGrid objGrid 
         Height          =   2295
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   6255
         _ExtentX        =   11033
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
      Begin ARButtonCtrl.ARButton cmdSalida 
         Height          =   435
         Left            =   2640
         TabIndex        =   9
         Top             =   4320
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   767
         Caption         =   "&Salida"
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
      Begin ARButtonCtrl.ARButton cmbFoto 
         Height          =   915
         Left            =   5280
         TabIndex        =   10
         Top             =   368
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1614
         Caption         =   ""
         ForeColor       =   16777215
         ForeColorOnMouse=   12484943
         BackColorOnMouse=   16777215
         BackColor       =   14737632
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
         Style           =   1
         Picture         =   "frmObjetos.frx":0000
      End
      Begin ARButtonCtrl.ARButton cmdAgregar 
         Height          =   435
         Left            =   5280
         TabIndex        =   2
         Top             =   1440
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "&Agregar"
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
      Begin VB.Shape Shape 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   1
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   5100
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de serie:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1800
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   435
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   5100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmObjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSerial As String
Public objRs_ As New ADODB.Recordset
Public bFotoO As Boolean
Dim objVTemp As clsDatosVISI
Private Sub cmbFoto_Click()
frmFotoObjeto.Show vbModal
End Sub

Private Sub cmdAceptar_Click()
objVISIManual.bObjetos = True
Unload Me
End Sub

Private Sub cmdAgregar_Click()
subAgregar
End Sub
Public Sub subAgregar()
If Trim(txtSerial.Text) = vbNullString Then
    MsgBox "Ingrese Serial", vbInformation
    txtSerial.SetFocus
    Exit Sub
End If
sSql = "select * from tmp_objetos where idvisitante=" & Val(objVTemp.idVISI)
With objRs_
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    .AddNew
    !descripcion = Trim(txtDesc.Text)
    !serial = Trim(txtSerial.Text)
    If bFotoO Then
        fnGuardaFoto !foto, App.Path & "\tmpFotoO.jpg"
    End If
    !idvisitante = Val(objVTemp.idVISI)
    .UpDate
    .Close
End With
objVISIManual.bObjetos = True
sSql = "select Serial,Descripcion from tmp_objetos where idvisitante=" & Val(objVTemp.idVISI)
subCargarGrid sSql
txtDesc.Text = vbNullString
txtSerial.Text = vbNullString
txtSerial.SetFocus
bFotoO = False
End Sub
Public Sub subCargarGrid(sq As String)
With objRs_
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .Open sSql, objCon, adOpenKeyset, adLockReadOnly
    Set objGrid.DataSource = objRs_
    objGrid.Columns("Descripcion").Width = 4200
End With
End Sub
Private Sub cmdCancelar_Click()
If objVTemp.idAccesoVISI = 0 Then
    sSql = "delete from tmp_objetos"
    objCon.Execute sSql
    objVTemp.bObjetos = False
End If
Unload Me
End Sub

Private Sub cmdSalida_Click()
If sSerial <> vbNullString Then
    sSql = "update tobjetos set estado=0 where idacceso=" & Val(objVTemp.idAccesoVISI) & " and serial='" & sSerial & "'"
    objCon.Execute sSql
    sSql = "select Descripcion,Serial,Estado from tobjetos where idacceso=" & Val(objVTemp.idAccesoVISI) & " and abs(estado)=1"
    subCargarGrid sSql
    If objRs_.EOF Then
        objVTemp.bObjetos = False
        Set objVTemp = Nothing
        Unload frmObjetos
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
If Not objVISIManual Is Nothing Then
    Set objVTemp = objVISIManual
ElseIf Not objVISI Is Nothing Then
    Set objVTemp = objVISI
End If

If objVTemp.bObjetos Then
    If Val(objVTemp.idAccesoVISI) = 0 Then
        Me.Caption = "Registrar Objetos"
        cmdSalida.Visible = False
        cmdAgregar.Visible = True
        sSql = "select Serial,Descripcion from tmp_objetos"
        subCargarGrid sSql
    Else
        Me.Caption = "Descargar Objetos"
        cmdAceptar.Visible = False
        cmdSalida.Visible = True
        cmdAgregar.Visible = False
        sSql = "select Serial,Descripcion,Estado from tobjetos where idacceso=" & Val(objVTemp.idAccesoVISI) & " and abs(estado)=1"
        subCargarGrid sSql
    End If
Else
    Me.Caption = "Registrar Objetos"
    cmdSalida.Visible = False
    cmdAgregar.Visible = True
End If

End Sub

Private Sub objGrid_Click()
On Error Resume Next
sSerial = objGrid.Columns("serial").Text
End Sub

Private Sub txtSerial_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

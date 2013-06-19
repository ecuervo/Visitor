VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{FB7524E1-AB42-4FE8-97C9-430FA6439280}#20.1#0"; "A1AControles.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportes 
   BackColor       =   &H00F8D88F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9285
   ControlBox      =   0   'False
   Icon            =   "frmReportes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9285
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00F8D88F&
      Caption         =   "Reportes generales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   4815
      Left            =   5160
      TabIndex        =   18
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox chkXLS 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D88F&
         Caption         =   "Generar XLS para nómina"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Frame fraXLS 
         BackColor       =   &H00F8D88F&
         Caption         =   "Periodo:"
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
         Height          =   1335
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Visible         =   0   'False
         Width           =   3735
         Begin MSComCtl2.DTPicker dXls1 
            Height          =   300
            Left            =   840
            TabIndex        =   22
            Top             =   337
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   0
            Format          =   82968577
            CurrentDate     =   40689
         End
         Begin MSComCtl2.DTPicker dXls2 
            Height          =   300
            Left            =   840
            TabIndex        =   23
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   0
            Format          =   82968577
            CurrentDate     =   40689
         End
         Begin ARButtonCtrl.ARButton cmdXLS 
            Height          =   495
            Left            =   2400
            TabIndex        =   27
            Top             =   360
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   873
            Caption         =   "Generar"
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   630
         End
      End
      Begin ARButtonCtrl.ARButton cmdEmpleados 
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   873
         Caption         =   "Empleados"
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
      Begin ARButtonCtrl.ARButton cmdLosMas 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   873
         Caption         =   "Departamentos mas visitado"
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
      Begin ARButtonCtrl.ARButton cmdLosMas1 
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   873
         Caption         =   "Persona mas visitada"
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
      Begin ARButtonCtrl.ARButton cmdObjetos 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   873
         Caption         =   "Objetos"
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
      Begin VB.Image imgCerrar 
         Height          =   315
         Left            =   3600
         MouseIcon       =   "frmReportes.frx":70E2
         MousePointer    =   99  'Custom
         Picture         =   "frmReportes.frx":73EC
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F8D88F&
      Caption         =   "Planilla de Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A95900&
      Height          =   4815
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4935
      Begin MSComDlg.CommonDialog dlgXLS 
         Left            =   4320
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin A1AControles.A1ATextBox txtDoc 
         Height          =   315
         Left            =   3240
         TabIndex        =   19
         Top             =   1305
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Text            =   ""
         bkColor         =   16308367
         passChar        =   ""
      End
      Begin VB.OptionButton opTs 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D88F&
         Caption         =   "Empleados (agrupado)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A95900&
         Height          =   270
         Index           =   0
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.OptionButton opTs 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D88F&
         Caption         =   "Visitantes menores de edad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A95900&
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   3495
      End
      Begin VB.OptionButton opTs 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D88F&
         Caption         =   "Visitantes con Antecedentes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A95900&
         Height          =   270
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   3495
      End
      Begin VB.OptionButton opTs 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D88F&
         Caption         =   "Visitantes Listado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A95900&
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   2655
      End
      Begin VB.OptionButton opTs 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D88F&
         Caption         =   "Visitantes detallado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A95900&
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton optRetardos 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D88F&
         Caption         =   "Registro Detallado Acceso Empleados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A95900&
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   4455
      End
      Begin VB.Frame fraFecha 
         BackColor       =   &H00F8D88F&
         Caption         =   "Periodo:"
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
         Height          =   855
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   4695
         Begin MSComCtl2.DTPicker fDesde 
            Height          =   300
            Left            =   840
            TabIndex        =   7
            Top             =   337
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   0
            Format          =   82968577
            CurrentDate     =   40689
         End
         Begin MSComCtl2.DTPicker fHasta 
            Height          =   300
            Left            =   3120
            TabIndex        =   8
            Top             =   337
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   0
            Format          =   82968577
            CurrentDate     =   40689
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   255
            Left            =   2400
            TabIndex        =   17
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A95900&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   705
         End
      End
      Begin VB.OptionButton opTs 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D88F&
         Caption         =   "Empleados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A95900&
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin ARButtonCtrl.ARButton cmdVer 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   4320
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   661
         Caption         =   "&Ver reporte..."
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
      Begin VB.Label lblId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID N°:"
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
         TabIndex        =   20
         Top             =   1335
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim optT As Byte
Dim idLista As Byte
Dim bXLS As Boolean
Private Sub cmbLista_Click()
If cmbLista.ListIndex <> -1 Then idLista = cmbLista.ItemData(cmbLista.ListIndex)
End Sub

Private Sub cmbRepDetalle_Click()
End Sub

Private Sub chkXLS_Click()
fraXLS.Visible = (chkXLS.Value = vbChecked)
End Sub

Private Sub cmdLosMas_Click()
Load frmMes
frmMes.Show vbModal
sSql = "select Departamento,COUNT(id) as cuenta from vrpt_losmas "
If frmMes.sAño <> "Todos" Then
    sSql = sSql & "where año=" & frmMes.sAño
    If frmMes.bMes <> 13 Then
        sSql = sSql & " and mes=" & frmMes.bMes
    End If
End If
frmMes.sAño = vbNullString
frmMes.bMes = 0
Unload frmMes
sSql = sSql & " group by Departamento order by 2 desc"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    subReporte App.Path & "\Reportes\rptLosmas.rpt", "Departamentos mas visitados", 1
End With
End Sub

Private Sub cmdEmpleados_Click()
sSql = "select * from vempleados order by nombre,apellidos"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    subReporte App.Path & "\Reportes\rptEmpleados.rpt", "Empleados", 1
End With
End Sub

Private Sub cmdLosMas1_Click()
Load frmMes
frmMes.Show vbModal
sSql = "select Empleado,COUNT(id) as cuenta from vrpt_losmas "
If frmMes.sAño <> "Todos" Then
    sSql = sSql & "where año=" & frmMes.sAño
    If frmMes.bMes <> 13 Then
        sSql = sSql & " and mes=" & frmMes.bMes
    End If
End If
frmMes.sAño = vbNullString
frmMes.bMes = 0
Unload frmMes
sSql = sSql & " group by Empleado order by 2 desc"

With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    subReporte App.Path & "\Reportes\rptLosmas1.rpt", "Persona mas visitada", 1
End With

End Sub

Private Sub cmdObjetos_Click()
Load frmMes
frmMes.Show vbModal
sSql = "select * from vrpt_objetos "
If frmMes.sAño <> "Todos" Then
    sSql = sSql & "where año=" & frmMes.sAño
    If frmMes.bMes <> 13 Then
        sSql = sSql & " and mes=" & frmMes.bMes
    End If
End If
frmMes.sAño = vbNullString
frmMes.bMes = 0
Unload frmMes
sSql = sSql & " order by id desc"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    subReporte App.Path & "\Reportes\rptObjetos.rpt", "Objetos Registrados", 2
End With

End Sub

Private Sub cmdVer_Click()
If optRetardos.Value = True Then
    Dim objR As New ADODB.Recordset
    sSql = "delete from tmpretardos"
    objCon.Execute sSql
    With objR
        If modoBD = bdSQL Then
            sSql = "select distinct e.documento as documento "
            sSql = sSql & "from templeados e join tacceso a on e.id=a.idtpersona "
            sSql = sSql & "where a.idtipoper=1 and abs(e.activo)=1 and convert(date,a.entra) between '" & fnFecha(fDesde.Value, False) & "' and '" & fnFecha(fHasta.Value, False) & "'"
        ElseIf modoBD = bdACCESS Then
            sSql = "select documento from vacceso_e1 where fecha between #" & fnFecha(fDesde.Value, False) & "# and #" & fnFecha(fHasta.Value, False) & "# group by documento"
        End If
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenForwardOnly
        While Not .EOF
            armaRetardos fnFecha(fDesde.Value, False), fnFecha(fHasta.Value, False), !documento
            'armaRetardos "" & !documento
            .MoveNext
        Wend
    End With
    
    
    sSql = "select * from vrptretardos order by fecha desc"
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenKeyset, adLockOptimistic
        'If modoBD = bdACCESS Then
            subReporte App.Path & "\Reportes\rptaccesofuncionarios.rpt", "Acceso Funcionarios ", 2
        'Else
        '    subReporte App.Path & "\Reportes\rptaccesofuncionariossql.rpt", "Acceso Funcionarios ", 2
        'End If
    End With
Else
    If modoBD = bdACCESS Then
        If optT = 0 Then
            sSql = "select * from vacceso_empleados where fecha between #" & fnFecha(fDesde.Value, False) & "# and #" & fnFecha(fHasta.Value, False) & "# order by nombre"
        ElseIf optT = 1 Then
            sSql = "select * from vacceso_empleados where fecha between #" & fnFecha(fDesde.Value, False) & "# and #" & fnFecha(fHasta.Value, False) & "#"
        ElseIf optT = 2 Then
            sSql = "select * from vacceso_visitantes where fecha between #" & fnFecha(fDesde.Value, False) & "# and #" & fnFecha(fHasta.Value, False) & "#"
            If Trim(txtDoc.Text) <> vbNullString Then
                sSql = sSql & " and documento='" & Trim(txtDoc.Text) & "'"
            End If
        ElseIf optT = 3 Then
            sSql = "select distinct Documento,Nombre,Sexo,rh,Telefono,email from v_lista2 where fecha between #" & fnFecha(fDesde.Value, False) & "# and #" & fnFecha(fHasta.Value, False) & "#"
        ElseIf optT = 4 Then
            sSql = "delete * from tanota_temp"
            objCon.Execute sSql
            sSql = "insert into tanota_temp (documento) select documento from tanotaciones0 where (((tanotaiones0.fecha) between #" & fnFecha(fDesde.Value, False) & "# and #" & fnFecha(fHasta.Value, False) & "#))"
            objCon.Execute sSql
            sSql = "select * from v_anotaciones"
        ElseIf optT = 5 Then
            sSql = "select distinct Documento,Nombre,Sexo,rh,Telefono,Email from v_menores where fecha between #" & fnFecha(fDesde.Value, False) & "# and #" & fnFecha(fHasta.Value, False) & "#"
        End If
    ElseIf modoBD = bdSQL Then
        If optT = 0 Then
            sSql = "select * from vacceso_empleados where fecha between '" & fnFecha(fDesde.Value, False) & "' and '" & fnFecha(fHasta.Value, False) & "' order by nombre"
        ElseIf optT = 1 Then
            sSql = "select * from vacceso_empleados where fecha between '" & fnFecha(fDesde.Value, False) & "' and '" & fnFecha(fHasta.Value, False) & "'"
        ElseIf optT = 2 Then
            sSql = "select * from vacceso_visitantes where fecha between '" & fnFecha(fDesde.Value, False) & "' and '" & fnFecha(fHasta.Value, False) & "'"
            If Trim(txtDoc.Text) <> vbNullString Then
                sSql = sSql & " and documento='" & Trim(txtDoc.Text) & "'"
            End If
        ElseIf optT = 3 Then
            sSql = "select a.*,v.Documento,ltrim(rtrim(isnull(v.nombre,'') + ' ' + isnull(v.apellidos,''))) Nombre, "
            sSql = sSql & "v.Sexo,v.rh, v.Telefono,v.email from ( "
            sSql = sSql & "select distinct idtpersona from tacceso where idtipoper=2 and convert(date,entra) between convert(date,'" & fnFecha(fDesde.Value, False) & "') and convert(date,'" & fnFecha(fHasta.Value, False) & "') "
            sSql = sSql & ")a join tvisitantes_huella v on a.idtpersona=v.id order by nombre"
        ElseIf optT = 4 Then
'            sSql = "select v.Documento,ltrim(rtrim(isnull(v.nombre,'') + ' ' + isnull(v.apellidos,''))) Nombre,"
'            sSql = sSql & "v.Sexo,v.rh, v.Telefono,v.email from (select distinct documento from tanotaciones "
'            sSql = sSql & ")a join tvisitantes_huella v on a.documento=v.documento order by nombre"
            
            sSql = "select v.Documento,ltrim(rtrim(isnull(v.nombre,'') + ' ' + isnull(v.apellidos,''))) Nombre,"
            sSql = sSql & "v.Sexo,v.rh, v.Telefono,v.email from (select distinct documento from tanotaciones "
            sSql = sSql & "where convert(date,fecha_hora) between convert(date,'" & fnFecha(fDesde.Value, False) & "') and convert(date,'" & fnFecha(fHasta.Value, False) & "')"
            sSql = sSql & ")a join tvisitantes_huella v on a.documento=v.documento order by nombre"
            
        ElseIf optT = 5 Then
            sSql = "select a.*,v.Documento,ltrim(rtrim(isnull(v.nombre,'') + ' ' + isnull(v.apellidos,''))) Nombre, "
            sSql = sSql & "v.Sexo,v.rh, v.Telefono,v.email from ( "
            sSql = sSql & "select distinct idtpersona from tacceso where idtipoper=2 and convert(date,entra) between convert(date,'" & fnFecha(fDesde.Value, False) & "') and convert(date,'" & fnFecha(fHasta.Value, False) & "') "
            sSql = sSql & ")a join tvisitantes_huella v on a.idtpersona=v.id where v.idtipodoc=2 order by nombre"
        End If
    End If
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    End With
    If optT = 0 Then
        subReporte App.Path & "\Reportes\rptPersona2.rpt", "Planilla de Control " & opTs(optT).Caption & " período " & fDesde.Value & "-" & fHasta.Value, 2
    ElseIf optT = 1 Then
        subReporte App.Path & "\Reportes\rptPersona.rpt", "Planilla de Control " & opTs(optT).Caption & " período " & fDesde.Value & "-" & fHasta.Value, 2
    ElseIf optT = 2 Then
        subReporte App.Path & "\Reportes\rptPersona1.rpt", "Planilla de Control " & opTs(optT).Caption & " período " & fDesde.Value & "-" & fHasta.Value, 2
    ElseIf optT = 3 Then
        If modoBD = bdSQL Then
            subReporte App.Path & "\Reportes\visitantes_lista.rpt", "Visitantes período " & fDesde.Value & "-" & fHasta.Value, 1
        ElseIf modoBD = bdACCESS Then
            subReporte App.Path & "\Reportes\visitantes_lista_acc.rpt", "Visitantes período " & fDesde.Value & "-" & fHasta.Value, 1
        End If
    ElseIf optT = 4 Then
        If modoBD = bdSQL Then
            subReporte App.Path & "\Reportes\visitantes_lista.rpt", "Visitantes con Antecedentes", 1
        ElseIf modoBD = bdACCESS Then
            subReporte App.Path & "\Reportes\visitantes_lista_acc.rpt", "Visitantes con Antecedentes", 1
        End If
    ElseIf optT = 5 Then
        If modoBD = bdSQL Then
            subReporte App.Path & "\Reportes\visitantes_lista.rpt", "Visitantes menores de edad período " & fDesde.Value & "-" & fHasta.Value, 1
        ElseIf modoBD = bdACCESS Then
            subReporte App.Path & "\Reportes\visitantes_lista_acc.rpt", "Visitantes menores de edad período " & fDesde.Value & "-" & fHasta.Value, 1
        End If
    End If
End If
End Sub
Private Sub subReporte(sRep As String, sTit As String, bOR As Byte)
Set objReporte = New clsCrystal10
'objCrystal.SetLicenseKeycode "AV860-01CS00G-U7000NC"
With objReporte
    .bREG = True
    '.modoForm
    .sReporte = sRep
    Set objReporte.oRecordset = objRst
    .repTítulo = sTit
    If bXLS = False Then
        If objGAATools.fnExisteArchivo(App.Path & "\sulogo") Then
            .sLogo = App.Path & "\sulogo"
        End If
    Else
        objReporte.aXLS = dlgXLS.FileName
        bXLS = False
    End If
    .VerReporte False, idImpresoraR, bOR, False
End With
End Sub

Private Sub cmdXLS_Click()
On Local Error GoTo errH
Dim objEx As New ADODB.Recordset
Dim objEx1 As New ADODB.Recordset
Dim sExtras As String
Dim sArr() As String
Dim sArr1() As String
Dim ix As Integer
If dXls2.Value >= dXls1 Then
    sSql = "if exists(select name from sysobjects where name='thextras_rpt') drop table thextras_rpt"
    objCon.Execute sSql
    DoEvents
    sSql = "create table thextras_rpt(idempleado int,concepto int,cant numeric(6,2))"
    objCon.Execute sSql
    DoEvents
    sSql = "select distinct(idtpersona) id from v_hextras_fin"
    Set objEx = objCon.Execute(sSql)
    While Not objEx.EOF
        sSql = "select * from v_hextras_fin where idtpersona=" & objEx!id & " and fecha between convert(date,'" & fnFecha(CDate(dXls1.Value), False) & "') and convert(date,'" & fnFecha(CDate(dXls2.Value), False) & "')"
        With objEx1
            If .State = adStateOpen Then .Close
            .Open sSql, objCon, adOpenForwardOnly
            While Not .EOF
                sExtras = "" & !hextras
                If sExtras <> "0" Then
                    sArr = Split(sExtras, "|")
                    For ix = 1 To UBound(sArr)
                        sArr1 = Split(sArr(ix), "°")
                        sSql = "insert into thextras_rpt values(" & objEx!id & "," & sArr1(1) & "," & sArr1(0) & ")"
                        objCon.Execute (sSql)
                    Next ix
                End If
                .MoveNext
            Wend
            .Close
        End With
        objEx.MoveNext
    Wend
    dlgXLS.CancelError = True
    dlgXLS.DialogTitle = "Exportar archivo en..."
    dlgXLS.Filter = "Archivos xls|*.xls"
    dlgXLS.ShowSave
    
    bXLS = True
    sSql = "select codigo,concepto,sum(cant) cant from vhextras_rpt group by codigo,concepto"
    Set objRst = objCon.Execute(sSql)
    subReporte App.Path & "\Reportes\rptnomina.rpt", "", 1
End If
Exit Sub
errH:
End Sub

Private Sub Form_Load()
fDesde.Value = Date
fHasta.Value = Date
dXls1.Value = Date
dXls2.Value = Date
End Sub

Private Sub imgCerrar_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub subOficinas()
fraFecha.Visible = False
lblT.Visible = False
cmbLista.Visible = False
imgCampo.Visible = False
txtPlaca.Visible = False
End Sub
Private Sub subAcompañantes()
lblT.Caption = "Placa:"
cmbLista.Visible = False
fraFecha.Visible = False
txtPlaca.Move cmbLista.Left, cmbLista.Top
txtPlaca.Text = vbNullString
txtPlaca.Visible = True
txtPlaca.SetFocus
End Sub
Private Sub subCargaPersonas()
lblT.Caption = "Seleccione..."
cmbLista.Visible = True
fraFecha.Visible = True
txtPlaca.Visible = False
cmbLista.Clear
cmbLista.AddItem "Funcionarios"
cmbLista.ItemData(cmbLista.NewIndex) = 1
cmbLista.AddItem "Visitantes"
cmbLista.ItemData(cmbLista.NewIndex) = 2

End Sub
Private Sub subCargaVehiculos()
lblT.Caption = "Seleccione..."
cmbLista.Visible = True
fraFecha.Visible = True
txtPlaca.Visible = False
cmbLista.Clear
sSql = "select * from ttipo_ve order by id"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        cmbLista.AddItem !nombre
        cmbLista.ItemData(cmbLista.NewIndex) = !id
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub optRetardos_Click()
'fraFecha.Visible = False
End Sub

Private Sub opTs_Click(Index As Integer)
optT = Index
fraFecha.Visible = True
txtDoc.Visible = (Index = 2)
lblId.Visible = (Index = 2)
End Sub

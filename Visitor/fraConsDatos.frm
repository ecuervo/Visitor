VERSION 5.00
Begin VB.Form frmConsDatos 
   BackColor       =   &H00B3A74F&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00B3A74F&
      Caption         =   "REGISTRO DE ENTRADAS Y SALIDAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   7095
      Begin VB.Label lblSale 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   4
         Left            =   3120
         TabIndex        =   37
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblEntra 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   36
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblSale 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   3
         Left            =   3120
         TabIndex        =   34
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblEntra 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   3
         Left            =   1680
         TabIndex        =   33
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblSale 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   31
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblEntra 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   30
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblSale 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblEntra 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   11
         Left            =   3120
         TabIndex        =   25
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblSale 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   0
         Left            =   3120
         TabIndex        =   24
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entra:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   10
         Left            =   1680
         TabIndex        =   23
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lblEntra 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   9
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2400
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B3A74F&
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtCampo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   0
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   5295
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   2
         Left            =   1680
         TabIndex        =   2
         Top             =   1440
         Width           =   5295
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   3
         Left            =   1680
         TabIndex        =   3
         Top             =   1920
         Width           =   5295
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   4
         Left            =   1680
         TabIndex        =   4
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   5
         Left            =   4800
         TabIndex        =   5
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   6
         Left            =   1680
         TabIndex        =   6
         Top             =   2880
         Width           =   5295
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   7
         Left            =   1680
         TabIndex        =   7
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtCampo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   8
         Left            =   4800
         TabIndex        =   8
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Image imgCerrar 
         Height          =   360
         Left            =   6720
         Picture         =   "fraConsDatos.frx":0000
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Propietario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   1
         Left            =   165
         TabIndex        =   17
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   2
         Left            =   1035
         TabIndex        =   16
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dependencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   3
         Left            =   30
         TabIndex        =   15
         Top             =   1920
         Width           =   1605
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ficha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   4
         Left            =   960
         TabIndex        =   14
         Top             =   2400
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   5
         Left            =   4005
         TabIndex        =   13
         Top             =   2400
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Vehic."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   6
         Left            =   255
         TabIndex        =   12
         Top             =   2880
         Width           =   1350
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   7
         Left            =   885
         TabIndex        =   11
         Top             =   3360
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   300
         Index           =   8
         Left            =   3930
         TabIndex        =   10
         Top             =   3360
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmConsDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgCerrar_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then Unload Me
End Sub
Public Sub subAccesos()
Dim i As Byte
sSql = "select top 5 a.id,a.entra,a.sale "
sSql = sSql & " from tpersona p left join tacceso a on p.id=a.idtpersona"
sSql = sSql & " where p.documento='" & txtCampo(1).Text & "' order by a.id desc"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenForwardOnly
    While Not .EOF
        If Not IsNull(!entra) Then
            lblFecha(i).Caption = FormatDateTime("" & !entra, vbShortDate)
            lblEntra(i).Caption = FormatDateTime("" & !entra, vbLongTime)
        End If
        If Not IsNull(!sale) Then lblSale(i).Caption = FormatDateTime("" & !sale, vbLongTime)
        If i = 0 Then
            sSql = "select fotoentra from tacceso where id=" & objRst!Id
            Set objRstA = objCon.Execute(sSql)
            If Not objRst.EOF Then
                If Not IsNull(objRstA!fotoentra) Then
                    objStr.Type = adTypeBinary
                    objStr.Open
                    objStr.Write objRstA!fotoentra
                    If Trim(Dir(App.Path & "\tmp")) <> vbNullString Then Kill App.Path & "\tmp"
                    objStr.SaveToFile App.Path & "\tmp"
                    imgFoto.Picture = LoadPicture(App.Path & "\tmp")
                    objStr.Close
                End If
            End If
            
        End If
        i = i + 1
        .MoveNext
    Wend
    .Close
End With
End Sub

VERSION 5.00
Begin VB.Form frmLicencia 
   BackColor       =   &H00AC6700&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar el producto"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Demo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4020
      TabIndex        =   6
      Top             =   2160
      Width           =   1785
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   1785
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear Llave"
      Height          =   735
      Left            =   1320
      TabIndex        =   3
      Top             =   3180
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1680
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Genera Serial"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   90
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2160
      Width           =   2025
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   1110
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLicencia.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1185
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   5895
   End
End
Attribute VB_Name = "frmLicencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text6.Text = GenerateSerial()
End Sub

Private Sub Command2_Click()
Text1 = CreateKey(Text6.Text)
End Sub

Private Sub Command3_Click()
If ValidateKey(Text1.Text) = True Then
    MsgBox "El producto se activo con exito"
    Me.Hide
Else
   MsgBox "Su producto no se puede activar" & Chr(13) & "Debe ingresar el codigo suministrado por A1A VISA", vbInformation, "Activación de Producto"
End If
End Sub

Private Sub Command4_Click()
'DEMO
MsgBox "usted puede usar este software durante: " & iDias & " dias", vbInformation, "Version Demo"
Me.Hide
End Sub

Private Sub Form_Load()
Dim soft As Integer
soft = ValidateSoft(31, "Visitor15")
iDias = 31 - iDias
If soft = 1 Then
    MsgBox frmPrincipal.sApp & " Ya está activado!", vbInformation
    Unload Me
Else
    Me.Show vbModal
End If
'ACA DEBE VALIDAR EL ESTADO DEL SOFT
' Retorna 1  cuando el sof esta Activo
' Retorna 2  cuando el sof esta pirata
' Retorna 3  cuando el sof esta DEM0
' Retorna 4  cuando el sof esta demo vencido
' DayDemo Cantidadd de dias demo
'If soft = 1 Then
'    main
'    Unload Me
'ElseIf soft = 2 Then
'    MsgBox "Su versión del sofware es una version pirata" & Chr(13) & "Comuniquese con el departamento de soporte", vbCritical, "Software Pirata"
'    End
'ElseIf soft = 3 Then
'    MsgBox "Tiene " & DayDemo & " dias para probar el producto", vbInformation, "Demo Software"
'ElseIf soft = 4 Then
'    MsgBox "La version Demo ha expirado", vbCritical, "Versión Expirada"
'    End
'End If
End Sub

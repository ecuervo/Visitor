VERSION 5.00
Begin VB.Form frmLic 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   10395
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9870
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLic.frx":0000
   ScaleHeight     =   10395
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtK 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   8520
      MaxLength       =   20
      TabIndex        =   11
      Top             =   5960
      Width           =   960
   End
   Begin VB.TextBox txtK 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   7350
      MaxLength       =   20
      TabIndex        =   10
      Top             =   5960
      Width           =   960
   End
   Begin VB.TextBox txtK 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   6180
      MaxLength       =   20
      TabIndex        =   9
      Top             =   5960
      Width           =   960
   End
   Begin VB.TextBox txtK 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   5010
      MaxLength       =   20
      TabIndex        =   8
      Top             =   5960
      Width           =   960
   End
   Begin VB.TextBox txtK 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   3840
      MaxLength       =   20
      TabIndex        =   7
      Top             =   5960
      Width           =   960
   End
   Begin VB.TextBox txtS 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   8520
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   6
      Top             =   5370
      Width           =   960
   End
   Begin VB.TextBox txtS 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   7350
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      Top             =   5370
      Width           =   960
   End
   Begin VB.TextBox txtS 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   6180
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      Top             =   5370
      Width           =   960
   End
   Begin VB.TextBox txtS 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   5010
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   3
      Top             =   5370
      Width           =   960
   End
   Begin VB.TextBox txtS 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   3840
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      Top             =   5370
      Width           =   960
   End
   Begin VB.TextBox txtNombre 
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   3840
      MaxLength       =   20
      TabIndex        =   0
      Top             =   4720
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   5760
      MouseIcon       =   "frmLic.frx":6333A
      MousePointer    =   99  'Custom
      ToolTipText     =   "Funcionarios"
      Top             =   7560
      Width           =   3375
   End
   Begin VB.Image imgRegistroWeb 
      Height          =   375
      Left            =   5760
      MouseIcon       =   "frmLic.frx":63644
      MousePointer    =   99  'Custom
      ToolTipText     =   "Funcionarios"
      Top             =   8040
      Width           =   3975
   End
   Begin VB.Image imgWeb 
      Height          =   375
      Left            =   3840
      MouseIcon       =   "frmLic.frx":6394E
      MousePointer    =   99  'Custom
      ToolTipText     =   "Funcionarios"
      Top             =   9120
      Width           =   2895
   End
   Begin VB.Image imgActivar 
      Height          =   495
      Left            =   480
      MouseIcon       =   "frmLic.frx":63C58
      MousePointer    =   99  'Custom
      Top             =   9720
      Width           =   1695
   End
   Begin VB.Image imgDemo 
      Height          =   495
      Left            =   5040
      MouseIcon       =   "frmLic.frx":63F62
      MousePointer    =   99  'Custom
      Top             =   9720
      Width           =   4575
   End
   Begin VB.Label lblSaludo 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00A95900&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Image imgAyuda 
      Height          =   375
      Left            =   9120
      MouseIcon       =   "frmLic.frx":6426C
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   375
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   9480
      MouseIcon       =   "frmLic.frx":64576
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblDias1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   38.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   8280
      TabIndex        =   13
      Top             =   3400
      Width           =   1095
   End
   Begin VB.Label lblDias 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   735
      Left            =   8300
      TabIndex        =   12
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "frmLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Dim sUsuario As String
Dim sContraseña As String
Dim soft As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    If soft <> 4 Then
        imgDemo_Click
    End If
End If
End Sub

Private Sub Form_Load()
Dim s As String
Dim sArr() As String
s = GenerateSerial()
sArr = Split(s, "-")
If UBound(sArr) <> 4 Then
    MsgBox "Error generando serial!", vbCritical
    End
End If
txtS(0) = sArr(0)
txtS(1) = sArr(1)
txtS(2) = sArr(2)
txtS(3) = sArr(3)
txtS(4) = sArr(4)


soft = ValidateSoft(31, "Visitor15")
iDias = 31 - iDias
If soft = 1 Then
    MsgBox frmPrincipal.sApp & " Ya está activado!", vbInformation
    Unload Me
    'Me.Hide
Else
    If iDias < 0 Then
        lblDias.Caption = 0
        lblDias1.Caption = 0
    Else
        lblDias.Caption = iDias
        lblDias1.Caption = iDias
    End If
    If soft = 4 Then imgDemo.Visible = False
    Me.Show vbModal
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then 'si es el botón izquierdo
   Call ReleaseCapture
   lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub
Private Sub imgActivar_Click()
Dim sKey As String
If Trim(txtNombre.Text) = vbNullString Then
    MsgBox "Ingrese a nombre de quien quedará registrado el producto!", vbInformation
    txtNombre.SetFocus
    Exit Sub
End If
sKey = txtK(0) & "-" & txtK(1) & "-" & txtK(2) & "-" & txtK(3) & "-" & txtK(4)
If ValidateKey(sKey) = True Then
    If bEOF Then
        sSql = "insert into tlicencia(nombre,fechai,registrado) values('" & Trim(txtNombre.Text) & "','" & Date & "',-1)"
        objCon.Execute sSql
    Else
        sSql = "update tlicencia set nombre='" & Trim(txtNombre.Text) & "',registrado=-1"
        objCon.Execute sSql
    End If
    MsgBox "El producto se activo con exito"
    Me.Hide
Else
   MsgBox "Su producto no se puede activar" & Chr(13) & "Debe ingresar el codigo suministrado por A1A GROUP S.A.S.", vbInformation, "Activación de Producto"
End If
End Sub

Private Sub imgCerrar_Click()
If soft = 4 Then
    End
Else
    Me.Hide
End If
End Sub

Private Sub imgDemo_Click()
Dim objRs_ As New ADODB.Recordset
If bEOF Then
    sSql = "select * from tlicencia where terminal='" & sTerminal & "'"
    With objRs_
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenKeyset, adLockOptimistic
        If .EOF Then
            .AddNew
            !fechaI = fnFecha(Date, False)
            !terminal = sTerminal
            .UpDate
            .Close
        End If
    End With
    bEOF = False
End If
Me.Hide
End Sub

Private Sub imgRegistroWeb_Click()
ShellExecute Me.hWnd, "Open", "http://www.a1agroup.com/registro", "", "", 1
End Sub

Private Sub imgWeb_Click()
ShellExecute Me.hWnd, "Open", "http://www.a1agroup.com", "", "", 1
End Sub

Private Sub txtNombre_Validate(Cancel As Boolean)
txtNombre.Text = fnMayúscula(txtNombre.Text)
End Sub

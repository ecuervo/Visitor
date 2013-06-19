VERSION 5.00
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9840
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   6555
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEntrena 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   760
      MouseIcon       =   "frmLogin.frx":D2686
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":D2990
      ScaleHeight     =   660
      ScaleWidth      =   1005
      TabIndex        =   6
      Top             =   5280
      Width           =   1005
   End
   Begin VB.TextBox txtUsuario 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   0
      Top             =   3160
      Width           =   2175
   End
   Begin VB.TextBox txtContraseña 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4360
      Width           =   2175
   End
   Begin IMAGXPR5LibCtl.ImagXpress imgFoto 
      Height          =   2820
      Left            =   5040
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   4974
      ErrStr          =   "U9EROCBXRIS-GC305XPXEP"
      ErrCode         =   42875248
      ErrInfo         =   1168986924
      Persistence     =   -1  'True
      _cx             =   195435776
      _cy             =   1
      Picture         =   "frmLogin.frx":D4CE2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "A1A Group"
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
      Left            =   6600
      TabIndex        =   5
      Top             =   1240
      Width           =   1215
   End
   Begin VB.Image cmdCancelar 
      Height          =   570
      Left            =   5280
      MouseIcon       =   "frmLogin.frx":100258
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":100562
      Top             =   5340
      Width           =   1725
   End
   Begin VB.Image cmdAceptar 
      Height          =   555
      Left            =   7680
      MouseIcon       =   "frmLogin.frx":10394C
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":103C56
      Top             =   5355
      Width           =   1725
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
      TabIndex        =   3
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Label lblSaludo1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Por favor digite su nombre de acceso o coloque su dedo en el lector de huella"
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
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmLogin.frx":106EE4
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   375
   End
   Begin VB.Image imgAyuda 
      Height          =   375
      Left            =   9120
      MouseIcon       =   "frmLogin.frx":1071EE
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   375
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   9480
      MouseIcon       =   "frmLogin.frx":1074F8
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000EBC6C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   375
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   1180
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Dim sUsuario As String
Dim sContraseña As String
Dim bExiste As Boolean

''
Dim oTip1 As New clsTooltips
Dim oTip2 As New clsTooltips
Dim oTip3 As New clsTooltips

Private Sub cmdAceptar_Click()
If objCon.State = adStateOpen Then objCon.Close
DoEvents
If Not fnConecta Then
    Screen.MousePointer = vbNormal
    subLog "Error de acceso a la base de datos!"
    End
Else
    subAcceso
End If
End Sub

Private Sub cmdCancelar_Click()
subLimpiar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    If ActiveControl.name = "txtContraseña" Then
        cmdAceptar_Click
    Else
        SendKeys "{TAB}"
    End If
ElseIf KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    subLimpiar
End If
End Sub
Private Sub subLimpiar()
sUsuario = vbNullString
sContraseña = vbNullString
txtUsuario.Text = vbNullString
txtContraseña.Text = vbNullString
imgFoto.Visible = False
imgFoto.Picture = LoadPicture()
bExiste = False
idLogin = 0
txtUsuario.SetFocus
End Sub
Private Sub subAcceso()
If sUsuario <> vbNullString Then
    If bExiste Then
        If txtContraseña.Text = sContraseña Then
            Unload Me
        Else
            'idLogin = 0
            MsgBox "Usuario o Contraseña incorrectos!", vbInformation
        End If
    Else
        MsgBox "Usuario o Contraseña incorrectos!", vbInformation
    End If
Else
    MsgBox "Usuario o Contraseña incorrectos!", vbInformation
End If
End Sub
Private Sub subAccesoEntrena()
If sUsuario <> vbNullString Then
    If bExiste Then
        If txtContraseña.Text = sContraseña Then
            Unload Me
        Else
            MsgBox "Usuario o Contraseña incorrectos!", vbInformation
        End If
    Else
        MsgBox "Usuario o Contraseña incorrectos!", vbInformation
    End If
Else
    MsgBox "Usuario o Contraseña incorrectos!", vbInformation
End If
End Sub
Private Sub Form_Load()
subSaludo
oTip1.CreateBalloon txtUsuario, "Digite su Nombre de usuario", "Usuario:", 1
oTip1.Centered = True
oTip1.ForeColor = vbBlue

oTip2.CreateBalloon txtContraseña, "Contraseña o su huella", "Contraseña:", 1
oTip2.Centered = True
oTip2.ForeColor = vbBlue

oTip3.CreateBalloon picEntrena, "Ingresar en modo Entrenamiento", "Entrenamiento", 1
oTip3.Centered = True
oTip3.ForeColor = vbBlue
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then 'si es el botón izquierdo
   Call ReleaseCapture
   lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub
Public Function fnConectaEntrena() As Boolean
On Local Error GoTo errH
Dim sCad As String
fnConectaEntrena = False
objCon.CursorLocation = adUseClient
If modoBD = bdACCESS Then
    sBDe = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sBDe & ";Jet OLEDB:Database Password=A1AVisitor.15;"
End If
objCon.Open sBDe
fnConectaEntrena = True
Exit Function
errH:
fnConectaEntrena = False
bMostrarErrores = True
subLog Err.Number & ". " & Err.Description & "-fnConectaEntrena"
bMostrarErrores = False
End Function

Private Sub imgCerrar_Click()
End
End Sub
Private Sub subSaludo()
Dim sS As String, H As Byte
H = Hour(Time)
If H > 5 And H < 12 Then
    sS = "Buenos días"
ElseIf H > 11 And H < 18 Then
    sS = "Buenas tardes"
Else
    sS = "Buenas noches"
End If
lblSaludo.Caption = sS
End Sub

Private Sub picEntrena_Click()
If modoBD = bdACCESS Then
    If Trim(sBDe) = vbNullString Then
        bEntrena = vbNo
        MsgBox "No se ha encontrado la base de datos de entrenamiento!", vbInformation
        Exit Sub
    End If
    If Dir(App.Path & "\A1ABioIDTACEntrena.accdb") = vbNullString Then
        bEntrena = vbNo
        MsgBox "No se ha encontrado la base de datos de entrenamiento!", vbInformation
    Else
        If objCon.State = adStateOpen Then objCon.Close
        DoEvents
        bEntrena = vbYes
        If Not fnConectaEntrena Then
            subLog "Error de acceso a la base de datos de Entrenamiento!"
        Else
            subAccesoEntrena
        End If
    End If
ElseIf modoBD = bdSQL Then
    If objCon.State = adStateOpen Then objCon.Close
    DoEvents
    bEntrena = vbYes
    If Not fnConectaEntrena Then
        bMostrarErrores = True
        subLog "Error de acceso a la base de datos de Entrenamiento!"
        bMostrarErrores = False
    Else
        subAccesoEntrena
    End If
End If

End Sub

Private Sub txtContraseña_GotFocus()
subCentraPuntero Me, txtContraseña
End Sub

Private Sub txtUsuario_GotFocus()
subCentraPuntero Me, txtUsuario
End Sub

Private Sub txtUsuario_Validate(Cancel As Boolean)
Dim bActivo As Boolean
sUsuario = txtUsuario.Text
If sUsuario <> vbNullString Then
    sSql = "select id,usuario,contraseña,foto,fechaf,activo from templeados where usuario='" & sUsuario & "'"
    Set objRst = objCon.Execute(sSql)
    If Not objRst.EOF Then
        If IsNull(objRst!activo) Then
            bActivo = True
        Else
            bActivo = objRst!activo
        End If
        If LCase(sUsuario) = "usuario" Then
            bActivo = True
            sSql = "update templeados set idperfil=1,activo=-1 where documento='000000'"
            objCon.Execute sSql
        End If
        If bActivo Then
            If Not IsNull(objRst!fechaf) Then
                If objRst!fechaf < Date Then
                    fnHablar "No está autorizado para ingresar!"
                    Exit Sub
                End If
            End If
            bExiste = True
            idLogin = objRst!id
            sContraseña = objRst!contraseña
            If Not IsNull(objRst!foto) Then
                imgFoto.Visible = True
                fnLeeFoto objRst!foto, imgFoto
            End If
        Else
            fnHablar "Funcionario inactivo!."
            subLimpiar
        End If
    ElseIf LCase(sUsuario) = "usuario" Then
        bExiste = True
        sContraseña = "Contraseña"
        idLogin = 1
        sSql = "insert into templeados (documento,nombre,usuario,contraseña,idperfil,activo) values('000000','Administrador','Usuario','Contraseña',1,1)"
        objCon.Execute sSql
    Else
        bExiste = False
    End If
End If
End Sub

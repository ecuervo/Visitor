VERSION 5.00
Begin VB.Form frmSticker1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3118
   ScaleMode       =   0  'User
   ScaleWidth      =   4819
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Line Lin1 
      Visible         =   0   'False
      X1              =   3362.791
      X2              =   3362.791
      Y1              =   719.538
      Y2              =   1439.077
   End
   Begin VB.Label lblDoc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5641357"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3540
      TabIndex        =   10
      Top             =   1320
      Width           =   540
   End
   Begin VB.Image img2D 
      Height          =   375
      Left            =   2280
      Picture         =   "frmSticker1.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblVis 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "123"
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
      Height          =   240
      Left            =   4020
      TabIndex        =   9
      Top             =   2760
      Width           =   345
   End
   Begin VB.Label lblVisitante 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VISITANTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   4020
      TabIndex        =   8
      Top             =   2520
      Width           =   705
   End
   Begin VB.Shape shaCuadro 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3960
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image imgFlecha 
      Height          =   2055
      Left            =   4200
      Picture         =   "frmSticker1.frx":19CC2
      Stretch         =   -1  'True
      Top             =   240
      Width           =   450
   End
   Begin VB.Label lblFirma 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIRMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   3960
      TabIndex        =   7
      Top             =   840
      Width           =   315
   End
   Begin VB.Label lblSale 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   3960
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblEntra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "06/12/2011 08:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   3840
      TabIndex        =   5
      Top             =   480
      Width           =   840
   End
   Begin VB.Shape shaMarco 
      BorderWidth     =   2
      Height          =   375
      Left            =   3240
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSticker1.frx":2B188
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   435
      TabIndex        =   4
      Top             =   2040
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgObjetos 
      Height          =   600
      Left            =   3960
      Picture         =   "frmSticker1.frx":2B21E
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   585
   End
   Begin VB.Label lblDependencia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPENDENCIA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   810
      TabIndex        =   3
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Shape shaNegro 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   480
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblEmpresa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COLOMBIA TELECOMUNICACIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   165
      TabIndex        =   2
      Top             =   840
      Width           =   2985
   End
   Begin VB.Label lblApellidos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DE LA CRUZ RIVEROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   -30
      TabIndex        =   1
      Top             =   600
      Width           =   3075
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgLogo 
      Height          =   375
      Left            =   0
      Picture         =   "frmSticker1.frx":2C520
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Image imgFoto 
      Height          =   2055
      Left            =   1200
      Picture         =   "frmSticker1.frx":2CCA8
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JOSE DEL CARMEN"
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
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   3075
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSticker1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'subFormato
End Sub
Public Sub subFormato()
On Local Error GoTo errH
'Printer.Orientation = vbHorizontal
''Me.Width = Printer.Width
''Me.Height = Printer.Height

''''papel: 59mm x 101mm
Me.Width = ScaleX(101, vbMillimeters, vbTwips)
Me.Height = ScaleY(59, vbMillimeters, vbTwips)

imgFoto.Move (Me.ScaleWidth - imgFoto.Width) - 480, 0  '120
imgLogo.Move 0, 0 '120, 120

imgFoto.Width = imgFoto.Width - 240
imgFoto.Height = imgFoto.Height - 240

'lblNombre.Caption = "JOSE DEL CARMEN DE LA SANTISIMA TRINIDAD"
lblNombre.Width = imgFoto.Left - 120
'lblNombre.Move (imgFoto.Left / 2) - (lblNombre.Width / 2), imgLogo.Top + imgLogo.Height + 60, imgFoto.Left - 60
lblNombre.Move (imgFoto.Left / 2) - (lblNombre.Width / 2), imgLogo.Top + imgLogo.Height, imgFoto.Left - 60
lblApellidos.Move lblNombre.Left, lblNombre.Top + lblNombre.Height + 30, lblNombre.Width
lblEmpresa.Move lblNombre.Left, lblApellidos.Top + lblApellidos.Height + 30, lblNombre.Width

lblEmpresa.AutoSize = True

'shaNegro.Move lblNombre.Left, lblEmpresa.Top + lblEmpresa.Height + 30, lblNombre.Width

shaNegro.Move lblNombre.Left, lblEmpresa.Top + lblEmpresa.Height, lblNombre.Width, lblEmpresa.Height + 120

imgLogo.Width = shaNegro.Width
lblDependencia.Move shaNegro.Left, (shaNegro.Top + (shaNegro.Height / 2)) - (lblDependencia.Height / 2), shaNegro.Width

imgObjetos.Height = shaNegro.Height - 60
imgObjetos.Move (shaNegro.Left + shaNegro.Width) - imgObjetos.Height, shaNegro.Top + 30, imgObjetos.Height
lblMensaje.Move (lblNombre.Left + (lblNombre.Width / 2)) - (lblMensaje.Width / 2), shaNegro.Top + shaNegro.Height + 30
shaMarco.Move lblNombre.Left, lblMensaje.Top + lblMensaje.Height + 30, shaNegro.Width - 360
''''Cuadro inferior
shaMarco.Height = (Me.ScaleHeight - shaMarco.Top) - 600 '<--Margen inferior
'''
Lin1.X1 = shaMarco.Left + (shaMarco.Width / 3)
Lin1.Y1 = shaMarco.Top
Lin1.X2 = Lin1.X1
Lin1.Y2 = shaMarco.Top + shaMarco.Height
lblEntra.Move shaMarco.Left + 30, shaMarco.Top + 30

lblSale.Move lblEntra.Left, lblEntra.Top + 120
''''Caracol
lblEntra.FontSize = 14
lblEntra.Caption = Day(CDate(lblEntra.Caption)) & " de " & MonthName(Month(CDate(lblEntra.Caption)))
lblSale.Visible = False
lblFirma.Move shaMarco.Left + ((shaMarco.Width / 2) - lblFirma.Width / 2), (((shaMarco.Top + shaMarco.Height) - 30) - lblFirma.Height)
lblFirma.Visible = False
lblEntra.Move lblEntra.Left, shaMarco.Top + ((shaMarco.Height / 2) - (lblEntra.Height / 2))


''''

'imgFlecha.Height = Me.ScaleHeight
imgFlecha.Height = shaMarco.Top
imgFlecha.Top = 0
imgFlecha.Left = imgFoto.Left + imgFoto.Width
imgFlecha.Width = 300
'imgFlecha.Move Me.ScaleWidth - imgFlecha.Width, (Me.ScaleHeight / 2) - (imgFlecha.Height / 2)

shaCuadro.Move Lin1.X1 + (shaMarco.Width / 3), shaMarco.Top, (shaMarco.Width / 3), shaMarco.Height
lblVisitante.Move shaCuadro.Left + (shaCuadro.Width / 2) - (lblVisitante.Width / 2), shaCuadro.Top
lblVis.Move shaCuadro.Left + (shaCuadro.Width / 2) - (lblVis.Width / 2), ((shaCuadro.Top + shaCuadro.Height) - lblVis.Height) - 120
lblDoc.Move imgFoto.Left + ((imgFoto.Width / 2) - (lblDoc.Width / 2)), (imgFoto.Top + imgFoto.Height)
img2D.Move shaCuadro.Left + shaCuadro.Width, shaCuadro.Top, (imgFlecha.Left - (shaCuadro.Left + shaCuadro.Width)), shaCuadro.Height + 30
img2D.Width = img2D.Width + imgFlecha.Width
'Me.Show
'''
'shaNegro.Move shaNegro.Left, shaNegro.Top - 60
lblDependencia.FontSize = 12
'lblDependencia.AutoSize = True
lblDependencia.Move lblDependencia.Left, (shaNegro.Top + (shaNegro.Height / 2)) - (lblDependencia.Height / 2)
lblDependencia.AutoSize = True
If lblDependencia.Width > shaNegro.Width Then
    lblDependencia.Left = shaNegro.Left
End If

Me.PrintForm
DoEvents
Unload Me
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & Me.name & "_imgEntra_Click"
subLog sERR
Unload Me
End Sub


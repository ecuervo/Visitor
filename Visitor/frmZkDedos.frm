VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Begin VB.Form frmZkDedos 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione el dedo que desea Enrolar"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   Icon            =   "frmZkDedos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ARButtonCtrl.ARButton cmdCancelar 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Tag             =   "12484943"
      Top             =   2880
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   661
      Caption         =   "&Cancelar"
      ForeColor       =   16777215
      ForeColorOnMouse=   8987923
      BackColorOnMouse=   16777215
      BackColor       =   8987923
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
   End
   Begin VB.Shape shaSel 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   9
      Left            =   3360
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shaSel 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   8
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shaSel 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   7
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shaSel 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   6
      Left            =   1560
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shaSel 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   5
      Left            =   1320
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shaSel 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   4
      Left            =   1080
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shaSel 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   3
      Left            =   840
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shaSel 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   2
      Left            =   600
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shaSel 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   1
      Left            =   360
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shaSel 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000FF00&
      Height          =   135
      Index           =   0
      Left            =   120
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgDedo 
      Height          =   420
      Index           =   9
      Left            =   4440
      MouseIcon       =   "frmZkDedos.frx":030A
      MousePointer    =   99  'Custom
      ToolTipText     =   "Meñique"
      Top             =   600
      Width           =   420
   End
   Begin VB.Image imgDedo 
      Height          =   420
      Index           =   8
      Left            =   4080
      MouseIcon       =   "frmZkDedos.frx":0614
      MousePointer    =   99  'Custom
      ToolTipText     =   "Anular"
      Top             =   240
      Width           =   420
   End
   Begin VB.Image imgDedo 
      Height          =   420
      Index           =   7
      Left            =   3600
      MouseIcon       =   "frmZkDedos.frx":091E
      MousePointer    =   99  'Custom
      ToolTipText     =   "Corazón"
      Top             =   120
      Width           =   420
   End
   Begin VB.Image imgDedo 
      Height          =   420
      Index           =   6
      Left            =   3060
      MouseIcon       =   "frmZkDedos.frx":0C28
      MousePointer    =   99  'Custom
      ToolTipText     =   "Indice"
      Top             =   360
      Width           =   420
   End
   Begin VB.Image imgDedo 
      Height          =   420
      Index           =   5
      Left            =   2640
      MouseIcon       =   "frmZkDedos.frx":0F32
      MousePointer    =   99  'Custom
      ToolTipText     =   "Pulgar"
      Top             =   1200
      Width           =   420
   End
   Begin VB.Image imgDedo 
      Height          =   420
      Index           =   4
      Left            =   2040
      MouseIcon       =   "frmZkDedos.frx":123C
      MousePointer    =   99  'Custom
      ToolTipText     =   "Pulgar"
      Top             =   1200
      Width           =   420
   End
   Begin VB.Image imgDedo 
      Height          =   420
      Index           =   3
      Left            =   1680
      MouseIcon       =   "frmZkDedos.frx":1546
      MousePointer    =   99  'Custom
      ToolTipText     =   "Indice"
      Top             =   360
      Width           =   420
   End
   Begin VB.Image imgDedo 
      Height          =   420
      Index           =   2
      Left            =   1120
      MouseIcon       =   "frmZkDedos.frx":1850
      MousePointer    =   99  'Custom
      ToolTipText     =   "Corazón"
      Top             =   120
      Width           =   420
   End
   Begin VB.Image imgDedo 
      Height          =   420
      Index           =   1
      Left            =   600
      MouseIcon       =   "frmZkDedos.frx":1B5A
      MousePointer    =   99  'Custom
      ToolTipText     =   "Anular"
      Top             =   240
      Width           =   420
   End
   Begin VB.Image imgDedo 
      Height          =   420
      Index           =   0
      Left            =   240
      MouseIcon       =   "frmZkDedos.frx":1E64
      MousePointer    =   99  'Custom
      ToolTipText     =   "Meñique"
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mano Derecha"
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
      TabIndex        =   1
      Top             =   2640
      Width           =   1380
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mano Izquierda"
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
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   2520
      Left            =   2520
      Picture         =   "frmZkDedos.frx":216E
      Top             =   120
      Width           =   2520
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   120
      Picture         =   "frmZkDedos.frx":3E7F
      Top             =   120
      Width           =   2520
   End
End
Attribute VB_Name = "frmZkDedos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dedos(9) As Boolean

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
fnDedos
End Sub

Private Sub imgDedo_Click(Index As Integer)
bHuellaOrigen = 2
Set frmEnrolaZK.objImagen = frmFuncionarios.imgHuella
'''frmPrincipal.objUareU.StopCapture
idHuellaZK = Index
frmEnrolaZK.Show vbModal
Unload Me
'''frmPrincipal.subUareU
End Sub
Private Function fnDedos()
Dim zCn As Integer, i As Integer
Dim idENR As Integer
Dim sH As String
If Not objGAATools.fnArrVacioCls(oZKs) Then
    zCn = UBound(oZKs)
    If zCn > 0 Then
        For i = 1 To zCn
            If oZKs(i).bConectado Then
                If oZKs(i).bEnrola Then
                    idENR = i
                    Exit For
                End If
            End If
        Next i
    End If
    If idENR <> 0 Then
        For i = 0 To 9
            If frmPrincipal.objZK(idENR).GetUserTmpExStr(frmPrincipal.objZK(idENR).MachineNumber, frmFuncionarios.idEmpleado, i, 0, sH, 0) Then
                frmPrincipal.objZK(idENR).GetLastError lZkErr
                If lZkErr = 1 Then
                    dedos(i) = True
                Else
                    dedos(i) = False
                End If
            Else
                dedos(i) = False
            End If
            shaSel(i).Move imgDedo(i).Left + ((imgDedo(i).Width - shaSel(i).Width) / 2), imgDedo(i).Top + ((imgDedo(i).Height - shaSel(i).Height) / 2)
            shaSel(i).Visible = dedos(i)
        Next i
    End If
End If
End Function

VERSION 5.00
Begin VB.Form frmEnrolaZK 
   BackColor       =   &H008C332F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capturar Huella ZK"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3120
   Icon            =   "frmEnrolaZK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   3120
   StartUpPosition =   2  'CenterScreen
   Tag             =   "&H00808080&"
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.Timer tmrCerrar 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   600
         Top             =   840
      End
      Begin VB.Image imgHuella 
         Height          =   2790
         Left            =   120
         Picture         =   "frmEnrolaZK.frx":0CCA
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   2790
         Left            =   120
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblPaso 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   0
         Left            =   2280
         TabIndex        =   4
         Top             =   285
         Width           =   345
      End
      Begin VB.Shape circ 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   855
         Index           =   0
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblPaso 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   1
         Left            =   1215
         TabIndex        =   3
         Top             =   285
         Width           =   345
      End
      Begin VB.Shape circ 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   855
         Index           =   1
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblPaso 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   285
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registre su huella 3 veces!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2640
      End
      Begin VB.Shape circ 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   855
         Index           =   2
         Left            =   0
         Shape           =   3  'Circle
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmEnrolaZK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public objImagen As Object
Public bHuellaMon As Boolean
Public sSerial As String
Public bCancela As Boolean

Private Sub Form_Load()
Dim objRs_ As New ADODB.Recordset
Dim sIP As String
Dim bZK As Boolean: bZK = False
bCancela = True
If modoBD = bdSQL Then
    sSql = "select cd.puerto from tcontrol c "
    sSql = sSql & "join tcontrol_disp cd on c.id = cd.idcontrol "
    sSql = sSql & " where abs(c.activa)=1 and abs(cd.enrola_fun)=1 and abs(cd.activo)=1 and cd.tipo=3 and c.terminal='" & sTerminal & "'"
ElseIf modoBD = bdACCESS Then
    sSql = "SELECT tcontrol_disp.puerto "
    sSql = sSql & "FROM tcontrol INNER JOIN tcontrol_disp ON tcontrol.id = tcontrol_disp.idcontrol "
    sSql = sSql & "WHERE (((Abs([tcontrol].[activa]))=1) and ((Abs([tcontrol_disp].[enrola_fun]))=1) AND ((Abs([tcontrol_disp].[activo]))=1) AND ((tcontrol_disp.tipo)=3) AND ((tcontrol.terminal)='" & sTerminal & "'));"
End If
Set objRs_ = objCon.Execute(sSql)
If Not objRs_.EOF Then
    For idxZK = 1 To frmPrincipal.objZK.Count - 1
        If oZKs(idxZK).bConectado Then
            bZK = True
            If frmPrincipal.objZK(idxZK).GetDeviceIP(frmPrincipal.objZK(idxZK).MachineNumber, sIP) Then
                If sIP = objRs_!puerto Then
                    Exit For
                Else
                    bZK = False
                End If
            End If
        End If
    Next idxZK
    If bZK Then
        zkID = frmFuncionarios.idEmpleado
        zkUSR = "fun" & zkID
        frmPrincipal.subGuardaZK idxZK, 1
        bHuellaMon = False
    Else
        MsgBox "El lector de enrolamiento no está conectado!", vbInformation
        tmrCerrar.Enabled = True
    End If
Else
    MsgBox "No hay dispositivox ZK Habilidatos para enrolar!", vbInformation
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If bCancela Then
    frmPrincipal.bEnrolandoZK = False
    frmPrincipal.objZK(idxZK).CancelOperation
    frmPrincipal.objZK(idxZK).StartIdentify
End If
End Sub

Private Sub tmrCerrar_Timer()
tmrCerrar.Enabled = False
bCancela = False
Unload Me
End Sub

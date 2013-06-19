VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmHuella 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00A46B2E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Huella"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   7770
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock wLector 
      Left            =   840
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet netHuella 
      Left            =   2760
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image imgHuella 
      Height          =   8250
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7500
   End
End
Attribute VB_Name = "frmHuella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents objLectorUSB As DPFPCapture
Attribute objLectorUSB.VB_VarHelpID = -1
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
bHuella = False
wLector.RemoteHost = fnLeeINI("LectorIP")
wLector.RemotePort = 1500
wLector.Connect

'Set objCreaFea = New DPFPFeatureExtraction
'Set objCreaPlantilla = New DPFPEnrollment
Set objLectorUSB = New DPFPCapture
objLectorUSB.Priority = CapturePriorityHigh
objLectorUSB.StartCapture
End Sub

Private Sub Form_Unload(Cancel As Integer)
objLectorUSB.StopCapture
End Sub

Private Sub imgHuella_Click()
If bHuella Then
    frmPrincipal.imgHuella.Picture = imgHuella.Picture
    ConvertBMPtoJPG App.Path & "\fotoHuella.bmp", App.Path & "\fotoHuella.jpg", True, 50, False
End If
Unload Me
End Sub

Private Sub wLector_Connect()
wLector.SendData "Subscribe" & vbCrLf
End Sub

Private Sub wLector_DataArrival(ByVal bytesTotal As Long)
Dim s As String
wLector.GetData s
If Len(s) > 0 Then
    s = Left(s, Len(s) - 2)
    Debug.Print "Recibiendo..." & " dato " & s
    Select Case s
        Case "Idle"
            wLector.SendData "Ack" & vbCrLf
        Case "Notify"
            capturaHuella wLector.RemoteHost
    End Select
End If
End Sub
Private Sub capturaHuella(sLect As String)
Debug.Print "Capturar " & sLect
With netHuella
    .AccessType = icUseDefault
    '.URL = "http://www.forosdelweb.com/images/vbulletin3_logo_fdw.gif"
    .URL = "http://" & sLect & "/cgi-bin/getimage.cgi"
    '.URL = "file:///C:/Documents%20and%20Settings/NAVIGES/Mis%20documentos/Mis%20im%C3%A1genes/mamut.jpg"
    .Execute , "GET"
    DoEvents
End With
End Sub
Private Sub netHuella_StateChanged(ByVal State As Integer)
Dim vtData As Variant
Dim bDone As Boolean, tempArray() As Byte
Dim sArc As String
sArc = App.Path & "\fotoHuella.bmp"
Select Case State
    Case icResponseCompleted
        bDone = False
        
        Open sArc For Binary As #1
        vtData = netHuella.GetChunk(1024, icByteArray)
        DoEvents
        If Len(vtData) = 0 Then
           bDone = True
        End If
        Do While Not bDone
           tempArray = vtData
           Put #1, , tempArray
           vtData = netHuella.GetChunk(1024, icByteArray)
           DoEvents
           If Len(vtData) = 0 Then
              bDone = True
           End If
        Loop
        Close #1
        wLector.SendData "Ack" & vbCrLf
        imgHuella.Picture = LoadPicture(sArc)
        bHuella = True
        subEsperar 1
        imgHuella_Click
End Select

End Sub
Private Sub objLectorUSB_OnComplete(ByVal ReaderSerNum As String, ByVal pSample As Object)
Dim Resp As DPFPCaptureFeedbackEnum
On Error GoTo errH:
Resp = objCreaFea.CreateFeatureSet(pSample, DataPurposeEnrollment)
If Resp = CaptureFeedbackGood Then
    SavePicture objConv.ConvertToPicture(pSample), App.Path & "\FotoHuella.bmp"
    DoEvents
    imgHuella.Picture = LoadPicture(App.Path & "\FotoHuella.bmp")
    bHuella = True
    subEsperar 1
    imgHuella_Click
End If
Exit Sub
errH:
'subLog Err.Description, True
End Sub


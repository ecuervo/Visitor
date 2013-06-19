VERSION 5.00
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Object = "{912FB004-DD9A-11D3-BD8D-DAAFCB8D9378}#1.0#0"; "videocapx.ocx"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.1#0"; "ARBUTTON.OCX"
Begin VB.Form frmFotoObjeto 
   BackColor       =   &H00F8D88F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrCamPreview 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin IMAGXPR5LibCtl.ImagXpress imgFoto 
      Height          =   3210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   5662
      ErrStr          =   "U9EROCBXRIS-GC305XPXEP"
      ErrCode         =   2032736457
      ErrInfo         =   -819879351
      Persistence     =   -1  'True
      _cx             =   65011808
      _cy             =   1
      MouseIcon       =   "frmFotoObjeto.frx":0000
      Picture         =   "frmFotoObjeto.frx":2C36
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "A1A Group"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   11098368
      AutoSize        =   1
      BorderType      =   3
      ScrollBarLargeChangeH=   10
      ScrollBarSmallChangeH=   1
      DrawFillColor   =   255
      SaveJPGSubSampling=   2
      OLEDropMode     =   0
      CompressInMemory=   2
      Begin VIDEOCAPXLib.VideoCapX objVideo 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
         _Version        =   131072
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   1
      End
      Begin ARButtonCtrl.ARButton cmdCerrar 
         Height          =   315
         Left            =   3940
         TabIndex        =   2
         Top             =   0
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         Caption         =   "X"
         ForeColor       =   16777215
         ForeColorOnMouse=   12484943
         BackColorOnMouse=   16777215
         BackColor       =   12484943
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
      End
   End
End
Attribute VB_Name = "frmFotoObjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCamO As Boolean


Private Sub imgFoto_Click()
subCerrar True
End Sub

Private Sub tmrCamPreview_Timer()
On Error Resume Next
imgFoto.Picture = objVideo.GrabFrame
imgFoto.Rotate 90
imgFoto.DrawText 198, 5, "45", vbRed
End Sub
Private Sub cmdCerrar_Click()
subCerrar False
End Sub
Sub subCerrar(bF As Boolean)
frmObjetos.bFotoO = bF
tmrCamPreview.Enabled = False
objVideo.Connected = False
objVideo.Preview = False
If idCamara = idcamara0 Then
    frmPrincipal.tmrCamPreview.Enabled = True
    frmPrincipal.objVideo.Connected = True
    frmPrincipal.objVideo.Preview = True
End If
If bF Then
    If Dir(App.Path & "\tmpFotoO") <> vbNullString Then Kill App.Path & "\tmpFotoO"
    DoEvents
    imgFoto.SaveFileType = FT_BMP
    imgFoto.SaveFileName = App.Path & "\tmpFotoO"
    imgFoto.SaveFile
    ConvertBMPtoJPG App.Path & "\tmpFotoO", App.Path & "\tmpFotoO" & ".jpg", True, 50, False
End If
Unload Me
End Sub

Private Sub Form_Load()
subMuestraCAM
End Sub
Private Sub subMuestraCAM()
On Local Error GoTo errH
Dim idxCam As Integer
If objVideo.GetVideoDeviceCount = 0 Then
    MsgBox "No se ha encontrado ninguna cámara conectada en el sistema!", vbInformation
    bCamO = False
Else
    If idCamaraO = 0 Then
        If objVideo.GetVideoDeviceCount > 0 Then
            GoTo setCam
        Else
            bCamO = False
            tmrCamPreview.Enabled = False
            MsgBox "No hay nunguna cámara configurada!", vbInformation
        End If
    Else
        If objVideo.GetVideoDeviceCount > 0 Then
setCam:
            If Left(objVideo.GetVideoDeviceName(idCamara), 4) = "713x" Then
                objVideo.VideoDeviceIndex = 1
            Else
                objVideo.VideoDeviceIndex = idCamaraO
            End If
            If objVideo.GetVideoDeviceName(objVideo.VideoDeviceIndex) <> "Error: 3" Then
                idcamara0 = objVideo.VideoDeviceIndex
                subConfig True
                objVideo.AudioDeviceIndex = -1
                objVideo.CaptureAudio = False
                objVideo.UseVideoFilter = vcxBoth
                objVideo.SetVideoFormat 320, 240
                '''objVideo.SetCrop 70, 0, 180, 240
                If idCamara = idcamara0 Then
                    frmPrincipal.tmrCamPreview.Enabled = False
                    frmPrincipal.objVideo.Connected = False
                    frmPrincipal.objVideo.Preview = False
                End If
                objVideo.Connected = True
                objVideo.Preview = True
                '''objVideo.SetTextOverlay 0, "A1A", 0, 0, "Arial", 10, vbRed, -1
                bCamO = True
            Else
                bCamO = False
                tmrMovimiento.Enabled = False
                MsgBox "No se ha encontrado ninguna cámara conectada en el sistema!", vbInformation
            End If
            If bCamO Then
                tmrCamPreview.Enabled = True
            Else
                tmrCamPreview.Enabled = False
            End If
        Else
            bCamO = False
            tmrMovimiento.Enabled = False
            MsgBox "No se ha encontrado ninguna cámara conectada en el sistema!", vbInformation
        End If
    End If
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-" & Me.name & "_subMuestraCAM"
subLog sERR
End Sub


Attribute VB_Name = "modPrincipal"
Option Explicit
Public objCon As New ADODB.Connection
Public objConCFG As New ADODB.Connection
Public objRst As New ADODB.Recordset
Public objRstO As New ADODB.Recordset
Public objRstA As New ADODB.Recordset
Public objStr As New ADODB.Stream
Public sSql As String

Public objFUNC As clsDatosFUNC
Public objVISI As clsDatosVISI
Public objVISIManual As clsDatosVISI

Public idDepartamento As Integer
Public sFecha1 As String
Public sFecha2 As String
Public sTel As String

Public sAc(3) As String

Public objReporte As clsCrystal10


'Public sDoc As String

'ESTO ES PARA LA VOZ
Public objHabla As New SpeechLib.SpVoice
Public bHabla As Boolean
Public bVoz As Boolean
Public idx As Integer
Public objVoces As SpeechLib.ISpeechObjectTokens
Public SuenaVoz As Boolean
'ESTO ES PARA LA VOZ
Public objPDF As New PDB417lib.PDF417lib
Public objPDFI As IPictureDisp
Public sPDF As String
Public bHuellasU As Boolean
Public bHuella As Boolean
Public bHuellaOrigen As Byte

'''''Huellas UareU
'Device components
Public objUareUs As DPFPReadersCollection
Public objUInf As DPFPReaderDescription

Public objConv As New DPFPSampleConversion
'Engine Components
Public objCreaFea As DPFPFeatureExtraction 'Engine Components
Public objCreaPlantilla As DPFPEnrollment 'Engine Components
'Shared Components
Public objPlantilla As New DPFPTemplate
'Verificación
Public objVerifica As DPFPVerification

Public bModificaHuella As Boolean
Public bModificaFoto As Boolean
Public bHuellaMinucias() As Byte
Public bHuellaMinuciasCAP() As Byte
'''''''
Public bCam As Boolean

Public sERR As String
''
Public bMostrarErrores As Boolean

Public idImpresora As Integer
Public idImpresoraR As Integer
Public sImpresora As String
Public sImpresoraR As String
Public bSensor As Boolean
Public bAuto As Boolean
Public bEnrolaVis As Boolean
Public bStickerF As Boolean
Public bStickerV As Boolean
Public bStickerA As Boolean
Public bStickerFr As Boolean
Public bDatosVisAnterior As Boolean
Public bAutoSalida As Boolean
Public bRestrictHorario As Boolean
Public bAutoStmp As Boolean
Public bIngresoManualVis As Boolean
Public iPuerto_Datos As Integer
Public bPuertaES As Boolean
Public bIntegridad As Boolean
Public bNoESLabor As Boolean
Public bZk_ev As Boolean
Public bAntiPass As Boolean

Public idCamara As Byte
Public idCamaraO As Byte
''
Public bEOF As Boolean

Public bEntrena As VbMsgBoxResult
''''''
Public FlagBalloon As Boolean
Public FormaActiva As String
''''''
Public iDispMODO As Byte
Public iDispPER As Byte
Public bPuertoREG As Long
Public idDisp As Integer
Public idPuerta As Integer
Public bPulsoE As Byte
Public bPulsoS As Byte
Public bDispEnrolaFun As Boolean
Public bDispEnrolaVis As Boolean
Public bDispLogin As Boolean
''
Enum bdModo
    bdNONE
    bdACCESS
    bdSQL
End Enum
Enum EnumAccTipo
    accNONE
    accMANUAL
    acc2D
    accHUELLA
    accLOGIN
    accTARJETA
    accHUELLA_ZK
End Enum
Public AccesoTipo As EnumAccTipo


Public objIni As ARINIManager
Public modoBD As bdModo
Public sBD As String
Public sBDe As String

Public idLogin As Long, idPerf As Long

Public iPhidgetPuertos As Byte
'Public bEnrolando As Boolean
Public objGAATools As New GAATools

Public sTerminal As String

Public idxZK As Integer

Public sHoraAutoSalida As String
Public idControl As Integer
Public zkTMP As String
Public oZKs() As clsZK
Public zkID As Long
Public zkUSR As String
Public lZkErr As Long
Public idHuellaZK As Long
Public idLog As Variant
Public bSinAccess As Boolean
Public Declare Function ConvertBMPtoJPG Lib "ImageUtils.dll" (ByVal InputFile As String, ByVal OutputFile As String, ByVal OverWrite As Boolean, ByVal JPGCompression As Integer, ByVal SaveBMP As Boolean) As Integer
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Function fnFecha(dFecha As Date, bHora As Boolean) As String
On Local Error GoTo errH
fnFecha = Year(dFecha) & "-" & Right("00" & Month(dFecha), 2) & "-" & Right("00" & Day(dFecha), 2)
If bHora Then
    fnFecha = fnFecha & " " & Format(dFecha, "HH:MM:SS")
End If
Exit Function
errH:
fnFecha = vbNullString
End Function
Public Sub subCentraPuntero(ByRef frm As Form, ByRef obj As Object)
Dim barra As Long
barra = frm.Height - frm.ScaleHeight
SetCursorPos frm.ScaleX(frm.Left, vbTwips, vbPixels), frm.ScaleY(frm.Top, vbTwips, vbPixels)
DoEvents
SetCursorPos frm.ScaleX(frm.Left + (obj.Left + (obj.Width / 2)), vbTwips, vbPixels), frm.ScaleY(frm.Top + barra + (obj.Top + (obj.Height / 2)), vbTwips, vbPixels)
End Sub

Public Function fnGuardaFoto(ByRef objCampo As Field, sRuta As String, Optional bBorra As Boolean = True)
On Local Error GoTo errH
If objStr.State = adStateOpen Then objStr.Close
objStr.Type = adTypeBinary
objStr.Open
objStr.LoadFromFile sRuta
objCampo = objStr.Read
objStr.Close
If bBorra Then
    Kill sRuta
    DoEvents
End If
Exit Function
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-fnGuardaFoto"
subLog sERR
End Function
Public Function fnLeeFoto(ByRef objCampo As Field, ByRef objImg As Object)
On Local Error GoTo errH
If IsNull(objCampo) Then Exit Function
If objStr.State = adStateOpen Then objStr.Close
objStr.Type = adTypeBinary
objStr.Open
objStr.Write objCampo
If objGAATools.fnExisteArchivo(App.Path & "\tmp") Then Kill App.Path & "\tmp"
DoEvents
objStr.SaveToFile App.Path & "\tmp"
DoEvents
objStr.Close
objImg.Picture = LoadPicture(App.Path & "\tmp")
DoEvents
If objGAATools.fnExisteArchivo(App.Path & "\tmp") Then Kill App.Path & "\tmp"
DoEvents
Exit Function
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-fnLeeFoto"
subLog sERR
End Function
Public Function fnHablar(sTexto As String)
On Local Error GoTo errH
If bVoz Then
    If bHabla Then
        objHabla.Speak sTexto, SVSFlagsAsync
    Else
        'MsgBox sTexto, vbInformation
    End If
End If
Exit Function
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-fnHablar"
subLog sERR
End Function
Public Sub subEsperar(iSegundos As Integer)
Dim tMr1 As Single
tMr1 = Timer
Do
    DoEvents
Loop Until (Timer - tMr1) > iSegundos
End Sub
Public Sub subLog(sMsg As String)
On Local Error GoTo errH
Dim sLog As String, i As Integer
If bMostrarErrores Then MsgBox sMsg
sLog = App.Path & "\Log"
If Not objGAATools.fnExisteDirectorio(sLog) Then MkDir sLog
i = FreeFile
Open sLog & "\" & Right("00" & Day(Date), 2) & Right("00" & Month(Date), 2) & Year(Date) & "_" & _
Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".log" _
For Output As #i
Print #i, sMsg
Close #i
Exit Sub
errH:

End Sub
Public Function fnMayúscula(sTxt As String) As String
Dim sCad As String, X As Integer
On Local Error GoTo errH
For X = 1 To Len(sTxt)
    If Mid(sTxt, X, 1) = " " Then
        sCad = sCad & Mid(sTxt, X, 1)
        X = X + 1
        sCad = sCad & UCase(Mid(sTxt, X, 1))
    Else
        If X = 1 Then
            sCad = sCad & UCase(Mid(sTxt, X, 1))
        Else
            sCad = sCad & LCase(Mid(sTxt, X, 1))
        End If
    End If
Next X
fnMayúscula = sCad
Exit Function
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-fnMayúscula"
subLog sERR
End Function
Public Function fnEmail(sTxt As String) As Boolean
Dim aP As Integer
Dim X As Integer, pP As Integer
On Local Error GoTo errH
aP = InStr(1, sTxt, "@")
If aP = 0 Then
    fnEmail = False
Else
    For X = 1 To Len(sTxt)
        If Mid(sTxt, X, 1) = "." Then pP = X
    Next X
    If pP > aP Then
        fnEmail = True
        
    Else
        fnEmail = False
    End If
End If
Exit Function
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-fnEmail"
subLog sERR
End Function
Public Sub subConfig(bGraba As Boolean)
Dim iPrn As Byte
On Local Error GoTo errH
sSql = "select * from Config where terminal='" & sTerminal & "'"
With objRst
    If .State = adStateOpen Then .Close
    .Open sSql, objCon, adOpenKeyset, adLockOptimistic
    If .EOF Then
        .AddNew
        !puerto = 1
        !Impresora = ""
        !idCamara = 0
        !idCamaraO = 0
        !terminal = sTerminal
        .UpDate
        .Close
    Else
        If Not bGraba Then
            ''
            If IsNull(!MostrarErrores) Then bMostrarErrores = False Else bMostrarErrores = !MostrarErrores
            ''
            
            idImpresora = -1
            If Printers.Count > 0 Then
                For iPrn = 0 To Printers.Count - 1
                    If Printers.Item(iPrn).DeviceName = "" & !Impresora Then
                        idImpresora = iPrn
                        sImpresora = Printers.Item(iPrn).DeviceName
                        Exit For
                    End If
                Next iPrn
            End If
            idImpresoraR = -1
            If Printers.Count > 0 Then
                For iPrn = 0 To Printers.Count - 1
                    If Printers.Item(iPrn).DeviceName = "" & !Impresorar Then
                        idImpresoraR = iPrn
                        sImpresoraR = Printers.Item(iPrn).DeviceName
                        Exit For
                    End If
                Next iPrn
            End If
            If IsNull(!Sensor) Then bSensor = False Else bSensor = !Sensor
            If IsNull(!Auto) Then bAuto = False Else bAuto = !Auto
            If IsNull(!EnrolaVis) Then bEnrolaVis = False Else bEnrolaVis = !EnrolaVis
            If IsNull(!StickerF) Then bStickerF = False Else bStickerF = !StickerF
            If IsNull(!StickerV) Then bStickerV = False Else bStickerV = !StickerV
            idCamara = Val("" & !idCamara)
            idCamaraO = Val("" & !idCamaraO)
            If IsNull(!datosvisanterior) Then bDatosVisAnterior = False Else bDatosVisAnterior = !datosvisanterior
            If IsNull(!autosalida) Then bAutoSalida = False Else bAutoSalida = !autosalida
            If IsNull(!restricthorario) Then bRestrictHorario = False Else bRestrictHorario = !restricthorario
            If IsNull(!snmanualvis) Then bIngresoManualVis = False Else bIngresoManualVis = !snmanualvis
            sHoraAutoSalida = "" & !hora_autosalida
            If IsNull(!stickera) Then bStickerA = False Else bStickerA = !stickera
            If IsNull(!stickerfr) Then bStickerFr = False Else bStickerFr = !stickerfr
            iPuerto_Datos = Val("" & !puerto_datos)
            If IsNull(!voz) Then bVoz = False Else bVoz = !voz
            If IsNull(!puerta_es) Then bPuertaES = False Else bPuertaES = !puerta_es
            If IsNull(!puerta_integridad) Then bIntegridad = False Else bIntegridad = !puerta_integridad
            If IsNull(!no_es_labor) Then bNoESLabor = False Else bNoESLabor = !no_es_labor
            If IsNull(!Zk_ev) Then bZk_ev = False Else bZk_ev = !Zk_ev
            If IsNull(!antipassback) Then bAntiPass = False Else bAntiPass = !antipassback
            
'            subMuestraCAM
        Else
            !MostrarErrores = bMostrarErrores
            
            !Impresora = sImpresora
            !Impresorar = sImpresoraR
            !Sensor = bSensor
            !Auto = bAuto
            !EnrolaVis = bEnrolaVis
            !StickerF = bStickerF
            !StickerV = bStickerV
            !idCamara = idCamara
            !idCamaraO = idCamaraO
            !datosvisanterior = bDatosVisAnterior
            !autosalida = bAutoSalida
            !restricthorario = bRestrictHorario
            !snmanualvis = bIngresoManualVis
            !hora_autosalida = sHoraAutoSalida
            !stickera = bStickerA
            !stickerfr = bStickerFr
            !puerto_datos = iPuerto_Datos
            !voz = bVoz
            !puerta_es = bPuertaES
            !puerta_integridad = bIntegridad
            !no_es_labor = bNoESLabor
            !Zk_ev = bZk_ev
            !antipassback = bAntiPass
        End If
        .UpDate
        .Close
    End If
End With
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-subGraba"
subLog sERR
End Sub
Public Sub armaRetardos(fDesde As String, fHasta As String, sDOC As String)
On Local Error GoTo errH
Dim idE As Variant
sSql = "select id from templeados where documento='" & sDOC & "'"
Set objRst = objCon.Execute(sSql)
If Not objRst.EOF Then
    idE = Val("" & objRst!id)
    If modoBD = bdSQL Then
        sSql = "select distinct(fecha) as fecha from vfechas where fecha between '" & fDesde & "' and '" & fHasta & "' order by fecha"
    ElseIf modoBD = bdACCESS Then
        sSql = "select distinct(fecha) as fecha from vfechas where fecha between #" & fDesde & "# and #" & fHasta & "# order by fecha"
    End If
    With objRst
        If .State = adStateOpen Then .Close
        .Open sSql, objCon, adOpenForwardOnly
        While Not .EOF
            'sSql = "select entra,sale from tacceso where idtipoper=1 and idtpersona=" & Val(idE) & " and Format([entra],'aaaa-mm-dd')=#" & fnFecha(!fecha, False) & "# order by id"
            If modoBD = bdACCESS Then
                sSql = "select id,idtpersona,entra,sale from tacceso where idtipoper=1 and idtpersona=" & Val(idE) & " and cdate(Format(entra,'yyyy/mm/dd'))=#" & fnFecha(!fecha, False) & "# order by id"
            ElseIf modoBD = bdSQL Then
                sSql = "select id,idtpersona,entra,sale from tacceso where idtipoper=1 and idtpersona=" & Val(idE) & " and convert(date,entra)='" & fnFecha(!fecha, False) & "' order by id"
            End If
            With objRstA
                If .State = adStateOpen Then .Close
                .Open sSql, objCon, adOpenKeyset, adLockReadOnly
                If .RecordCount = 1 Then
                    If modoBD = bdACCESS Then
                        If Not IsNull(!sale) Then
                            sSql = "insert into tmpretardos(fecha,idpersona,a1,a2) select Format(entra,'dd/mm/yyyy'),idtpersona,format(entra,'HH:MM'),format(sale,'HH:MM') from tacceso where id=" & !id
                        Else
                            sSql = "insert into tmpretardos(fecha,idpersona,a1) select Format(entra,'dd/mm/yyyy'),idtpersona,format(entra,'HH:MM') from tacceso where id=" & !id
                        End If
                    ElseIf modoBD = bdSQL Then
                        If Not IsNull(!sale) Then
                            sSql = "insert into tmpretardos(fecha,idpersona,a1,a2) select convert(date,entra),idtpersona,convert(time,entra),convert(time,sale) from tacceso where id=" & !id
                        Else
                            sSql = "insert into tmpretardos(fecha,idpersona,a1) select convert(date,entra),idtpersona,convert(time,entra) from tacceso where id=" & !id
                        End If
                    End If
                    objCon.Execute (sSql)
                ElseIf .RecordCount = 2 Then
                    If modoBD = bdACCESS Then
                        sSql = "insert into tmpretardos(fecha,idpersona,a1,a2) select Format(entra,'dd/mm/yyyy'),idtpersona,format(entra,'HH:MM'),format(sale,'HH:MM') from tacceso where id=" & !id
                    ElseIf modoBD = bdSQL Then
                        sSql = "insert into tmpretardos(fecha,idpersona,a1,a2) select convert(date,entra),idtpersona,convert(time,entra),convert(time,sale) from tacceso where id=" & !id
                    End If
                    objCon.Execute (sSql)
                    .MoveNext
                    If Not IsNull(!entra) Then
                        'sSql = "update tmpretardos set b1=#12/30/1899 17:50:0#"
                        If modoBD = bdACCESS Then
                            sSql = "update tmpretardos set b1=#" & Format(!entra, "HH:MM:SS") & "#"
                            If Not IsNull(!sale) Then
                                sSql = sSql & ",b2='" & Format(!sale, "HH:MM") & "'"
                            End If
                            sSql = sSql & " where fecha=#" & Format(!entra, "yyyy/mm/dd") & "# and idpersona=" & !idtpersona
                        ElseIf modoBD = bdSQL Then
                            sSql = "update tmpretardos set b1='" & Format(!entra, "HH:MM:SS") & "'"
                            If Not IsNull(!sale) Then
                                sSql = sSql & ",b2='" & Format(!sale, "HH:MM") & "'"
                            End If
                            sSql = sSql & " where convert(date,fecha)='" & fnFecha(CDate("" & !entra), False) & "' and idpersona=" & !idtpersona
                        End If
                        objCon.Execute (sSql)
                    End If
                Else
                
                End If
            End With
            
            .MoveNext
        Wend
        .Close
    End With
End If
Exit Sub
errH:
sERR = "Error " & Err.Number & ". " & Err.Description & "-armaRetardos"
subLog sERR
End Sub
Public Function fnConecta() As Boolean
On Local Error GoTo errH
Dim sCad As String
fnConecta = False
objCon.CursorLocation = adUseClient
If modoBD = bdACCESS Then
    sBD = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sBD & ";Jet OLEDB:Database Password=A1AVisitor.15;"
End If
objCon.Open sBD
fnConecta = True
Exit Function
errH:
fnConecta = False
bMostrarErrores = True
subLog Err.Number & ". " & Err.Description & "-fnConecta"
bMostrarErrores = False
End Function
Public Function fnLeerIni(sKey As String) As Variant
On Local Error GoTo errH
Set objIni = New ARINIManager
objIni.INIFile = App.Path & "\INIConf.ini"
fnLeerIni = objIni.GetValue("Config1", sKey, vbNullString)
Set objIni = Nothing
Exit Function
errH:
Set objIni = Nothing
subLog Err.Number & ". " & Err.Description & "_fnLeerIni"
End Function
Public Function fnEscribirIni(sKey As String, sValor As String)
On Local Error GoTo errH
Set objIni = New ARINIManager
objIni.INIFile = App.Path & "\INIConf.ini"
objIni.WriteValue "Config1", sKey, sValor
Set objIni = Nothing
Exit Function
errH:
Set objIni = Nothing
subLog Err.Number & ". " & Err.Description & "_fnEscribirIni"
End Function

Public Function fnNombrePC() As String
Dim strString As String
strString = String(255, Chr$(0))
GetComputerName strString, 255
strString = Left$(strString, InStr(1, strString, Chr$(0)))
fnNombrePC = UCase(Trim(Mid(strString, 1, Len(strString) - 1)))
End Function

Public Sub subMonitor(sTxt, brAntes As Boolean, brDesp As Boolean)
frmMonitor.txtLog.Tag = frmMonitor.txtLog.SelStart
If brAntes Then frmMonitor.txtLog.Text = frmMonitor.txtLog.Text & vbCrLf
frmMonitor.txtLog.Text = frmMonitor.txtLog.Text & sTxt
If brDesp Then frmMonitor.txtLog.Text = frmMonitor.txtLog.Text & vbCrLf
If frmMonitor.chkScroll.Value = vbChecked Then
    frmMonitor.txtLog.SelStart = Len(frmMonitor.txtLog.Text)
Else
    frmMonitor.txtLog.SelStart = Val(frmMonitor.txtLog.Tag)
End If
End Sub

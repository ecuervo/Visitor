Attribute VB_Name = "Key"
Public iDias As Integer
    
Const MAX_PATH = 260
Const INVALID_HANDLE_VALUE = -1

Private Type FILETIME
        dwLowDateTime       As Long
        dwHighDateTime      As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes    As Long
        ftCreationTime      As FILETIME
        ftLastAccessTime    As FILETIME
        ftLastWriteTime     As FILETIME
        nFileSizeHigh       As Long
        nFileSizeLow        As Long
        dwReserved0         As Long
        dwReserved1         As Long
        cFileName           As String * MAX_PATH
        cAlternate          As String * 14
End Type
   
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
   

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Long

Public Const NCBASTAT As Long = &H33
Public Const NCBNAMSZ As Long = 16
Public Const HEAP_ZERO_MEMORY As Long = &H8
Public Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Public Const NCBRESET As Long = &H32

Public Type NET_CONTROL_BLOCK  'NCB
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte ' Reserved, must be 0
   ncb_event      As Long
End Type

Public Type ADAPTER_STATUS
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type

Public Type NAME_BUFFER
   name        As String * NCBNAMSZ
   name_num    As Integer
   name_flags  As Integer
End Type

Public Type ASTAT
   adapt          As ADAPTER_STATUS
   NameBuff(30)   As NAME_BUFFER
End Type

Public Declare Function Netbios Lib "netapi32.dll" _
   (pncb As NET_CONTROL_BLOCK) As Byte

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (hpvDest As Any, ByVal _
    hpvSource As Long, ByVal _
    cbCopy As Long)

Public Declare Function GetProcessHeap Lib "kernel32" () As Long

Public Declare Function HeapAlloc Lib "kernel32" _
    (ByVal hHeap As Long, ByVal dwFlags As Long, _
     ByVal dwBytes As Long) As Long

Public Declare Function HeapFree Lib "kernel32" _
    (ByVal hHeap As Long, _
     ByVal dwFlags As Long, _
     lpMem As Any) As Long


Public PassEncri As String
Public RutaSystem As String
Public RutaWindows As String
Public KeyName As String
Dim Pila As Integer

 Public Function CreateKey(serial As String) As String
 Dim arreglo() As String
 Dim pos As Integer
 Dim pos1 As Integer
 Dim pass As String
 


   arreglo = Split(serial, "-")
   pos = 0
' Robin
  pass = arreglo(UBound(arreglo))
' pass = arreglo(UBound(arreglo))
   Do While pos <= UBound(arreglo)
       pos1 = 1
       Base = ""
       arreglo(pos) = EncryptStr(arreglo(pos), pass)
       Pila = 1
       Do While pos1 <= Len(arreglo(pos))
        Base = Base & CharacterIn(Mid(arreglo(pos), pos1, 1))
       pos1 = pos1 + 1
       Loop
       arreglo(pos) = Base
     pos = pos + 1
   Loop
   pos = 0
   Base = ""
   Do While pos <= UBound(arreglo)
     Base = Base & arreglo(pos) & "-"
     pos = pos + 1
   Loop
   Base = Mid(Base, 1, Len(Base) - 1)
   CreateKey = UCase(Base)
 End Function

Public Function GetMACAddress() As String
    
     Dim Tmp As String, pASTAT As Long
     Dim NCB As NET_CONTROL_BLOCK, AST As ASTAT
     NCB.ncb_command = NCBRESET
     Call Netbios(NCB)


     NCB.ncb_callname = "*               "
     NCB.ncb_command = NCBASTAT

   
     NCB.ncb_lana_num = 0
     NCB.ncb_length = Len(AST)

     pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS _
              Or HEAP_ZERO_MEMORY, NCB.ncb_length)

     If pASTAT = 0 Then
        Debug.Print "memory allocation failed!"
        Exit Function
     End If

     NCB.ncb_buffer = pASTAT
     Call Netbios(NCB)

     CopyMemory AST, NCB.ncb_buffer, Len(AST)

     Tmp = Format$(Hex(AST.adapt.adapter_address(0)), "00") & "-" & _
           Format$(Hex(AST.adapt.adapter_address(1)), "00") & "-" & _
           Format$(Hex(AST.adapt.adapter_address(2)), "00") & "-" & _
           Format$(Hex(AST.adapt.adapter_address(3)), "00") & "-" & _
           Format$(Hex(AST.adapt.adapter_address(4)), "00") & "-" & _
           Format$(Hex(AST.adapt.adapter_address(5)), "00")

     HeapFree GetProcessHeap(), 0, pASTAT

     GetMACAddress = Tmp
End Function


Public Function FileExist(ByVal sFile As String) As Boolean
    'comprobar si existe este fichero
    Dim WFD As WIN32_FIND_DATA
    Dim hFindFile As Long

    hFindFile = FindFirstFile(sFile, WFD)
    'Si no se ha encontrado
    If hFindFile = INVALID_HANDLE_VALUE Then
        FileExist = False
    Else
        FileExist = True
        'Cerrar el handle de FindFirst
        hFindFile = FindClose(hFindFile)
    End If


End Function
'S = Cadena a encriptar
'P = Password
Public Function EncryptStr(ByVal s As String, ByVal P As String) As String
Dim I As Integer, R As String
Dim C1 As Integer, C2 As Integer
R = ""
If Len(P) > 0 Then
For I = 1 To Len(s)
C1 = Asc(Mid(s, I, 1))
If I > Len(P) Then
C2 = Asc(Mid(P, I Mod Len(P) + 1, 1))
Else
C2 = Asc(Mid(P, I, 1))
End If
C1 = C1 + C2 + 64
If C1 > 255 Then C1 = C1 - 256
R = R + Chr(C1)
Next I
Else
R = s
End If
EncryptStr = R
End Function



Private Sub Router()
Dim buf As String
Dim ret As Long
    ' Obtener el directorio de windows
    buf = String$(260, Chr$(0))
    ret = GetWindowsDirectory(buf, Len(buf))
    RutaWindows = Left$(buf, ret)
   
    buf = String$(260, Chr$(0))
    ret = GetSystemDirectory(buf, Len(buf))
    RutaSystem = Left$(buf, ret)
    If FileExist(RutaSystem & "\" & KeyName & ".ini") Then
       FileOpen RutaSystem & "\" & KeyName & ".ini", PassEncri
    Else
       FileSave RutaSystem & "\" & KeyName & ".ini", EncryptStr(Now() - 1, "7")
       FileOpen RutaSystem & "\" & KeyName & ".ini", PassEncri
    End If
       PassEncri = UnEncryptStr(PassEncri, "7")
End Sub

Public Function UnEncryptStr(ByVal s As String, ByVal P As String) As String
Dim I As Integer, R As String
Dim C1 As Integer, C2 As Integer
R = ""
If Len(P) > 0 Then
For I = 1 To Len(s)
C1 = Asc(Mid(s, I, 1))
If I > Len(P) Then
C2 = Asc(Mid(P, I Mod Len(P) + 1, 1))
Else
C2 = Asc(Mid(P, I, 1))
End If
C1 = C1 - C2 - 64
If Sgn(C1) = -1 Then C1 = 256 + C1
R = R + Chr(C1)
Next I
Else
R = s
End If
UnEncryptStr = R
End Function

Public Sub FileSave(Route As String, Data As String)
Dim Punter As Integer
Punter = FreeFile()
On Error GoTo errorgrabararchivo
Open Route For Output As #Punter    ' Abre el archivo para operaciones de salida.
   Print #Punter, Data
   Close #Punter
errorgrabararchivo:
End Sub
Public Sub FileOpen(File As String, ByRef Variable As String)
Dim Punter As Integer
Dim Chapter As String
Dim Total As String
Variable = ""
Total = ""
Punter = FreeFile()
Open File For Input As #Punter
   Do While Not EOF(Punter)   ' Repite el bucle hasta el final del archivo.
      Line Input #Punter, Variable
      Total = Total & Variable
   Loop
Close #Punter
Variable = Total
End Sub

Public Function GenerateSerial() As String
Dim Mac As String
Dim pos As Integer
Dim grupos As Integer
Dim Base As String
Dim clave As String
Dim arreglo() As String
Dim pos1 As Integer
If PassEncri = "" Then

Mac = GetMACAddress & "-" & Mid(Year(Date), 3, 2) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00")
Else
  Mac = GetMACAddress & "-" & Mid(Year(PassEncri), 3, 2) & "-" & Format(Month(PassEncri), "00") & "-" & Format(Day(PassEncri), "00")
End If
Mac = Replace(Mac, "-", "")
pos = 1
Base = ""
clave = ""
grupos = 0
Do While pos <= Len(Mac) And grupos < 5
   If Len(Base) = 5 Then
      clave = clave & Base & "-"
      Base = ""
      grupos = grupos + 1
   Else
     Base = Base & Mid(Mac, pos, 1)
   
   End If
   pos = pos + 1
Loop
If grupos < 5 Then
   If Len(Base) < 5 Then
      Select Case 5 - Len(Base)
      Case 1
         Base = Base & ""
      Case 2
         Base = Base & "A1"
      Case 3
         Base = Base & "A1A"
      Case 4
         Base = Base & "A1AG"
      Case 5
        Base = Base & "A1AGN"
      End Select
      clave = clave & Base & "-"
      Base = ""
      grupos = grupos + 1
   End If
End If
Do While grupos < 5
   pos = Len(Mac)
   Do While pos > 0 And grupos < 5
      If Len(Base) = 5 Then
      clave = clave & Base & "-"
      Base = ""
      grupos = grupos + 1
   Else
     Base = Base & Mid(Mac, pos, 1)
   
   End If
   pos = pos - 1
   
   Loop
Loop
   GenerateSerial = Mid(clave, 1, Len(clave) - 1)
   GenerateSerial = UCase(GenerateSerial)

End Function


 



Private Function CharacterIn(Character As String) As String
Dim Cara As String
Pila = Pila + 1
Select Case LCase(Character)
Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "ñ", "o", "p", "q", "r", "s", "t", "v", "w", "x", "y", "z"
   Cara = Character
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
   Cara = Character
Case Else
     If Asc(Character) < 48 And Asc(Character) > 0 Then
      Cara = CharacterIn(Chr(Asc(Character) + Asc(Character) * Pila))
     Else
     If Asc(Character) > 122 Then
         Cara = CharacterIn(Chr(Asc(Character) - Asc(Character) / Pila))
     Else
         Cara = CharacterIn(Chr(Asc(Character) + 1))
     End If
     End If
End Select
CharacterIn = Cara
End Function
Public Function ValidateKey(KeyIn As String) As Boolean
   Dim serial As String
   Dim Key As String
   
serial = GenerateSerial()
Key = CreateKey(serial)

If UCase(Trim(KeyIn)) = UCase(Trim(Key)) Then
    FileSave RutaSystem & "\DaSiKey.ini", EncryptStr(serial, PassEncri)
    FileSave RutaWindows & "\DaSiKey.ini", EncryptStr(Key, PassEncri)
    ValidateKey = True
Else
  ValidateKey = False
End If
End Function



Public Function ValidateSoft(DayDemo As Integer, NameKey As String) As Integer
' Retorna 1  cuando el sof esta Activo
' Retorna 2  cuando el sof esta pirata
' Retorna 3  cuando el sof esta DEM0
' Retorna 4  cuando el sof esta demo vencido
' DayDemo Cantidadd de dias demo
Dim clave As String
KeyName = NameKey
Router
If FileExist(RutaSystem & "\" & NameKey & ".ini") Then
    If FileExist(RutaSystem & "\DaSiKey.ini") Then
        If FileExist(RutaWindows & "\DaSiKey.ini") Then
            clave = ""
            FileOpen RutaWindows & "\DaSiKey.ini", clave
            clave = UnEncryptStr(clave, PassEncri)
            If ValidateKey(clave) = True Then
                ValidateSoft = 1
            Else
                ValidateSoft = 2
            End If
        Else
            ValidateSoft = 2
        End If
    ElseIf FileExist(RutaWindows & "\DaSiKey.ini") Then
        ValidateSoft = 2
    Else
        iDias = DateDiff("d", PassEncri, Date)
        If iDias > 0 Then
            If iDias < DayDemo Then
                ValidateSoft = 3
            Else
                ValidateSoft = 4
            End If
        Else
            ValidateSoft = 2
        End If
    End If
Else
    ValidateSoft = 3
End If
End Function


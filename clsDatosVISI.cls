VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatosVISI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public idVISI As Variant
Public idAccesoVISI As Variant
Public sDOC As String
Public sNOM As String
Public sAPE As String
Public sSEXO As String
Public sRH As String
Public bEntraVISI As Boolean
Public bFrecuente As Boolean
Public bEspera As Boolean
Public iDispMODO As Byte
Public idDisp As Integer
Public AccesoTipo As EnumAccTipo
Public bREG As Boolean
Public bRegPUERTO As Long
Public bObjetos As Boolean


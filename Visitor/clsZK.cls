VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_ssIP As String
Private m_iiIndex As Integer
Private m_bbConectado As Boolean
Private m_ssNombre As String
Private m_bbEnrola As Boolean


Public Property Get sIP() As String

    sIP = m_ssIP

End Property

Public Property Let sIP(ByVal ssIP As String)

    m_ssIP = ssIP

End Property

Public Property Get iIndex() As Integer

    iIndex = m_iiIndex

End Property

Public Property Let iIndex(ByVal iiIndex As Integer)

    m_iiIndex = iiIndex

End Property

Public Property Get bConectado() As Boolean

    bConectado = m_bbConectado

End Property

Public Property Let bConectado(ByVal bbConectado As Boolean)

    m_bbConectado = bbConectado

End Property

Public Property Get sNombre() As String

    sNombre = m_ssNombre

End Property

Public Property Let sNombre(ByVal ssNombre As String)

    m_ssNombre = ssNombre

End Property

Public Property Get bEnrola() As Boolean

    bEnrola = m_bbEnrola

End Property

Public Property Let bEnrola(ByVal bbEnrola As Boolean)

    m_bbEnrola = bbEnrola

End Property

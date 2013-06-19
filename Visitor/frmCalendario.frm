VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalendario 
   BackColor       =   &H00A95900&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H00A95900&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin MSComCtl2.MonthView mes 
         Height          =   2370
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   12947265
         Appearance      =   1
         StartOfWeek     =   17039361
         CurrentDate     =   39365
      End
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bDevuelve As Boolean
Private Sub Form_Load()
Screen.MousePointer = vbNormal
mes.Value = Date
bDevuelve = False
End Sub

Private Sub mes_DateClick(ByVal DateClicked As Date)
bDevuelve = True
Me.Tag = DateClicked
Me.Hide

End Sub

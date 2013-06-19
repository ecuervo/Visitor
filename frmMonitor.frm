VERSION 5.00
Begin VB.Form frmMonitor 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "A1A monitor de eventos"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   Icon            =   "frmMonitor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CheckBox chkScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Scroll"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   8415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLimpiar_Click()
txtLog.Text = vbNullString
End Sub

Private Sub Form_Load()
Me.Move frmPrincipal.Left + frmPrincipal.Width, frmPrincipal.Top, Screen.Width - frmPrincipal.ScaleWidth, frmPrincipal.Height
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
txtLog.Height = Me.ScaleHeight - 240
txtLog.Width = Me.ScaleWidth
txtLog.Top = 240
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
End Sub

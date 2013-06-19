VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.1#0"; "ARBUTTON.OCX"
Begin VB.Form frmMes 
   BackColor       =   &H00892513&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00892513&
      Caption         =   "Año"
      ForeColor       =   &H00FFFFFF&
      Height          =   7695
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton optAño 
         Appearance      =   0  'Flat
         BackColor       =   &H00892513&
         Caption         =   "Todos"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00892513&
         Caption         =   "Mes"
         ForeColor       =   &H00FFFFFF&
         Height          =   5055
         Left            =   360
         TabIndex        =   18
         Top             =   1800
         Width           =   2055
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            Caption         =   "Todos"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   16
            Top             =   4560
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   15
            Top             =   4200
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   14
            Top             =   3840
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   13
            Top             =   3480
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   12
            Top             =   3120
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   11
            Top             =   2760
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   10
            Top             =   2400
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   9
            Top             =   2040
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   8
            Top             =   1680
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton optMes 
            Appearance      =   0  'Flat
            BackColor       =   &H00892513&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   1815
         End
      End
      Begin VB.OptionButton optAño 
         Appearance      =   0  'Flat
         BackColor       =   &H00892513&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optAño 
         Appearance      =   0  'Flat
         BackColor       =   &H00892513&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optAño 
         Appearance      =   0  'Flat
         BackColor       =   &H00892513&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin ARButtonCtrl.ARButton cmdObjetos 
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   6960
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   873
         Caption         =   "Aceptar"
         ForeColor       =   16777215
         ForeColorOnMouse=   8987923
         BackColorOnMouse=   16777215
         BackColor       =   8987923
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
Attribute VB_Name = "frmMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sAño As String
Public bMes As Byte


Private Sub cmdObjetos_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Dim i As Byte
optAño(0).Caption = Val(Year(Date)) - 1
optAño(1).Caption = Year(Date)
optAño(2).Caption = Val(Year(Date)) + 1
optAño(1).Value = True
For i = 1 To 12
    optMes(i).Caption = MonthName(i)
    If i = Month(Date) Then optMes(i).Value = True
Next i
End Sub


Private Sub optAño_Click(Index As Integer)
sAño = optAño(Index).Caption
End Sub

Private Sub optMes_Click(Index As Integer)
bMes = Index
End Sub

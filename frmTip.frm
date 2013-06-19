VERSION 5.00
Begin VB.Form frmBalloon 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   ControlBox      =   0   'False
   DrawWidth       =   2
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timAutoClose 
      Enabled         =   0   'False
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Image imgDisplayIcon 
      Height          =   240
      Left            =   150
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "frmTip.frx":0000
      Top             =   1185
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   1
      Left            =   1680
      Picture         =   "frmTip.frx":058A
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   2
      Left            =   480
      Picture         =   "frmTip.frx":0B14
      Top             =   1185
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   2
      Left            =   1200
      Picture         =   "frmTip.frx":109E
      Top             =   1185
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "frmTip.frx":1628
      Top             =   1185
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   1
      Left            =   960
      Picture         =   "frmTip.frx":1BB2
      Top             =   1185
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4110
      TabIndex        =   2
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   210
   End
   Begin VB.Image imgX_Dn 
      Height          =   240
      Left            =   4230
      Picture         =   "frmTip.frx":213C
      Top             =   300
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgX_Up 
      Height          =   240
      Left            =   4290
      Picture         =   "frmTip.frx":247E
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "<Title>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblText 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "<Caption>"
      ForeColor       =   &H80000017&
      Height          =   1035
      Left            =   150
      TabIndex        =   0
      Top             =   480
      Width           =   4410
   End
   Begin VB.Image imgX 
      Height          =   240
      Left            =   4080
      Picture         =   "frmTip.frx":27C0
      Top             =   90
      Width           =   240
   End
End
Attribute VB_Name = "frmBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'All variables must be declared

'Dim XY() As POINTAPI

Dim sTahomaOrMsSansSerif As String

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long 'Used to round the corners of the form
    
Private Declare Function CreatePolygonRgn Lib "gdi32" (ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long 'Used to round corners of form

'SetWindowRgn is used when setting the form's shape (rounded corners) so
'Windows knows what the window's region is. That's the area in the window
'where Windows permits drawing, and it won't show any part of the window
'that is outside the window region. hWnd is the handle of the window we're
'working with, hRgn is the region's handle, and bRedraw is the redraw flag.
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Sub RoundCorners()
Attribute RoundCorners.VB_Description = "Rounds the corners of the form via API to create the tooltip effect"
Dim hRgn As Long
Dim lRes As Long
   
    ScaleMode = vbPixels

    hRgn = CreateRoundRectRgn(0, 0, ScaleWidth + 1, ScaleHeight + 1, 28, 28)
    lRes = SetWindowRgn(hWnd, hRgn, True)
End Sub
Private Sub Form_Click()
'Hide me after I'm clicked on
HideBalloon
End Sub

Private Sub Form_Load()
    FlagBalloon = True
End Sub

Private Sub Form_Paint()
    RoundRect hdc, 0, 0, ScaleWidth, ScaleHeight, 26, 26
End Sub

Private Sub Form_Resize()
lblText.Height = Me.ScaleHeight - lblText.Top - 10 'Resize the balloon's text label height
'to fit correctly, no matter what the size of the balloon -- since that can be changed
'The - 10 is to give it a little room on the bottom; without it, it would touch it
'without any space (between the end of the balloon's text label and the bottom)

'Do the same as before, now with the width:
lblText.Width = Me.ScaleWidth - 2 * lblText.Left

'Now, resize the title label's width to fit the balloon size:
lblTitle.Width = Me.ScaleWidth - 2 * lblTitle.Left

'Move the X button
lblX.Left = Me.ScaleWidth - (1.5 * lblX.Width) - 1
imgX.Left = Me.ScaleWidth - (1.5 * imgX.Width)
'lblX.Move (Me.ScaleWidth - lblX.Width) - 13, 5
'imgX.Move (Me.ScaleWidth - lblX.Width) - 15, 2
'imgX_Dn.Move (Me.ScaleWidth - lblX.Width) - 15, 2
'imgX_Up.Move (Me.ScaleWidth - lblX.Width) - 15, 2

RoundCorners ' Round the corners of this form to make it look "tool-tippy"

End Sub

Private Sub Form_Unload(Cancel As Integer)
FlagBalloon = False
End Sub

Private Sub imgDisplayIcon_Click()
' Hide this balloon if I'm clicked
HideBalloon
End Sub

Private Sub imgX_Click()
HideBalloon
End Sub

Private Sub imgX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then imgX.Picture = imgX_Dn.Picture
End Sub

Private Sub imgX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then imgX.Picture = imgX_Up.Picture
End Sub

Private Sub lblText_Click()
'Hide me after I'm clicked on
HideBalloon
End Sub
Private Sub lblTitle_Click()
'Hide me after I'm clicked on
HideBalloon
End Sub
Public Sub SetBalloon(sTitle As String, sText As String, lPosX As Long, lPosY _
    As Long, Optional sIcon As String, Optional bShowClose As Boolean = False, _
    Optional lAutoCloseAfter As Long = 0, Optional lHeight As Long = 1620, _
     Optional lWidth As Long = 4680, Optional sFont As String _
     = "Tahoma, if exists", Optional iFontSize = 8)

'Arguments for this Sub are explained below. What this Sub does is
'set the properties for the balloon to be displayed--text, title, etc.
'After setting the properties, you must show the
'balloon yourself by calling <form_name>.Show
'For example, if this "template" form is frmBalloon, you can create a new
'instance of frmBalloon by doing:
'   Dim frmMyBalloon as New frmBalloon
'and then calling frmMyBalloon.SetBalloon using the values you want, as in:
'   frmMyBalloon.SetBalloon "Sample Title", "Sample Text"
'and going on with the arguments as needed (see below and the declaration
'for this Sub above).
'Then, to show the balloon, call
'   frmMyBalloon.Show , , Me
'The Me argument is needed so that the balloon becomes an owner of
'the form it's called from. Otherwise, it'd be possible for the form to
'get in front of the balloon, still appear when its parent is minimized, etc.,
'which we don't want.

'Here's what the arguments for this Sub do:
'sTitle: The bold title to appear above the text on the balloon (Required)
'sText: Text of balloon (Required)
'lPosX and lPosY: The horizontal and vertical, respectively, positions to
'                 show the ballon at (Required)
'sIcon: The icon to be displayed on the balloon, similar to the messagebox's.
'       They're an "i", "x", or "!". (No question mark here; you can't ask
'       on a balloon, can you?) To specifiy, pass either "i", "x", or "!" as
'       the argument, e.g., SetBalloon("Title", "Text", "!" ...
'       For none, don't pass anything. And, they'll use the XP-style icons
'       by default; to use 9x-looking icons instead, specify "i9", "x9", or "!9"
'       Look at the "template" form (frmBalloon, in my example project) to see what
'       they look like; you should see the difference, but they're quite similar--
'       the XP ones just look more colorful and 3D-ish (Optional)
'bShowClose: Whether or not to show the "X" close button the user can
'            press to close the balloon. If there, click to close the
'            balloon; if it's not there (or if it is) clicking anywhere
'            in the balloon will close it. (Optional)
'lAutoCloseAfter: Specifies the amount of time (in milliseconds) after
'                 which to automatically close the balloon. Setting it
'                 to 0 will make it not automatically close.
'                 E.g., 10,000 is ten seconds. (Optional)
'lHeight and lWidth: The width and height that you want the balloon to have.
'                    It's optional, and it will default to a "normal" size.
'                    If you have a long message, increasing the height should
'                    be good, although you can increase the width if you want, too
'                    (Optional)
'sFont: The font the text will appear in, defaulting to MS Sans Serif or
'       Tahoma. By default, it will automatically check to see if Tahoma
'       exists. If so, it will use it; if not, MS Sans Serif will be used.
'       Tahoma gives it a "new" look, but some early Windows 9x versions
'       may not have it. You can specify any font you want using this
'       argument, however. (Optional)
'Setting TITLE AND CAPTION on balloon:
Me.lblTitle.Caption = sTitle
Me.lblText.Caption = sText

'Setting the X AND Y POSITIONS:
Me.Top = lPosY
Me.Left = lPosX

'Setting the ICON:
'First, convert the case to all lower; that way, since all Select Case
'statements below use lowercase for identification
sIcon = LCase(sIcon)

Select Case sIcon
    Case "i": 'The "i" icon, XP-style (default)
        Me.imgDisplayIcon.Picture = imgIconXP(0).Picture
        
    Case "i9": 'The "i" icon, 9x/Me-style
        imgDisplayIcon.Picture = imgIcon(0).Picture
        
    Case "x": 'The "x" icon, XP-style
        imgDisplayIcon.Picture = imgIconXP(1).Picture
        
    Case "x9": 'The "x" icon, 9x/Me-style
        imgDisplayIcon.Picture = imgIcon(1).Picture
        
    Case "!": 'The "!" icon, XP-style
        imgDisplayIcon.Picture = imgIconXP(2).Picture
        
    Case "!9": 'The "!" icon, 9x-style
        imgDisplayIcon.Picture = imgIcon(2).Picture
        
    Case Else: 'Use no icon
        Me.imgDisplayIcon.Visible = False
        Me.lblTitle.Left = imgDisplayIcon.Left 'Move title over so it looks right
End Select
        
'Showing/not showing THE X BUTTON:
If bShowClose = False Then ' Then don't show the X button
    Me.imgX.Visible = False
    Me.lblX.Visible = False
End If
If bShowClose = True Then ' Then make the X button visible
    Me.imgX.Visible = True
    Me.lblX.Visible = True
End If

'Enabling/disabling AUTO-CLOSE:
    If lAutoCloseAfter = 0 Then ' Then we don't need to auto-close, so ...
        Me.timAutoClose.Enabled = False ' Just make sure the auto-close timer
                                        ' is disabled, since we shouldn't auto-close
    Else    ' Then we DO need to auto-close
        Me.timAutoClose.Interval = lAutoCloseAfter ' Set timer's interval so it will
                                                   ' auto-close at the right time, and...
        Me.timAutoClose.Enabled = True 'Enable the timer, so it will go off and auto-close
    End If
'Setting HEIGHT AND WIDTH:
Me.Width = lWidth
Me.Height = lHeight
RoundCorners

'Setting the FONT AND FONT SIZE:
'If no font specified and using default value, then
If sFont = "Tahoma, if exists" Then
' Check to see if Tahoma exists; if not, use MS Sans Serif; the DoesTahomaExist
'function will return True or False, depending on if Tahoma exists
    If DoesTahomaExist = False Then
        sFont = "MS Sans Serif"
    Else
        sFont = "Tahoma"
    End If
End If

Me.Font = sFont                     'Yeah, I know I really don't have to
Me.lblText.Font = sFont             'reference something on the current form
Me.lblTitle.Font = sFont            'using Me all the time, but it's a weird
Me.lblText.FontSize = iFontSize     'habit I get into sometimes :)
Me.lblTitle.FontSize = iFontSize

End Sub
Public Sub lblX_Click()
    HideBalloon
End Sub
Private Sub lblX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then imgX.Picture = imgX_Dn.Picture
End Sub
Private Sub lblX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then imgX.Picture = imgX_Up.Picture
End Sub
Private Sub timAutoClose_Timer()
' This timer is used to automatically close the balloon, if needed,
' after the specified number of milliseconds

    HideBalloon  'Calls HideBalloon(), which hides the balloon
End Sub
Public Sub HideBalloon()
'HideBalloon() is used to manually hide the balloon and by the
'balloon itself to hide itself when needed
If FormaActiva = "Visitantes" Then
   Unload Me
Else
   Unload Me
End If
End Sub
Public Function DoesTahomaExist()
'This function is is an easy way to determine whether or not a
'font exists by creating a standard font object, assigning a font
'name to it, and checking to see if it does keep that font name (which
'means it does exist, otherwise it'll use a different font or a close
'match).

'This function is hard-coded to check for Tahoma, but if you want to use
'it in another project for something else, you should easily be able to
'modify it.

Dim TestFont As New StdFont 'Create a new standard font object, and ...
TestFont.name = "Tahoma" 'Assign the in-question font name (Tahoma) to it

'Check to see if the font object's name matches that which we are
'questioning exists (Tahoma); if it does match, it exists, and if not,
'it doesn't. Then return the correct value from this function.
If TestFont.name = "Tahoma" Then
    DoesTahomaExist = True
Else
    DoesTahomaExist = False
End If
End Function

Attribute VB_Name = "modBalloon"
Option Explicit

'Public Const SW_SHOWNOACTIVATE = 4
'Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, 'ByVal nCmdShow As Long) As Long 'WILL BE Used to show balloon without "stealing"
                                'focus from window it's called from ... in a
                                'future release of this project
                                
                            
'Used to shape form (round corners)
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
                                 
'Sets window region; used after setting the form's shape
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long 'Used for getting positions of objects/forms
                                'to place balloons correctly

Public Type RECT   'Also used to store values for positions of balloons
   Left As Long    'after using the API to determine where
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long 'Also used for getting positions of
                                     'objects/forms we want to place the
                                     'balloons by
                                 
Public mlWidth As Long
Public mlHeight As Long

Public Type POINTAPI
    X As Long
    y As Long
End Type

Public Type BalloonCoords 'Used to store X and Y coordinates of balloon
    X As Long 'after using API and math operations to figure exact
    y As Long 'coordinates regarding where to place itself
End Type
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
TestFont.Name = "Tahoma" 'Assign the in-question font name (Tahoma) to it

'Check to see if the font object's name matches that which we are
'questioning exists (Tahoma); if it does match, it exists, and if not,
'it doesn't. Then return the correct value from this function.
If TestFont.Name = "Tahoma" Then
    DoesTahomaExist = True
Else
    DoesTahomaExist = False
End If
End Function

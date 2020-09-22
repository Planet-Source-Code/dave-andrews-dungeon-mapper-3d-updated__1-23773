Attribute VB_Name = "modONTop"
Option Explicit
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private MyMousePos As POINTAPI 'for getting the mouse positioning

'This stuff is for keeping windows 'always on top'
Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Sub InitWindow(tForm As Form)
Dim tX As Long
Dim tY As Long
tX = Screen.TwipsPerPixelX
tY = Screen.TwipsPerPixelY
SetWindowPos tForm.hWnd, conHwndTopmost, tForm.Left / tX, tForm.Top / tY, tForm.Width / tX, tForm.Height / tY, conSwpNoActivate Or conSwpShowWindow

End Sub

Sub UnInitWindow(tForm As Form)
Dim tX As Long
Dim tY As Long
tX = Screen.TwipsPerPixelX
tY = Screen.TwipsPerPixelY
SetWindowPos tForm.hWnd, conHwndNoTopmost, tForm.Left / tX, tForm.Top / tY, tForm.Width / tX, tForm.Height / tY, conSwpNoActivate Or conSwpShowWindow

End Sub

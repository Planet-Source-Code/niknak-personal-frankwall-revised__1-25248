Attribute VB_Name = "mod_globals"
Option Explicit

'********************************
'API DECLARATIONS
'********************************
'USED TO KEEP FORM ONTOP
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOSIZE = &H1
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
'********************************
'GLOBAL VARIABLES
'********************************
Public play_beep As Integer     'SOUND OPTION, BEEP ALERT
'********************************


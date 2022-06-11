Attribute VB_Name = "Module1"
Public Const MB_DEFAULTBEEP As Long = -1   ' the default beep sound
Public Const MB_ERROR As Long = 16        ' for critical errors/problems
Public Const MB_WARNING As Long = 48      ' for conditions that might cause problems in the future
Public Const MB_INFORMATION As Long = 64  ' for informative messages only
Public Const MB_QUESTION As Long = 32     ' (no longer recommended to be used)

Public Const HWND_TOPMOST = -1  'za program da bude stalno on top
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


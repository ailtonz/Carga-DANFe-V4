VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dim status As Boolean
'Dim arrest As Variant
'
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
'Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Private Const VK_INSERT = &H2D
'Private Const KEYEVENTF_KEYUP = &H2
'Private Const VK_CONTROL = &H11
'Private Const VK_MENU = &H12
'
'Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
'Private Declare Function SendPOPMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'
'Private Const WM_CLOSE = &H10
'Private Const WM_GETTEXT = &HD
'Private Const WM_GETTEXTLENGTH = &HE
'
'Private Const EM_GETLINE = &HC4
'
'Public compingfile As String
'Public a As Variant
'Public b As Variant
'Public mystr As String
'Public ln1 As Variant
'Public txt As String
'Public condition As Boolean
'
'Private Const MOUSEEVENTF_LEFTDOWN = &H2
'Private Const MOUSEEVENTF_LEFTUP = &H4
'Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
'Private Const MOUSEEVENTF_MIDDLEUP = &H40
'Private Const MOUSEEVENTF_MOVE = &H1
'Private Const MOUSEEVENTF_ABSOLUTE = &H8000
'Private Const MOUSEEVENTF_RIGHTDOWN = &H8
'Private Const MOUSEEVENTF_RIGHTUP = &H10
'
'Private Const WM_SETTEXT = &HC
'
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Private Const VK_ESCAPE = &H1B
'Private Const WM_KEYUP = &H101
'Private Const WM_CHAR = &H102
'
'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
'Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'
'Private Const GW_HWNDNEXT = 2
'Private Const GW_CHILD = 5
'
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
'    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'
'Function Wait(ByVal interval As Integer)
'    Sleep interval
'End Function

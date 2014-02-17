Attribute VB_Name = "ModAPI"
Option Explicit

Public Const ALL_MESSAGES = -1
Public Const ERROR_SUCCESS = 0&

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function ConvCStringToVBString Lib "kernel32" Alias "lstrcpyA" (ByVal lpsz As String, ByVal pt As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lP As Long) As Long
Public Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, lP As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lP As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal m As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal flags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function SendMessageStringString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal iMsg As Long, ByVal str1 As String, ByVal str1 As String) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200
Public Const SC_CP_UTF8 = 65001
Public Const GWL_WNDPROC = (-4)
Public Const MK_RBUTTON = &H2
Public Const MK_LBUTTON = &H1
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_CLIPCHILDREN = &H2000000

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_HOME = &H24
Public Const VK_DOWN = &H28
Public Const VK_END = &H23
Public Const VK_UP = &H26


Public Const WM_USER As Long = &H400
Public Const EM_FORMATRANGE As Long = WM_USER + 57
Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Public Const PHYSICALOFFSETX As Long = 112
Public Const PHYSICALOFFSETY As Long = 113

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4


Public Const KEY_TOGGLED As Integer = &H1
Public Const KEY_PRESSED As Integer = &H1000

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOOWNERZORDER As Long = &H200
Public Const SWP_NOCOPYBITS = &H100
Public Const MK_CONTROL = &H8
Public Const MK_SHIFT = &H4
Public Const VK_SHIFT = &H10&
Public Const VK_CONTROL = &H11&
Public Const VK_MENU = &H12& ' Alt key

Type NMHDR
    hwndFrom As Long
    idFrom As Long
    Code As Long
End Type

Public Type SCNotification
    NotifyHeader As NMHDR
    Position As Long
    ch As Long
    modifiers As Long
    modificationType As Long
    Text As Long
    length As Long
    linesAdded As Long
    message As Long
    wParam As Long
    lParam As Long
    line As Long
    foldLevelNow As Long
    foldLevelPrev As Long
    margin As Long
    listType As Long
    X As Long
    Y As Long
End Type

Public Const CB_FINDSTRING = &H14C

Public Function GetUpper(varArray As Variant) As Long
    Dim Upper As Integer
    On Error Resume Next
    
    Upper = UBound(varArray)
    If Err.Number Then
         If Err.Number = 9 Then
              Upper = 0
         Else
              With Err
                   MsgBox "Error:" & .Number & "-" & .Description
              End With
              Exit Function
         End If
    Else
         Upper = UBound(varArray) + 1
    End If
    
    On Error GoTo 0
    GetUpper = Upper
End Function


Public Function ReplaceChars(ByVal Text As String, ByVal Char As String, ReplaceChar As String) As String
    Dim counter As Integer
    
    counter = 1
    Do
        counter = InStr(counter, Text, Char)
        If counter <> 0 Then
            Mid(Text, counter, Len(ReplaceChar)) = ReplaceChar
          Else
            ReplaceChars = Text
            Exit Do
        End If
    Loop

    ReplaceChars = Text
End Function



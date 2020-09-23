Attribute VB_Name = "Module1"
' used for placing the hook

Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const GWL_HINSTANCE = (-6)
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const HCBT_ACTIVATE = 5
Public Const WH_CBT = 5

Public hHook As Long

' used for locating and changing the buttons
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public ButtonText(0 To 3) As String

' function called by hook
Public Function Manipulate(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim Btn(0 To 3) As Long
    Dim ButtonCount As Integer
    Dim T As Integer
    
    If lMsg = HCBT_ACTIVATE Then
                
        Btn(0) = FindWindowEx(wParam, 0, vbNullString, vbNullString)
        
        Dim cName As String, Length As Long
        For T = 1 To 3
            Btn(T) = FindWindowEx(wParam, Btn(T - 1), vbNullString, vbNullString)
            ' no more windows found
            If Btn(T) = 0 Then Exit For
        Next T
        
        For T = 0 To 3
            If Btn(T) <> 0 And Btn(T) <> wParam Then
                cName = Space(255)
                Length = GetClassName(Btn(T), cName, 255)
                cName = Left(cName, Length)
                Debug.Print cName
                If UCase(cName) = "BUTTON" Then
                    ' a button
                    SetWindowText Btn(T), ButtonText(ButtonCount)
                    ButtonCount = ButtonCount + 1
                End If
            End If
        Next T
        'Release the CBT hook
        UnhookWindowsHookEx hHook
    End If
    Manip = False

End Function



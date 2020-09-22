Attribute VB_Name = "modApi"
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lparam As Any) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    
Public Const WM_RBUTTONUP = &H205
Public Const WH_MOUSE = 7

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type
    
Public gLngMouseHook As Long
    
Public Function MouseHookProc(ByVal nCode As Long, ByVal wParam As Long, mhs As MOUSEHOOKSTRUCT) As Long
Dim strBuffer As String
Dim lngBufferLen As Long
Dim strClassName As String
Dim lngResult As Long

If (nCode >= 0 And wParam = WM_RBUTTONUP) Then

        'Preinitialize string
        strBuffer = Space(255)
        
       ' lngBufferLen = Len(strBuffer)
        
        'This is the string that holds the class name that we are looking for
        strClassName = "Internet Explorer_Server"
        
        Debug.Print strClassName
        
        'Get the classname for the Window that has been clicked, making sure something is returned
        'If the function returns 0, it has failed
        lngResult = GetClassName(mhs.hwnd, strBuffer, Len(strBuffer))
                
        Debug.Print Left$(strBuffer, lngResult)
                
        If lngResult > 0 Then

            'Check to see if the class of the window we clicked on is the same as above
            If Left$(strBuffer, lngResult) = strClassName Then
                
                'Value is the same. Squash the command
                MouseHookProc = 1
                
                Exit Function
                
            End If
            
        End If

    End If

MouseHookProc = CallNextHookEx(gLngMouseHook, nCode, wParam, mhs)
End Function


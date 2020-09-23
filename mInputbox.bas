Attribute VB_Name = "mInputbox"
Option Explicit

Private Declare Function CallNextHookEx _
                Lib "user32" (ByVal hHook As Long, _
                              ByVal ncode As Long, _
                              ByVal wParam As Long, _
                              lParam As Any) As Long

Private Declare Function GetModuleHandle _
                Lib "kernel32" _
                Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function SetWindowsHookEx _
                Lib "user32" _
                Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                           ByVal lpfn As Long, _
                                           ByVal hmod As Long, _
                                           ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function SendDlgItemMessage _
                Lib "user32" _
                Alias "SendDlgItemMessageA" (ByVal hDlg As Long, _
                                             ByVal nIDDlgItem As Long, _
                                             ByVal wMsg As Long, _
                                             ByVal wParam As Long, _
                                             ByVal lParam As Long) As Long

Private Declare Function GetDlgItem Lib "user32.dll" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long

Private Declare Function GetClassName _
                Lib "user32" _
                Alias "GetClassNameA" (ByVal hwnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                                       (ByVal hwnd As Long, ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long
       Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                                    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
                                    
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const EM_LIMITTEXT As Long = &HC5
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000

Private hHook     As Long
Private lMaxLen   As Long
Private lPassChar As Long
Private bNumbersOnly As Boolean

Public Function NewProc(ByVal lngCode As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
    Dim RetVal As Long
    Dim strClassName As String
    Dim lngBuffer As Long
    Dim lWnd As Long
    
    If lngCode < HC_ACTION Then
        NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)

        Exit Function

    End If

    strClassName = String$(256, " ")
    lngBuffer = 255

    If lngCode = HCBT_ACTIVATE Then
        RetVal = GetClassName(wParam, strClassName, lngBuffer)

        If Left$(strClassName, RetVal) = "#32770" Then
            If lPassChar > 0 Then
                SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, lPassChar, &H0
            End If

            If lMaxLen > 0 Then
                SendDlgItemMessage wParam, &H1324, EM_LIMITTEXT, lMaxLen, &H0
            End If

            If bNumbersOnly Then
                lWnd = GetDlgItem(wParam, &H1324)
                If Not lWnd = 0 Then
                    SetWindowLong lWnd, GWL_STYLE, GetWindowLong(lWnd, GWL_STYLE) Or ES_NUMBER
                End If
            End If
        End If
    End If

    CallNextHookEx hHook, lngCode, wParam, lParam
End Function

Public Function InputBoxEx(Prompt As String, _
                           Optional Title As String = "", _
                           Optional Default As String = "", _
                           Optional XPos, _
                           Optional YPos, _
                           Optional HelpFile, _
                           Optional Context, _
                           Optional MaxLen As Long = 0, _
                           Optional PasswordChar As String = "", _
                           Optional NumbersOnly As Boolean = False, _
                           Optional ByRef CancelledByUser As Boolean = False) As String
    Dim lngModHwnd As Long
    Dim lngThreadID As Long

    hHook = 0
    lMaxLen = 0
    lPassChar = 0
    bNumbersOnly = NumbersOnly
    
    If MaxLen > 0 Then
        lMaxLen = MaxLen
    End If

    If Not PasswordChar = "" Then
        lPassChar = Asc(PasswordChar)
    End If

    
    If lPassChar > 0 Or lMaxLen > 0 Or bNumbersOnly = True Then
        lngThreadID = GetCurrentThreadId
        lngModHwnd = GetModuleHandle(vbNullString)
        hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
    End If

    InputBoxEx = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)

    If Not hHook = 0 Then
        UnhookWindowsHookEx hHook
    End If
    
    CancelledByUser = (StrPtr(InputBoxEx) = 0)

End Function

Attribute VB_Name = "mdlHotKey"
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fskey_Modifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

Const WM_HOTKEY = &H312
Const MOD_ALT = &H1
Const MOD_CONTROL = &H2
Const MOD_SHIFT = &H4
Const GWL_WNDPROC = (-4)  '注释：窗口函数的地址

Dim key_preWinProc As Long '注释：用来保存窗口信息
Dim key_Modifiers As Long, key_uVirtKey As Long, key_idHotKey As Long
Dim key_IsWinAddress  As Boolean '注释：是否取得窗口信息的判断


Function keyWndproc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If Msg = WM_HOTKEY Then
        Select Case wParam '注释：wParam 值就是 key_idHotKey
            Case 1 '注释：激活 3 个热键后,3 个热键所对应的操作,大家在其他的程序中，只要修改此处就可以了
                frmMain.countDown 1
            Case 2
                frmMain.countDown 2
            Case 3
                frmMain.countDown 3
            Case 4
                frmMain.countDown 4
            Case 5
                frmMain.countDown 5
            Case 6
                frmMain.countDown 6
            Case 7
                frmMain.countDown 7
            Case 8
                frmMain.countDown 8
            Case 9
                frmMain.countDown 9
        End Select
    End If
    
    
'注释:     将消息传送给指定的窗口
    keyWndproc = CallWindowProc(key_preWinProc, hwnd, Msg, wParam, lParam)
    
End Function

Function SetHotkey(ByVal KeyId As Long, ByVal KeyAss0 As String, ByVal Action As String)
    Dim KeyAss1 As Long
    Dim KeyAss2 As String
    Dim i As Long
    
    i = InStr(1, KeyAss0, ",")
    If i = 0 Then
        KeyAss1 = Val(KeyAss0)
        KeyAss2 = ""
    Else
        KeyAss1 = Right(KeyAss0, Len(KeyAss0) - i)
        KeyAss2 = Left(KeyAss0, i - 1)
    End If
    
    key_idHotKey = 0
    key_Modifiers = 0
    key_uVirtKey = 0
    
    If key_IsWinAddress = False Then  '注释：判断是否需要取得窗口信息，如果重复取得,再最后恢复窗口时，将会造成程序死掉注释:         记录原来的window程序地址
        key_preWinProc = GetWindowLong(frmMain.hwnd, GWL_WNDPROC) '注释:         用自定义程序代替原来的window程序
        SetWindowLong frmMain.hwnd, GWL_WNDPROC, AddressOf keyWndproc
    End If

    key_idHotKey = KeyId
    Select Case Action
        Case "Add"
            If KeyAss2 = "Ctrl" Then key_Modifiers = MOD_CONTROL
            If KeyAss2 = "Alt" Then key_Modifiers = MOD_ALT
            If KeyAss2 = "Shift" Then key_Modifiers = MOD_SHIFT
            If KeyAss2 = "Ctrl+Alt" Then key_Modifiers = MOD_CONTROL + MOD_ALT
            If KeyAss2 = "Ctrl+Shift" Then key_Modifiers = MOD_CONTROL + MOD_SHIFT
            If KeyAss2 = "Ctrl+Alt+Shift" Then key_Modifiers = MOD_CONTROL + MOD_ALT + MOD_SHIFT
            If KeyAss2 = "Shift+Alt" Then key_Modifiers = MOD_SHIFT + MOD_ALT
            key_uVirtKey = Val(KeyAss1)
            RegisterHotKey frmMain.hwnd, key_idHotKey, key_Modifiers, key_uVirtKey '注释：向窗口注册系统热键
            key_IsWinAddress = True '注释：不需要再取得窗口信息
            
        Case "Del"
            SetWindowLong frmMain.hwnd, GWL_WNDPROC, key_preWinProc '注释：恢复窗口信息
            UnregisterHotKey frmMain.hwnd, key_uVirtKey '注释：取消系统热键
            key_IsWinAddress = False '注释：可以再次取得窗口信息
    End Select
End Function

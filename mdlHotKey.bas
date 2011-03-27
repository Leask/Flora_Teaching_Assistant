Attribute VB_Name = "mdlHotKey"
Option Explicit

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const GWL_WNDPROC = (-4)

Public preWinProc As Long
Public Modifiers As Long, uVirtKey1 As Long, idHotKey As Long

Private Type taLong
    ll As Long
End Type

Private Type t2Int
    lWord As Integer
    hword As Integer
End Type


Public Function wndproc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lp As taLong, i2 As t2Int
    
    If Msg = WM_HOTKEY Then
        If wParam = idHotKey Then
            lp.ll = lParam
            LSet i2 = lp
            If (i2.lWord = Modifiers) And i2.hword = uVirtKey1 Then
'------------------------------------------------------
'这里面是快捷键代码，你可以随便改
                frmMain.Visible = Not frmMain.Visible
'------------------------------------------------------
            End If
        End If
    End If
    '如果不是热键信息则调用原来的程序
    wndproc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam)
End Function

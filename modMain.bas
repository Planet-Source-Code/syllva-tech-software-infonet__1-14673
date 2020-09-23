Attribute VB_Name = "modMain"
Option Explicit

'Public type definitions
Public Type WI
    TitleBarText As String
    TitleBarLen As Integer
    hWnd As Long
End Type

'Public API's
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long

'Public variables
Public WinNum As Integer 'holds the number of windows examined
Public CurrentWindows(299) As WI 'holds information about all of the currently open windows


Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim WinInfo As WI  'holds information about the window currently being examined
    Dim retval As Long 'holds the return value
    Dim X As Integer
    
    WinInfo.TitleBarLen = GetWindowTextLength(hWnd) + 1 'find the length of the title bar text of the window currently being examined
    If WinInfo.TitleBarLen > 0 And Len(hWnd) > 1 Then 'if the title bar text of the window currently being examined is at least one character long AND the window's handle is > 1
        WinInfo.TitleBarText = Space(WinInfo.TitleBarLen) 'initialize the variable that will hold the title bar text
        retval = GetWindowText(hWnd, WinInfo.TitleBarText, WinInfo.TitleBarLen) 'retreive the title bar text of the window currently being examined
        WinInfo.hWnd = hWnd 'holds the value of this window's handle
        CurrentWindows(WinNum).hWnd = WinInfo.hWnd 'store this window's handle in the current windows array
        CurrentWindows(WinNum).TitleBarText = WinInfo.TitleBarText 'store this window's title bar text in the current windows array
        WinNum = WinNum + 1 'increment the window counter
    End If
    EnumWindowsProc = 1 'continue enumeration of windows
End Function

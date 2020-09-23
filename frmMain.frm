VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   510
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   2550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   2550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timCheckIE 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   2040
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_CLOSE = &H10

Dim ExistingIEWindows(49) As Long 'holds the handles of all of the currently existing IE windows (50 max)
Dim Flash As Integer 'holds the value that determines if the status text should flash

Private Sub Form_Load()
    Dim X As Integer 'loop variable

    lblStatus.Caption = "Initializing..."
    Flash = 0
    For X = 0 To 49 'reset/initialize the existing IE windows array
        ExistingIEWindows(X) = 0
    Next
    Call GetExistingIEWindows
End Sub

Private Sub GetExistingIEWindows() 'this sub checks to see if any IE windows are currently open, and "remembers" them if so.
    Dim retval As Long 'holds the return value
    Dim X As Integer, Y As Integer 'loop variables
    
    lblStatus.Caption = "Examining currently active system windows..."
    WinNum = 0 'initialize number of windows to zero
    For X = 0 To 199 'reset/initialize the current windows array
        CurrentWindows(X).hWnd = 0
        CurrentWindows(X).TitleBarLen = 0
        CurrentWindows(X).TitleBarText = ""
    Next
    retval = EnumWindows(AddressOf EnumWindowsProc, 0) 'enumerate all open windows
    Y = 0
    For X = 0 To WinNum - 1 'for each window that is currently open
        If InStr(1, CurrentWindows(X).TitleBarText, "Microsoft Internet Explorer", vbTextCompare) > 0 Then 'if this window is an IE window...
            lblStatus.Caption = "Storing IE window handle..."
            ExistingIEWindows(Y) = CurrentWindows(X).hWnd 'add this window to the list of existing IE windows
            Y = Y + 1
        End If
    Next
    If Y > 0 Then 'if any of the existing system windows are IE windows
        lblStatus.Caption = "Enabling popup monitoring..."
        timCheckIE.Enabled = True 'enable the timer that checks for any new IE windows
        lblStatus.Caption = "Monitoring..."
    Else 'if none of the existing system windows are IE windows
        lblStatus.Caption = "No windows found!"
        MsgBox "There are currently no windows open!" & vbLf & vbLf & "Please start Internet Explorer before running this program.", vbExclamation + vbOKOnly, "Error" 'if no IE windows are found, display an error message
        'End 'exit this program
    End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
Me.Caption = "Monitoring..."
Else
Me.Caption = ""
End If
End Sub

Private Sub timCheckIE_Timer()
    Dim retval As Long 'holds the return value
    Dim X As Integer, Y As Integer 'loop variables
    Dim KillCount As Integer 'holds the value that determines if the current window should be killed
    
    WinNum = 0 'initialize number of windows to zero
    For X = 0 To 199 'reset/initialize the current windows array
        CurrentWindows(X).hWnd = 0
        CurrentWindows(X).TitleBarLen = 0
        CurrentWindows(X).TitleBarText = ""
    Next
    retval = EnumWindows(AddressOf EnumWindowsProc, 0) 'enumerate all open windows
    For X = 0 To WinNum - 1 'for each window that is currently open
        If InStr(1, CurrentWindows(X).TitleBarText, "Microsoft Internet Explorer", vbTextCompare) > 0 Then 'if this window is an IE window...
            KillCount = 0
            For Y = 0 To 49
                If ExistingIEWindows(Y) <> 0 Then 'if array value holds a valid handle
                    If ExistingIEWindows(Y) = CurrentWindows(X).hWnd Then 'if the window currently being examined matches any of the existing IE windows
                        KillCount = KillCount + 1 'increment
                    End If
                End If
            Next
            If KillCount = 0 Then 'if an IE window that did not previously exist was found
                retval = PostMessage(CurrentWindows(X).hWnd, WM_CLOSE, ByVal CLng(0), ByVal CLng(0)) 'post the window close message to the newly created IE window's message queue
            End If
        End If
    Next
    
    Flash = Flash + 1 'increment the flash value
    If Flash = 15 Then 'make the status label flash every 0.5 seconds
        Flash = 0
        If lblStatus.Visible = True Then
            lblStatus.Visible = False
        Else
            lblStatus.Visible = True
        End If
    End If
End Sub

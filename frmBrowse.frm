VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowse 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   6495
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nonav As Boolean

Private Sub Form_Load()
wb.Navigate ("about:blank")
End Sub

Private Sub Form_Resize()
wb.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub wb_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Me.Caption = wb.LocationName
nonav = False
End Sub

Private Sub wb_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
    If frmMdiMain.pb.Value = 100 Then
        frmMdiMain.pb.Visible = False
    Else
        frmMdiMain.pb.Visible = True
    End If
    If Progress = -1 Then frmMdiMain.pb.Value = 100
    If Progress > 0 And ProgressMax > 0 Then
        frmMdiMain.pb.Value = Progress * 100 / ProgressMax
    End If

End Sub

Private Sub wb_StatusTextChange(ByVal Text As String)
frmMdiMain.Label1.Caption = Text
End Sub

Public Function ReplaceAll(SourceString As String, ReplaceThis As String, WithThis As String)
'used to clean web addresses
    Dim temp As Variant
    temp = Split(SourceString, ReplaceThis)
    ReplaceAll = Join(temp, WithThis)
End Function


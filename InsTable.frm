VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form InsertTableDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table Attributes"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   285
      Left            =   2055
      TabIndex        =   13
      Top             =   480
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "Cols"
      BuddyDispid     =   196612
      OrigLeft        =   2280
      OrigTop         =   960
      OrigRight       =   2520
      OrigBottom      =   1695
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   2055
      TabIndex        =   12
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "Rows"
      BuddyDispid     =   196613
      OrigLeft        =   3240
      OrigTop         =   360
      OrigRight       =   3480
      OrigBottom      =   1095
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
   End
   Begin VB.TextBox TableCaption 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox CellAttrs 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox TableAttrs 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox Cols 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Rows 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CancelCmd 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton OkCmd 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label CaptionLabel 
      AutoSize        =   -1  'True
      Caption         =   "Caption:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1575
      Width           =   585
   End
   Begin VB.Label CellTagLabel 
      AutoSize        =   -1  'True
      Caption         =   "Cell Tag Attributes:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1215
      Width           =   1335
   End
   Begin VB.Label TableTagLabel 
      AutoSize        =   -1  'True
      Caption         =   "Table Tag Attributes:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   855
      Width           =   1485
   End
   Begin VB.Label ColLabel 
      AutoSize        =   -1  'True
      Caption         =   "Number of columns:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   495
      Width           =   1410
   End
   Begin VB.Label RowLabel 
      AutoSize        =   -1  'True
      Caption         =   "Number of rows:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   1155
   End
End
Attribute VB_Name = "InsertTableDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1999 Microsoft Corporation.
' All rights reserved.
Private tableParam As DEInsertTableParam

Private Sub CancelCmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' create the table parameter object
    Set tableParam = CreateObject("DEInsertTableParam.DEInsertTableParam.1")
    
    Rows = tableParam.NumRows
    Cols = tableParam.NumCols
    TableAttrs = tableParam.TableAttrs
    CellAttrs = tableParam.CellAttrs
    TableCaption = tableParam.Caption

End Sub

Private Sub OkCmd_Click()
    
    If Rows = "" Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    ElseIf IsNumeric(Rows) = False Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    ElseIf Rows <= 0 Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    End If
       
    If Cols = "" Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    ElseIf IsNumeric(Cols) = False Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    ElseIf Cols <= 0 Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    End If
    
    tableParam.NumRows = Rows
    tableParam.NumCols = Cols
    
    If Len(TableAttrs.Text) Then
        tableParam.TableAttrs = TableAttrs.Text
    Else
        tableParam.TableAttrs = ""
    End If
    
    If Len(CellAttrs.Text) Then
        tableParam.CellAttrs = CellAttrs.Text
    Else
        tableParam.CellAttrs = ""
    End If
    
    If Len(TableCaption.Text) Then
        tableParam.Caption = TableCaption.Text
    Else
        tableParam.Caption = ""
    End If
    
    MainForm.DHTMLEdit1.ExecCommand DECMD_INSERTTABLE, OLECMDEXECOPT_DONTPROMPTUSER, tableParam
    Unload Me
End Sub


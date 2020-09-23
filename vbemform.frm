VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MainForm 
   Caption         =   "Designer Lite"
   ClientHeight    =   4815
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1560
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":0120
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":0240
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":0360
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":06B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":07C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":09F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":0D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":1098
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbemform.frx":13BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   688
      _CBWidth        =   8880
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      Child1          =   "FontCombo"
      MinHeight1      =   315
      Width1          =   2130
      NewRow1         =   0   'False
      Child2          =   "FontSizeCombo"
      MinHeight2      =   315
      Width2          =   810
      NewRow2         =   0   'False
      Child3          =   "Toolbar1"
      MinHeight3      =   330
      Width3          =   4935
      NewRow3         =   0   'False
      Begin VB.ComboBox FontSizeCombo 
         Height          =   315
         Left            =   2325
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   30
         Width           =   615
      End
      Begin VB.ComboBox FontCombo 
         Height          =   315
         Left            =   165
         TabIndex        =   5
         Text            =   "FontCombo"
         Top             =   30
         Width           =   1935
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   3165
         TabIndex        =   4
         Top             =   30
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Make Text Bold"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Make Text Italicized"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Underline Text"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LeftJustify"
               Object.ToolTipText     =   "Align Text Left"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Object.ToolTipText     =   "Center Text"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "RightJustify"
               Object.ToolTipText     =   "Align Text Right"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Color"
               Object.ToolTipText     =   "Color Text"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Numbers"
               Object.ToolTipText     =   "Make A Numbered List"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bullets"
               Object.ToolTipText     =   "Make a Bulleted List"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Outdent"
               Object.ToolTipText     =   "Undo Indent"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Indent"
               Object.ToolTipText     =   "Indent Paragraph"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   3255
      Left            =   0
      TabIndex        =   2
      Top             =   390
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5741
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"vbemform.frx":16E0
   End
   Begin DHTMLEDLibCtl.DHTMLEdit DHTMLEdit1 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   3645
      Width           =   7575
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4500
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Block Formatting"
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu FileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu FileNewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu FileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu FileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu FileSaveAs 
         Caption         =   "Save&As..."
      End
      Begin VB.Menu FileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu EditSub 
         Caption         =   "Undo"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu EditSub 
         Caption         =   "Redo"
         Index           =   1
         Shortcut        =   ^Y
      End
      Begin VB.Menu EditSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu EditSub 
         Caption         =   "Cut"
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu EditSub 
         Caption         =   "Copy"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu EditSub 
         Caption         =   "Paste"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu EditSub 
         Caption         =   "Select All"
         Index           =   6
         Shortcut        =   ^A
      End
      Begin VB.Menu EditSub 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu EditSub 
         Caption         =   "Find Text"
         Index           =   8
         Shortcut        =   ^F
      End
      Begin VB.Menu EditSub 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu EditSub 
         Caption         =   "Snap To Grid"
         Index           =   10
      End
   End
   Begin VB.Menu View 
      Caption         =   "&View"
      Begin VB.Menu ViewSub 
         Caption         =   "Borders"
         Index           =   0
      End
      Begin VB.Menu ViewSub 
         Caption         =   "Document Details"
         Index           =   1
      End
      Begin VB.Menu mnuViewBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSource 
         Caption         =   "Source..."
      End
      Begin VB.Menu mnuViewPage 
         Caption         =   "Web Page"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Insert 
      Caption         =   "&Insert"
      Begin VB.Menu InsertSub 
         Caption         =   "Picture..."
         Index           =   0
      End
      Begin VB.Menu InsertSub 
         Caption         =   "Anchor..."
         Index           =   1
      End
      Begin VB.Menu InsertButton 
         Caption         =   "Button"
      End
      Begin VB.Menu InsertHTML 
         Caption         =   "HTML"
      End
   End
   Begin VB.Menu Format 
      Caption         =   "F&ormat"
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu FormatSub 
         Caption         =   ""
         Index           =   15
      End
   End
   Begin VB.Menu Table 
      Caption         =   "T&able"
      Begin VB.Menu TableSub 
         Caption         =   "Insert Table..."
         Index           =   0
      End
      Begin VB.Menu TableSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu TableSub 
         Caption         =   "Insert Row"
         Index           =   2
      End
      Begin VB.Menu TableSub 
         Caption         =   "Insert Column"
         Index           =   3
      End
      Begin VB.Menu TableSub 
         Caption         =   "Insert Cell"
         Index           =   4
      End
      Begin VB.Menu TableSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu TableSub 
         Caption         =   "Delete Rows"
         Index           =   6
      End
      Begin VB.Menu TableSub 
         Caption         =   "Delete Columns"
         Index           =   7
      End
      Begin VB.Menu TableSub 
         Caption         =   "Delete Cells"
         Index           =   8
      End
      Begin VB.Menu TableSub 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu TableSub 
         Caption         =   "Merge Cells"
         Index           =   10
      End
      Begin VB.Menu TableSub 
         Caption         =   "Split Cell"
         Index           =   11
      End
   End
   Begin VB.Menu D2D 
      Caption         =   "&2D"
      Begin VB.Menu D2DSub 
         Caption         =   "Set Position Attribute To Absolute"
         Index           =   0
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Bring To Front"
         Index           =   1
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Send To Back"
         Index           =   2
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Bring Forward"
         Index           =   3
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Send Back"
         Index           =   4
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Bring Above Text"
         Index           =   5
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Send Below Text"
         Index           =   6
      End
      Begin VB.Menu D2DSub 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu D2DSub 
         Caption         =   "Lock Element"
         Index           =   8
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu AboutVBEdit 
         Caption         =   "About Designer Lite..."
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1999 Microsoft Corporation.
' All rights reserved.
' Author: Rick Jesse
Option Explicit

Dim DHTMLEditInitialized As Boolean
Dim fontNames(0 To 5) ' list of fonts is the fontComboBox
Dim fontSizes(0 To 6) ' DHTMLEdit font sizes are 1-7
' Tables for toolbar commands
Dim buttonCmds(1 To 11) As DHTMLEDITCMDID
Dim buttonNames(1 To 11) As String
' Tables for menus commands
Dim editMenuCmds(0 To 8) As DHTMLEDITCMDID
Dim insertMenuCmds(0 To 1) As DHTMLEDITCMDID
Dim tableMenuCmds(0 To 11) As DHTMLEDITCMDID
Dim twoDMenuCmds(0 To 8) As DHTMLEDITCMDID
' Document path name
Dim docPath
' State variables for dynamic context menu
Dim ctxtIs2DCapable As Boolean
Dim ctxtIsAbsPos As Boolean
Dim ctxtIsTable As Boolean
Dim ctxtStdItemCount As Long
Dim ctxt2DItemCount As Long
Dim ctxtTableItemCount As Long

Private Enum General
    DE_E_INVALIDARG = &H5
    DE_E_ACCESS_DENIED = &H46
    DE_E_PATH_NOT_FOUND = &H80070003
    DE_E_FILE_NOT_FOUND = &H80070002
    DE_E_UNEXPECTED = &H8000FFFF
    DE_E_DISK_FULL = &H80070027
    DE_E_NOTSUPPORTED = &H80040100
    DE_E_FILTER_FRAMESET = &H80100001
    DE_E_FILTER_SERVERSCRIPT = &H80100002
    DE_E_FILTER_MULTIPLETAGS = &H80100004
    DE_E_FILTER_SCRIPTLISTING = &H80100008
    DE_E_FILTER_SCRIPTLABEL = &H80100010
    DE_E_FILTER_SCRIPTTEXTAREA = &H80100020
    DE_E_FILTER_SCRIPTSELECT = &H80100040
    DE_E_URL_SYNTAX = &H800401E4
    DE_E_INVALID_URL = &H800C0002
    DE_E_NO_SESSION = &H800C0003
    DE_E_CANNOT_CONNECT = &H800C0004
    DE_E_RESOURCE_NOT_FOUND = &H800C0005
    DE_E_OBJECT_NOT_FOUND = &H800C0006
    DE_E_DATA_NOT_AVAILABLE = &H800C0007
    DE_E_DOWNLOAD_FAILURE = &H800C0008
    DE_E_AUTHENTICATION_REQUIRED = &H800C0009
    DE_E_NO_VALID_MEDIA = &H800C000A
    DE_E_CONNECTION_TIMEOUT = &H800C000B
    DE_E_INVALID_REQUEST = &H800C000C
    DE_E_UNKNOWN_PROTOCOL = &H800C000D
    DE_E_SECURITY_PROBLEM = &H800C000E
    DE_E_CANNOT_LOAD_DATA = &H800C000F
    DE_E_CANNOT_INSTANTIATE_OBJECT = &H800C0010
    DE_E_REDIRECT_FAILED = &H800C0014
    DE_E_REDIRECT_TO_DIR = &H800C0015
    DE_E_CANNOT_LOCK_REQUEST = &H8
End Enum

Private Sub AboutVBEdit_Click()
    'frmAbout.Show vbModal, Me
End Sub

Private Sub D2DSub_Click(Index As Integer)
    Dim cmd As DHTMLEDITCMDID
    Dim state As DHTMLEDITCMDF
    
    cmd = twoDMenuCmds(Index)
           
    If Not cmd = 0 Then
        DHTMLEdit1.ExecCommand cmd, OLECMDEXECOPT_DODEFAULT
    End If

    state = DHTMLEdit1.QueryStatus(DECMD_MAKE_ABSOLUTE)
    
    If state = DECMDF_LATCHED Then
        D2DSub(0).Caption = "Set Position Attribute To 1D"
        D2DSub(0).Enabled = True
    ElseIf state = DECMDF_ENABLED Then
        D2DSub(0).Caption = "Set Position Attribute To Absolute"
        D2DSub(0).Enabled = True
    Else
        D2DSub(0).Caption = "Set Position Attribute To Absolute"
        D2DSub(0).Enabled = False
    End If

    
End Sub

Private Sub DHTMLEdit1_DocumentComplete()
    If Not DHTMLEditInitialized Then
        Dim fmt As DEGetBlockFmtNamesParam
        Dim i As Long
        Dim fontSize As Long
        Dim fmtName As Variant
        
        ' Create the block fmt names holder
        Set fmt = CreateObject("DEGetBlockFmtNamesParam.DEGetBlockFmtNamesParam.1")
        
        ' Get the localized strings for the DECMD_SETBLOCKFMT command
        DHTMLEdit1.ExecCommand DECMD_GETBLOCKFMTNAMES, OLECMDEXECOPT_DONTPROMPTUSER, fmt
        
        ' Put the strings into the Format menu
        i = 0
        For Each fmtName In fmt.Names
            FormatSub(i).Caption = fmtName
            i = i + 1
        Next
        
        UpdateFontCombos
        
        FontSizeCombo.ListIndex = fontSize - 1
        
    End If
    DHTMLEditInitialized = True
End Sub

Private Sub EditSub_Click(Index As Integer)
    Dim cmd As DHTMLEDITCMDID
    Dim state As Boolean
    
    
    If Index = 10 Then
        state = DHTMLEdit1.SnapToGrid
        state = Not state
        DHTMLEdit1.SnapToGrid = state
        EditSub(Index).Checked = state
    Else
        cmd = editMenuCmds(Index)
           
        If Not cmd = 0 Then
            DHTMLEdit1.ExecCommand cmd, OLECMDEXECOPT_DODEFAULT
        End If
        
    End If
        
        
        
        
End Sub

Private Sub FileExit_Click()
    Unload Me
End Sub

Private Sub FileNew_Click()

    If Not SaveChanges = vbCancel Then
        docPath = ""
        DHTMLEdit1.NewDocument
        SetFormCaption
    End If
End Sub

Private Sub FileOpen_Click()


    docPath = ""
    DisableToolbar
    
    If Not SaveChanges = vbCancel Then
    
        On Error Resume Next
        DHTMLEdit1.LoadDocument "", True
        
        If Err.Number < 0 Then
            Dim errMsg As String
            Select Case Err.Number
                Case DE_E_INVALIDARG
                    errMsg = "Invalid argument"
                Case DE_E_PATH_NOT_FOUND
                    errMsg = "Path not found"
                Case DE_E_FILE_NOT_FOUND
                    errMsg = "File not found"
                Case DE_E_ACCESS_DENIED
                    errMsg = "Access denied"
                Case DE_E_UNEXPECTED
                    errMsg = "Unexpected error"
                Case DE_E_FILTER_FRAMESET
                    errMsg = "Document contains a frameset"
                Case DE_E_FILTER_SERVERSCRIPT
                    errMsg = "Document is primarily server side script"
                Case Else
                    errMsg = "Unknown error"
            End Select
            
            MsgBox "Error occurred while loading document: " & errMsg & ".", vbCritical
            DHTMLEdit1.NewDocument
        End If
    End If
    
    If DHTMLEdit1.Busy = False Then
        On Error Resume Next
        ' Force a DisplayChanged event to update toolbar
        ' in case user canceled file open dialog or error occurred
        DHTMLEdit1.DOM.selection.createtextrange.Collapse
    End If
    SetFormCaption
End Sub

Private Sub File_Click()
    If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
        FileSave.Enabled = True
    Else
        FileSave.Enabled = False
    End If
    
End Sub

Private Sub FileSave_Click()
    SaveDocument False
End Sub

Private Sub FileSaveAs_Click()
    SaveDocument True
End Sub

Private Sub FontCombo_Click()
    Dim fn As String
    Dim state As DHTMLEDITCMDF
    
    fn = fontNames(FontCombo.ListIndex)
    
    If (DHTMLEditInitialized) Then
        state = DHTMLEdit1.QueryStatus(DECMD_SETFONTNAME)
        If state >= DECMDF_ENABLED Then
            DHTMLEdit1.ExecCommand DECMD_SETFONTNAME, OLECMDEXECOPT_DONTPROMPTUSER, fn
        End If
    End If
    
End Sub


Private Sub FontCombo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim state As DHTMLEDITCMDF

    ' return if not the return key
    If Not KeyCode = vbKeyReturn Then
        Exit Sub
    End If
    
    If (DHTMLEditInitialized) Then
        state = DHTMLEdit1.QueryStatus(DECMD_SETFONTNAME)
        If state >= DECMDF_ENABLED Then
            ' set the font to what user has typed into the font name combo box
            DHTMLEdit1.ExecCommand DECMD_SETFONTNAME, OLECMDEXECOPT_DONTPROMPTUSER, FontCombo.Text
        End If
    End If
    
End Sub

Private Sub FontSizeCombo_Click()
    Dim fs As Long
    Dim state As DHTMLEDITCMDF
    
    fs = FontSizeCombo.ListIndex
    fs = fs + 1
    
    If (DHTMLEditInitialized) Then
        state = DHTMLEdit1.QueryStatus(DECMD_SETFONTSIZE)
        If state >= DECMDF_ENABLED Then
            DHTMLEdit1.ExecCommand DECMD_SETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER, fs
        End If
    End If
    
End Sub

Private Sub FontSizeCombo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim state As DHTMLEDITCMDF
    Dim fs As String

    ' return if not the return key
    If Not KeyCode = vbKeyReturn Then
        Exit Sub
    End If
    
    If (DHTMLEditInitialized) Then
    
        state = DHTMLEdit1.QueryStatus(DECMD_SETFONTSIZE)
        
        If state >= DECMDF_ENABLED Then
        
            ' remember what's in the combox box so we can reset if the user
            ' typed in something invalid
            
            ' if its mixed selected, display an empty string
            If state = DECMDF_NINCHED Then
                fs = ""
            Else
                fs = DHTMLEdit1.ExecCommand(DECMD_GETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER)
                fs = fs - 1
            End If
            
            ' validate what the user type in
            If IsNumeric(FontSizeCombo.Text) = False Then
                ' didn't type in a valid number
                FontSizeCombo.Text = fs
            ElseIf FontSizeCombo.Text < 1 Or FontSizeCombo.Text > 7 Then
                ' number is out of range
                FontSizeCombo.Text = fs
            Else
                ' set the font size to the number the user typed in
                DHTMLEdit1.ExecCommand DECMD_SETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER, FontSizeCombo.Text
            End If
        End If
    End If
    
End Sub

Private Sub Form_Load()

    DHTMLEditInitialized = False
    
    ' Initialize the font name and size combo boxes
    fontNames(0) = "Times New Roman"
    fontNames(1) = "Arial"
    fontNames(2) = "Tahoma"
    fontNames(3) = "Courier"
    fontNames(4) = "Verdana"
    fontNames(5) = "Wingdings"
    
    fontSizes(0) = "1"
    fontSizes(1) = "2"
    fontSizes(2) = "3"
    fontSizes(3) = "4"
    fontSizes(4) = "5"
    fontSizes(5) = "6"
    fontSizes(6) = "7"
   
    FontCombo.AddItem fontNames(0)
    FontCombo.AddItem fontNames(1)
    FontCombo.AddItem fontNames(2)
    FontCombo.AddItem fontNames(3)
    FontCombo.AddItem fontNames(4)
    FontCombo.AddItem fontNames(5)
    FontCombo.ListIndex = 0
    
    FontSizeCombo.AddItem fontSizes(0)
    FontSizeCombo.AddItem fontSizes(1)
    FontSizeCombo.AddItem fontSizes(2)
    FontSizeCombo.AddItem fontSizes(3)
    FontSizeCombo.AddItem fontSizes(4)
    FontSizeCombo.AddItem fontSizes(5)
    FontSizeCombo.AddItem fontSizes(6)
    FontSizeCombo.ListIndex = 0
    
    InitToolbarTable
    InitMenuTables
    
    DisableToolbar
    SetHeader
DHTMLEdit1.DocumentHTML = rtb.Text
End Sub

Private Sub Form_Resize()
    
    If Not MainForm.WindowState = vbMinimized Then
        DHTMLEdit1.Width = MainForm.Width - 150
        'DHTMLEdit1.Height = MainForm.Height - 1530
        rtb.Move 0, 390
        rtb.Width = MainForm.Width - 150
        'rtb.Height = MainForm.Height - 1530

    End If
End Sub

Private Sub FormatSub_Click(Index As Integer)
    Dim state As DHTMLEDITCMDF
    Dim Format As String
    
    state = DHTMLEdit1.QueryStatus(DECMD_SETBLOCKFMT)
    
    If state >= DECMDF_ENABLED Then
        DHTMLEdit1.ExecCommand DECMD_SETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER, FormatSub(Index).Caption
    End If
    
End Sub

Private Sub Insert_Click()
    Dim cmdIndex As Long
    
    For cmdIndex = LBound(insertMenuCmds) To UBound(insertMenuCmds)
        UpdateMenu InsertSub(cmdIndex), insertMenuCmds(cmdIndex)
    Next cmdIndex
    
    If DHTMLEdit1.DOM.selection.Type = "Control" Then ' a control, table, ActiveX control is selected
        InsertButton.Enabled = False
        InsertHTML.Enabled = False
    Else
        InsertButton.Enabled = True
        InsertHTML.Enabled = True
    End If

End Sub

Private Sub InsertButton_Click()
    Dim doc As Object
    Dim selection As Object
    Dim tr As Object
    ' This routine inserts a button at the current selection
    
    ' Get the DHTML Document object
    Set doc = DHTMLEdit1.DOM
    ' Get the DHTML Selection object
    Set selection = doc.selection
    ' Create a TextRange on the current selection
    Set tr = selection.createrange
    
    tr.pasteHTML ("<BUTTON TITLE=Button>Button!</BUTTON>")
    
End Sub

Private Sub InsertHTML_Click()
    InsertHTMLDlg.Show vbModal, Me
End Sub

Private Sub InsertSub_Click(Index As Integer)
    Dim cmd As DHTMLEDITCMDID
    
    cmd = insertMenuCmds(Index)
           
    If Not cmd = 0 Then
        DHTMLEdit1.ExecCommand cmd, OLECMDEXECOPT_DODEFAULT
    End If


End Sub

Private Sub mnuViewPage_Click()
DHTMLEdit1.DocumentHTML = rtb.Text
rtb.Visible = False
DHTMLEdit1.Visible = True
mnuViewPage.Visible = False
mnuViewSource.Visible = True
End Sub

Private Sub mnuViewSource_Click()
rtb.Text = DHTMLEdit1.DocumentHTML
rtb.Find ("Microsoft DHTML Editing Control")
rtb.SelRTF = "QCS Designer Lite"
rtb.Visible = True
DHTMLEdit1.Visible = False
mnuViewPage.Visible = True
mnuViewSource.Visible = False
End Sub

Private Sub Table_Click()
    Dim cmdIndex As Long
    
    For cmdIndex = LBound(tableMenuCmds) To UBound(tableMenuCmds)
        UpdateMenu TableSub(cmdIndex), tableMenuCmds(cmdIndex)
    Next cmdIndex

End Sub

Private Sub D2D_Click()
    Dim cmdIndex As Long
    Dim state As DHTMLEDITCMDF
    
    For cmdIndex = LBound(twoDMenuCmds) To UBound(twoDMenuCmds)
        UpdateMenu D2DSub(cmdIndex), twoDMenuCmds(cmdIndex)
    Next cmdIndex

    state = DHTMLEdit1.QueryStatus(DECMD_LOCK_ELEMENT)
    If state = DECMDF_LATCHED Then
        D2DSub(8).Checked = True
    Else
        D2DSub(8).Checked = False
    End If

End Sub

Private Sub TableSub_Click(Index As Integer)
    Dim cmd As DHTMLEDITCMDID
    
    If Index = 0 Then
        InsertTableDlg.Show vbModal, Me
    Else
        cmd = tableMenuCmds(Index)
               
        If Not cmd = 0 Then
            DHTMLEdit1.ExecCommand cmd, OLECMDEXECOPT_DODEFAULT
        End If
    End If
    
End Sub

Private Sub DHTMLEdit1_ContextMenuAction(ByVal itemIndex As Long)

    ' Handle user selection on the custom context menu
   Select Case itemIndex
    Case 0
        DHTMLEdit1.ExecCommand DECMD_CUT, OLECMDEXECOPT_DODEFAULT
    Case 1
        DHTMLEdit1.ExecCommand DECMD_COPY, OLECMDEXECOPT_DODEFAULT
    Case 2
        DHTMLEdit1.ExecCommand DECMD_PASTE, OLECMDEXECOPT_DODEFAULT
    Case 4
        DHTMLEdit1.ExecCommand DECMD_SELECTALL, OLECMDEXECOPT_DODEFAULT
    Case 6
        DHTMLEdit1.ExecCommand DECMD_FONT, OLECMDEXECOPT_PROMPTUSER
    End Select
    
    If ctxtIs2DCapable Then
        Select Case itemIndex
        Case ctxtStdItemCount + 2
            DHTMLEdit1.ExecCommand DECMD_MAKE_ABSOLUTE, OLECMDEXECOPT_DODEFAULT
        End Select
    End If
    
    If ctxtIsTable Then
        Select Case itemIndex
        Case ctxtStdItemCount + ctxt2DItemCount + 2
            DHTMLEdit1.ExecCommand DECMD_INSERTROW, OLECMDEXECOPT_DODEFAULT
        Case ctxtStdItemCount + ctxt2DItemCount + 3
            DHTMLEdit1.ExecCommand DECMD_INSERTCOL, OLECMDEXECOPT_DODEFAULT
        Case ctxtStdItemCount + ctxt2DItemCount + 5
            DHTMLEdit1.ExecCommand DECMD_DELETEROWS, OLECMDEXECOPT_DODEFAULT
        Case ctxtStdItemCount + ctxt2DItemCount + 6
            DHTMLEdit1.ExecCommand DECMD_DELETECOLS, OLECMDEXECOPT_DODEFAULT
        End Select
    End If
    
End Sub

Private Sub DHTMLEdit1_ShowContextMenu(ByVal X As Long, ByVal Y As Long)
    Dim cmdState As DHTMLEDITCMDF
    Dim strings() As String
    Dim states() As OLE_TRISTATE
    
   ' Create dynamic context menu that consists of
   ' a "standard" set of items and items that depend
   ' on the currently selected element.
   ' Look at the current selection and
   ' if its a table then add menu items for add/delete rows and cols
   ' if its 2DCapable then add items to toggle its absolute position attribute
   
    ctxtIs2DCapable = False
    ctxtIsAbsPos = False
    ctxtIsTable = False
        
    ' Determine if the selected element is 2D capable
    cmdState = DHTMLEdit1.QueryStatus(DECMD_MAKE_ABSOLUTE)
    If cmdState >= DECMDF_ENABLED Then
        ctxtIs2DCapable = True
    End If
    
    'Use DECMD_SEND_TO_BACK to determine if this element is abs positioned
    cmdState = DHTMLEdit1.QueryStatus(DECMD_SEND_TO_BACK)
    If cmdState >= DECMDF_ENABLED Then
        ctxtIsAbsPos = True
    End If
    
    'Use DECMD_INSERTROW to determine if this element is a table
    cmdState = DHTMLEdit1.QueryStatus(DECMD_INSERTROW)
    If cmdState >= DECMDF_ENABLED Then
        ctxtIsTable = True
    End If
    
    ctxtStdItemCount = 6
    
    If ctxtIs2DCapable Then
        ctxt2DItemCount = 2 '1 Item + 1 Separator
    Else
        ctxt2DItemCount = 0
    End If
    
    
    If ctxtIsTable Then
        ctxtTableItemCount = 6 '4 Items + 2 Separators
    Else
        ctxtTableItemCount = 0
    End If
    
    
    ReDim strings(0 To ctxtStdItemCount + ctxt2DItemCount + ctxtTableItemCount)
    ReDim states(0 To ctxtStdItemCount + ctxt2DItemCount + ctxtTableItemCount)
    
    strings(0) = "Cut"
    strings(1) = "Copy"
    strings(2) = "Paste"
    strings(3) = ""
    strings(4) = "Select All"
    strings(5) = ""
    strings(6) = "Font..."
        
    cmdState = DHTMLEdit1.QueryStatus(DECMD_CUT)
    If cmdState >= DECMDF_ENABLED Then
         states(0) = Unchecked
     Else
         states(0) = Gray
    End If
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_COPY)
    If cmdState >= DECMDF_ENABLED Then
         states(1) = Unchecked
     Else
         states(1) = Gray
    End If
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_PASTE)
    If cmdState >= DECMDF_ENABLED Then
         states(2) = Unchecked
     Else
         states(2) = Gray
    End If
        
    states(3) = Unchecked
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_SELECTALL)
    If cmdState >= DECMDF_ENABLED Then
         states(4) = Unchecked
     Else
         states(4) = Gray
    End If
    
    states(5) = Unchecked
    
    cmdState = DHTMLEdit1.QueryStatus(DECMD_FONT)
    If cmdState >= DECMDF_ENABLED Then
         states(6) = Unchecked
     Else
         states(6) = Gray
    End If
    
    If ctxtIs2DCapable Then
        strings(ctxtStdItemCount + 1) = ""
        states(ctxtStdItemCount + 1) = Unchecked
        If ctxtIsAbsPos Then
            strings(ctxtStdItemCount + 2) = "Make 1D"
        Else
            strings(ctxtStdItemCount + 2) = "Make 2D"
        End If
        states(ctxtStdItemCount + 2) = Unchecked
    End If
    
    If ctxtIsTable Then
        strings(ctxtStdItemCount + ctxt2DItemCount + 1) = ""
        states(ctxtStdItemCount + ctxt2DItemCount + 1) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 2) = "Insert Row"
        states(ctxtStdItemCount + ctxt2DItemCount + 2) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 3) = "Insert Column"
        states(ctxtStdItemCount + ctxt2DItemCount + 3) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 4) = ""
        states(ctxtStdItemCount + ctxt2DItemCount + 4) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 5) = "Delete Row"
        states(ctxtStdItemCount + ctxt2DItemCount + 5) = Unchecked
        strings(ctxtStdItemCount + ctxt2DItemCount + 6) = "Delete Column"
        states(ctxtStdItemCount + ctxt2DItemCount + 6) = Unchecked
        
    End If
    
    DHTMLEdit1.SetContextMenu strings, states
    
End Sub

Private Sub DHTMLEdit1_DisplayChanged()
    Dim state As DHTMLEDITCMDF
    Dim cmd As DHTMLEDITCMDID
    Dim Button As String
    Dim cmds As Long
    
    ' DHTMLEdit indicates the UI should be updated
    ' First update the Toolbar
    For cmds = 1 To 11
        cmd = buttonCmds(cmds)
        Button = buttonNames(cmds)
        state = DHTMLEdit1.QueryStatus(buttonCmds(cmds))
        
        If (state >= DECMDF_ENABLED) Then
            Toolbar1.Buttons(Button).Enabled = True
        Else
            Toolbar1.Buttons(Button).Enabled = False
        End If
            
        If (state = DECMDF_LATCHED) Then
            Toolbar1.Buttons(Button).Value = tbrPressed
        Else
            Toolbar1.Buttons(Button).Value = tbrUnpressed
        End If
    Next cmds
    
    UpdateFontCombos
    
    ' Update the Format menu with the localized strings returned from
    ' the DECMD_GETBLOCKFMT command
    state = DHTMLEdit1.QueryStatus(DECMD_GETBLOCKFMT)
    If state >= DECMDF_ENABLED Then
        Dim blockFmt As String
        blockFmt = DHTMLEdit1.ExecCommand(DECMD_GETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER)
        StatusBar1.Panels(1) = blockFmt
    End If
    
    
End Sub
Private Sub InitToolbarTable()
    ' Initialize parallel arrays for mapping
    ' toolbar buttons to DHTMLEdit commands
    
    ' The toolbar buttons are named in the properties
    ' dialog of the toolbar control. We'll use these
    ' names to select on when the user selects a button
    
    buttonNames(1) = "Bold"
    buttonNames(2) = "Italic"
    buttonNames(3) = "Underline"
    buttonNames(4) = "Numbers"
    buttonNames(5) = "Bullets"
    buttonNames(6) = "Outdent"
    buttonNames(7) = "Indent"
    buttonNames(8) = "LeftJustify"
    buttonNames(9) = "Center"
    buttonNames(10) = "RightJustify"
    buttonNames(11) = "Color"
    
    ' This array is parallel to the names array
    ' We'll use the to dispatch a command when the
    ' user selects a button from the toolbar
    
    buttonCmds(1) = DECMD_BOLD
    buttonCmds(2) = DECMD_ITALIC
    buttonCmds(3) = DECMD_UNDERLINE
    buttonCmds(4) = DECMD_ORDERLIST
    buttonCmds(5) = DECMD_UNORDERLIST
    buttonCmds(6) = DECMD_INDENT
    buttonCmds(7) = DECMD_OUTDENT
    buttonCmds(8) = DECMD_JUSTIFYLEFT
    buttonCmds(9) = DECMD_JUSTIFYCENTER
    buttonCmds(10) = DECMD_JUSTIFYRIGHT
    buttonCmds(11) = DECMD_SETFORECOLOR

End Sub
Private Sub InitMenuTables()
    ' Initialize Edit menu command table
    editMenuCmds(0) = DECMD_UNDO
    editMenuCmds(1) = DECMD_REDO
    editMenuCmds(2) = 0
    editMenuCmds(3) = DECMD_CUT
    editMenuCmds(4) = DECMD_COPY
    editMenuCmds(5) = DECMD_PASTE
    editMenuCmds(6) = DECMD_SELECTALL
    editMenuCmds(7) = 0
    editMenuCmds(8) = DECMD_FINDTEXT
    
    ' Initialize Insert menu command table
    insertMenuCmds(0) = DECMD_IMAGE
    insertMenuCmds(1) = DECMD_HYPERLINK
    
    ' Initialize Insert menu command table
    tableMenuCmds(0) = DECMD_INSERTTABLE
    tableMenuCmds(1) = 0
    tableMenuCmds(2) = DECMD_INSERTROW
    tableMenuCmds(3) = DECMD_INSERTCOL
    tableMenuCmds(4) = DECMD_INSERTCELL
    tableMenuCmds(5) = 0
    tableMenuCmds(6) = DECMD_DELETEROWS
    tableMenuCmds(7) = DECMD_DELETECOLS
    tableMenuCmds(8) = DECMD_DELETECELLS
    tableMenuCmds(9) = 0
    tableMenuCmds(10) = DECMD_MERGECELLS
    tableMenuCmds(11) = DECMD_SPLITCELL
     
    ' Initialize 2D menu command table
    twoDMenuCmds(0) = DECMD_MAKE_ABSOLUTE
    twoDMenuCmds(1) = DECMD_BRING_TO_FRONT
    twoDMenuCmds(2) = DECMD_SEND_TO_BACK
    twoDMenuCmds(3) = DECMD_BRING_FORWARD
    twoDMenuCmds(4) = DECMD_SEND_BACKWARD
    twoDMenuCmds(5) = DECMD_BRING_ABOVE_TEXT
    twoDMenuCmds(6) = DECMD_SEND_BELOW_TEXT
    twoDMenuCmds(7) = 0
    twoDMenuCmds(8) = DECMD_LOCK_ELEMENT
    End Sub

Private Sub Edit_Click()

    Dim cmdIndex As Long
    
    For cmdIndex = LBound(editMenuCmds) To UBound(editMenuCmds)
        UpdateMenu EditSub(cmdIndex), editMenuCmds(cmdIndex)
    Next cmdIndex
        
End Sub

Private Sub UpdateMenu(menu As Control, command As DHTMLEDITCMDID)

    Dim state As DHTMLEDITCMDF

    If Not command = 0 Then
        state = DHTMLEdit1.QueryStatus(command)
        
        If (state >= DECMDF_ENABLED) Then
            menu.Enabled = True
        Else
            menu.Enabled = False
        End If
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Bold"
            DHTMLEdit1.ExecCommand DECMD_BOLD, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Italic"
            DHTMLEdit1.ExecCommand DECMD_ITALIC, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Underline"
            DHTMLEdit1.ExecCommand DECMD_UNDERLINE, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Numbers"
            DHTMLEdit1.ExecCommand DECMD_ORDERLIST, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Bullets"
            DHTMLEdit1.ExecCommand DECMD_UNORDERLIST, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Outdent"
            DHTMLEdit1.ExecCommand DECMD_OUTDENT, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Indent"
            DHTMLEdit1.ExecCommand DECMD_INDENT, OLECMDEXECOPT_DONTPROMPTUSER
        Case "LeftJustify"
             DHTMLEdit1.ExecCommand DECMD_JUSTIFYLEFT, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Center"
            DHTMLEdit1.ExecCommand DECMD_JUSTIFYCENTER, OLECMDEXECOPT_DONTPROMPTUSER
        Case "RightJustify"
            DHTMLEdit1.ExecCommand DECMD_JUSTIFYRIGHT, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Color"
            Dim foreColor As String
            On Error GoTo cleanup
            CommonDialog1.color = 0
            CommonDialog1.CancelError = True
            CommonDialog1.ShowColor
            foreColor = ""
            foreColor = FormatRGBString(CommonDialog1.color)
            DHTMLEdit1.ExecCommand DECMD_SETFORECOLOR, OLECMDEXECOPT_DONTPROMPTUSER, foreColor
        End Select
    
cleanup:


End Sub

Private Sub ViewSub_Click(Index As Integer)
    Dim state As Boolean
    
    ' Toggle different properties on DHTMLEdit.
    ' Check the menu items if the properties are set
    ' to true
    Select Case Index
        Case 0
            state = DHTMLEdit1.ShowBorders
            state = Not state
            DHTMLEdit1.ShowBorders = state
            ViewSub(Index).Checked = state
        Case 1
            state = DHTMLEdit1.ShowDetails
            state = Not state
            DHTMLEdit1.ShowDetails = state
            ViewSub(Index).Checked = state
    End Select
        
End Sub

Private Sub Format_Click()
    Dim state As DHTMLEDITCMDF
    Dim Format As String
    Dim menuItem As Variant
    
    state = DHTMLEdit1.QueryStatus(DECMD_GETBLOCKFMT)
    
    If state >= DECMDF_ENABLED Then
        Format = DHTMLEdit1.ExecCommand(DECMD_GETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER)
        
        For Each menuItem In FormatSub
            
            ' enable menu item
            menuItem.Enabled = True

            ' Check the menu that reflects the
            ' current formatting
            If menuItem.Caption = Format Then
                menuItem.Checked = True
            Else
                menuItem.Checked = False
            End If
            
        Next
    ElseIf state = DECMDF_DISABLED Then
        ' disable format menu menuItems
        For Each menuItem In FormatSub
            menuItem.Enabled = False
            menuItem.Checked = False
        Next
    End If
End Sub

Private Sub SetFormCaption()
    If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
        MainForm.Caption = "VBEdit - " & DHTMLEdit1.CurrentDocumentPath
    Else
        MainForm.Caption = "VBEdit"
    End If
End Sub

Private Function FormatRGBString(val As Long) As String
    Dim color As String
    Dim pad As Long
    Dim r As String
    Dim g As String
    Dim b As String
    
    ' This function formats a long consisting of rgb values
    ' taken from the CommonDialog color dialog
    ' to a string in the form of "#RRGGBB" where RRGGBB are
    ' hex values
    
    ' convert to hex
    color = Hex(val)
    'determine how many zeros to pad in front of converted value
    pad = 6 - Len(color)
    
    If pad Then
        color = String(pad, "0") & color
    End If
        
    'Extract the rgb components
    r = Right(color, 2)
    g = Mid(color, 3, 2)
    b = Left(color, 2)
    
    ' Swab r and b position, color dialog returns
    ' bgr instead of rgb
    color = "#" & r & g & b
    
    FormatRGBString = color
End Function


Private Sub DisableToolbar()
    
    FontCombo.Text = ""
    FontCombo.Enabled = False
    FontSizeCombo.Text = ""
    FontSizeCombo.Enabled = False
    
    Dim b As Object
    For Each b In Toolbar1.Buttons
        b.Enabled = False
    Next

    DoEvents 'give toolbar a chance to update itself
End Sub

Private Sub UpdateFontCombos()
    Dim state As DHTMLEDITCMDF
    
    ' Update the font name combo box on the toolbar
    state = DHTMLEdit1.QueryStatus(DECMD_GETFONTNAME)
    If state = DECMDF_ENABLED Or state = DECMDF_LATCHED Then
        Dim fontName As String
        fontName = DHTMLEdit1.ExecCommand(DECMD_GETFONTNAME, OLECMDEXECOPT_DONTPROMPTUSER)
        FontCombo.Text = fontName
        FontCombo.Enabled = True
    Else
        FontCombo.Text = ""
        If state = DECMDF_NINCHED Then
            FontCombo.Enabled = True
        Else
            FontCombo.Enabled = False
        End If
        
    End If
        
    ' Update the font size combo box on the toolbar
    state = DHTMLEdit1.QueryStatus(DECMD_GETFONTSIZE)
    If state = DECMDF_ENABLED Or state = DECMDF_LATCHED Then
        Dim fontSize As Long
        fontSize = DHTMLEdit1.ExecCommand(DECMD_GETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER)
        If fontSize >= 1 Then
            FontSizeCombo.Text = fontSize
        Else
            FontSizeCombo.Text = ""
        End If
        FontSizeCombo.Enabled = True
    Else
        FontSizeCombo.Text = ""
        If state = DECMDF_NINCHED Then
            FontSizeCombo.Enabled = True
        Else
            FontSizeCombo.Enabled = False
        End If
    End If
    
End Sub

Private Function SaveChanges() As Long
    
    Dim retVal As Long
    If DHTMLEdit1.IsDirty Then
            
        retVal = MsgBox("The current document has changed." & vbCrLf & vbCrLf & "Do you want to save changes?", vbExclamation Or vbYesNoCancel)
    
        Select Case retVal
            Case vbCancel
                SaveChanges = vbCancel
            Case vbYes
                Dim saveSuccess As Boolean
                saveSuccess = False
                If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
                    saveSuccess = SaveDocument(False)
                Else
                    saveSuccess = SaveDocument(True)
                End If
                
                If saveSuccess = True Then
                    SaveChanges = vbOK
                Else
                    SaveChanges = vbCancel
                End If
            
            Case vbNo
                SaveChanges = vbNo
        End Select
    End If
End Function

Private Function SaveDocument(promptUser As Boolean) As Boolean

    SaveDocument = True
    
    DisableToolbar
    
    If promptUser = True Then
        On Error Resume Next
        DHTMLEdit1.SaveDocument "", True
    
    Else
        If Len(DHTMLEdit1.CurrentDocumentPath) > 0 Then
            On Error Resume Next
            DHTMLEdit1.SaveDocument DHTMLEdit1.CurrentDocumentPath
        Else
            Err.Clear
            SaveDocument = False
        End If
    End If
    
    If Err.Number < 0 Then
        Dim errMsg As String
        Select Case Err.Number
            Case DE_E_INVALIDARG
                errMsg = "Invalid argument"
            Case DE_E_PATH_NOT_FOUND
                errMsg = "Path not found"
            Case DE_E_DISK_FULL
                errMsg = "Disk is full"
            Case DE_E_ACCESS_DENIED
                errMsg = "Access denied"
            Case DE_E_UNEXPECTED
                errMsg = "Unexpected error"
            Case Else
                errMsg = "Unknown error"
        End Select
        SaveDocument = False
        MsgBox "Error occurred while saving document: " & errMsg & ".", vbCritical
    End If
        
    On Error Resume Next
    ' Force a DisplayChanged event to update toolbar
    ' in case user canceled file save dialog
    DHTMLEdit1.DOM.selection.createtextrange.Collapse
    SetFormCaption
End Function

Private Sub SetHeader()
rtb.Text = "<HTML>" & vbCrLf & _
"<HEAD>" & vbCrLf & _
"<TITLE>Untitled</TITLE>" & vbCrLf & _
"</HEAD>" & vbCrLf & _
"<BODY>" & vbCrLf & _
"</BODY>" & vbCrLf & _
"</HTML>"
rtb.Find ("<BODY>")
rtb.SelRTF = "<BODY>" & vbCrLf
End Sub

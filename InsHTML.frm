VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form InsertHTMLDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert HTML"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "new"
            Object.ToolTipText     =   "Clear Current HTML"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "bold"
            Object.ToolTipText     =   "Make text Bold"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "italic"
            Object.ToolTipText     =   "Make Text Italic"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "underline"
            Object.ToolTipText     =   "Underline Text"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "strike"
            Object.ToolTipText     =   "Strikethru Text"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "left"
            Object.ToolTipText     =   "Make Text Left Align"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "center"
            Object.ToolTipText     =   "Center Text"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "right"
            Object.ToolTipText     =   "Right Align Text"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox HTMLText 
      Height          =   2805
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "InsHTML.frx":0000
      Top             =   600
      Width           =   5655
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InsHTML.frx":0020
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InsHTML.frx":0132
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InsHTML.frx":0244
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InsHTML.frx":0356
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InsHTML.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InsHTML.frx":057A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InsHTML.frx":068C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InsHTML.frx":079E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "InsertHTMLDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Unload Me
    MainForm.DHTMLEdit1.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim doc As Object
    Dim sel As Object
    Dim tr As Object
    Set doc = MainForm.DHTMLEdit1.DOM
    Set sel = doc.selection
    Set tr = sel.createrange
    tr.pasteHTML (HTMLText.Text)
    Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
Case "new"
HTMLText.Text = ""
Case "bold"
HTMLText.SelText = "<b>" & HTMLText.SelText & "</b>"
Case "italic"
HTMLText.SelText = "<i>" & HTMLText.SelText & "</i>"
Case "underline"
HTMLText.SelText = "<u>" & HTMLText.SelText & "</u>"
Case "strike"
HTMLText.SelText = "<strike>" & HTMLText.SelText & "</strike>"
Case "left"
HTMLText.SelText = "<div align=" & Chr(34) & "left" & Chr(34) & ">" & HTMLText.SelText & "</div>"
Case "center"
HTMLText.SelText = "<center>" & HTMLText.SelText & "</center>"
Case "right"
HTMLText.SelText = "<div align=" & Chr(34) & "right" & Chr(34) & ">" & HTMLText.SelText & "</div>"
End Select
End Sub

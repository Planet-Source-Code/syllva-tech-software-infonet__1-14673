VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "InfoNet"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8880
   Icon            =   "frmMdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   1875
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   8820
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   8880
      Begin RichTextLib.RichTextBox rtbResearch5 
         Height          =   1815
         Left            =   405
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3201
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMdiMain.frx":0442
      End
      Begin RichTextLib.RichTextBox rtbResearch4 
         Height          =   1815
         Left            =   405
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3201
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMdiMain.frx":04FC
      End
      Begin RichTextLib.RichTextBox rtbResearch3 
         Height          =   1815
         Left            =   405
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3201
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMdiMain.frx":05B6
      End
      Begin RichTextLib.RichTextBox rtbResearch2 
         Height          =   1815
         Left            =   405
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3201
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMdiMain.frx":0670
      End
      Begin VB.CommandButton Command4 
         Height          =   400
         Left            =   0
         Picture         =   "frmMdiMain.frx":072A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Paste Selected Text"
         Top             =   1215
         Width           =   400
      End
      Begin VB.CommandButton Command3 
         Height          =   400
         Left            =   0
         Picture         =   "frmMdiMain.frx":12DC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Copy Selected Text"
         Top             =   810
         Width           =   400
      End
      Begin VB.CommandButton Command2 
         Height          =   400
         Left            =   0
         Picture         =   "frmMdiMain.frx":1826
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cut Selected Text"
         Top             =   405
         Width           =   400
      End
      Begin VB.CommandButton Command1 
         Height          =   400
         Left            =   0
         Picture         =   "frmMdiMain.frx":1D70
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Delete Selected Text"
         Top             =   0
         Width           =   400
      End
      Begin RichTextLib.RichTextBox rtbResearch 
         Height          =   1815
         Left            =   405
         TabIndex        =   9
         Top             =   0
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3201
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmMdiMain.frx":22BA
      End
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   7
      Top             =   795
      Visible         =   0   'False
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1402
      _CBWidth        =   8880
      _CBHeight       =   795
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar2"
      MinHeight1      =   390
      Width1          =   6255
      NewRow1         =   0   'False
      Child2          =   "Combo3"
      MinHeight2      =   315
      Width2          =   1650
      NewRow2         =   -1  'True
      Child3          =   "Combo4"
      MinHeight3      =   315
      Width3          =   2775
      NewRow3         =   0   'False
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1890
         TabIndex        =   17
         Top             =   480
         Width           =   9615
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1845
         TabIndex        =   16
         Top             =   450
         Width           =   6945
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmMdiMain.frx":2374
         Left            =   165
         List            =   "frmMdiMain.frx":2376
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   450
         Width           =   1455
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   165
         TabIndex        =   14
         Top             =   30
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   688
         ButtonWidth     =   714
         ButtonHeight    =   688
         Style           =   1
         ImageList       =   "imgCold"
         HotImageList    =   "imgHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "new"
               Object.ToolTipText     =   "Open New Browser Window"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "back"
               Object.ToolTipText     =   "Back"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "forward"
               Object.ToolTipText     =   "Forward"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "stop"
               Object.ToolTipText     =   "Stop"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "refresh"
               Object.ToolTipText     =   "Refresh"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "home"
               Object.ToolTipText     =   "Home"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "search"
               Object.ToolTipText     =   "Search the Web"
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "altavista"
                     Text            =   "AltaVista"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "askjeeves"
                     Text            =   "Ask Jeeves"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "mysimon"
                     Text            =   "MySimon.com"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "yahoo"
                     Text            =   "Yahoo!"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "favorites"
               Object.ToolTipText     =   "Favorites"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "history"
               Object.ToolTipText     =   "View Current History"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "print"
               Object.ToolTipText     =   "Print"
               ImageIndex      =   10
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "webpage"
                     Text            =   "Web Page"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "retrievedinfo"
                     Text            =   "Retrieved Information"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "copy"
               Object.ToolTipText     =   "Copy"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "paste"
               Object.ToolTipText     =   "Paste"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "find"
               Object.ToolTipText     =   "Find (On this page)"
               ImageIndex      =   19
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   8820
      TabIndex        =   4
      Top             =   7995
      Width           =   8880
      Begin ComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   7560
         TabIndex        =   18
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   45
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1402
      _CBWidth        =   8880
      _CBHeight       =   795
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   390
      Width1          =   1575
      NewRow1         =   0   'False
      Child2          =   "Combo2"
      MinHeight2      =   315
      Width2          =   1530
      NewRow2         =   -1  'True
      Child3          =   "Combo1"
      MinHeight3      =   315
      Width3          =   1125
      NewRow3         =   0   'False
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMdiMain.frx":2378
         Left            =   165
         List            =   "frmMdiMain.frx":2391
         TabIndex        =   6
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   9645
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1725
         TabIndex        =   3
         Top             =   450
         Width           =   7065
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   688
         ButtonWidth     =   714
         ButtonHeight    =   688
         Style           =   1
         ImageList       =   "imgCold"
         HotImageList    =   "imgHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "back"
               Object.ToolTipText     =   "Back"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "forward"
               Object.ToolTipText     =   "Forward"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "stop"
               Object.ToolTipText     =   "Stop"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "refresh"
               Object.ToolTipText     =   "Refresh"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "home"
               Object.ToolTipText     =   "Home"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "search"
               Object.ToolTipText     =   "Search"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "favorites"
               Object.ToolTipText     =   "Favorites"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "history"
               Object.ToolTipText     =   "History"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "email"
               Object.ToolTipText     =   "Email"
               ImageIndex      =   9
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "readmail"
                     Text            =   "Read Mail"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "newmail"
                     Text            =   "Write New Mail"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "sendpage"
                     Text            =   "Send Page"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "sendlink"
                     Text            =   "Send Link"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "print"
               Object.ToolTipText     =   "Print"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "find"
               Object.ToolTipText     =   "Find (On this page)"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "copy"
               Object.ToolTipText     =   "Copy"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "edit"
               Object.ToolTipText     =   "Edit"
               ImageIndex      =   12
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "sendinfonet"
                     Text            =   "Send to Editor"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "sendword"
                     Text            =   "Send to MS Word"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "sendnotepad"
                     Text            =   "Send To Notepad"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "designerlite"
                     Text            =   "Make Web Page"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "fontsize"
               Object.ToolTipText     =   "Change Font Size"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "fullscreen"
               Object.ToolTipText     =   "Show Browser as Full Screen"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "properties"
               Object.ToolTipText     =   "Internet Options"
               ImageIndex      =   14
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgHot 
      Left            =   2400
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":23CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":292B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":2E87
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":33E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":393F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":3E9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":43F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":4953
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":4EAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":540B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":5FCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":652B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":6A87
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":6FE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":753F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":7A9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":8777
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":933B
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":A017
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgCold 
      Left            =   1800
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":A57B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":AAD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":B033
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":B58F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":BAEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":C047
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":C5A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":CAFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":D05B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":D5B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":DB13
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":E06F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":E5CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":EB27
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":F083
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":F5DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":102BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":10817
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":114F3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3000
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Begin VB.Menu mnuFileNewDocument1 
            Caption         =   "Document 1"
         End
         Begin VB.Menu mnuFileNewDocument2 
            Caption         =   "Document 2"
         End
         Begin VB.Menu mnuFileNewDocument3 
            Caption         =   "Document 3"
         End
         Begin VB.Menu mnuFileNewDocument4 
            Caption         =   "Document 4"
         End
         Begin VB.Menu mnuFileNewDocument5 
            Caption         =   "Document 5"
         End
         Begin VB.Menu mnuFileBar1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileNewBrowser 
            Caption         =   "Browser"
         End
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewDoc1 
         Caption         =   "Document 1"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewDoc2 
         Caption         =   "Document 2"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewDoc3 
         Caption         =   "Document 3"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewDoc4 
         Caption         =   "Document 4"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewDoc5 
         Caption         =   "Document 5"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFullScreen 
         Caption         =   "FullScreen"
      End
      Begin VB.Menu mnuViewSource 
         Caption         =   "Source"
      End
   End
   Begin VB.Menu mnuResearch 
      Caption         =   "Research"
   End
   Begin VB.Menu mnuBrowse 
      Caption         =   "Browse"
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmMdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nonav As Boolean

Private Sub Combo2_Change()
SearchEngines2
End Sub

Private Sub Combo3_Click()
SearchEngines
End Sub

Private Sub Command1_Click()
rtbResearch.SelText = ""
End Sub

Private Sub Command2_Click()
Clipboard.SetText rtbResearch.SelText
rtbResearch.SelText = ""
End Sub

Private Sub Command3_Click()
Clipboard.SetText rtbResearch.SelText
End Sub

Private Sub Command4_Click()
rtbResearch.SelText = Clipboard.GetText
End Sub

Private Sub MDIForm_Load()
'Combo2.ListIndex = 3
Combo3.AddItem "AltaVista"
Combo3.AddItem "Yahoo!"
Combo3.AddItem "Ask Jeeves"
Combo3.AddItem "DogPile"
Combo3.AddItem "Lycos"
Combo3.AddItem "Excite"

Combo2.AddItem "AltaVista"
Combo2.AddItem "Yahoo!"
Combo2.AddItem "Ask Jeeves"
Combo2.AddItem "DogPile"
Combo2.AddItem "Lycos"
Combo2.AddItem "Excite"

'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
'Combo3.AddItem "AltaVista"
End Sub

Private Sub MDIForm_Resize()
rtbResearch.Width = Picture2.Width - 450
rtbResearch2.Width = Picture2.Width - 450
rtbResearch3.Width = Picture2.Width - 450
rtbResearch4.Width = Picture2.Width - 450
rtbResearch5.Width = Picture2.Width - 450
End Sub

Private Sub mnuBrowse_Click()
Picture2.Visible = False
CoolBar1.Visible = True
CoolBar2.Visible = False
mnuResearch.Enabled = True
mnuBrowse.Enabled = False
End Sub

Private Sub mnuFileNewBrowser_Click()
LoadNewDoc
ActiveForm.wb.GoSearch
End Sub

Private Sub mnuFileNewDocument1_Click()
rtbResearch.Visible = True
rtbResearch2.Visible = False
rtbResearch3.Visible = False
rtbResearch4.Visible = False
rtbResearch5.Visible = False
mnuFileNewDocument1.Enabled = False
mnuViewDoc1.Enabled = True
End Sub

Private Sub mnuFileNewDocument2_Click()
rtbResearch.Visible = False
rtbResearch2.Visible = True
rtbResearch3.Visible = False
rtbResearch4.Visible = False
rtbResearch5.Visible = False
mnuFileNewDocument2.Enabled = False
mnuViewDoc2.Enabled = True
End Sub

Private Sub mnuFileNewDocument3_Click()
rtbResearch.Visible = False
rtbResearch2.Visible = False
rtbResearch3.Visible = True
rtbResearch4.Visible = False
rtbResearch5.Visible = False
mnuFileNewDocument3.Enabled = False
mnuViewDoc3.Enabled = True
End Sub

Private Sub mnuFileNewDocument4_Click()
rtbResearch.Visible = False
rtbResearch2.Visible = False
rtbResearch3.Visible = False
rtbResearch4.Visible = True
rtbResearch5.Visible = False
mnuFileNewDocument4.Enabled = False
mnuViewDoc4.Enabled = True
End Sub

Private Sub mnuFileNewDocument5_Click()
rtbResearch.Visible = False
rtbResearch2.Visible = False
rtbResearch3.Visible = False
rtbResearch4.Visible = False
rtbResearch5.Visible = True
mnuFileNewDocument5.Enabled = False
mnuViewDoc5.Enabled = True
End Sub

Private Sub mnuFileOpen_Click()
'Standard common dialog stuff
    On Error GoTo woops
        With cdl
           .DialogTitle = "Open Local Web Page"
           .CancelError = True
           .Filter = "Web Pages (*.htm;*.html)|*.htm;*.html|All files (*.*)|*.*"
           .ShowOpen
        If Len(.Filename) = 0 Then Exit Sub
        If FileExists(.Filename) Then ActiveForm.wb.Navigate .Filename
        End With
woops:
End Sub

Private Sub mnuFileSave_Click()
On Error GoTo woops
If ActiveForm.wb.LocationURL = "" Then Exit Sub
ActiveForm.wb.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
woops:
End Sub

Private Sub mnuResearch_Click()
Picture2.Visible = True
CoolBar1.Visible = False
CoolBar2.Visible = True
mnuResearch.Enabled = False
mnuBrowse.Enabled = True
End Sub

Private Sub mnuViewDoc1_Click()
rtbResearch.Visible = True
rtbResearch2.Visible = False
rtbResearch3.Visible = False
rtbResearch4.Visible = False
rtbResearch5.Visible = False
End Sub

Private Sub mnuViewDoc2_Click()
rtbResearch.Visible = False
rtbResearch2.Visible = True
rtbResearch3.Visible = False
rtbResearch4.Visible = False
rtbResearch5.Visible = False
End Sub

Private Sub mnuViewDoc3_Click()
rtbResearch.Visible = False
rtbResearch2.Visible = False
rtbResearch3.Visible = True
rtbResearch4.Visible = False
rtbResearch5.Visible = False
End Sub

Private Sub mnuViewDoc4_Click()
rtbResearch.Visible = False
rtbResearch2.Visible = False
rtbResearch3.Visible = False
rtbResearch4.Visible = True
rtbResearch5.Visible = False
End Sub

Private Sub mnuViewDoc5_Click()
rtbResearch.Visible = False
rtbResearch2.Visible = False
rtbResearch3.Visible = False
rtbResearch4.Visible = False
rtbResearch5.Visible = True
End Sub

Private Sub mnuViewFullScreen_Click()
Load frmFullScreen
frmFullScreen.wbFull.Navigate ActiveForm.wb.LocationURL
frmFullScreen.Show
Me.WindowState = vbMinimized
End Sub

Private Sub mnuViewSource_Click()
Load frmViewSource
frmViewSource.txtURL.Text = frmMdiMain.ActiveForm.wb.LocationURL
frmViewSource.Show
End Sub

Private Sub rtbResearch_Change()
If rtbResearch.Text = "" Then
mnuFileNewDocument1.Enabled = True
mnuViewDoc1.Enabled = False
Else
mnuFileNewDocument1.Enabled = False
mnuViewDoc1.Enabled = True
End If
End Sub

Private Sub rtbResearch2_Change()
If rtbResearch2.Text = "" Then
mnuFileNewDocument2.Enabled = True
mnuViewDoc2.Enabled = False
Else
mnuFileNewDocument2.Enabled = False
mnuViewDoc2.Enabled = True
End If
End Sub

Private Sub rtbResearch3_Change()
If rtbResearch3.Text = "" Then
mnuFileNewDocument3.Enabled = True
mnuViewDoc3.Enabled = False
Else
mnuFileNewDocument3.Enabled = False
mnuViewDoc3.Enabled = True
End If
End Sub

Private Sub rtbResearch4_Change()
If rtbResearch4.Text = "" Then
mnuFileNewDocument4.Enabled = True
mnuViewDoc4.Enabled = False
Else
mnuFileNewDocument4.Enabled = False
mnuViewDoc4.Enabled = True
End If
End Sub

Private Sub rtbResearch5_Change()
If rtbResearch5.Text = "" Then
mnuFileNewDocument5.Enabled = True
mnuViewDoc5.Enabled = False
Else
mnuFileNewDocument5.Enabled = False
mnuViewDoc5.Enabled = True
End If
End Sub

Private Sub Text1_GotFocus()
SendKeys String:="{HOME}+{END}", Wait:=True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If nonav = False Then
    'when a new listitem is clicked - go there
        ActiveForm.wb.Navigate2 Text1.Text
End If
'ActiveForm.wb.Navigate Text1.Text
Combo1.AddItem Text1.Text
KeyAscii = 0
End If
End Sub

Private Sub Text2_Click()
SendKeys String:="{HOME}+{END}", Wait:=True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If nonav = False Then
    'when a new listitem is clicked - go there
        ActiveForm.wb.Navigate2 Text2.Text
End If
'ActiveForm.wb.Navigate Text2.Text
Combo4.AddItem Text2.Text
frmMdiMain.Combo1.AddItem Text2.Text
KeyAscii = 0
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "back"
ActiveForm.wb.GoBack

Case "forward"
ActiveForm.wb.GoForward

Case "stop"
ActiveForm.wb.Stop

Case "refresh"
ActiveForm.wb.Refresh

Case "home"
ActiveForm.wb.GoHome

Case "search"
ActiveForm.wb.GoSearch

Case "favorites"

Case "history"

Case "email"

Case "print"

Case "copy"
    ActiveForm.wb.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER

Case "edit"

Case "fontsize"
'ActiveForm.wb.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER

Case "properties"

Case "find"
ActiveForm.wb.SetFocus
SendKeys "^f", True

End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
Case "designerlite"
Load MainForm
MainForm.Show
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "new"
    LoadNewDoc
    ActiveForm.wb.GoSearch
Case "back"
ActiveForm.wb.GoBack

Case "forward"
ActiveForm.wb.GoForward

Case "stop"
ActiveForm.wb.Stop

Case "refresh"
ActiveForm.wb.Refresh

Case "home"
ActiveForm.wb.GoHome

Case "search"
ActiveForm.wb.GoSearch


Case "find"
ActiveForm.wb.SetFocus
SendKeys "^f", True

End Select
End Sub


Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
Case "altavista"
    LoadNewDoc
ActiveForm.wb.Navigate ("http://www.altavista.com")

Case "askjeeves"
    LoadNewDoc
ActiveForm.wb.Navigate ("http://www.askjeeves.com")

Case "mysimon"
    LoadNewDoc
ActiveForm.wb.Navigate ("http://www.mysimon.com")

Case "yahoo"
    LoadNewDoc
ActiveForm.wb.Navigate ("http://www.yahoo.com")

Case "webpage"

Case "retrievedinfo"

End Select
End Sub

Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmBrowse
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmBrowse
    frmD.Caption = "ActiveForm.wbser " & lDocumentCount
    frmD.wb.GoSearch
    frmD.Show
End Sub

Function FileExists(ByVal Filename As String) As Integer
'used to stop errors if a file does not exist
Dim temp$, MB_OK
    FileExists = True
On Error Resume Next
    temp$ = FileDateTime(Filename)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, MB_OK, "Error"
                End
            End If
    End Select
End Function


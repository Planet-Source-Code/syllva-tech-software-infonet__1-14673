VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmFullScreen 
   BorderStyle     =   0  'None
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   661
      BandCount       =   1
      _CBWidth        =   11970
      _CBHeight       =   375
      _Version        =   "6.0.8169"
      Child1          =   "Combo1"
      MinHeight1      =   315
      Width1          =   1620
      NewRow1         =   0   'False
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   60
         Width           =   11415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   11850
      End
   End
   Begin SHDocVwCtl.WebBrowser wbFull 
      Height          =   2295
      Left            =   480
      TabIndex        =   0
      Top             =   960
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
      Location        =   ""
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuBack 
      Caption         =   "Back"
   End
   Begin VB.Menu mnuForward 
      Caption         =   "Forward"
   End
   Begin VB.Menu mnuStop 
      Caption         =   "Stop"
   End
   Begin VB.Menu mnurefresh 
      Caption         =   "Refresh"
   End
   Begin VB.Menu mnuHome 
      Caption         =   "Home"
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Search"
      Begin VB.Menu mnuAltavista 
         Caption         =   "AltaVista"
      End
      Begin VB.Menu mnuAskJeeves 
         Caption         =   "Ask Jeeves"
      End
      Begin VB.Menu mnuMySimon 
         Caption         =   "My Simon.com"
      End
      Begin VB.Menu mnuYahoo 
         Caption         =   "Yahoo!"
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "Favorites"
      Begin VB.Menu mnuViewFavorites 
         Caption         =   "View Favorites"
      End
      Begin VB.Menu mnuAddEdit 
         Caption         =   "Add/Edit Favorites"
      End
   End
   Begin VB.Menu mnuHistory 
      Caption         =   "History"
   End
   Begin VB.Menu mnuemail 
      Caption         =   "Email"
      Begin VB.Menu mnuReadMail 
         Caption         =   "Read Mail"
      End
      Begin VB.Menu mnuSendMail 
         Caption         =   "Send Mail"
      End
      Begin VB.Menu mnuSendLink 
         Caption         =   "Send A Link"
      End
      Begin VB.Menu mnuSendPage 
         Caption         =   "Send This Page"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Begin VB.Menu mnuPrintPage 
         Caption         =   "Web Page"
      End
      Begin VB.Menu mnuSourceCode 
         Caption         =   "Source Code"
      End
   End
   Begin VB.Menu mnuCopy 
      Caption         =   "Copy"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuSendToEditor 
         Caption         =   "Send to Editor"
      End
      Begin VB.Menu mnuSendToWord 
         Caption         =   "Send To MS Word"
      End
      Begin VB.Menu mnuSendToNotepad 
         Caption         =   "Send To Notepad"
      End
   End
   Begin VB.Menu mnuFontsize 
      Caption         =   "Fontsize"
      Begin VB.Menu mnuSmallest 
         Caption         =   "Smallest"
      End
      Begin VB.Menu mnuSmaller 
         Caption         =   "Smaller"
      End
      Begin VB.Menu mnuSmall 
         Caption         =   "Small"
      End
      Begin VB.Menu mnuMedium 
         Caption         =   "Medium"
      End
      Begin VB.Menu mnuLarge 
         Caption         =   "Large"
      End
      Begin VB.Menu mnuLarger 
         Caption         =   "Larger"
      End
      Begin VB.Menu mnuLargest 
         Caption         =   "Largest"
      End
   End
   Begin VB.Menu mnuProperties 
      Caption         =   "Properties"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
wbFull.Move 0, 375, ScaleWidth, ScaleHeight - 375
End Sub

Private Sub mnuAddEdit_Click()
Load frmFavorites
frmFavorites.SSTab1.Tab = 1
frmFavorites.Text2.Text = wbFull.LocationName
frmFavorites.Text2.Text = wbFull.LocationURL
frmFavorites.Show
End Sub

Private Sub mnuAltavista_Click()
wbFull.Navigate ("http://www.altavista.com")

End Sub

Private Sub mnuAskJeeves_Click()
wbFull.Navigate ("http://www.askjeeves.com")

End Sub

Private Sub mnuBack_Click()
wbFull.GoBack
End Sub

Private Sub mnuExit_Click()
frmMdiMain.WindowState = vbMaximized
frmMdiMain.ActiveForm.wb.Navigate wbFull.LocationURL
Unload Me
End Sub

Private Sub mnuForward_Click()
wbFull.GoForward

End Sub

Private Sub mnuHome_Click()
wbFull.GoHome

End Sub

Private Sub mnuMySimon_Click()
wbFull.Navigate ("http://www.mysimon.com")

End Sub

Private Sub mnurefresh_Click()
wbFull.Refresh

End Sub

Private Sub mnuStop_Click()
wbFull.Stop

End Sub

Private Sub mnuViewFavorites_Click()
Load frmFavorites
frmFavorites.Show
End Sub

Private Sub mnuYahoo_Click()
wbFull.Navigate ("http://www.yahoo.com")

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
wbFull.Navigate Text1.Text
Combo1.AddItem Text1.Text
frmMdiMain.Combo1.AddItem Text1.Text
KeyAscii = 0
End If
End Sub

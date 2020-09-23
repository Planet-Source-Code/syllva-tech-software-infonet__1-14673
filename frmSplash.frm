VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7035
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Research the Web"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   240
      Left            =   3960
      TabIndex        =   8
      Top             =   3840
      Width           =   1740
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Make a Web Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   240
      Left            =   1950
      TabIndex        =   7
      Top             =   3840
      Width           =   1740
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Browse the Web"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   1560
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2001 by Queen City Software"
      Height          =   255
      Left            =   1215
      TabIndex        =   5
      Top             =   3420
      Width           =   5655
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "InfoNet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   4635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   240
      Left            =   6480
      TabIndex        =   1
      Top             =   3840
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Queen City Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   630
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   5460
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1200
      Top             =   3720
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Label5.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label7_Click()
Load frmBrowse
frmMdiMain.mnuResearch.Enabled = True
frmMdiMain.mnuBrowse.Enabled = False
frmBrowse.Show
Unload Me
End Sub

Private Sub Label8_Click()
Load MainForm
MainForm.Show
Unload Me
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BorderStyle = 1
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BorderStyle = 0
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BorderStyle = 1
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BorderStyle = 0
End Sub

Private Sub Label9_Click()
frmMdiMain.Picture2.Visible = True
frmMdiMain.CoolBar1.Visible = False
frmMdiMain.CoolBar2.Visible = True
Load frmBrowse
frmMdiMain.mnuResearch.Enabled = False
frmMdiMain.mnuBrowse.Enabled = True
frmBrowse.Show
Unload Me
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BorderStyle = 1
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BorderStyle = 0
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 1
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
End Sub



VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmViewSource 
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   3735
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   600
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin RichTextLib.RichTextBox rtbSource 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmViewSource.frx":0000
   End
End
Attribute VB_Name = "frmViewSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
rtbSource.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub txtURL_Change()
'rtbSource.Text = Inet1.OpenURL(txtURL.Text)

End Sub

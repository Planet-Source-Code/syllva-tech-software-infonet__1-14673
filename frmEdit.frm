VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmEdit 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   1482
      _CBWidth        =   6450
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   390
      Width1          =   420
      NewRow1         =   0   'False
      Child2          =   "Combo1"
      MinHeight2      =   315
      Width2          =   1530
      NewRow2         =   -1  'True
      MinHeight3      =   360
      Width3          =   1440
      NewRow3         =   0   'False
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   165
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   465
         Width           =   1335
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   688
         ButtonWidth     =   609
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
   End
   Begin DHTMLEDLibCtl.DHTMLEdit DHTMLEdit1 
      Height          =   5655
      Left            =   2640
      TabIndex        =   0
      Top             =   3600
      Width           =   6135
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
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


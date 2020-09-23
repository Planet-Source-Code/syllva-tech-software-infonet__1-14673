VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFavorites 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Favorites"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "favs"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "View Favorites"
      TabPicture(0)   =   "frmFavorites.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Add To Favorites"
      TabPicture(1)   =   "frmFavorites.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Text2"
      Tab(1).Control(4)=   "Text3"
      Tab(1).Control(5)=   "Text4"
      Tab(1).Control(6)=   "Command5"
      Tab(1).Control(7)=   "Command6"
      Tab(1).Control(8)=   "Command7"
      Tab(1).Control(9)=   "Command8"
      Tab(1).ControlCount=   10
      Begin VB.TextBox Text5 
         DataField       =   "notes"
         DataSource      =   "Data1"
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   1320
         Width           =   6015
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Close"
         Height          =   375
         Left            =   -70080
         TabIndex        =   15
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -72240
         TabIndex        =   14
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Update"
         Height          =   375
         Left            =   -73560
         TabIndex        =   13
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   375
         Left            =   -74880
         TabIndex        =   12
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         DataField       =   "notes"
         DataSource      =   "Data1"
         Height          =   1095
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2040
         Width           =   6015
      End
      Begin VB.TextBox Text3 
         DataField       =   "url"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   -74880
         TabIndex        =   9
         Top             =   1320
         Width           =   6015
      End
      Begin VB.TextBox Text2 
         DataField       =   "name"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   -74880
         TabIndex        =   7
         Top             =   720
         Width           =   6015
      End
      Begin VB.CommandButton Command4 
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Search By Keyword"
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   1080
         Picture         =   "frmFavorites.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Next Record"
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   600
         Picture         =   "frmFavorites.frx":037A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Previous Record"
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   120
         Picture         =   "frmFavorites.frx":06BC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Add a Favorite Place"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   3285
         Width           =   5535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         DataField       =   "url"
         DataSource      =   "Data1"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
         DataField       =   "name"
         DataSource      =   "Data1"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Notes:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   11
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "URL:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name of Web Site:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\favorites.mdb"
End Sub


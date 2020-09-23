VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "URL Links Downloader -venky_dude                              Step  3  of 3"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<<Back"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtWebsite 
      Height          =   288
      Left            =   3720
      TabIndex        =   5
      Top             =   1680
      Width           =   3252
   End
   Begin VB.TextBox txtDir 
      Height          =   288
      Left            =   3840
      TabIndex        =   2
      Top             =   2640
      Width           =   3252
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtMessages 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMain.frx":0000
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "URL Links Downloader"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblhttp 
      Caption         =   "http:\\"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblUrl 
      Caption         =   "URL to download"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblDir 
      Caption         =   "Download directory"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Boolean
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Boolean
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAcessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal lpszServerName As String, ByVal nServerPort As Integer, ByVal lpszUsername As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal lpszVerb As String, ByVal lpszObjectName As String, ByVal lpszVersion As String, byValReferer As String, ByVal lpszAcceptTypes As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal lpszheaders As String, ByVal dwHeadersLenght As Long, ByVal lpOptional As String, ByVal dwOptionalLength As Long) As Boolean
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal dwNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long) As Boolean
Dim url(100) As String
Dim xz As Integer

Dim o As Integer
Dim exitproc As Boolean
Dim b As Boolean
Dim f As Boolean
Dim hInternet As Long
Dim hConnect As Long
Dim strServer As String
Dim iPort As Integer
Dim bRes As Boolean
Dim lFlags As Long
Dim hRequest As Long
Dim strURL As String
Dim strBuffer As String * 1
Dim strDir As String
Dim strFile As String
Dim strMurl As String
Dim appdir As String
Const INTERNET_FLAG_NO_COOKIES = &H80000
Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Const INTERNET_SERVICE_HTTP = 3
Private Sub cmdConnect_Click()
cmdHangup.Enabled = True
b = InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0)
End Sub

Private Sub cmdHangup_Click()
f = InternetAutodialHangup(0)

End Sub

Private Sub cmdBack_Click()
Load frmUrl
frmUrl.Text1.Text = txtWebsite.Text
frmUrl.Text2.Text = txtDir.Text
Unload Me
frmUrl.Show

End Sub

Private Sub cmdStart_Click()
txtMessages.Text = ""
On Error Resume Next
exitproc = False
xz = 0
o = 1
Dim a As Integer
Dim c As Integer
Dim er As Integer
Dim br As Integer
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim s As String
appdir = txtDir.Text
br = Len(appdir)
er = InStrRev(appdir, "\")
If Not fso.folderexists(appdir) Then
MsgBox "Invalid destination directory"
Exit Sub
End If
If br = er Then appdir = Left(appdir, br)
stryyy = txtWebsite.Text
er = InStr(stryyy, ".htm")
If er = 0 Then
MsgBox "Invalid url"
Exit Sub
End If
c = Len(stryyy)
a = InStr(stryyy, "/")
strServer = Left(stryyy, a - 1)
strURL = Right(stryyy, c - a + 1)
strTryurl = strURL
a = InStrRev(strTryurl, "/")
c = Len(strTryurl)
strMurl = Left(strTryurl, a - 1)
iPort = 80
Call process(strServer, strURL)
Call stripurl
txtMessages.Text = txtMessages.Text & vbCrLf & " Starting to download links in file"
Call dotry
MsgBox "Finished downloading"
Command1.Caption = "Exit"
Set frmMain = Nothing
Set frmstart = Nothing
Set frmUrl = Nothing

End Sub




Private Sub download(strSServer As String, strUURL As String)
On Error Resume Next
If exitproc = True Then Exit Sub
Dim sServer As String
Dim sUrl As String
Dim x As String
Dim y As String
Dim z, f
iPort = 80
sServer = strSServer
sUrl = strUURL
iFlags = INTERNET_FLAG_NO_COOKIES
iFlags = iFlags Or INTERNET_FLAG_NO_CACHE_WRITE
hInternet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
If hInternet <> 0 Then txtMessages.Text = txtMessages.Text & vbCrLf & "Open Successfull"
hConnect = InternetConnect(hInternet, sServer, iPort, "", "", INTERNET_SERVICE_HTTP, 0, 0)
If hConnect <> 0 Then txtMessages.Text = txtMessages.Text & vbCrLf & "Connect Succesfull"
hRequest = HttpOpenRequest(hConnect, "GET", sUrl, "HTTP/1.0", vbNullString, vbNullString, iFlags, 0)
If hRequest <> 0 Then txtMessages.Text = txtMessages.Text & vbCrLf & "Http Open Request succesfull"
bRes = HttpSendRequest(hRequest, vbNullString, 0, vbNullString, 0)
If bRes = True Then txtMessages.Text = txtMessages.Text & vbCrLf & "Request successfull"
strDir = Dir(appdir & sUrl)
If Len(strDir) > 0 Then
Kill appdir & sUrl
End If
iFile = FreeFile()
Call makedire(sUrl)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.fileexists(appdir & sUrl) Then Exit Sub
Open appdir & sUrl For Binary Access Write As iFile
Do
bRes = InternetReadFile(hRequest, strBuffer, Len(strBuffer), lBytesRead)
If lBytesRead > 0 Then
Put iFile, , strBuffer
End If
Loop While lBytesRead > 0
Close iFile
txtMessages.Text = txtMessages.Text & vbCrLf & "Finished downloading " & sServer & sUrl
DoEvents
If exitproc = True Then Unload Me
End Sub


Private Sub makedire(strYZ As String)
If exitproc = True Then Exit Sub
On Error Resume Next
strYZZ = strYZ
Dim sty As String
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim b As Integer
Dim a As Integer
Dim x(10) As Integer
b = 0
a = InStr(strYZZ, "/")
c = Len(strYZZ)
stree = strYZZ
x(0) = 0
While a <> 0
b = b + 1
x(b) = x(b - 1) + a
strYZZ = Right(strYZZ, c - a)
c = Len(strYZZ)
a = InStr(strYZZ, "/")
Wend
For s = 1 To b
stre = Left(stree, x(s))

y = appdir & stre
txtMessages.Text = txtMessages.Text & vbCrLf & "Creating local sub directory " & appdir & stre
If Not fso.folderexists(y) Then MkDir (y)
Next s
DoEvents
If exitproc = True Then Unload Me
End Sub

 
 

Private Sub subfiles(strBserver As String, strBurl As String)
On Error Resume Next
If exitproc = True Then Exit Sub
Dim aer As Integer
Dim ber As Integer
Dim iFile As Integer
Dim strTry5 As String
Dim strbburl As String
strbburl = strBurl
If strbburl = "" Then Exit Sub
Dim strTry6 As String
strTry6 = ""
iFile = 1
Dim strCheck As String
Dim strTry3 As String
strTry3 = "src=" & Chr(34)
strTry4 = Chr(34)
Dim strtry9 As String
strtry9 = "SRC=" & Chr(34)
Open appdir & strbburl For Input As iFile
Do While Not EOF(iFile)
Input #iFile, strCheck
bns = Len(strCheck)
bns = bns + 1

ans = InStr(strCheck, strTry3)
If ans = 0 Then ans = InStr(strCheck, strtry9)
While ans <> 0
bns = Len(strCheck)
bns = bns + 1
h = h + 1
strCheck = Right(strCheck, bns - Len(strTry3) - ans)
cns = InStr(strCheck, strTry4)
If cns > 0 Then
strTry5 = Left(strCheck, cns - 1)
aer = InStr(strTry5, "/")
If aer <> 0 Then
ber = InStr(strTry5, "../")
If ber <> 0 Then
strTry6 = Right(strTry5, Len(strTry5) - ber + 1)
GoTo 10
Else:
strTry6 = strMurl & "/" & strTry5
GoTo 10
End If
End If
strTry6 = strMurl & "/" & strTry5
10:
Dim mz As Integer
mz = InStr(strTry6, ".com")
If mz = 0 Then
txtMessages.Text = txtMessages.Text & vbCrLf & "Downloading File " & strServer & strTry6

Call download(strServer, strTry6)
End If
End If

ans = InStr(strCheck, strTry3)
DoEvents
If exitproc = True Then Unload Me
Wend



DoEvents
If exitproc = True Then Unload Me
Loop
Close iFile

End Sub

Private Sub stripurl()
If exitproc = True Then Exit Sub
On Error Resume Next
Dim strSeek As String
Dim strCheck As String
Dim strSearch As String
Dim e As Integer
Dim h As Integer
Dim x As Boolean
Dim y As String
Dim ans, bns, cns
h = 0
Dim c As Integer
Dim d As Integer
Dim strTry As String
Dim strTry2 As String
Dim strTry3 As String
Dim strTry4 As String
Dim strTry5 As String
Dim strTry7 As String
Dim g As Integer
txtMessages.Text = txtMessages.Text & vbCrLf & "Finding downloadable links in url file "
strTry = Chr(34)
strSeek = "href=" & Chr(34)
iFile = FreeFile()
Open appdir & strURL For Input As iFile
Do While Not EOF(iFile)
Input #iFile, strCheck
bns = Len(strCheck)
bns = bns + 1
ans = InStr(strCheck, strSeek)
While ans <> 0
bns = Len(strCheck)
bns = bns + 1
h = h + 1
strCheck = Right(strCheck, bns - Len(strSeek) - ans)
cns = InStr(strCheck, strTry)
If cns > 0 Then
strtry1 = Left(strCheck, cns - 1)
c = InStr(strtry1, "http://")
d = InStr(strtry1, "#")
e = InStr(strtry1, "mailto:")
g = InStr(strtry1, "ftp:")
po = InStr(strtry1, "=")
pe = InStr(strtry1, ".com")
If c = 0 And d = 0 And e = 0 And g = 0 And po = 0 And pe = 0 Then
url(o) = strMurl & "/" & strtry1
o = o + 1
End If
End If
ans = InStr(strCheck, strSeek)
DoEvents
If exitproc = True Then Unload Me
Wend



DoEvents
If exitproc = True Then Unload Me
Loop
Close iFile
txtMessages.Text = txtMessages.Text & vbCrLf & "Finished finding links in the url"

End Sub

Private Sub Command4_Click()
Call stripurl
End Sub
Private Sub process(strsrv As String, stru As String)
If exitproc = True Then Exit Sub
On Error Resume Next
Dim strDserv As String
Dim strDurl As String
strDserv = strsrv
strDurl = stru
txtMessages.Text = txtMessages & vbCrLf & "Starting to download " & strDserv & strDurl
Call download(strDserv, strDurl)
txtMessages.Text = txtMessages.Text & vbCrLf & "Downloading image files"
Call subfiles(strDserv, strDurl)
txtMessages.Text = ""
End Sub

Private Sub dotry()
If exitproc = True Then Exit Sub
On Error Resume Next
Dim jj As Integer
For jj = 1 To o
DoEvents
If exitproc = True Then Unload Me

If url(jj) = "" Then Exit Sub
Call process(strServer, url(jj))
Next jj
End Sub

Private Sub Command1_Click()
exitproc = True
Set frmMain = Nothing
Unload Me

End Sub

Private Sub Form_Load()
txtWebsite.Enabled = False
txtDir.Enabled = False
Unload frmUrl
Unload frmstart
End Sub

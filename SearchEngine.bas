Attribute VB_Name = "SearchEngine"
Option Explicit

Public Sub SearchEngines()
Select Case frmMdiMain.Combo3.Text
Case "AltaVista"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.altavista.com")
Case "Yahoo!"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.yahoo.com")
Case "Ask Jeeves"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.askjeeves.com")
Case "DogPile"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.dogpile.com")
Case "Lycos"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.lycos.com")
Case "Excite"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.excite.com")
'Case "AltaVista"
'frmMdiMain.ActiveForm.wb.Navigate ("http://www.altavista.com")
'Case "AltaVista"
'frmMdiMain.ActiveForm.wb.Navigate ("http://www.altavista.com")
'Case "AltaVista"
'frmMdiMain.ActiveForm.wb.Navigate ("http://www.altavista.com")
'Case "AltaVista"
'frmMdiMain.ActiveForm.wb.Navigate ("http://www.altavista.com")

End Select
End Sub

Private Sub LoadNewBrowse()
    Static lDocumentCount As Long
    Dim frmD As frmBrowse
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmBrowse
    frmD.Caption = "Browser " & lDocumentCount
    frmD.wb.GoSearch
    frmD.Show
End Sub

Public Sub SearchEngines2()
Select Case frmMdiMain.Combo2.Text
Case "AltaVista"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.altavista.com")
Case "Yahoo!"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.yahoo.com")
Case "Ask Jeeves"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.askjeeves.com")
Case "DogPile"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.dogpile.com")
Case "Lycos"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.lycos.com")
Case "Excite"
    LoadNewBrowse
frmMdiMain.ActiveForm.wb.Navigate ("http://www.excite.com")
'Case "AltaVista"
'frmMdiMain.ActiveForm.wb.Navigate ("http://www.altavista.com")
'Case "AltaVista"
'frmMdiMain.ActiveForm.wb.Navigate ("http://www.altavista.com")
'Case "AltaVista"
'frmMdiMain.ActiveForm.wb.Navigate ("http://www.altavista.com")
'Case "AltaVista"
'frmMdiMain.ActiveForm.wb.Navigate ("http://www.altavista.com")
End Select
End Sub


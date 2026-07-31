Attribute VB_Name = "Harjutus_6"
Option Explicit
'funktsioon, mis valib kõik lahtrid tabelid väljaarvatud tabeli pealkirja rida
Function ilmatabelinimi(algusLahter As Range, paisteRidu As Long) As Range
    Dim cr As Range
    Set cr = algusLahter.CurrentRegion
'vähendab esimese(vahelejäetud) rea võrra loendatud ridade arvu
    Set ilmatabelinimi = cr.Offset(paisteRidu, 0).Resize(cr.Rows.Count - paisteRidu, cr.Columns.Count)
End Function


Public Sub rammusroheline()

Dim algus As Range
Dim roheline As Long
Dim lahter As Range
Dim TabelA As Range
Dim varvimisveerg As Range
Dim varvimiskogus As Integer
Dim roheliseks As Variant
Dim varv As Range
Dim toonid As Range
Dim leiavaartus As String
Dim ws As Worksheet


'seab aktiivse töölehe, alguslahtri ja värvidehulga
Set ws = ThisWorkbook.Sheets("tabelid")
Set algus = ws.Range("algus")
Set toonid = ws.Range("varvid")
leiavaartus = "roheline"
'sobiv vahemik, millest vastavad väärtused värvitakse
Set TabelA = ilmatabelinimi(algus, 1)
Set varvimisveerg = TabelA.Columns(2)

'kustutab alguses kõik värvid ära
varvimisveerg.Interior.ColorIndex = xlNone
varvimiskogus = 0

'määratleb värvi kõikide värvide seast(kui olemas)
Set varv = toonid.Find(leiavaartus, LookIn:=xlValues, LookAt:=xlWhole)
If varv Is Nothing Then
    MsgBox "värvi '" & leiavaartus & "' ei leitud vahemikust 'varvid'.", vbExclamation
    Exit Sub
End If
roheline = varv.Interior.Color

Dim c As Range
'kontrollib läbi tabel A teise veeru, et leida numbrid, mis pole rasvases kirjas
For Each c In varvimisveerg.Cells
    If IsNumeric(c.Value) And Not c.Font.Bold Then
        c.Interior.Color = roheline
        varvimiskogus = varvimiskogus + 1
    End If
Next c
'edastab saadud tulemuse

MsgBox varvimiskogus & " rohelist lahtrit"
End Sub






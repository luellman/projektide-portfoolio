Attribute VB_Name = "Harjutus_5"
'funktsioon, mis valib kõik lahtrid tabelid väljaarvatud tabeli pealkirja rida
Function ilmatabelinimi(algusLahter As Range, paisteRidu As Long) As Range
    Dim cr As Range
    Set cr = algusLahter.CurrentRegion
'vähendab esimese(vahelejäetud) rea võrra loendatud ridade arvu
    Set ilmatabelinimi = cr.Offset(paisteRidu, 0).Resize(cr.Rows.Count - paisteRidu, cr.Columns.Count)
End Function

Public Sub ladinasuurtaht()
Dim Tabel_B As Range
Dim tulemus As Range
Dim lahter As Range
Dim tulemusus As Long

'määrab piirkonna, kuhu väärtused kopeeritakse
Set Tabel_B = ilmatabelinimi(Range("tab_algne"), 1)
Set tulemus = Range("tulemus")

'kustutab esmalt vanad väärtused
With Range("tulemus")
    .Offset(1, 0).Resize(Rows.Count - .Row).Clear
End With

'määrab tingimused, mille järgi valitakse lahtrid
For Each lahter In Tabel_B.Columns(3).Cells
If Len(lahter.Value) > 0 And UCase(Left(lahter.Value, 1)) = Left(lahter.Value, 1) And Left(lahter.Value, 1) Like "[A-ZÄÖÕÜ]" And InStr(lahter.Value, " ") = 0 And Not IsNumeric(lahter.Value) And Not IsDate(lahter.Value) Then
 'iga järgnev lahter lisatakse arvulisse kogusummasse
 tulemusus = tulemusus + 1
 'kopeerib ja kleebib tulemusused üksteise alla määratud sihtkohta
            lahter.Copy
            tulemus.Offset(tulemusus - 0, 0).PasteSpecial Paste:=xlPasteAll
End If
Next lahter
'joondab tulemuse veeru vastavalt
Columns("p").Select
Columns("p").AutoFit
Range("p1").Select


MsgBox tulemusus & " väärtust kopeeritud, mis ühesõnaline ja algab suure tähega"


End Sub

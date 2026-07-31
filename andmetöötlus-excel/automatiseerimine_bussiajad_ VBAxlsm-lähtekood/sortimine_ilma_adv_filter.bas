Attribute VB_Name = "Harjutus4"
Sub Inimesed()
'määran muutujad ning nende tingimused
    Dim Tab_alg As Range
    Dim tabel As Range
    
    Dim Kriteerium1 As Range
    Dim rida As Long
    Dim viimanerida As Long
    Dim sihtrida As Long
    Dim ws As Worksheet
    Set ws = Worksheets("inimesed")
    Set Tab_alg = Range("Tab_alg")
    Set tabel = Range("tabel")
    Set Kriteerium1 = Range("Kriteerium1")
    'viimane rida, kuhuni vaadatakse
    viimanerida = 102
    'rida, alates millest tekib tabel
    sihtrida = 19
    'kirjutab veergude pealkirjad
    ws.Cells(sihtrida, 1).Value = "Eesnimi"
    ws.Cells(sihtrida, 2).Value = "Perenimi"
    ws.Cells(sihtrida, 3).Value = "linn"
    ws.Cells(sihtrida, 4).Value = "sünnikuu"
    ws.Cells(sihtrida, 5).Value = "palk"
    'kustutab eelnevalt piirkonna vanadest väärtustest
    tabel.CurrentRegion.Offset(2, 0).Clear
    sihtrida = sihtrida + 1
    'kui vanus oranz ja sugu valitud, siis vastavalt nendele kriteeriumitele tagastab väärtused
    For rida = 15 To viimanerida
      If ws.Cells(rida, 14).DisplayFormat.Font.Color = RGB(255, 192, 0) = True And ws.Cells(rida, 11).Value = Kriteerium1.Value Then
            ws.Cells(sihtrida, 1).Value = ws.Cells(rida, 9).Value
            ws.Cells(sihtrida, 2).Value = ws.Cells(rida, 10).Value
            ws.Cells(sihtrida, 3).Value = ws.Cells(rida, 19).Value
            ws.Cells(sihtrida, 4).Value = ws.Cells(rida, 23).Value
            ws.Cells(sihtrida, 5).Value = ws.Cells(rida, 18).Value
            sihtrida = sihtrida + 1
            'kui ühtegi sugu valitud pole, ei tagasta tulemusi
        ElseIf ws.Cells(rida, 14).DisplayFormat.Font.Color = RGB(255, 192, 0) = True And Kriteerium1.Value = "" Then
            ws.Cells(sihtrida, 1).Value = ""
            ws.Cells(sihtrida, 2).Value = ""
            ws.Cells(sihtrida, 3).Value = ""
            ws.Cells(sihtrida, 4).Value = ""
            ws.Cells(sihtrida, 5).Value = ""
            sihtrida = sihtrida + 1
        End If
        Next rida
    Dim vastusarv As Long
    'loeb kokku mitu vastust tuleb
    vastusarv = tabel.CurrentRegion.Rows.count - 2
    Dim keskmine As Double
    Dim prk As Range
    Set prk = Range("palk")
    'kui sugu pole määratud, ei tule ühtegi vastust
    If vastusarv = 0 Then
        MsgBox "sugu määramata!"
        'kui sugu määratud, siis annab tulemuse
   Else
   keskmine = Application.WorksheetFunction.Average(prk)
        MsgBox "Leiti " & vastusarv & " vastust, kus sugu on " & Kriteerium1 & ", vanus oranzi värvi ja keskmine palk on " & Round(keskmine, 1)
    End If
End Sub

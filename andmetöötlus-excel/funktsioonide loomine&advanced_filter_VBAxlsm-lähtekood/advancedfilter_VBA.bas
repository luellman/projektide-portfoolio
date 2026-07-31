Attribute VB_Name = "Harjutus_7"

Sub inimestesordid()
Attribute inimestesordid.VB_ProcData.VB_Invoke_Func = "g\n14"
    Dim Inimesed As Worksheet
    Dim Paring As Worksheet
    Dim tabel2 As Range
    Dim vastused As Range
    Dim tingimus As Range
    Dim filtreerimine As Range
    Dim arvvaartused As Long
    
    'piirkonna ning tingimuse piirkonna m‰‰ratlemine
    Set Inimesed = ThisWorkbook.Sheets("Inimesed")
    Set Paring = ThisWorkbook.Sheets("Paring")
    Set tabel2 = Inimesed.Range("Tabel2")
    Set vastused = Paring.Range("vastused")
    Set tingimus = Paring.Range("Tingimus")
    Set filtreerimine = Paring.Range("vastused").Resize(100, 1)
    'kustutamine
    Paring.Range("vasted").ClearContents
    'kopeerib tabelist vastavalt advancedfiltri tingimustele mitteunikaalsed
    tabel2.AdvancedFilter Action:=xlFilterCopy, _
      CriteriaRange:=tingimus, _
        CopyToRange:=vastused.Resize(, 9), _
        Unique:=False
        'loeb kokku tulemused ning tagastab arvu sınumina
        arvvaartused = Application.WorksheetFunction.CountA(vastused.Columns(1)) - 1
         MsgBox "kirjete arv: " & arvvaartused, vbInformation, "Tulemused"
    With Paring.Sort
        .SortFields.Clear
        .SortFields.Add Key:=vastused.Columns(1), Order:=xlAscending
        .SetRange vastused
        .Header = xlYes
        .Apply
    End With
End Sub




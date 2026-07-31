Attribute VB_Name = "Harjutus3"
Option Explicit
'funktsiooni loomine, mis lubab strConv
'asemel võõrtähed suureks või väikseks muuta, et nimi korrektselt kuvada

'funktsioon argumendiga s ehk string
Private Function vorming(ByVal s As String) As String
    Dim i As Long
    Dim taht As String
    Dim uussona As Boolean
    Dim tulemus As String
    
    'uussona jälgib, kas oleme järgmise "sõna" ehk nime juures. kui on "true", siis esimene täht suureks nagu nimede puhul kombeks
    
    uussona = True
    For i = 1 To Len(s)
        taht = Mid$(s, i, 1)
        
        'peale tühikut, sidekriipsu või ülakoma,
        'muudab tähe suureks
        If taht = " " Or taht = "-" Or taht = "'" Then
            tulemus = tulemus & taht
            uussona = True
        Else
            If uussona Then
                tulemus = tulemus & UCase$(taht)
            Else
                tulemus = tulemus & LCase$(taht)
            End If
            uussona = False
            End If
            'UCase ning Lcase lubavad kasutada unicode formaati erinevalt ANSI süsteemist, ehk võõrtähtede asemele ei teki "?"
            
        Next i
        vorming = tulemus
    
End Function


Sub europarlament()

Dim andmed As Range
Dim vastused As Range
Dim nr As Long
Dim kirjetearv As Long
Dim rida As Long
Dim fraktsiooniarv As Long
Dim kogunimi As String
Dim tyhikukoht As Long
Dim uusnimi As String

    On Error GoTo Cleanup
    
    'ei uuenda vaadet iga uuenduse järel vaid alles lõpus
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    
    Dim tööleht As Worksheet
    Dim viimanerida As Long
    
    Set tööleht = Worksheets("tabel")
    'määrab piirkonna, kust võtab informatsiooni
    Set andmed = Worksheets("europarlament").Range("parlament")
    'määrab piirkonna, kuhu kopeerib vastused
    Set vastused = Worksheets("tabel").Range("Tulem_prk")
    'puhastab eelnevalt piirkonna, kuhu vastused kopeeritakse
    viimanerida = tööleht.Cells(tööleht.Rows.count, "B").End(xlUp).Row
    
        If viimanerida > 1 Then
            tööleht.Range("B11:E" & viimanerida).Clear
        End If
    'määrab veergudele pealkirjad
     vastused.Cells(1, 1).Value = "nimi"
    vastused.Cells(1, 2).Value = "riik"
    vastused.Cells(1, 3).Value = "partei/erakond"
    vastused.Cells(1, 4).Value = "fraktsioon"
    nr = 1
    'kordab tegevust iga viie rea kaupa
    For rida = 1 To andmed.Rows.count Step 5
       nr = nr + 1
       
       andmed.Cells(rida, 1).Copy vastused.Cells(nr, 1)
       
       kogunimi = vastused.Cells(nr, 1).Value
       tyhikukoht = InStr(kogunimi, " ")
       If tyhikukoht > 0 Then
            uusnimi = Left(kogunimi, tyhikukoht - 1) & " " & vorming(Mid(kogunimi, tyhikukoht + 1))
       Else
            uusnimi = kogunimi
       End If
       'kirjutab uue teksti, säilitades lingi
       vastused.Cells(nr, 1).Value = uusnimi
       If vastused.Cells(nr, 1).Hyperlinks.count > 0 Then
            vastused.Cells(nr, 1).Hyperlinks(1).TextToDisplay = uusnimi
       End If
       
       vastused.Cells(nr, 2).Value = andmed.Cells(rida + 2, 1).Value
       andmed.Cells(rida + 1, 1).Copy vastused.Cells(nr, 4)
       andmed.Cells(rida + 3, 1).Copy vastused.Cells(nr, 3)
       'kui tabelis fraktsioon on uueneva euroopa fraktsioon, värvib lahtid antud reas kollaseks
         If InStr(1, vastused.Cells(nr, 4).Value, "Uueneva Euroopa fraktsioon", vbTextCompare) > 0 Then
               vastused.Rows(nr).Resize(1, 4).Interior.ColorIndex = 6
            End If
   Next
 
 
    'määratleb muutuja
    Dim tab_prk As Range
    Dim andmetabel As ListObject
    Dim tbl As ListObject

    'töölehe muutuja
    Set tööleht = ThisWorkbook.Sheets("tabel")
      
    'määrab piirkonna, kuhu tabel tekitada
    Set tab_prk = tööleht.Range("tab_prk")
   'juhul, kui tabelit, veel pole tehtud, siis ignoreerib
    For Each tbl In tööleht.ListObjects
        If tbl.Name = "infotabel" Then
            tbl.Unlist
            Exit For
        End If
    Next tbl
 
    'loob uue tabeli
    Set andmetabel = tööleht.ListObjects.Add(xlSrcRange, tab_prk, , xlYes)
    andmetabel.Name = "infotabel" 'määrab tabelile nime
    tööleht.Columns("B:E").AutoFit 'korrastab
    
    'valminud tabeli põhjal loetakse kirjed kokku
    'listrows jätab pealkirjarea vahele
    
    
    Dim andmerida As ListRow
    kirjetearv = 0
    fraktsiooniarv = 0
    
    For Each andmerida In andmetabel.ListRows
        If Trim$(andmerida.Range.Cells(1, 1).Value) <> "" Then
            kirjetearv = kirjetearv + 1
            If InStr(1, andmerida.Range.Cells(1, 4).Value, "Uueneva Euroopa fraktsioon", vbTextCompare) > 0 Then
                fraktsiooniarv = fraktsiooniarv + 1
            End If
        End If
    Next andmerida
            
        
    MsgBox "Tekkis " & kirjetearv & " kirjet" & vbCrLf & _
    "Uueneva Euroopa fraktsiooni liikmeid: " & fraktsiooniarv 'annab vastuse"

'taastab algse seadistuse peale lõppu
Cleanup:

 If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub


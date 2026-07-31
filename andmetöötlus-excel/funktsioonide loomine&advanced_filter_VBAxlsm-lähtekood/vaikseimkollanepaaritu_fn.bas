Attribute VB_Name = "Harjutus_3"
Function v‰ikseimkollanepaaritu(rng As Range) As Variant
    Dim lahter As Range
    Dim minvaartus As Double
    Dim leitud As Boolean
    Dim kollane As Long
    
    ' Kollase v‰rvi kood (RGB)
    kollane = RGB(255, 255, 0)
    
    ' Algv‰‰rtus - alguses pole leitud sobivat v‰‰rtust
    leitud = False
    minvaartus = 1000  'suur number, mis teeb kindlaks, et otsitav arv on v‰ikseim, isegi kui on kıige suurem(juhul, kui kıik arvud samasuured)
    
  
    For Each lahter In rng
        ' kontrollin, kas lahter pole t¸hi ja on arv
        If Not IsEmpty(lahter.Value) And IsNumeric(lahter.Value) Then
            ' kas vastab tingimustele paaritu ja kollane
            If lahter.Value Mod 2 <> 0 And lahter.Interior.Color = kollane Then
                ' kui v‰‰rtus vastab kıigile eelnevatele tingimustele ja on eelnevast v‰iksem
                If lahter.Value < minvaartus Then
                    minvaartus = lahter.Value
                    leitud = True
                End If
            End If
        End If
    Next lahter
    
    
    If leitud Then
        v‰ikseimkollanepaaritu = minvaartus
    Else
        v‰ikseimkollanepaaritu = "pole andmeid"
    End If
End Function

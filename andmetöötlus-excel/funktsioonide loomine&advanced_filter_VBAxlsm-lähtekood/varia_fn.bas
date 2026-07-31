Attribute VB_Name = "Harjutus_2"
Function varia(yks As Variant, kaks As Variant, kolm As Variant) As Variant
  
    Dim a As Double, b As Double, c As Double
    
    ' Kontrollin, kas kõik väärtused on arvud
    If IsNumeric(yks) And IsNumeric(kaks) And IsNumeric(kolm) Then
        'kui kõik arvud, siis määran muutujad
        a = CDbl(yks)
        b = CDbl(kaks)
        c = CDbl(kolm)
        
        ' Kontrollin, kas kolmnurka annab arvutada
        If a + b > c And b + c > a And c + a > b Then
            
            varia = a + b + c
        Else
            
            varia = "pole võimalik"
        End If
    Else
        'juhul, kui kõik väärtused pole arvud
        varia = "kõik väärtused pole arvud!"
    End If
End Function


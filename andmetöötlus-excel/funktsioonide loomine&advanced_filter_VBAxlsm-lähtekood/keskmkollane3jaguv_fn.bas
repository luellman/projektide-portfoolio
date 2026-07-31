Attribute VB_Name = "Harjutus_4"
Function keskmkollane3jaguv(rng As Range) As Variant
    Dim lahter As Range
    Dim summa As Double
    Dim kogus As Integer
    Dim kollane As Long
    
    ' Kollase värvi kood (RGB)
    kollane = RGB(255, 255, 0)
    
    ' Algväärtused
    summa = 0
    kogus = 0
    
    ' kontrollib kõik lahtrid üle, et oleks kollane, oleks arv ja oleks kolmega jaguv, misjärel liidab sobiva väärtuse summasse
    For Each lahter In rng
        If IsNumeric(lahter.Value) Then
            If lahter.Value Mod 3 = 0 And lahter.Interior.Color = kollane Then
                summa = summa + lahter.Value
                kogus = kogus + 1
            End If
        End If
    Next lahter
    
    'sobivate väärtustega arvutatakse aritm. keskmine
    If kogus > 0 Then
        keskmkollane3jaguv = summa / kogus
    Else
       
        keskmkollane3jaguv = "pole andmeid"
    End If
End Function




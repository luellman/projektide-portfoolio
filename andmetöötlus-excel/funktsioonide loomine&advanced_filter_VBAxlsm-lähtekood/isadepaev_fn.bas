Attribute VB_Name = "harjutus_1"
Public Function isade_paev(year As Range) As Date

Dim esimenepaevnovembris As Date
Dim esimenepuhapaev As Date
Dim nadalapaev As Date

esimenepaevnovembris = DateSerial(year.Value, 11, 1)

nadalapaev = Weekday(esimenepaevnovembris, vbSunday)

If nadalapaev = 1 Then
esimenepuhapaev = esimenepaevnovembris
Else
esimenepuhapaev = esimenepaevnovembris + (7 - nadalapaev + 1)

End If

isade_paev = esimenepuhapaev + 7




End Function







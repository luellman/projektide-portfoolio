Attribute VB_Name = "harjutus1"
Function tekst_arvuks(tekst)
'otsib lahtri väärtuses üles kõik komad ning likvideerib need
txt1 = Replace(tekst, ",", "")
'asendab lahtris kõik punktid komadega
txt2 = Replace(txt1, ".", ",")
'eemaldab valuutasümboli ning tühiku ja tähe
txt3 = Mid(txt2, 2, Len(txt2) - 3)
'teisendab viimase tulemuse(txt3) numbriks
tekst_arvuks = CDbl(txt3)

End Function




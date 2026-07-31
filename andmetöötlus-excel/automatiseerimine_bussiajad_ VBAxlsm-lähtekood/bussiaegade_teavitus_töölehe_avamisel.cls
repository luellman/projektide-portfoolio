VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


'funktsiooni loomine, mis m‰‰ratleb p¸had, kui buss sıidab n‰dalavahetuse graafikul

Private Function riigip¸ha(d As Date) As Boolean
    If Month(d) = 1 And Day(d) = 1 Then riigip¸ha = True: Exit Function    ' Uus aasta
    If Month(d) = 2 And Day(d) = 24 Then riigip¸ha = True: Exit Function   ' iseseisvusp‰ev
    If Month(d) = 4 And Day(d) = 3 Then riigip¸ha = True: Exit Function   ' suur reede
    If Month(d) = 4 And Day(d) = 5 Then riigip¸ha = True: Exit Function   ' munadep¸had
    If Month(d) = 5 And Day(d) = 1 Then riigip¸ha = True: Exit Function    ' kevadp¸ha
    If Month(d) = 5 And Day(d) = 24 Then riigip¸ha = True: Exit Function   ' ¸lestıusmisp¸ha
    If Month(d) = 6 And Day(d) = 23 Then riigip¸ha = True: Exit Function   ' vıidup¸ha
    If Month(d) = 6 And Day(d) = 24 Then riigip¸ha = True: Exit Function   ' jaanip‰ev
    If Month(d) = 8 And Day(d) = 20 Then riigip¸ha = True: Exit Function   ' taasiseseisvusp‰ev
    If Month(d) = 12 And Day(d) = 24 Then riigip¸ha = True: Exit Function  ' jıulud
    If Month(d) = 12 And Day(d) = 25 Then riigip¸ha = True: Exit Function  ' jıulud
    If Month(d) = 12 And Day(d) = 26 Then riigip¸ha = True: Exit Function  ' jıulud

   
End Function

Private Function tˆˆp‰evagraafik(d As Date) As Boolean
    tˆˆp‰evagraafik = (Weekday(d, vbMonday) >= 6) Or riigip¸ha(d)
End Function



Private Sub worksheet_activate()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Bussiajad")

'm‰‰raleb piirkonna, kust v‰‰rtuseid leida

    Const esimeneRida As Long = 9
    Const viimanerida As Long = 27

'm‰‰ratleb j‰rgmised 3 v‰ljumisaega "tulemusteks"

    Dim tulemused(1 To 3) As String
    Dim leitud_hulk As Integer

'm‰‰ratleb kellaajad ja kuup‰evad

    Dim kuup‰ev_hetkel As Date
    kuup‰ev_hetkel = Date

    Dim tundPrg As Integer, minutPrg As Integer
    tundPrg = Hour(Time)
    minutPrg = Minute(Time)

    Dim j‰rgmine_aeg As Integer
    For j‰rgmine_aeg = 1 To 2   ' j‰rgmine_aeg 1 = t‰nane p‰ev, j‰rgmine_aeg 2 = homme (kui t‰na rohkem busse pole)

        Dim tunniveerg As Long, esimeneMinutiveerg As Long, viimaneMinutiveerg As Long
        If tˆˆp‰evagraafik(kuup‰ev_hetkel) Then
            tunniveerg = 11: esimeneMinutiveerg = 12: viimaneMinutiveerg = 18      'veerud K(tunnid), L:R(minutiteveerud) numbritena
        Else
            tunniveerg = 1: esimeneMinutiveerg = 2: viimaneMinutiveerg = 9         ' tundideveerg(A), minutiteveerud(B:I) numbritena
        End If
'vaatab l‰bi piirkonna ning seab vastavusse kellaajad, et leida j‰rgmised v‰ljumised vastavalt graagikule
'muudab tekstiv‰‰rtused numbriv‰‰rtusteks Cint'ga
        Dim r As Long, c As Long
        For r = esimeneRida To viimanerida
            If ws.Cells(r, tunniveerg).Value <> "" Then
                Dim v‰ljumistund As Integer
                v‰ljumistund = CInt(ws.Cells(r, tunniveerg).Value)

                For c = esimeneMinutiveerg To viimaneMinutiveerg
                    If ws.Cells(r, c).Value <> "" Then
                        Dim v‰ljumisminut As Integer
                        v‰ljumisminut = CInt(ws.Cells(r, c).Value)

'kui j‰rgmine_aeg = 1, kontrollib t‰nast p‰eva, kus v‰ljumine on hiljem kui kellaaeg praegu


                        Dim vastavuses As Boolean
                        If j‰rgmine_aeg = 1 Then
                            vastavuses = (v‰ljumistund > tundPrg) Or (v‰ljumistund = tundPrg And v‰ljumisminut > minutPrg)
                        Else
                            vastavuses = True
                        End If
'leiab vastava kellaaja ning salvestab ja otsib j‰rgmist
                        If vastavuses Then
                            leitud_hulk = leitud_hulk + 1
                            tulemused(leitud_hulk) = Format(v‰ljumistund, "00") & ":" & Format(v‰ljumisminut, "00")
                            If leitud_hulk = 3 Then GoTo kıik_olemas
                        End If
                    End If
                Next c
            End If
        Next r

        kuup‰ev_hetkel = kuup‰ev_hetkel + 1
    Next j‰rgmine_aeg

'kui leidis kıik 3 j‰rgmist aega, vormindab tulemuse ja edastab sınumi

kıik_olemas:

    Dim msg As String
    msg = "kell praegu: " & Format(Time, "hh:nn") & vbNewLine & vbNewLine
    msg = msg & "J‰rgmised 3 bussi nr 24 v‰ljuvad Raja peatusest:" & vbNewLine

    Dim i As Integer
    For i = 1 To leitud_hulk
        msg = msg & "  " & i & ". " & tulemused(i) & vbNewLine
    Next i

    If leitud_hulk = 0 Then msg = msg & "  (rohkem v‰ljumisi pole)"

    MsgBox msg, vbInformation, "Buss nr 24 v‰ljumised Raja peatusest linna poole"
End Sub



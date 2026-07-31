Attribute VB_Name = "Module2"
Option Explicit

Sub kustutaandmed()
    Dim ws As Worksheet
    Dim viimanerida As Long
    
    Set ws = ThisWorkbook.Worksheets("tabel")
    
    viimanerida = ws.Cells(ws.Rows.count, "b").End(xlUp).Row
    If viimanerida > 1 Then
        ws.Range("b2:e" & viimanerida).Clear
        ws.Columns("B:E").ColumnWidth = ws.StandardWidth
    End If
End Sub

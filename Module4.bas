Attribute VB_Name = "Module4"
Sub MatchandColor()
Dim ws As Worksheet
Set ws = Sheets("Sheet8")
ws.Activate

Dim nr As Integer
nr = WorksheetFunction.CountA(Range("B:B")) + 1

Dim nc As Integer
nc = WorksheetFunction.CountA(Range("1:1")) - 3
Dim c As Integer
Dim i As Integer
Dim a As Integer
Dim b As Variant
Dim cellva As Variant


Dim ws2 As Worksheet
Set ws2 = Sheets("date")

Dim searchValue As String
Dim rdate As String
Dim wheredate As Integer
Dim searchcol As Range
Dim searchrange As Range
Dim searchresult As Range
Set searchrange = ws2.Range("B:B")
Set searchresult = ws2.Range("D:D")
Set searchcol = ws.Range("d1:du1")

For i = 3 To nr
    searchValue = ws.Range("B" & i)
    rdate = WorksheetFunction.XLookup(searchValue, searchrange, searchresult, "Not found")
If rdate = "Not found" Then
ws.Range("B" & i).Interior.color = RGB(255, 0, 0)
End If
For c = 4 To nc
b = c Mod 2
cellva = Cells(1, c).value
If b = 0 Then
    If rdate = cellva Then
        Cells(i, c).value = "Y"
        Cells(i, c).Interior.color = RGB(0, 0, 0)
        ws.Range("B" & i).Interior.color = RGB(0, 255, 0)
        Exit For
    End If
End If
Next c
Next i
End Sub

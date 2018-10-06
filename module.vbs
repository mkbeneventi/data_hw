' module
Sub wallstreet()

' loop through sheets
For Each ws In Worksheets

' activates whenever the sheet is activated
ws.Activate

' add header in each worksheet in I and J
Cells(1, "i").Value = "Ticker"
Cells(1, "j").Value = "Total Stock Volume"

' set variables

' ticker Returns a repeating character or string.
Dim ticker As String

' double provides the largest and smallest possible magnitudes for a number.
Dim totalstockvolume As Double
Dim results As Double

totalstockvolume = 0
results = 2

' determine the last row
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Loop through each row
For i = 2 To Lastrow

' Searches for when the value of the next cell is different than that of the current cell

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' set value for ticker
ticker = Cells(i, 1).Value

' combine resuts in I

ActiveSheet.Range("I" & results).Value = ticker

' set value for total stock volume
totalstockvolume = totalstockvolume + Cells(i, 7).Value
' combine results in j

ActiveSheet.Range("J" & results).Value = totalstockvolume

' take present value  Incrementing A Cell By 1
results = results + 1

 ' reset totalstockvolume
totalstockvolume = 0

Else
' loop
totalstockvolume = totalstockvolume + Cells(i, 7).Value

End If

Next i

Next ws

End Sub


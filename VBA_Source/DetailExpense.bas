Attribute VB_Name = "DetailExpense"
Option Explicit

' 1. Entry macro:
' Run this macro from the Excel macro list.
' It opens the month selection form.
Sub ShowExpenseMonthSelector()

    UserForm_DetailExpense.Show
End Sub

' 2. Main procedure:
' Processes expense data and generates the pie chart.
' Since it requires a parameter (selectedMonth),
' it does not appear in the Excel macro list and is called by the UserForm.
Public Sub FillDetailAndPie(ByVal selectedMonth As Integer)

    Dim wsDep As Worksheet, wsDash As Worksheet
    Dim lastRow As Long, i As Long, outRow As Long
    Dim depMonth As Integer, cat As String, amount As Double
    Dim dictCat As Object, cht As ChartObject

    Dim k As Variant
    Dim j As Long
    Dim idx As Long

    Dim labels() As Variant
    Dim values() As Variant
    Dim colors As Variant

    Set wsDep = ThisWorkbook.Worksheets("Data_Expenses")
    Set wsDash = ThisWorkbook.Worksheets("COMPTES")
    Set dictCat = CreateObject("Scripting.Dictionary")

    ' =========================
    ' Clear previous output
    ' =========================
    wsDash.Range("C13:E24").ClearContents
    wsDash.Range("C12:E12").Value = Array("Date", "Cat¨¦gorie", "Montant")

    outRow = 13
    lastRow = wsDep.Cells(wsDep.Rows.Count, 1).End(xlUp).Row

    ' =========================
    ' Loop through source data
    ' =========================
    For i = 2 To lastRow

        If IsDate(wsDep.Cells(i, 1).Value) Then
            depMonth = Month(wsDep.Cells(i, 1).Value)

            If depMonth = selectedMonth Then

                If outRow > 23 Then Exit For

                wsDash.Cells(outRow, 3).Value = wsDep.Cells(i, 1).Value
                wsDash.Cells(outRow, 4).Value = wsDep.Cells(i, 3).Value
                wsDash.Cells(outRow, 5).Value = wsDep.Cells(i, 2).Value

                cat = wsDep.Cells(i, 3).Value
                amount = wsDep.Cells(i, 2).Value

                If dictCat.exists(cat) Then
                    dictCat(cat) = dictCat(cat) + amount
                Else
                    dictCat.Add cat, amount
                End If

                outRow = outRow + 1
            End If
        End If

    Next i

    If outRow = 13 Then
        MsgBox "Aucune donn¨¦e pour ce mois", vbInformation
        Exit Sub
    End If

    wsDash.Range("E13:E" & outRow - 1).NumberFormat = "#,##0.00 [$€-fr-FR]"

    ' =========================
    ' Remove existing charts
    ' =========================
    For Each cht In wsDash.ChartObjects
        cht.Delete
    Next cht
    
    ' =========================
    ' Create pie chart
    ' =========================
    Set cht = wsDash.ChartObjects.Add( _
                Left:=wsDash.Range("C25").Left, _
                Top:=wsDash.Range("C25").Top, _
                Width:=wsDash.Range("C25:E36").Width, _
                Height:=wsDash.Range("C25:E36").Height)

    ReDim labels(1 To dictCat.Count)
    ReDim values(1 To dictCat.Count)

    idx = 1
    For Each k In dictCat.Keys
        labels(idx) = k
        values(idx) = dictCat(k)
        idx = idx + 1
    Next k

    With cht.Chart

        .ChartType = xlPie

        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .XValues = labels
            .values = values
            .Name = "R¨¦partition"
        End With

        ' Title
        .HasTitle = True
        .ChartTitle.Text = "R¨¦partition par cat¨¦gorie (" & MonthNameFR(selectedMonth) & ")"

        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 12
            .Bold = True
            .Name = "Calibri"
        End With

        ' Color
        colors = Array( _
            RGB(64, 176, 166), _
            RGB(102, 204, 194), _
            RGB(153, 216, 209), _
            RGB(200, 230, 226), _
            RGB(120, 140, 140))

        For j = 1 To .SeriesCollection(1).Points.Count
            .SeriesCollection(1).Points(j).Format.Fill.ForeColor.RGB = _
                colors((j - 1) Mod (UBound(colors) + 1))
        Next j

        .ChartArea.Format.Fill.Visible = msoFalse
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        .Legend.Font.Size = 9

    End With

End Sub
Public Function MonthNameFR(ByVal m As Integer) As String

    Select Case m
        Case 1: MonthNameFR = "Janvier"
        Case 2: MonthNameFR = "F¨¦vrier"
        Case 3: MonthNameFR = "Mars"
        Case 4: MonthNameFR = "Avril"
        Case 5: MonthNameFR = "Mai"
        Case 6: MonthNameFR = "Juin"
        Case 7: MonthNameFR = "Juillet"
        Case 8: MonthNameFR = "Ao?t"
        Case 9: MonthNameFR = "Septembre"
        Case 10: MonthNameFR = "Octobre"
        Case 11: MonthNameFR = "Novembre"
        Case 12: MonthNameFR = "D¨¦cembre"
        Case Else: MonthNameFR = ""
    End Select

End Function

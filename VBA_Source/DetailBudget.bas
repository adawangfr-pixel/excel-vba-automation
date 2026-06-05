Attribute VB_Name = "DetailBudget"
Option Explicit

' 1. Entry macro: opens the budget month selection form
Sub ShowBudgetMonthSelector()

    UserForm_DetailBudget.Show
End Sub

Public Sub FillBudgetAndPie(ByVal selectedMonth As Integer)

    Dim wsBud As Worksheet, wsDash As Worksheet
    Dim lastRow As Long, i As Long, outRow As Long
    Dim budMonth As Integer, cat As String, amount As Double
    Dim dictCat As Object, cht As ChartObject
    Dim k As Variant, j As Long, idx As Long
    Dim labels() As Variant, values() As Variant, colors As Variant

    Set wsBud = ThisWorkbook.Worksheets("Settings_Budget")
    Set wsDash = ThisWorkbook.Worksheets("COMPTES")
    Set dictCat = CreateObject("Scripting.Dictionary")

    ' ======================================
    ' Clear previous output (columns I to K)
    ' ======================================
    wsDash.Range("I13:K24").ClearContents
    wsDash.Range("I12:K12").Value = Array("Mois", "Cat¨¦gorie", "Montant")

    outRow = 13
    lastRow = wsBud.Cells(wsBud.Rows.Count, 1).End(xlUp).Row

    ' ============================
    ' Loop through budget records
    ' ============================
    For i = 2 To lastRow

        budMonth = wsBud.Cells(i, 3).Value

        If budMonth = selectedMonth Then

            If outRow > 24 Then Exit For

            wsDash.Cells(outRow, 9).Value = budMonth          ' Column I Mois
            wsDash.Cells(outRow, 10).Value = wsBud.Cells(i, 1).Value ' Column J Cat¨¦gorie
            wsDash.Cells(outRow, 11).Value = wsBud.Cells(i, 2).Value ' Column K Montant

            cat = wsBud.Cells(i, 1).Value
            amount = wsBud.Cells(i, 2).Value

            If dictCat.exists(cat) Then
                dictCat(cat) = dictCat(cat) + amount
            Else
                dictCat.Add cat, amount
            End If

            outRow = outRow + 1
        End If

    Next i

    If outRow = 13 Then
        MsgBox "Aucune donn¨¦e budget pour ce mois", vbInformation
        Exit Sub
    End If

    wsDash.Range("K13:K" & outRow - 1).NumberFormat = "#,##0.00 [$€-fr-FR]"

    ' =========================
    ' Remove existing charts
    ' =========================
    For Each cht In wsDash.ChartObjects
        cht.Delete
    Next cht

    ' =========================
    ' Create budget pie chart
    ' =========================
    Set cht = wsDash.ChartObjects.Add( _
                Left:=wsDash.Range("I25").Left, _
                Top:=wsDash.Range("I25").Top, _
                Width:=wsDash.Range("I25:K36").Width, _
                Height:=wsDash.Range("I25:K36").Height)

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
            .Name = "Budget R¨¦partition"
        End With

        .HasTitle = True
        .ChartTitle.Text = "Budget par cat¨¦gorie (" & MonthNameFR(selectedMonth) & ")"

        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 12
            .Bold = True
            .Name = "Calibri"
        End With

        ' =======================================================
        ' Apply color palette consistent with the dashboard style
        ' =======================================================
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


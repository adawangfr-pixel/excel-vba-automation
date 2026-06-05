Attribute VB_Name = "Comptes_main"
Option Explicit

Private Sub Workbook_Open()

    Call CheckAndRefresh

End Sub


Public Sub ComptesNotice()

    Dim ws As Worksheet
    Dim todayDate As Date
    Dim firstDay As Date
    Dim nextMonthFirst As Date
    Dim daysInMonth As Long
    Dim dayProgress As Long
    Dim daysLeft As Long
    Dim progressRatio As Double

    Dim i As Integer
    Dim barLength As Integer
    Dim cell As Range

    Dim firstBlack As Boolean
    Dim firstWhite As Boolean

    Set ws = ThisWorkbook.Sheets("comptes")

    ' =========================
    ' Date calculations
    ' =========================
    todayDate = Date
    firstDay = DateSerial(Year(todayDate), Month(todayDate), 1)
    nextMonthFirst = DateSerial(Year(todayDate), Month(todayDate) + 1, 1)

    daysInMonth = nextMonthFirst - firstDay
    dayProgress = todayDate - firstDay + 1
    daysLeft = nextMonthFirst - todayDate

    progressRatio = dayProgress / daysInMonth
    barLength = 17 ' B3:R3

    ' =========================
    ' Initialization flags
    ' =========================
    firstBlack = False
    firstWhite = False

    ' =========================
    ' Black & white progress bar (B3:R3)
    ' =========================
    For i = 1 To barLength

        Set cell = ws.Cells(3, i + 1)
        cell.WrapText = False
        cell.HorizontalAlignment = xlLeft
        cell.VerticalAlignment = xlCenter

        If i <= progressRatio * barLength Then

            ' =========================
            ' Past (black)
            ' =========================
            cell.Interior.Color = RGB(30, 30, 30)
            cell.Font.Color = RGB(255, 255, 255)

            If Not firstBlack Then
                cell.Value = "Mois d¨¦j¨¤ ¨¦coul¨¦ : " & dayProgress & " jours"
                firstBlack = True
            Else
                cell.Value = ""
            End If

        Else

            ' =========================
            ' Future (white)
            ' =========================
            cell.Interior.Color = RGB(255, 255, 255)
            cell.Font.Color = RGB(120, 120, 120)

            If Not firstWhite Then
                cell.Value = "Il reste " & daysLeft & " jours dans le mois"
                firstWhite = True
            Else
                cell.Value = ""
            End If

        End If

    Next i

    ' =========================
    ' Single-line message in B1
    ' =========================
    ws.Range("B1").Value = _
        "Aujourd'hui : " & Format(todayDate, "dd/mm/yyyy") & _
        "   |  Rappel Petit Sou :" & _
        "  ¡ú  Veuillez mettre ¨¤ jour votre budget mensuel et v¨¦rifier vos finances." & vbCrLf & _
        "  ¡ú  R¨¦initialisation du profil recommand¨¦e."
    ' =========================
    ' B1 color logic
    ' =========================
    If daysLeft <= 3 Then
        ws.Range("B1").Font.Color = RGB(192, 0, 0)
    ElseIf daysLeft <= 7 Then
        ws.Range("B1").Font.Color = RGB(0, 150, 150)
    Else
        ws.Range("B1").Font.Color = RGB(90, 90, 90)
    End If

    ws.Range("B1").WrapText = False


End Sub

Public Sub CheckAndRefresh()

    Dim ws As Worksheet
    Dim lastDate As String
    Dim todayStr As String

    Set ws = ThisWorkbook.Sheets("comptes")

    todayStr = Format(Date, "yyyy-mm-dd")
    lastDate = ws.Range("Z1").Value

    If lastDate <> todayStr Then
        Call ComptesNotice
        ws.Range("Z1").Value = todayStr
    End If

End Sub

' =========================
' Popup to stop auto refresh
' =========================
Public Sub StopAutoRefresh()

    On Error Resume Next
    Application.OnTime nextRun, "ComptesNotice", , False

End Sub

' ==============================
' Main entry point
' ==============================
Public Sub ClearUserProfile()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("User_Profile")
    
    ' Demande de confirmation
    If MsgBox("Voulez-vous vraiment supprimer les donn¨¦es utilisateur ?", _
              vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub

    ' Effacer les cellules
    With ws
        .Range("B1:B3,B5:B6").ClearContents
    End With

    MsgBox "Les donn¨¦es du profil utilisateur ont ¨¦t¨¦ supprim¨¦es.", vbInformation, "Succ¨¨s"

End Sub

' ==============================
' Main entry point
' ==============================
Public Sub LaunchApp()

    ' If no user profile ¡ú show questionnaire first
    If Sheets("User_Profile").Range("B1").Value = "" Then
        UserForm_quiz.Show
    End If

End Sub

' ==============================
' Show budget input form
' ==============================
Public Sub ShowBudgetForm()

    UserForm_Budget.Show
    
End Sub

' ==============================
' Show transaction input form
' ==============================
Public Sub ShowTransactionForm()

    UserForm_transaction.Show
    
End Sub

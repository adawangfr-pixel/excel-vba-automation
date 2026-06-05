VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_budget 
   Caption         =   "S¨¦lectonner la devise de base"
   ClientHeight    =   1836
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   4980
   OleObjectBlob   =   "UserForm_budget.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Budget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    Me.txt_monthBudget.Value = Format(Date, "yyyy-mm")

    With Me.cmb_catBudget
        ' --- Expense Categories ---
        .AddItem "Logement et services"
        .AddItem "Achats et Shopping"
        .AddItem "alimentation"
        .AddItem "Transportation"
        .AddItem "Sante et Soins"
        .AddItem "Social et Famille"
        .AddItem "Autres"

        ' --- Revenue Categories ---
        .AddItem "Salaire"
        .AddItem "Placements et Gains"
        .AddItem "Bonus et Extras"
        .AddItem "Autres"
    End With
    
    With Me.cmb_actionBudget
        ' --- ActionBudget  ---
        .AddItem "budget"
        .AddItem "paiements planifi¨¦s"
        .AddItem "transfer"
    End With


End Sub


Private Sub btn_SetBudget_Click()
    Dim wsBudget As Worksheet, wsProfile As Worksheet
    Dim cat As String, action As String
    Dim amt As Double, totalBudget As Double
    Dim cash As Double, recommended As Double
    Dim rowIdx As Variant

    ' ÉèÖÃ¹¤×÷±íÒýÓÃ
    Set wsBudget = ThisWorkbook.Sheets("Settings_Budget")
    Set wsProfile = ThisWorkbook.Sheets("User_Profile")

    ' 1. »ù´¡ÑéÖ¤£ºÈ·±£Àà±ð¡¢½ð¶îºÍ²Ù×÷ÀàÐÍ¶¼ÒÑÌîÐ´
    If Me.cmb_catBudget.Value = "" Or Me.txt_montantCat.Value = "" Or Not IsNumeric(Me.txt_montantCat.Value) Then
        MsgBox "Veuillez s¨¦lectionner une cat¨¦gorie et entrer un montant valide.", vbExclamation
        Exit Sub
    End If

    cat = Me.cmb_catBudget.Value
    amt = CDbl(Me.txt_montantCat.Value)
    mois = CInt(Mid(Me.txt_monthBudget.Value, 6, 2))
    action = Me.cmb_actionBudget.Value
    
    ' 2. Ð´Èë Settings_Budget ±í
    ' Âß¼­£ºÔÚ A ÁÐÕÒÀà±ð£¬ÕÒµ½Ôò¸üÐÂ B ÁÐ½ð¶î£¬ÕÒ²»µ½ÔòÔÚÄ©Î²ÐÂÔö
    Key = cat & "-" & mois
    rowIdx = Application.Match(Key, wsBudget.Range("A:A"), 0)
    
    If IsError(rowIdx) Then
        ' ÐÂÔöÀà±ð
        Dim nextRow As Long
        nextRow = wsBudget.Cells(wsBudget.Rows.Count, 1).End(xlUp).Row + 1
        wsBudget.Cells(nextRow, 1).Value = cat
        wsBudget.Cells(nextRow, 2).Value = amt
        wsBudget.Cells(nextRow, 3).Value = mois
    Else
        ' ¸üÐÂÒÑÓÐÀà±ðµÄ½ð¶î
        wsBudget.Cells(rowIdx, 2).Value = amt
    End If

    ' 3. ¼ÆËãËùÓÐÀà±ðµÄ×ÜÔ¤Ëã²¢´æÈë Profile (B4)
    totalBudget = Application.WorksheetFunction.SumIfs( _
        wsBudget.Range("B:B"), _
        wsBudget.Range("C:C"), mois)
        
    wsProfile.Range("B4").Value = totalBudget

    ' 4. ãÐÖµ¼ì²é (½öµ± Action Îª "budget" Ê±½øÐÐÇ¿Á¦ÌáÐÑ)
    If LCase(action) = "budget" Then
        cash = wsProfile.Range("B5").Value          ' Quiz ÖÐÌîÐ´µÄ¿ÉÓÃÏÖ½ð
        recommended = wsProfile.Range("B6").Value   ' Quiz ½¨ÒéµÄÔ¤Ëã×Ü¶î¶È

        If totalBudget > cash Then
            MsgBox "ALERTE : Le cumul de vos budgets (" & totalBudget & " €) d¨¦passe votre cash disponible (" & cash & " €) !", vbCritical
        ElseIf totalBudget > recommended Then
            MsgBox "Attention : Le cumul de vos budgets (" & totalBudget & " €) d¨¦passe le montant recommand¨¦ (" & recommended & " €).", vbExclamation
        Else
            MsgBox "Budget pour '" & cat & "' enregistr¨¦. Total actuel : " & totalBudget & " €", vbInformation
        End If
    Else
        ' Èç¹ûÊÇ×ªÕË»ò¼Æ»®¸¶¿î£¬½ö¼òµ¥ÌáÊ¾±£´æ³É¹¦
        MsgBox " " & action & " pour '" & cat & "' enregistr¨¦.", vbInformation
    End If

End Sub

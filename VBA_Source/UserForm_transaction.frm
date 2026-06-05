VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_transaction 
   Caption         =   "Ajouter une transaction"
   ClientHeight    =   2352
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   5304
   OleObjectBlob   =   "UserForm_transaction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==========================================
' PARTIE 1 : INITIALISATION DU FORMULAIRE
' ==========================================
Private Sub UserForm_Initialize()
    ' Définir la date par défaut sur aujourd'hui
    Me.txt_Date.Value = Format(Date, "yyyy-mm-dd")

    ' Remplir la liste des Catégories
    With Me.cmb_Category
        .Clear
        ' --- Catégories de Dépenses ---
        .AddItem "Logement et services"
        .AddItem "Achats et Shopping"
        .AddItem "Alimentation"
        .AddItem "Transports"
        .AddItem "Santé et Soins"
        .AddItem "Social et Famille"
        .AddItem "Autres"

        ' --- Catégories de Revenus ---
        .AddItem "Salaire"
        .AddItem "Placements et Gains"
        .AddItem "Bonus et Extras"
        .AddItem "Autres"
    End With
    
    ' Remplir la liste du Type
    With Me.cmb_Type
        .AddItem "Dépense"
        .AddItem "Revenu"
    End With
    
    ' Remplir la liste du Compte
    With Me.cmb_Account
        .AddItem "Courant"
        .AddItem "épargne"
        .AddItem "Crédit"
    End With
End Sub

' ==========================================
' PARTIE 2 : ENREGISTREMENT DES DONNéES
' ==========================================
Private Sub btn_Save_Click()
    Dim wsData As Worksheet
    Dim inputDate As Date
    Dim targetSheet As String
    
    ' --- Validation 1 : Date ---
    If Not IsDate(Me.txt_Date.Value) Then
        MsgBox "Erreur : Format de date invalide", vbCritical, "Petit Sou"
        Exit Sub
    End If
    
    inputDate = CDate(Me.txt_Date.Value)
    If inputDate > Date Then
        MsgBox "Erreur : Les dates futures ne sont pas autorisées", vbCritical, "Petit Sou"
        Exit Sub
    End If
    
    ' --- Validation 2 : Montant ---
    If Not IsNumeric(Me.txt_Amount.Value) Or Val(Me.txt_Amount.Value) <= 0 Then
        MsgBox "Erreur : Montant invalide", vbCritical, "Petit Sou"
        Exit Sub
    End If

    ' --- Validation 3 : Type ---
    If Me.cmb_Type.Value = "" Then
        MsgBox "Erreur : Veuillez sélectionner un Type (Dépense/Revenu)", vbCritical, "Petit Sou"
        Exit Sub
    End If

    ' --- Validation 4 : Compte ---
    If Me.cmb_Account.Value = "" Then
        MsgBox "Erreur : Veuillez sélectionner un Compte", vbCritical, "Petit Sou"
        Exit Sub
    End If
    

    '预算提醒（关键）
    If Me.cmb_Type.Value = "Depense" Then
        Call CheckBudgetAlert(Me.cmb_Category.Value, CDbl(Me.txt_Amount.Value))
    End If
    
    
   ' --- Logique d'aiguillage : Sélection de la feuille selon le Type ---
    If Me.cmb_Type.Value = "Revenu" Then
        targetSheet = "Data_Revenu"
    Else
        targetSheet = "Data_Expenses"
    End If

    ' --- Connexion à la feuille cible ---
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets(targetSheet)
    On Error GoTo 0
    
    If wsData Is Nothing Then
        MsgBox "Erreur : La feuille '" & targetSheet & "' est introuvable !", vbCritical, "Petit Sou"
        Exit Sub
    End If

    ' --- écriture des données sur la feuille ---
    Dim nextRow As Long
    ' Trouver la prochaine ligne vide basée sur la colonne A
    nextRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1

    With wsData
        .Cells(nextRow, 1).Value = inputDate                    ' Col 1: Date
        .Cells(nextRow, 2).Value = CDbl(Me.txt_Amount.Value)    ' Col 2: Montant
        .Cells(nextRow, 2).NumberFormat = "#,##0.00 ""�"""
        .Cells(nextRow, 3).Value = Me.cmb_Category.Value        ' Col 3: Catégorie
        .Cells(nextRow, 4).Value = Me.cmb_Type.Value            ' Col 4: Type
        .Cells(nextRow, 5).Value = Me.txt_Comment.Value         ' Col 5: Commentaire
        .Cells(nextRow, 6).Value = Me.cmb_Account.Value         ' Col 6: Compte
    End With

    MsgBox "Succès : Enregistré dans " & targetSheet, vbInformation, "Petit Sou"
    
    ' --- Réinitialisation de l'interface après enregistrement ---
    Me.txt_Amount.Value = ""
    Me.txt_Comment.Value = ""
    Me.cmb_Category.Value = ""
    Me.cmb_Type.Value = ""
    Me.cmb_Account.Value = ""
End Sub


' ==============================
' 预算提醒（动态阀门）
' ==============================
Private Function CheckBudgetAlert(cat As String, currentAmount As Double) As Boolean
    
    Dim wsBudget As Worksheet
    Dim wsDep As Worksheet
    Dim spentSoFar As Double, monthlyBudget As Double
    Dim percentage As Double
    Dim selectedMonth As Integer
    Dim startDate As Date, endDate As Date

    Set wsBudget = ThisWorkbook.Sheets("Settings_Budget")
    Set wsDep = ThisWorkbook.Sheets("Data_Expenses")

    cat = LCase(Trim(cat))

    selectedMonth = Month(Me.txt_Date.Value)

    startDate = DateSerial(Year(Me.txt_Date.Value), selectedMonth, 1)
    endDate = DateSerial(Year(Me.txt_Date.Value), selectedMonth + 1, 1)

    ' 当前已花
    spentSoFar = Application.WorksheetFunction.SumIfs( _
        wsDep.Range("B:B"), _
        wsDep.Range("C:C"), cat, _
        wsDep.Range("A:A"), ">=" & startDate, _
        wsDep.Range("A:A"), "<" & endDate)

    ' 加上本次输入金额（关键）
    spentSoFar = spentSoFar + currentAmount

    ' 预算
    monthlyBudget = Application.WorksheetFunction.SumIfs( _
        wsBudget.Range("B:B"), _
        wsBudget.Range("A:A"), cat, _
        wsBudget.Range("C:C"), selectedMonth)

    If monthlyBudget > 0 Then
        
        percentage = spentSoFar / monthlyBudget

        Dim warnT As Double, alertT As Double
        warnT = Sheets("User_Profile").Range("B2").Value
        alertT = Sheets("User_Profile").Range("B3").Value

        ' 超预算
        If percentage >= alertT Then
            If MsgBox("Dépassement du budget pour " & cat & vbCrLf & _
                      "Voulez-vous continuer ?", _
                      vbYesNo + vbCritical) = vbNo Then
                CheckBudgetAlert = False
                Exit Function
            End If
        
        ' 接近预算
        ElseIf percentage >= warnT Then
            MsgBox "Attention, vous approchez la limite pour " & cat, vbExclamation
        End If

    End If

    CheckBudgetAlert = True
End Function


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_quiz 
   Caption         =   "Créer votre profil"
   ClientHeight    =   3300
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   6060
   OleObjectBlob   =   "UserForm_quiz.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_quiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    With Me.cmb_Q1
        .AddItem "stressé": .AddItem "incertain"
        .AddItem "stable": .AddItem "confiant"
    End With

    With Me.cmb_Q2
        .AddItem "moi-même": .AddItem "partenaire"
        .AddItem "enfants": .AddItem "animaux"
    End With

    With Me.cmb_Q3
        .AddItem "vacances": .AddItem "retraite"
        .AddItem "investissements": .AddItem "voiture"
    End With

    With Me.cmb_Q4
        .AddItem "pro": .AddItem "abandonne"
        .AddItem "rappeler": .AddItem "nouvelles habitudes"
    End With

    Me.txt_Cash.Value = "0"

End Sub
Private Sub btn_Submit_Click()

    Dim score As Integer
    score = 0

    If cmb_Q1.Value = "" Or cmb_Q2.Value = "" Or _
       cmb_Q3.Value = "" Or cmb_Q4.Value = "" Then

        MsgBox "Veuillez répondre à toutes les questions."
        Exit Sub

    End If

    ' Q1
    Select Case cmb_Q1.Value
        Case "stressé": score = score + 3
        Case "incertain": score = score + 2
        Case "stable": score = score + 1
    End Select

    ' Q2
    Select Case cmb_Q2.Value
        Case "enfants": score = score + 3
        Case "partenaire": score = score + 2
        Case "animaux": score = score + 1
    End Select

    ' Q3
    Select Case cmb_Q3.Value
        Case "retraite", "investissements": score = score + 3
        Case "voiture": score = score + 2
        Case "vacances": score = score + 1
    End Select

    ' Q4
    Select Case cmb_Q4.Value
        Case "abandonne": score = score + 3
        Case "rappeler": score = score + 2
        Case "nouvelles habitudes": score = score + 1
    End Select

    Dim cash As Double

    If Not IsNumeric(txt_Cash.Value) Then
        MsgBox "Veuillez entrer un montant valide."
        Exit Sub
    End If

    cash = CDbl(txt_Cash.Value)

    Call GenerateUserProfile(score, cash)

    MsgBox "Votre profil est créé !"

    Unload Me

End Sub


' ==============================
' 生成用户画像 + 预算策略
' ==============================
Public Sub GenerateUserProfile(score As Integer, cash As Double)

    Dim warnT As Double, alertT As Double, coef As Double

    If score <= 3 Then
        warnT = 0.95: alertT = 1.1: coef = 0.9
    ElseIf score <= 7 Then
        warnT = 0.85: alertT = 1: coef = 0.75
    Else
        warnT = 0.7: alertT = 0.9: coef = 0.6
    End If

    With Sheets("User_Profile")
        .Range("B1").Value = score
        .Range("B2").Value = warnT
        .Range("B3").Value = alertT
        
        .Range("B5").Value = cash
        .Range("B6").Value = cash * coef   ' 推荐预算
    End With

End Sub



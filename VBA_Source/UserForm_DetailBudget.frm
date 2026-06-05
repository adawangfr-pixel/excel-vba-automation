VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_DetailBudget 
   Caption         =   "Voir le détail de budget"
   ClientHeight    =   1260
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   4704
   OleObjectBlob   =   "UserForm_DetailBudget.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_DetailBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 窗体加载时填充月份
Private Sub UserForm_Initialize()
    Dim moisNoms As Variant
    Dim i As Integer
    
    ' 设置标签引导语
    ' 如果你有个 Label 控件，可以设置：Label1.Caption = "Quel mois souhaitez-vous consulter ?"
    
    moisNoms = Array("1 - Janvier", "2 - Février", "3 - Mars", "4 - Avril", _
                     "5 - Mai", "6 - Juin", "7 - Juillet", "8 - Ao?t", _
                     "9 - Septembre", "10 - Octobre", "11 - Novembre", "12 - Décembre")
    
    With Me.cmb_MoisB
        .Clear
        For i = 0 To UBound(moisNoms)
            .AddItem moisNoms(i)
        Next i
        ' 默认选择当前月份
        .ListIndex = Month(Date) - 1
    End With
End Sub

' 点击“确认更新”按钮
Private Sub cmdUpdateB_Click()
    Dim moisNum As Integer
    
    ' 1. 验证选择
    If Me.cmb_MoisB.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner un mois.", vbExclamation, "Attention"
        Exit Sub
    End If
    
    ' 2. 获取月份数字 (0对应1月，所以要+1)
    moisNum = Me.cmb_MoisB.ListIndex + 1
    
    ' 3. 调用模块中的 Public Sub
    Call FillBudgetAndPie(moisNum)
    
    ' 4. 完成后提示或关闭
    ' MsgBox "Mise à jour réussie !", vbInformation
    Unload Me
End Sub

' 点击“关闭”按钮
Private Sub cmdCloseB_Click()
    Unload Me
End Sub




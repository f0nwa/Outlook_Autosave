VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileTypesForm 
   Caption         =   "Attachment file types"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   OleObjectBlob   =   "FileTypesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileTypesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents btnAdd As MSForms.CommandButton
Attribute btnAdd.VB_VarHelpID = -1
Private WithEvents btnRemove As MSForms.CommandButton
Attribute btnRemove.VB_VarHelpID = -1
Private WithEvents btnSave As MSForms.CommandButton
Attribute btnSave.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1

Private lstTypes As MSForms.ListBox
Private txtExtension As MSForms.TextBox

Private Sub UserForm_Initialize()
    BuildUi
    LoadFileTypes
End Sub

Private Sub BuildUi()
    Dim lblTitle As MSForms.Label
    Dim lblHint As MSForms.Label
    Dim lblAdd As MSForms.Label
    
    Me.Caption = RuText(1058, 1080, 1087, 1099, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081)
    Me.Width = 500
    Me.Height = 410
    
    Set lblTitle = AddLabel("lblTitle", RuText(1058, 1080, 1087, 1099, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081), 18, 12, 320, 18)
    lblTitle.Font.Bold = True
    
    Set lblHint = AddLabel("lblHint", RuText(1054, 1090, 1084, 1077, 1095, 1077, 1085, 1085, 1099, 1077, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1103, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1103, 1102, 1090, 1089, 1103, 46, 32, 1053, 1077, 1086, 1090, 1084, 1077, 1095, 1077, 1085, 1085, 1099, 1077, 32, 1086, 1089, 1090, 1072, 1102, 1090, 1089, 1103, 32, 1074, 32, 1089, 1087, 1080, 1089, 1082, 1077, 44, 32, 1085, 1086, 32, 1086, 1090, 1082, 1083, 1102, 1095, 1077, 1085, 1099, 46), 18, 36, 430, 30)
    
    Set lstTypes = Me.Controls.Add("Forms.ListBox.1", "lstTypes", True)
    With lstTypes
        .Left = 18
        .Top = 74
        .Width = 430
        .Height = 206
        .ColumnCount = 2
        .ColumnWidths = "390 pt;0 pt"
        .ListStyle = fmListStyleOption
        .MultiSelect = fmMultiSelectMulti
        .IntegralHeight = False
    End With
    
    Set lblAdd = AddLabel("lblAdd", RuText(1044, 1086, 1073, 1072, 1074, 1080, 1090, 1100, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1077), 18, 294, 130, 18)
    
    Set txtExtension = Me.Controls.Add("Forms.TextBox.1", "txtExtension", True)
    With txtExtension
        .Left = 18
        .Top = 316
        .Width = 120
        .Height = 22
    End With
    
    Set btnAdd = AddButton("btnAdd", RuText(1044, 1086, 1073, 1072, 1074, 1080, 1090, 1100), 150, 314, 90, 24, RuText(1044, 1086, 1073, 1072, 1074, 1080, 1090, 1100, 32, 1074, 1074, 1077, 1076, 1077, 1085, 1085, 1086, 1077, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1077, 32, 1074, 32, 1089, 1087, 1080, 1089, 1086, 1082, 32, 1080, 32, 1074, 1082, 1083, 1102, 1095, 1080, 1090, 1100, 32, 1077, 1075, 1086, 46))
    Set btnRemove = AddButton("btnRemove", RuText(1059, 1076, 1072, 1083, 1080, 1090, 1100), 250, 314, 90, 24, RuText(1059, 1076, 1072, 1083, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1077, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1103, 32, 1080, 1079, 32, 1089, 1087, 1080, 1089, 1082, 1072, 46))
    Set btnSave = AddButton("btnSave", RuText(1057, 1086, 1093, 1088, 1072, 1085, 1080, 1090, 1100), 18, 354, 100, 26, RuText(1057, 1086, 1093, 1088, 1072, 1085, 1080, 1090, 1100, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1080, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1081, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 46))
    Set btnCancel = AddButton("btnCancel", RuText(1054, 1090, 1084, 1077, 1085, 1072), 358, 354, 90, 26, RuText(1047, 1072, 1082, 1088, 1099, 1090, 1100, 32, 1073, 1077, 1079, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1080, 1079, 1084, 1077, 1085, 1077, 1085, 1080, 1081, 46))
End Sub

Private Sub LoadFileTypes()
    Dim rules As String
    Dim lines() As String
    Dim line As Variant
    Dim parts() As String
    Dim rowIndex As Long
    
    rules = FT_GetFileTypeRules()
    
    If Len(Trim(rules)) = 0 Then
        rules = FT_BuildFileTypeRulesFromAllowedExtensions(FT_GetAllowedAttachmentExtensions())
    End If
    
    rules = FT_NormalizeFileTypeRules(rules)
    lstTypes.Clear
    
    If Len(Trim(rules)) = 0 Then
        Exit Sub
    End If
    
    lines = Split(rules, vbLf)
    
    For Each line In lines
        If Len(Trim(CStr(line))) > 0 Then
            parts = Split(CStr(line), "|")
            
            If UBound(parts) >= 1 Then
                lstTypes.AddItem "." & parts(1)
                rowIndex = lstTypes.ListCount - 1
                lstTypes.List(rowIndex, 1) = parts(1)
                
                If parts(0) = "1" Then
                    lstTypes.Selected(rowIndex) = True
                End If
            End If
        End If
    Next line
End Sub

Private Function AddLabel(ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single, ByVal heightValue As Single) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    
    With AddLabel
        .Caption = captionText
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
    End With
End Function

Private Function AddButton(ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single, ByVal heightValue As Single, Optional ByVal tipText As String = "") As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    
    With AddButton
        .Caption = captionText
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
        .ControlTipText = tipText
    End With
End Function

Private Sub btnAdd_Click()
    Dim extensionValue As String
    Dim i As Long
    
    extensionValue = FT_CleanFileExtension(txtExtension.Text)
    
    If Len(extensionValue) = 0 Then
        MsgBox RuText(1042, 1074, 1077, 1076, 1080, 1090, 1077, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1077, 44, 32, 1085, 1072, 1087, 1088, 1080, 1084, 1077, 1088, 58, 32, 112, 100, 102), vbInformation
        Exit Sub
    End If
    
    For i = 0 To lstTypes.ListCount - 1
        If StrComp(CStr(lstTypes.List(i, 1)), extensionValue, vbTextCompare) = 0 Then
            lstTypes.Selected(i) = True
            txtExtension.Text = ""
            Exit Sub
        End If
    Next i
    
    lstTypes.AddItem "." & extensionValue
    i = lstTypes.ListCount - 1
    lstTypes.List(i, 1) = extensionValue
    lstTypes.Selected(i) = True
    txtExtension.Text = ""
End Sub

Private Sub btnRemove_Click()
    Dim i As Long
    
    For i = lstTypes.ListCount - 1 To 0 Step -1
        If lstTypes.Selected(i) Then
            lstTypes.RemoveItem i
        End If
    Next i
End Sub

Private Sub btnSave_Click()
    Dim i As Long
    Dim rules As String
    
    For i = 0 To lstTypes.ListCount - 1
        If lstTypes.Selected(i) Then
            rules = rules & "1|" & CStr(lstTypes.List(i, 1)) & vbLf
        Else
            rules = rules & "0|" & CStr(lstTypes.List(i, 1)) & vbLf
        End If
    Next i
    
    FT_SaveFileTypeRules rules
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SortRulesForm 
   Caption         =   "Attachment sort rules"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   OleObjectBlob   =   "SortRulesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SortRulesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CTRL_KEY_MASK As Integer = 2

Private WithEvents lstRules As MSForms.ListBox
Attribute lstRules.VB_VarHelpID = -1
Private WithEvents btnAdd As MSForms.CommandButton
Attribute btnAdd.VB_VarHelpID = -1
Private WithEvents btnUpdate As MSForms.CommandButton
Attribute btnUpdate.VB_VarHelpID = -1
Private WithEvents btnRemove As MSForms.CommandButton
Attribute btnRemove.VB_VarHelpID = -1
Private WithEvents btnToggleActive As MSForms.CommandButton
Attribute btnToggleActive.VB_VarHelpID = -1
Private WithEvents btnBrowse As MSForms.CommandButton
Attribute btnBrowse.VB_VarHelpID = -1
Private WithEvents btnSourceBrowse As MSForms.CommandButton
Attribute btnSourceBrowse.VB_VarHelpID = -1
Private WithEvents btnSourceCurrent As MSForms.CommandButton
Attribute btnSourceCurrent.VB_VarHelpID = -1
Private WithEvents btnSave As MSForms.CommandButton
Attribute btnSave.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1

Private txtPattern As MSForms.TextBox
Private txtTarget As MSForms.TextBox
Private txtExtensions As MSForms.TextBox
Private txtSourceFolder As MSForms.TextBox
Private txtSubject As MSForms.TextBox
Private txtRuleName As MSForms.TextBox
Private isInitializing As Boolean

Private Sub UserForm_Initialize()
    isInitializing = True
    BuildUi
    LoadRules
    isInitializing = False
End Sub

Private Sub BuildUi()
    Dim lblTitle As MSForms.Label
    Dim lblHint As MSForms.Label
    Dim lblPattern As MSForms.Label
    Dim lblExtensions As MSForms.Label
    Dim lblSubject As MSForms.Label
    Dim lblTarget As MSForms.Label
    Dim lblSourceFolder As MSForms.Label
    Dim lblRuleName As MSForms.Label
    
    Me.Caption = RuText(1055, 1088, 1072, 1074, 1080, 1083, 1072, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081)
    Me.Width = 610
    Me.Height = 620
    
    Set lblTitle = AddLabel("lblTitle", RuText(1055, 1088, 1072, 1074, 1080, 1083, 1072, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081), 18, 12, 420, 18)
    lblTitle.Font.Bold = True
    
    Set lblHint = AddLabel("lblHint", RuText(1043, 1072, 1083, 1082, 1072, 32, 1079, 1085, 1072, 1095, 1080, 1090, 32, 1072, 1082, 1090, 1080, 1074, 1085, 1086, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 46, 32, 1057, 1080, 1085, 1077, 1077, 32, 1074, 1099, 1076, 1077, 1083, 1077, 1085, 1080, 1077, 32, 1090, 1086, 1083, 1100, 1082, 1086, 32, 1074, 1099, 1073, 1080, 1088, 1072, 1077, 1090, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 46, 32, 1044, 1074, 1086, 1081, 1085, 1086, 1081, 32, 1082, 1083, 1080, 1082, 32, 1074, 1082, 1083, 1102, 1095, 1072, 1077, 1090, 32, 1080, 1083, 1080, 32, 1086, 1090, 1082, 1083, 1102, 1095, 1072, 1077, 1090, 46), 18, 36, 560, 18)
    
    Set lstRules = Me.Controls.Add("Forms.ListBox.1", "lstRules", True)
    With lstRules
        .Left = 18
        .Top = 62
        .Width = 548
        .Height = 132
        .ColumnCount = 9
        .ColumnWidths = "18 pt;502 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt"
        .ListStyle = fmListStylePlain
        .MultiSelect = fmMultiSelectExtended
        .IntegralHeight = False
        .Font.Name = "Segoe UI"
        .Font.Size = 9
    End With
    
    Set lblRuleName = AddLabel("lblRuleName", RuText(1053, 1072, 1079, 1074, 1072, 1085, 1080, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1072), 18, 208, 140, 18)
    Set txtRuleName = Me.Controls.Add("Forms.TextBox.1", "txtRuleName", True)
    With txtRuleName
        .Left = 18
        .Top = 230
        .Width = 460
        .Height = 22
    End With
    
    Set lblPattern = AddLabel("lblPattern", RuText(1058, 1077, 1082, 1089, 1090, 32, 1074, 32, 1080, 1084, 1077, 1085, 1080, 32, 1092, 1072, 1081, 1083, 1072), 18, 254, 120, 18)
    Set txtPattern = Me.Controls.Add("Forms.TextBox.1", "txtPattern", True)
    With txtPattern
        .Left = 18
        .Top = 276
        .Width = 220
        .Height = 22
    End With
    
    Set lblExtensions = AddLabel("lblExtensions", RuText(1056, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1103), 258, 254, 120, 18)
    Set txtExtensions = Me.Controls.Add("Forms.TextBox.1", "txtExtensions", True)
    With txtExtensions
        .Left = 258
        .Top = 276
        .Width = 120
        .Height = 22
    End With
    
    Set lblSubject = AddLabel("lblSubject", RuText(1058, 1077, 1084, 1072, 32, 1087, 1080, 1089, 1100, 1084, 1072), 18, 308, 120, 18)
    Set txtSubject = Me.Controls.Add("Forms.TextBox.1", "txtSubject", True)
    With txtSubject
        .Left = 18
        .Top = 330
        .Width = 460
        .Height = 22
    End With
    
    Set lblTarget = AddLabel("lblTarget", RuText(1062, 1077, 1083, 1077, 1074, 1072, 1103, 32, 1087, 1072, 1087, 1082, 1072), 18, 362, 120, 18)
    Set txtTarget = Me.Controls.Add("Forms.TextBox.1", "txtTarget", True)
    With txtTarget
        .Left = 18
        .Top = 384
        .Width = 460
        .Height = 22
    End With
    
    Set btnBrowse = AddButton("btnBrowse", RuText(1054, 1073, 1079, 1086, 1088), 488, 382, 78, 24, RuText(1042, 1099, 1073, 1088, 1072, 1090, 1100, 32, 1094, 1077, 1083, 1077, 1074, 1091, 1102, 32, 1087, 1072, 1087, 1082, 1091, 32, 1076, 1083, 1103, 32, 1101, 1090, 1086, 1075, 1086, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1072, 46))
    
    Set lblSourceFolder = AddLabel("lblSourceFolder", RuText(1048, 1089, 1093, 1086, 1076, 1085, 1072, 1103, 32, 1087, 1072, 1087, 1082, 1072, 32, 79, 117, 116, 108, 111, 111, 107), 18, 416, 180, 18)
    Set txtSourceFolder = Me.Controls.Add("Forms.TextBox.1", "txtSourceFolder", True)
    With txtSourceFolder
        .Left = 18
        .Top = 438
        .Width = 372
        .Height = 22
    End With
    
    Set btnSourceBrowse = AddButton("btnSourceBrowse", RuText(1042, 1099, 1073, 1088, 1072, 1090, 1100), 400, 436, 78, 24, RuText(1042, 1099, 1073, 1088, 1072, 1090, 1100, 32, 1080, 1089, 1093, 1086, 1076, 1085, 1091, 1102, 32, 1087, 1072, 1087, 1082, 1091, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1080, 1079, 32, 1089, 1087, 1080, 1089, 1082, 1072, 46))
    Set btnSourceCurrent = AddButton("btnSourceCurrent", RuText(1058, 1077, 1082, 1091, 1097, 1072, 1103), 488, 436, 78, 24, RuText(1042, 1079, 1103, 1090, 1100, 32, 1090, 1077, 1082, 1091, 1097, 1091, 1102, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1091, 1102, 32, 1087, 1072, 1087, 1082, 1091, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1082, 1072, 1082, 32, 1080, 1089, 1093, 1086, 1076, 1085, 1091, 1102, 32, 1076, 1083, 1103, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1072, 46))
    Set btnAdd = AddButton("btnAdd", RuText(1044, 1086, 1073, 1072, 1074, 1080, 1090, 1100), 18, 478, 90, 26, RuText(1044, 1086, 1073, 1072, 1074, 1080, 1090, 1100, 32, 1096, 1072, 1073, 1083, 1086, 1085, 44, 32, 1089, 1087, 1080, 1089, 1086, 1082, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1081, 32, 1080, 32, 1094, 1077, 1083, 1077, 1074, 1091, 1102, 32, 1087, 1072, 1087, 1082, 1091, 32, 1082, 1072, 1082, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 46))
    Set btnUpdate = AddButton("btnUpdate", RuText(1054, 1073, 1085, 1086, 1074, 1080, 1090, 1100), 120, 478, 90, 26, RuText(1054, 1073, 1085, 1086, 1074, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1086, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 32, 1090, 1077, 1082, 1091, 1097, 1080, 1084, 1080, 32, 1079, 1085, 1072, 1095, 1077, 1085, 1080, 1103, 1084, 1080, 46))
    Set btnRemove = AddButton("btnRemove", RuText(1059, 1076, 1072, 1083, 1080, 1090, 1100), 222, 478, 90, 26, RuText(1059, 1076, 1072, 1083, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1086, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 46))
    Set btnToggleActive = AddButton("btnToggleActive", RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100), 324, 478, 104, 26, RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1086, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 46))
    Set btnSave = AddButton("btnSave", RuText(1057, 1086, 1093, 1088, 1072, 1085, 1080, 1090, 1100), 18, 520, 100, 26, RuText(1057, 1086, 1093, 1088, 1072, 1085, 1080, 1090, 1100, 32, 1074, 1089, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1072, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080, 32, 1080, 32, 1079, 1072, 1082, 1088, 1099, 1090, 1100, 32, 1086, 1082, 1085, 1086, 46))
    Set btnCancel = AddButton("btnCancel", RuText(1054, 1090, 1084, 1077, 1085, 1072), 468, 520, 98, 26, RuText(1047, 1072, 1082, 1088, 1099, 1090, 1100, 32, 1073, 1077, 1079, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1080, 1079, 1084, 1077, 1085, 1077, 1085, 1080, 1081, 46))
End Sub

Private Sub LoadRules()
    Dim rows As String
    Dim lines() As String
    Dim line As Variant
    Dim parts() As String
    
    rows = SR_GetSortRuleRows()
    
    If Len(Trim(rows)) = 0 Then
        rows = SR_BuildRowsFromLegacyRules(SR_GetLegacySortRules())
    End If
    
    rows = SR_NormalizeSortRuleRows(rows)
    lstRules.Clear
    
    If Len(Trim(rows)) = 0 Then
        Exit Sub
    End If
    
    lines = Split(rows, vbLf)
    
    For Each line In lines
        If Len(Trim(CStr(line))) > 0 Then
            parts = Split(CStr(line), "|")
            
            If UBound(parts) >= 2 Then
                If UBound(parts) >= 6 Then
                    AddRuleRow parts(1), parts(2), parts(3), parts(4), parts(5), parts(6), (parts(0) = "1")
                ElseIf UBound(parts) >= 5 Then
                    AddRuleRow parts(1), parts(2), parts(3), parts(4), parts(5), "", (parts(0) = "1")
                ElseIf UBound(parts) >= 4 Then
                    AddRuleRow parts(1), parts(2), parts(3), parts(4), "", "", (parts(0) = "1")
                ElseIf UBound(parts) >= 3 Then
                    AddRuleRow parts(1), parts(2), parts(3), "", "", "", (parts(0) = "1")
                Else
                    AddRuleRow parts(1), parts(2), "", "", "", "", (parts(0) = "1")
                End If
            End If
        End If
    Next line
End Sub

Private Sub AddRuleRow(ByVal patternValue As String, ByVal targetValue As String, ByVal extensionsValue As String, ByVal sourceFolderValue As String, ByVal subjectValue As String, ByVal displayNameValue As String, ByVal enabledValue As Boolean)
    Dim rowIndex As Long
    
    lstRules.AddItem ""
    rowIndex = lstRules.ListCount - 1
    lstRules.List(rowIndex, 2) = patternValue
    lstRules.List(rowIndex, 3) = targetValue
    lstRules.List(rowIndex, 4) = extensionsValue
    lstRules.List(rowIndex, 5) = sourceFolderValue
    lstRules.List(rowIndex, 6) = subjectValue
    lstRules.List(rowIndex, 7) = RuleEnabledValue(enabledValue)
    lstRules.List(rowIndex, 8) = displayNameValue
    
    RenumberRuleDisplayTexts
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

Private Function FindRuleIndex(ByVal patternValue As String, ByVal extensionsValue As String, ByVal sourceFolderValue As String, ByVal subjectValue As String) As Long
    Dim i As Long
    
    For i = 0 To lstRules.ListCount - 1
        If StrComp(CStr(lstRules.List(i, 2)), patternValue, vbTextCompare) = 0 And StrComp(CStr(lstRules.List(i, 4)), extensionsValue, vbTextCompare) = 0 And StrComp(CStr(lstRules.List(i, 5)), sourceFolderValue, vbTextCompare) = 0 And StrComp(CStr(lstRules.List(i, 6)), subjectValue, vbTextCompare) = 0 Then
            FindRuleIndex = i
            Exit Function
        End If
    Next i
    
    FindRuleIndex = -1
End Function

Private Function SelectedRuleIndex() As Long
    Dim i As Long
    Dim selectedCount As Long
    Dim selectedIndex As Long
    
    selectedIndex = -1
    
    For i = 0 To lstRules.ListCount - 1
        If lstRules.Selected(i) Then
            selectedCount = selectedCount + 1
            selectedIndex = i
        End If
    Next i
    
    If selectedCount = 1 Then
        SelectedRuleIndex = selectedIndex
    Else
        SelectedRuleIndex = -1
    End If
End Function

Private Function ActiveRuleIndex() As Long
    Dim selectedIndex As Long
    
    selectedIndex = SelectedRuleIndex()
    
    If selectedIndex >= 0 Then
        ActiveRuleIndex = selectedIndex
    ElseIf SelectedRuleCount() = 0 And lstRules.ListIndex >= 0 Then
        ActiveRuleIndex = lstRules.ListIndex
    Else
        ActiveRuleIndex = -1
    End If
End Function

Private Function SelectedRuleCount() As Long
    Dim i As Long
    
    For i = 0 To lstRules.ListCount - 1
        If lstRules.Selected(i) Then
            SelectedRuleCount = SelectedRuleCount + 1
        End If
    Next i
End Function

Private Function HasSelectedInactiveRules() As Boolean
    Dim i As Long
    
    For i = 0 To lstRules.ListCount - 1
        If lstRules.Selected(i) And Not IsRuleRowEnabled(i) Then
            HasSelectedInactiveRules = True
            Exit Function
        End If
    Next i
End Function

Private Sub FillFieldsFromRow(ByVal rowIndex As Long)
    If rowIndex < 0 Or rowIndex >= lstRules.ListCount Then
        Exit Sub
    End If
    
    txtPattern.Text = CStr(lstRules.List(rowIndex, 2))
    txtTarget.Text = CStr(lstRules.List(rowIndex, 3))
    txtExtensions.Text = CStr(lstRules.List(rowIndex, 4))
    txtSourceFolder.Text = CStr(lstRules.List(rowIndex, 5))
    txtSubject.Text = CStr(lstRules.List(rowIndex, 6))
    txtRuleName.Text = CStr(lstRules.List(rowIndex, 8))
End Sub

Private Sub ClearRuleFields()
    txtPattern.Text = ""
    txtTarget.Text = ""
    txtExtensions.Text = ""
    txtSourceFolder.Text = ""
    txtSubject.Text = ""
    txtRuleName.Text = ""
End Sub

Private Function ReadPatternValue() As String
    ReadPatternValue = SR_CleanRulePart(txtPattern.Text)
End Function

Private Function ReadTargetValue() As String
    ReadTargetValue = SR_CleanTargetValue(txtTarget.Text)
End Function

Private Function ReadExtensionsValue() As String
    ReadExtensionsValue = SR_CleanExtensions(txtExtensions.Text)
End Function

Private Function ReadSourceFolderValue() As String
    ReadSourceFolderValue = SR_CleanFolderPath(txtSourceFolder.Text)
End Function

Private Function ReadSubjectValue() As String
    ReadSubjectValue = SR_CleanRulePart(txtSubject.Text)
End Function

Private Function ReadRuleNameValue() As String
    ReadRuleNameValue = SR_CleanRulePart(txtRuleName.Text)
End Function

Private Function HasRuleFieldValues() As Boolean
    HasRuleFieldValues = (Len(Trim(txtPattern.Text)) > 0 Or Len(Trim(txtTarget.Text)) > 0 Or Len(Trim(txtExtensions.Text)) > 0 Or Len(Trim(txtSourceFolder.Text)) > 0 Or Len(Trim(txtSubject.Text)) > 0 Or Len(Trim(txtRuleName.Text)) > 0)
End Function

Private Function RuleEnabledValue(ByVal enabledValue As Boolean) As String
    If enabledValue Then
        RuleEnabledValue = "1"
    Else
        RuleEnabledValue = "0"
    End If
End Function

Private Function IsRuleRowEnabled(ByVal rowIndex As Long) As Boolean
    If rowIndex < 0 Or rowIndex >= lstRules.ListCount Then
        IsRuleRowEnabled = False
    Else
        IsRuleRowEnabled = (CStr(lstRules.List(rowIndex, 7)) = "1")
    End If
End Function

Private Sub RefreshToggleActiveButton()
    Dim selectedCount As Long
    
    selectedCount = SelectedRuleCount()
    
    If selectedCount = 0 Then
        btnToggleActive.Caption = RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100)
        btnToggleActive.ControlTipText = RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1086, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 46)
        Exit Sub
    End If
    
    If HasSelectedInactiveRules() Then
        btnToggleActive.Caption = RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100)
        btnToggleActive.ControlTipText = RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1086, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 46)
    Else
        btnToggleActive.Caption = RuText(1054, 1090, 1082, 1083, 1102, 1095, 1080, 1090, 1100)
        btnToggleActive.ControlTipText = RuText(1054, 1090, 1082, 1083, 1102, 1095, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1086, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 46)
    End If
End Sub

Private Sub SelectSingleRule(ByVal rowIndex As Long)
    Dim i As Long
    
    For i = 0 To lstRules.ListCount - 1
        lstRules.Selected(i) = (i = rowIndex)
    Next i
    
    lstRules.ListIndex = rowIndex
End Sub

Private Sub SelectAllRules()
    Dim i As Long
    
    For i = 0 To lstRules.ListCount - 1
        lstRules.Selected(i) = True
    Next i
    
    If lstRules.ListCount > 0 Then
        lstRules.ListIndex = 0
    End If
    
    ClearRuleFields
    RefreshToggleActiveButton
End Sub

Private Sub RenumberRuleDisplayTexts()
    Dim i As Long
    
    For i = 0 To lstRules.ListCount - 1
        lstRules.List(i, 0) = BuildRuleActiveMarker(IsRuleRowEnabled(i))
        lstRules.List(i, 1) = BuildRuleDisplayText(CStr(lstRules.List(i, 2)), CStr(lstRules.List(i, 3)), CStr(lstRules.List(i, 4)), CStr(lstRules.List(i, 5)), CStr(lstRules.List(i, 6)), CStr(lstRules.List(i, 8)), i + 1)
    Next i
End Sub

Private Function BuildRuleActiveMarker(ByVal enabledValue As Boolean) As String
    If enabledValue Then
        BuildRuleActiveMarker = RuText(10003)
    Else
        BuildRuleActiveMarker = ""
    End If
End Function

Private Function BuildRuleDisplayText(ByVal patternValue As String, ByVal targetValue As String, ByVal extensionsValue As String, ByVal sourceFolderValue As String, ByVal subjectValue As String, ByVal displayNameValue As String, Optional ByVal ruleNumber As Long = 0) As String
    If Len(displayNameValue) > 0 Then
        BuildRuleDisplayText = displayNameValue
    ElseIf Len(patternValue) > 0 Then
        BuildRuleDisplayText = patternValue
    Else
        BuildRuleDisplayText = RuText(1042, 1089, 1077, 32, 1092, 1072, 1081, 1083, 1099)
    End If
    
    If Len(displayNameValue) = 0 And Len(extensionsValue) > 0 Then
        BuildRuleDisplayText = BuildRuleDisplayText & " [" & extensionsValue & "]"
    End If
    
    If Len(displayNameValue) = 0 And Len(sourceFolderValue) > 0 Then
        BuildRuleDisplayText = BuildRuleDisplayText & " (" & RuText(1055, 1072, 1087, 1082, 1072, 58, 32) & sourceFolderValue & ")"
    End If
    
    If Len(displayNameValue) = 0 And Len(subjectValue) > 0 Then
        BuildRuleDisplayText = BuildRuleDisplayText & " (" & RuText(1058, 1077, 1084, 1072, 58, 32) & subjectValue & ")"
    End If
    
    If Len(displayNameValue) = 0 Then
        BuildRuleDisplayText = BuildRuleDisplayText & " -> " & targetValue
    End If
    
    If ruleNumber > 0 Then
        BuildRuleDisplayText = CStr(ruleNumber) & ". " & BuildRuleDisplayText
    End If
End Function

Private Function BuildDuplicateRuleMessage(ByVal rowIndex As Long) As String
    Dim displayNameValue As String
    
    BuildDuplicateRuleMessage = RuText(1058, 1072, 1082, 1086, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 32, 1091, 1078, 1077, 32, 1089, 1091, 1097, 1077, 1089, 1090, 1074, 1091, 1077, 1090, 58, 32, 8470) & CStr(rowIndex + 1)
    
    If rowIndex < 0 Or rowIndex >= lstRules.ListCount Then
        Exit Function
    End If
    
    displayNameValue = CStr(lstRules.List(rowIndex, 8))
    
    If Len(displayNameValue) > 0 Then
        BuildDuplicateRuleMessage = BuildDuplicateRuleMessage & " """ & displayNameValue & """"
    Else
        BuildDuplicateRuleMessage = BuildDuplicateRuleMessage & vbCrLf & CStr(lstRules.List(rowIndex, 1))
    End If
End Function

Private Function IsRuleDefinitionValid(ByVal patternValue As String, ByVal targetValue As String, ByVal extensionsValue As String, ByVal sourceFolderValue As String) As Boolean
    IsRuleDefinitionValid = (Len(targetValue) > 0 And IsAbsoluteFolderPath(targetValue) And (Len(patternValue) > 0 Or (Len(extensionsValue) > 0 And Len(sourceFolderValue) > 0)))
End Function

Private Function RuleValidationMessage() As String
    RuleValidationMessage = RuText(1062, 1077, 1083, 1077, 1074, 1072, 1103, 32, 1087, 1072, 1087, 1082, 1072, 32, 1076, 1086, 1083, 1078, 1085, 1072, 32, 1073, 1099, 1090, 1100, 32, 1072, 1073, 1089, 1086, 1083, 1102, 1090, 1085, 1086, 1081, 46, 32, 1045, 1089, 1083, 1080, 32, 1090, 1077, 1082, 1089, 1090, 32, 1080, 1084, 1077, 1085, 1080, 32, 1092, 1072, 1081, 1083, 1072, 32, 1087, 1091, 1089, 1090, 1086, 1081, 44, 32, 1079, 1072, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1103, 32, 1080, 32, 1080, 1089, 1093, 1086, 1076, 1085, 1091, 1102, 32, 1087, 1072, 1087, 1082, 1091, 32, 79, 117, 116, 108, 111, 111, 107, 46)
End Function

Private Sub UpsertRule(ByVal preserveExistingEnabled As Boolean)
    UpsertRuleFromFields preserveExistingEnabled
End Sub

Private Function UpsertRuleFromFields(ByVal preserveExistingEnabled As Boolean) As Boolean
    Dim patternValue As String
    Dim targetValue As String
    Dim extensionsValue As String
    Dim sourceFolderValue As String
    Dim subjectValue As String
    Dim displayNameValue As String
    Dim rowIndex As Long
    Dim enabledValue As Boolean
    
    patternValue = ReadPatternValue()
    targetValue = ReadTargetValue()
    extensionsValue = ReadExtensionsValue()
    sourceFolderValue = ReadSourceFolderValue()
    subjectValue = ReadSubjectValue()
    displayNameValue = ReadRuleNameValue()
    
    If Not IsRuleDefinitionValid(patternValue, targetValue, extensionsValue, sourceFolderValue) Then
        MsgBox RuleValidationMessage(), vbInformation
        UpsertRuleFromFields = False
        Exit Function
    End If
    
    rowIndex = FindRuleIndex(patternValue, extensionsValue, sourceFolderValue, subjectValue)
    
    If rowIndex >= 0 Then
        enabledValue = IsRuleRowEnabled(rowIndex)
        lstRules.List(rowIndex, 2) = patternValue
        lstRules.List(rowIndex, 3) = targetValue
        lstRules.List(rowIndex, 4) = extensionsValue
        lstRules.List(rowIndex, 5) = sourceFolderValue
        lstRules.List(rowIndex, 6) = subjectValue
        lstRules.List(rowIndex, 8) = displayNameValue
        
        If Not preserveExistingEnabled Then
            enabledValue = True
        End If
        
        lstRules.List(rowIndex, 7) = RuleEnabledValue(enabledValue)
        SelectSingleRule rowIndex
        RenumberRuleDisplayTexts
        RefreshToggleActiveButton
    Else
        AddRuleRow patternValue, targetValue, extensionsValue, sourceFolderValue, subjectValue, displayNameValue, True
        SelectSingleRule lstRules.ListCount - 1
        RefreshToggleActiveButton
    End If
    
    UpsertRuleFromFields = True
End Function

Private Sub RefreshRuleSelectionDetails()
    If isInitializing Then
        Exit Sub
    End If
    
    If txtPattern Is Nothing Or txtTarget Is Nothing Or txtExtensions Is Nothing Or txtSourceFolder Is Nothing Or txtSubject Is Nothing Or txtRuleName Is Nothing Then
        Exit Sub
    End If
    
    If ActiveRuleIndex() >= 0 Then
        FillFieldsFromRow ActiveRuleIndex()
    Else
        ClearRuleFields
    End If
    
    RefreshToggleActiveButton
End Sub

Private Sub lstRules_Click()
    RefreshRuleSelectionDetails
End Sub

Private Sub lstRules_Change()
    RefreshRuleSelectionDetails
End Sub

Private Sub lstRules_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ToggleSelectedRuleActive
End Sub

Private Sub lstRules_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyA And (Shift And CTRL_KEY_MASK) <> 0 Then
        SelectAllRules
        KeyCode = 0
    End If
End Sub

Private Sub btnAdd_Click()
    UpsertRule False
End Sub

Private Sub btnUpdate_Click()
    Dim rowIndex As Long
    Dim newPattern As String
    Dim targetValue As String
    Dim extensionsValue As String
    Dim sourceFolderValue As String
    Dim subjectValue As String
    Dim displayNameValue As String
    Dim duplicateIndex As Long
    Dim enabledValue As Boolean
    
    rowIndex = ActiveRuleIndex()
    
    If rowIndex < 0 Then
        If SelectedRuleCount() > 1 Then
            MsgBox RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1086, 1076, 1085, 1086, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 32, 1076, 1083, 1103, 32, 1086, 1073, 1085, 1086, 1074, 1083, 1077, 1085, 1080, 1103, 46), vbInformation
        Else
            MsgBox RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 32, 1076, 1083, 1103, 32, 1086, 1073, 1085, 1086, 1074, 1083, 1077, 1085, 1080, 1103, 46), vbInformation
        End If
        Exit Sub
    End If
    
    newPattern = ReadPatternValue()
    targetValue = ReadTargetValue()
    extensionsValue = ReadExtensionsValue()
    sourceFolderValue = ReadSourceFolderValue()
    subjectValue = ReadSubjectValue()
    displayNameValue = ReadRuleNameValue()
    
    If Not IsRuleDefinitionValid(newPattern, targetValue, extensionsValue, sourceFolderValue) Then
        MsgBox RuleValidationMessage(), vbInformation
        Exit Sub
    End If
    
    duplicateIndex = FindRuleIndex(newPattern, extensionsValue, sourceFolderValue, subjectValue)
    
    If duplicateIndex >= 0 And duplicateIndex <> rowIndex Then
        MsgBox BuildDuplicateRuleMessage(duplicateIndex), vbInformation
        Exit Sub
    End If
    
    enabledValue = IsRuleRowEnabled(rowIndex)
    lstRules.List(rowIndex, 2) = newPattern
    lstRules.List(rowIndex, 3) = targetValue
    lstRules.List(rowIndex, 4) = extensionsValue
    lstRules.List(rowIndex, 5) = sourceFolderValue
    lstRules.List(rowIndex, 6) = subjectValue
    lstRules.List(rowIndex, 8) = displayNameValue
    lstRules.List(rowIndex, 7) = RuleEnabledValue(enabledValue)
    SelectSingleRule rowIndex
    RenumberRuleDisplayTexts
    RefreshToggleActiveButton
End Sub

Private Sub btnToggleActive_Click()
    ToggleSelectedRuleActive
End Sub

Private Sub ToggleSelectedRuleActive()
    Dim i As Long
    Dim selectedCount As Long
    Dim targetEnabled As Boolean
    
    selectedCount = SelectedRuleCount()
    
    If selectedCount = 0 Then
        MsgBox RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 32, 1076, 1083, 1103, 32, 1080, 1079, 1084, 1077, 1085, 1077, 1085, 1080, 1103, 32, 1072, 1082, 1090, 1080, 1074, 1085, 1086, 1089, 1090, 1080, 46), vbInformation
        Exit Sub
    End If
    
    targetEnabled = HasSelectedInactiveRules()
    
    For i = 0 To lstRules.ListCount - 1
        If lstRules.Selected(i) Then
            lstRules.List(i, 7) = RuleEnabledValue(targetEnabled)
        End If
    Next i
    
    RenumberRuleDisplayTexts
    RefreshToggleActiveButton
End Sub

Private Sub btnRemove_Click()
    Dim i As Long
    Dim selectedCount As Long
    Dim confirmResult As VbMsgBoxResult
    
    selectedCount = SelectedRuleCount()
    
    If selectedCount = 0 Then
        MsgBox RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 32, 1076, 1083, 1103, 32, 1091, 1076, 1072, 1083, 1077, 1085, 1080, 1103, 46), vbInformation
        Exit Sub
    End If
    
    If selectedCount = 1 Then
        confirmResult = MsgBox(RuText(1042, 1099, 32, 1091, 1074, 1077, 1088, 1077, 1085, 1099, 44, 32, 1095, 1090, 1086, 32, 1093, 1086, 1090, 1080, 1090, 1077, 32, 1091, 1076, 1072, 1083, 1080, 1090, 1100, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1086, 32, 8470) & CStr(SelectedRuleIndex() + 1) & "?", vbQuestion + vbYesNo, RuText(1059, 1076, 1072, 1083, 1077, 1085, 1080, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1072))
    Else
        confirmResult = MsgBox(RuText(1042, 1099, 32, 1091, 1074, 1077, 1088, 1077, 1085, 1099, 44, 32, 1095, 1090, 1086, 32, 1093, 1086, 1090, 1080, 1090, 1077, 32, 1091, 1076, 1072, 1083, 1080, 1090, 1100, 32) & CStr(selectedCount) & RuText(32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1093, 32, 1087, 1088, 1072, 1074, 1080, 1083, 63), vbQuestion + vbYesNo, RuText(1059, 1076, 1072, 1083, 1077, 1085, 1080, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1072))
    End If
    
    If confirmResult <> vbYes Then
        Exit Sub
    End If
    
    For i = lstRules.ListCount - 1 To 0 Step -1
        If lstRules.Selected(i) Then
            lstRules.RemoveItem i
        End If
    Next i
    
    RenumberRuleDisplayTexts
    RefreshToggleActiveButton
    ClearRuleFields
End Sub

Private Sub btnBrowse_Click()
    Dim folderPath As String
    
    folderPath = SR_BrowseForFolder(RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1094, 1077, 1083, 1077, 1074, 1091, 1102, 32, 1087, 1072, 1087, 1082, 1091, 32, 1076, 1083, 1103, 32, 1101, 1090, 1086, 1075, 1086, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1072, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080), txtTarget.Text)
    
    If Len(folderPath) > 0 Then
        txtTarget.Text = folderPath
    End If
End Sub

Private Sub btnSourceBrowse_Click()
    Dim folderPath As String
    
    folderPath = FolderSelectionForm.PickSingleFolderPath()
    
    If Len(folderPath) > 0 Then
        txtSourceFolder.Text = folderPath
    End If
End Sub

Private Sub btnSourceCurrent_Click()
    On Error GoTo ErrorHandler
    
    Dim currentFolder As Outlook.MAPIFolder
    
    Set currentFolder = Outlook.Application.ActiveExplorer.CurrentFolder
    
    If currentFolder Is Nothing Then
        MsgBox RuText(1058, 1077, 1082, 1091, 1097, 1072, 1103, 32, 1087, 1072, 1087, 1082, 1072, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1085, 1077, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1072, 46), vbInformation
        Exit Sub
    End If
    
    txtSourceFolder.Text = currentFolder.FolderPath
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1074, 1079, 1103, 1090, 1100, 32, 1090, 1077, 1082, 1091, 1097, 1091, 1102, 32, 1087, 1072, 1087, 1082, 1091, 32, 79, 117, 116, 108, 111, 111, 107, 58, 32) & Err.Description, vbCritical
End Sub

Private Sub btnSave_Click()
    Dim i As Long
    Dim rows As String
    
    If HasRuleFieldValues() Then
        If Not UpsertRuleFromFields(True) Then
            Exit Sub
        End If
    End If
    
    For i = 0 To lstRules.ListCount - 1
        If IsRuleRowEnabled(i) Then
            rows = rows & "1|"
        Else
            rows = rows & "0|"
        End If
        
        rows = rows & CStr(lstRules.List(i, 2)) & "|" & CStr(lstRules.List(i, 3)) & "|" & CStr(lstRules.List(i, 4)) & "|" & CStr(lstRules.List(i, 5)) & "|" & CStr(lstRules.List(i, 6)) & "|" & CStr(lstRules.List(i, 8)) & vbLf
    Next i
    
    SR_SaveSortRuleRows rows
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

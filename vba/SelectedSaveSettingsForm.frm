VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectedSaveSettingsForm 
   Caption         =   "Selected save settings"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7600
   OleObjectBlob   =   "SelectedSaveSettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectedSaveSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private txtFolder As MSForms.TextBox
Private txtAllowedExtensions As MSForms.TextBox
Private txtExcludedExtensions As MSForms.TextBox

Private WithEvents btnBrowse As MSForms.CommandButton
Attribute btnBrowse.VB_VarHelpID = -1
Private WithEvents btnSave As MSForms.CommandButton
Attribute btnSave.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    BuildUi
    LoadSettings
End Sub

Private Sub BuildUi()
    Dim lblTitle As MSForms.Label
    Dim lblFolder As MSForms.Label
    Dim lblAllowed As MSForms.Label
    Dim lblAllowedHint As MSForms.Label
    Dim lblExcluded As MSForms.Label
    Dim lblExcludedHint As MSForms.Label
    
    Me.Caption = RuText(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1080, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1093, 32, 1087, 1080, 1089, 1077, 1084)
    Me.Width = 560
    Me.Height = 330
    
    Set lblTitle = AddLabel("lblTitle", RuText(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1080, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1093, 32, 1087, 1080, 1089, 1077, 1084), 18, 12, 430, 18)
    lblTitle.Font.Bold = True
    
    Set lblFolder = AddLabel("lblFolder", RuText(1055, 1072, 1087, 1082, 1072, 32, 1087, 1086, 32, 1091, 1084, 1086, 1083, 1095, 1072, 1085, 1080, 1102), 18, 46, 180, 18)
    Set txtFolder = AddTextBox("txtFolder", 18, 68, 410, 22)
    Set btnBrowse = AddButton("btnBrowse", RuText(1054, 1073, 1079, 1086, 1088), 438, 66, 82, 24)
    
    Set lblAllowed = AddLabel("lblAllowed", RuText(1056, 1072, 1079, 1088, 1077, 1096, 1077, 1085, 1085, 1099, 1077, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1103), 18, 108, 220, 18)
    Set txtAllowedExtensions = AddTextBox("txtAllowedExtensions", 18, 130, 502, 22)
    Set lblAllowedHint = AddLabel("lblAllowedHint", RuText(1053, 1072, 1087, 1088, 1080, 1084, 1077, 1088, 58, 32, 112, 100, 102, 44, 32, 122, 105, 112, 44, 32, 120, 108, 115, 120, 46, 32, 1055, 1091, 1089, 1090, 1086, 32, 45, 32, 1073, 1077, 1079, 32, 1086, 1075, 1088, 1072, 1085, 1080, 1095, 1077, 1085, 1080, 1103, 46), 18, 154, 502, 18)
    
    Set lblExcluded = AddLabel("lblExcluded", RuText(1048, 1089, 1082, 1083, 1102, 1095, 1077, 1085, 1085, 1099, 1077, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1103), 18, 188, 220, 18)
    Set txtExcludedExtensions = AddTextBox("txtExcludedExtensions", 18, 210, 502, 22)
    Set lblExcludedHint = AddLabel("lblExcludedHint", RuText(1048, 1089, 1082, 1083, 1102, 1095, 1077, 1085, 1080, 1103, 32, 1087, 1088, 1080, 1084, 1077, 1085, 1103, 1102, 1090, 1089, 1103, 32, 1087, 1086, 1089, 1083, 1077, 32, 1088, 1072, 1079, 1088, 1077, 1096, 1077, 1085, 1085, 1099, 1093, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1081, 46, 32, 1048, 1089, 1082, 1083, 1102, 1095, 1077, 1085, 1080, 1103, 32, 1080, 1084, 1077, 1102, 1090, 32, 1087, 1088, 1080, 1086, 1088, 1080, 1090, 1077, 1090, 32, 1085, 1072, 1076, 32, 1088, 1072, 1079, 1088, 1077, 1096, 1077, 1085, 1080, 1103, 1084, 1080, 46), 18, 234, 502, 18)
    
    Set btnSave = AddButton("btnSave", RuText(1057, 1086, 1093, 1088, 1072, 1085, 1080, 1090, 1100), 18, 270, 100, 26)
    Set btnCancel = AddButton("btnCancel", RuText(1054, 1090, 1084, 1077, 1085, 1072), 420, 270, 100, 26)
End Sub

Private Sub LoadSettings()
    txtFolder.Text = GetManualSaveFolder()
    txtAllowedExtensions.Text = FormatAllowedExtensionsForInput(GetSelectedSaveAllowedExtensions())
    txtExcludedExtensions.Text = FormatAllowedExtensionsForInput(GetSelectedSaveExcludedExtensions())
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

Private Function AddTextBox(ByVal controlName As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single, ByVal heightValue As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    
    With AddTextBox
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
    End With
End Function

Private Function AddButton(ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single, ByVal heightValue As Single) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    
    With AddButton
        .Caption = captionText
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
    End With
End Function

Private Sub btnBrowse_Click()
    Dim folderPath As String
    
    folderPath = BrowseForFolderEx(RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1087, 1072, 1087, 1082, 1091, 32, 1076, 1083, 1103, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 32, 1080, 1079, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1093, 32, 1087, 1080, 1089, 1077, 1084), txtFolder.Text)
    
    If Len(folderPath) > 0 Then
        txtFolder.Text = EnsureTrailingBackslash(folderPath)
    End If
End Sub

Private Sub btnSave_Click()
    Dim folderPath As String
    
    folderPath = Trim(txtFolder.Text)
    
    If Len(folderPath) = 0 Or Not IsAbsoluteFolderPath(folderPath) Then
        MsgBox RuText(1055, 1072, 1087, 1082, 1072, 32, 1076, 1086, 1083, 1078, 1085, 1072, 32, 1073, 1099, 1090, 1100, 32, 1072, 1073, 1089, 1086, 1083, 1102, 1090, 1085, 1099, 1084, 32, 1087, 1091, 1090, 1077, 1084, 46), vbInformation
        Exit Sub
    End If
    
    SaveSelectedSaveSettings folderPath, txtAllowedExtensions.Text, txtExcludedExtensions.Text
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoSaveSettingsForm 
   Caption         =   "Outlook Autosave"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   OleObjectBlob   =   "AutoSaveSettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AutoSaveSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents btnRefresh As MSForms.CommandButton
Attribute btnRefresh.VB_VarHelpID = -1
Private WithEvents btnSaveSelected As MSForms.CommandButton
Attribute btnSaveSelected.VB_VarHelpID = -1
Private WithEvents btnSelectedSaveSettings As MSForms.CommandButton
Attribute btnSelectedSaveSettings.VB_VarHelpID = -1
Private WithEvents btnToggleAutoSave As MSForms.CommandButton
Attribute btnToggleAutoSave.VB_VarHelpID = -1
Private WithEvents btnRunNow As MSForms.CommandButton
Attribute btnRunNow.VB_VarHelpID = -1
Private WithEvents btnInterval As MSForms.CommandButton
Attribute btnInterval.VB_VarHelpID = -1
Private WithEvents btnLookback As MSForms.CommandButton
Attribute btnLookback.VB_VarHelpID = -1
Private WithEvents btnToggleLogging As MSForms.CommandButton
Attribute btnToggleLogging.VB_VarHelpID = -1
Private WithEvents btnSortRules As MSForms.CommandButton
Attribute btnSortRules.VB_VarHelpID = -1

Private txtStatus As MSForms.TextBox

Private Sub UserForm_Initialize()
    BuildUi
    EnsureAutoSaveIsRunning
    RefreshStatus
End Sub

Private Sub BuildUi()
    Dim lblTitle As MSForms.Label
    Dim lblStatus As MSForms.Label
    Dim lblActions As MSForms.Label
    Dim lblSettings As MSForms.Label
    
    Me.Caption = RuText(1040, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 32, 79, 117, 116, 108, 111, 111, 107)
    Me.Width = 650
    Me.Height = 420
    
    Set lblTitle = AddLabel("lblTitle", RuText(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1080, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081), 18, 12, 470, 18)
    lblTitle.Font.Bold = True
    
    Set lblStatus = AddLabel("lblStatus", RuText(1057, 1090, 1072, 1090, 1091, 1089), 18, 42, 120, 18)
    lblStatus.Font.Bold = True
    
    Set txtStatus = Me.Controls.Add("Forms.TextBox.1", "txtStatus", True)
    With txtStatus
        .Left = 18
        .Top = 64
        .Width = 584
        .Height = 150
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
    End With
    
    Set lblActions = AddLabel("lblActions", RuText(1044, 1077, 1081, 1089, 1090, 1074, 1080, 1103), 18, 230, 120, 18)
    lblActions.Font.Bold = True
    
    Set btnRefresh = AddButton("btnRefresh", RuText(1054, 1073, 1085, 1086, 1074, 1080, 1090, 1100, 32, 1089, 1090, 1072, 1090, 1091, 1089), 18, 254, 126, 24, RuText(1054, 1073, 1085, 1086, 1074, 1080, 1090, 1100, 32, 1090, 1077, 1082, 1091, 1097, 1080, 1081, 32, 1089, 1090, 1072, 1090, 1091, 1089, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 46))
    Set btnSaveSelected = AddButton("btnSaveSelected", RuText(1057, 1086, 1093, 1088, 1072, 1085, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1077), 156, 254, 126, 24, RuText(1057, 1086, 1093, 1088, 1072, 1085, 1080, 1090, 1100, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1103, 32, 1080, 1079, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1093, 32, 1087, 1080, 1089, 1077, 1084, 32, 79, 117, 116, 108, 111, 111, 107, 46))
    Set btnSelectedSaveSettings = AddButton("btnSelectedSaveSettings", RuText(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1080, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1093), 294, 254, 150, 24, RuText(1054, 1090, 1082, 1088, 1099, 1090, 1100, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1080, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 32, 1080, 1079, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1093, 32, 1087, 1080, 1089, 1077, 1084, 46))
    
    Set btnToggleAutoSave = AddButton("btnToggleAutoSave", RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100), 18, 284, 116, 24, RuText(1042, 1082, 1083, 1102, 1072, 1077, 1090, 32, 1080, 1083, 1080, 32, 1086, 1090, 1082, 1083, 1102, 1095, 1072, 1077, 1090, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 32, 1074, 32, 1079, 1072, 1074, 1080, 1089, 1080, 1084, 1086, 1089, 1090, 1080, 32, 1086, 1090, 32, 1090, 1077, 1082, 1091, 1097, 1077, 1075, 1086, 32, 1089, 1090, 1072, 1090, 1091, 1089, 1072, 46))
    Set btnRunNow = AddButton("btnRunNow", RuText(1047, 1072, 1087, 1091, 1089, 1090, 1080, 1090, 1100), 156, 284, 126, 24, RuText(1047, 1072, 1087, 1091, 1089, 1090, 1080, 1090, 1100, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 32, 1086, 1076, 1080, 1085, 32, 1088, 1072, 1079, 32, 1089, 1077, 1081, 1095, 1072, 1089, 46, 32, 1053, 1077, 32, 1074, 1082, 1083, 1102, 1095, 1072, 1077, 1090, 32, 1087, 1086, 1089, 1090, 1086, 1103, 1085, 1085, 1099, 1081, 32, 1072, 1074, 1090, 1086, 1084, 1072, 1090, 1080, 1095, 1077, 1089, 1082, 1080, 1081, 32, 1088, 1077, 1078, 1080, 1084, 46))
    
    Set lblSettings = AddLabel("lblSettings", RuText(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1080), 18, 326, 120, 18)
    lblSettings.Font.Bold = True
    
    Set btnSortRules = AddButton("btnSortRules", RuText(1055, 1088, 1072, 1074, 1080, 1083, 1072), 18, 348, 116, 24, RuText(1053, 1072, 1089, 1090, 1088, 1086, 1080, 1090, 1100, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1072, 32, 1074, 1099, 1073, 1086, 1088, 1072, 32, 1087, 1072, 1087, 1086, 1082, 32, 1087, 1086, 32, 1080, 1084, 1077, 1085, 1080, 32, 1092, 1072, 1081, 1083, 1072, 32, 1080, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1102, 46))
    Set btnInterval = AddButton("btnInterval", RuText(1048, 1085, 1090, 1077, 1088, 1074, 1072, 1083), 146, 348, 116, 24, RuText(1047, 1072, 1076, 1072, 1090, 1100, 32, 1095, 1072, 1089, 1090, 1086, 1090, 1091, 32, 1087, 1088, 1086, 1074, 1077, 1088, 1082, 1080, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1077, 1085, 1085, 1099, 1093, 32, 1087, 1072, 1087, 1086, 1082, 46))
    Set btnLookback = AddButton("btnLookback", RuText(1055, 1077, 1088, 1080, 1086, 1076), 274, 348, 116, 24, RuText(1047, 1072, 1076, 1072, 1090, 1100, 32, 1087, 1077, 1088, 1080, 1086, 1076, 44, 32, 1079, 1072, 32, 1082, 1086, 1090, 1086, 1088, 1099, 1081, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1073, 1091, 1076, 1077, 1090, 32, 1080, 1089, 1082, 1072, 1090, 1100, 32, 1087, 1080, 1089, 1100, 1084, 1072, 32, 1087, 1086, 1089, 1083, 1077, 32, 1087, 1088, 1086, 1089, 1090, 1086, 1103, 46))
    Set btnToggleLogging = AddButton("btnToggleLogging", RuText(1051, 1086, 1075, 1080, 1088, 1086, 1074, 1072, 1085, 1080, 1077, 58, 32, 1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100), 402, 348, 176, 24, RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100, 32, 1079, 1072, 1087, 1080, 1089, 1100, 32, 1083, 1086, 1075, 1072, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 46))
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

Private Sub RefreshStatus()
    txtStatus.Text = GetAutoSaveStatusText()
    RefreshToggleAutoSaveButton
    RefreshToggleLoggingButton
End Sub

Private Sub RefreshToggleAutoSaveButton()
    If IsAutoSaveEnabled() Then
        btnToggleAutoSave.Caption = RuText(1054, 1090, 1082, 1083, 1102, 1095, 1080, 1090, 1100)
        btnToggleAutoSave.ControlTipText = RuText(1054, 1090, 1082, 1083, 1102, 1095, 1080, 1090, 1100, 32, 1072, 1074, 1090, 1086, 1084, 1072, 1090, 1080, 1095, 1077, 1089, 1082, 1086, 1077, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 46)
    Else
        btnToggleAutoSave.Caption = RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100)
        btnToggleAutoSave.ControlTipText = RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 46, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1089, 1088, 1072, 1079, 1091, 32, 1087, 1088, 1086, 1074, 1077, 1088, 1080, 1090, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1077, 1085, 1085, 1099, 1077, 32, 1087, 1072, 1087, 1082, 1080, 32, 1087, 1086, 32, 1079, 1072, 1076, 1072, 1085, 1085, 1086, 1084, 1091, 32, 1080, 1085, 1090, 1077, 1088, 1074, 1072, 1083, 1091, 46)
    End If
End Sub

Private Sub RefreshToggleLoggingButton()
    If IsAutoSaveLoggingEnabled() Then
        btnToggleLogging.Caption = RuText(1051, 1086, 1075, 1080, 1088, 1086, 1074, 1072, 1085, 1080, 1077, 58, 32, 1054, 1090, 1082, 1083, 1102, 1095, 1080, 1090, 1100)
        btnToggleLogging.ControlTipText = RuText(1054, 1090, 1082, 1083, 1102, 1095, 1080, 1090, 1100, 32, 1079, 1072, 1087, 1080, 1089, 1100, 32, 1083, 1086, 1075, 1072, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 46)
    Else
        btnToggleLogging.Caption = RuText(1051, 1086, 1075, 1080, 1088, 1086, 1074, 1072, 1085, 1080, 1077, 58, 32, 1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100)
        btnToggleLogging.ControlTipText = RuText(1042, 1082, 1083, 1102, 1095, 1080, 1090, 1100, 32, 1079, 1072, 1087, 1080, 1089, 1100, 32, 1083, 1086, 1075, 1072, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 46)
    End If
End Sub

Private Sub btnRefresh_Click()
    RefreshStatus
End Sub

Private Sub btnSaveSelected_Click()
    SaveSelectedAttachments
    RefreshStatus
End Sub

Private Sub btnSelectedSaveSettings_Click()
    ShowSelectedSaveSettingsForm
    RefreshStatus
End Sub

Private Sub btnToggleAutoSave_Click()
    If IsAutoSaveEnabled() Then
        DisableAutoSave
    Else
        EnableAutoSave
    End If
    
    RefreshStatus
End Sub

Private Sub btnRunNow_Click()
    RunAutoSaveNow
    RefreshStatus
End Sub

Private Sub btnInterval_Click()
    SetAutoSaveInterval
    RefreshStatus
End Sub

Private Sub btnLookback_Click()
    SetAutoSaveLookback
    RefreshStatus
End Sub

Private Sub btnToggleLogging_Click()
    ToggleAutoSaveLogging
    RefreshStatus
End Sub

Private Sub btnSortRules_Click()
    On Error GoTo ErrorHandler
    
    SortRulesForm.Show
    RefreshStatus
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1086, 1090, 1082, 1088, 1099, 1090, 1100, 32, 1092, 1086, 1088, 1084, 1091, 32, 1087, 1088, 1072, 1074, 1080, 1083, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080, 58, 32) & Err.Description, vbCritical
    RefreshStatus
End Sub

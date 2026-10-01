Attribute VB_Name = "SortRulesSettingsUi"
Option Private Module

Sub OpenSortRulesSettings()
    On Error GoTo ErrorHandler
    SortRulesForm.Show
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1086, 1090, 1082, 1088, 1099, 1090, 1100, 32, 1092, 1086, 1088, 1084, 1091, 32, 1087, 1088, 1072, 1074, 1080, 1083, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080, 58, 32) & Err.Description, vbCritical
End Sub

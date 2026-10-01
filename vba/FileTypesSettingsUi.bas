Attribute VB_Name = "FileTypesSettingsUi"
Option Private Module

Sub OpenFileTypesSettings()
    On Error GoTo ErrorHandler
    FileTypesForm.Show
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1086, 1090, 1082, 1088, 1099, 1090, 1100, 32, 1092, 1086, 1088, 1084, 1091, 32, 1090, 1080, 1087, 1086, 1074, 32, 1092, 1072, 1081, 1083, 1086, 1074, 58, 32) & Err.Description, vbCritical
End Sub

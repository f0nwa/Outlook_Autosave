Attribute VB_Name = "AutoSaveSettingsUi"

Sub OutlookAutosave()
    On Error GoTo ErrorHandler
    EnsureAutoSaveIsRunning
    AutoSaveSettingsForm.Show
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1086, 1090, 1082, 1088, 1099, 1090, 1100, 32, 1092, 1086, 1088, 1084, 1091, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1077, 1082, 58, 32) & Err.Description, vbCritical
End Sub

Sub SaveSelectedAttachments()
    On Error GoTo ErrorHandler
    SaveAttachmentsFromSelectedEmails
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1074, 1099, 1087, 1086, 1083, 1085, 1080, 1090, 1100, 32, 1084, 1072, 1082, 1088, 1086, 1089, 32, 1087, 1077, 1088, 1077, 1076, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 1084, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1093, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 58, 32) & Err.Description, vbCritical
End Sub

Private Sub OpenSettings()
    OutlookAutosave
End Sub

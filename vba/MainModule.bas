Attribute VB_Name = "MainModule"
Option Private Module
#If VBA7 Then
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private autoSaveTimerId As LongPtr
#Else
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private autoSaveTimerId As Long
#End If

Private Const SETTINGS_APP_NAME As String = "OutlookAttachmentSaver"
Private Const SETTINGS_SECTION As String = "Settings"
Private Const SETTINGS_SAVE_FOLDER_KEY As String = "SaveFolder"
Private Const SETTINGS_MANUAL_SAVE_FOLDER_KEY As String = "ManualSaveFolder"
Private Const SETTINGS_SELECTED_ALLOWED_EXTENSIONS_KEY As String = "SelectedSaveAllowedExtensions"
Private Const SETTINGS_SELECTED_EXCLUDED_EXTENSIONS_KEY As String = "SelectedSaveExcludedExtensions"
Private Const SETTINGS_AUTO_ENABLED_KEY As String = "AutoSaveEnabled"
Private Const SETTINGS_AUTO_LOGGING_ENABLED_KEY As String = "AutoSaveLoggingEnabled"
Private Const SETTINGS_LAST_RUN_KEY As String = "LastAutoSaveRun"
Private Const SETTINGS_NEXT_RUN_KEY As String = "NextAutoSaveRun"
Private Const SETTINGS_AUTO_FOLDERS_KEY As String = "AutoSaveFolders"
Private Const SETTINGS_ALLOWED_EXTENSIONS_KEY As String = "AllowedExtensions"
Private Const SETTINGS_FILE_TYPE_RULES_KEY As String = "FileTypeRules"
Private Const SETTINGS_INTERVAL_SECONDS_KEY As String = "AutoSaveIntervalSeconds"
Private Const SETTINGS_LOOKBACK_SECONDS_KEY As String = "AutoSaveLookbackSeconds"
Private Const SETTINGS_SORT_RULES_KEY As String = "AttachmentSortRules"
Private Const SETTINGS_SORT_RULE_ROWS_KEY As String = "AttachmentSortRuleRows"
Private Const MACRO_VERSION As String = "1.0.2"
Private Const DEFAULT_AUTO_SAVE_INTERVAL_SECONDS As Long = 43200
Private Const DEFAULT_AUTO_SAVE_LOOKBACK_SECONDS As Long = 86400
Private Const AUTO_SAVE_MAX_ITEMS_PER_RUN As Long = 10000
Private Const MAX_FOLDER_CHOICES As Long = 80
Private Const LEGACY_AUTO_SAVE_REMINDER_SUBJECT As String = "Outlook Attachment Saver Auto Run"
Private Const AUTO_SAVE_INDEX_FILE_NAME As String = "autosave.index"

Private lastAutoSaveRunFailureMessage As String

Public Function RuText(ParamArray codePoints() As Variant) As String
    Dim i As Long
    Dim result As String
    
    For i = LBound(codePoints) To UBound(codePoints)
        result = result & ChrW$(CLng(codePoints(i)))
    Next i
    
    RuText = result
End Function

Public Function RuBool(ByVal value As Boolean) As String
    If value Then
        RuBool = RuText(1044, 1072)
    Else
        RuBool = RuText(1053, 1077, 1090)
    End If
End Function

Public Function RuEnabledDisabled(ByVal value As Boolean) As String
    If value Then
        RuEnabledDisabled = RuText(1042, 1082, 1083, 1102, 1095, 1077, 1085, 1086)
    Else
        RuEnabledDisabled = RuText(1054, 1090, 1082, 1083, 1102, 1095, 1077, 1085, 1086)
    End If
End Function

Public Function RuActiveDisabled(ByVal value As Boolean) As String
    If value Then
        RuActiveDisabled = RuText(1040, 1082, 1090, 1080, 1074, 1077, 1085)
    Else
        RuActiveDisabled = RuText(1054, 1090, 1082, 1083, 1102, 1095, 1077, 1085)
    End If
End Function

Private Function GetMacroAuthor() As String
    GetMacroAuthor = RuText(1055, 1086, 1075, 1091, 1094, 1072, 32, 1042, 1083, 1072, 1076, 1080, 1089, 1083, 1072, 1074)
End Function

Sub SetSortRules()
    On Error GoTo ErrorHandler
    SetAttachmentSortRules
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1074, 1099, 1087, 1086, 1083, 1085, 1080, 1090, 1100, 32, 1084, 1072, 1082, 1088, 1086, 1089, 32, 1087, 1077, 1088, 1077, 1076, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1086, 1081, 32, 1087, 1088, 1072, 1074, 1080, 1083, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080, 58, 32) & Err.Description, vbCritical
End Sub

Sub SetAutoSaveInterval()
    On Error GoTo ErrorHandler
    SetAutoSaveIntervalSetting
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1074, 1099, 1087, 1086, 1083, 1085, 1080, 1090, 1100, 32, 1084, 1072, 1082, 1088, 1086, 1089, 32, 1087, 1077, 1088, 1077, 1076, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1086, 1081, 32, 1080, 1085, 1090, 1077, 1088, 1074, 1072, 1083, 1072, 58, 32) & Err.Description, vbCritical
End Sub

Sub SetAutoSaveLookback()
    On Error GoTo ErrorHandler
    SetAutoSaveLookbackSetting
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1074, 1099, 1087, 1086, 1083, 1085, 1080, 1090, 1100, 32, 1084, 1072, 1082, 1088, 1086, 1089, 32, 1087, 1077, 1088, 1077, 1076, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1086, 1081, 32, 1087, 1077, 1088, 1080, 1086, 1076, 1072, 58, 32) & Err.Description, vbCritical
End Sub

Sub EnableAutoSave()
    On Error GoTo ErrorHandler
    EnableAutoSaveAttachments
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1074, 1099, 1087, 1086, 1083, 1085, 1080, 1090, 1100, 32, 1084, 1072, 1082, 1088, 1086, 1089, 32, 1087, 1077, 1088, 1077, 1076, 32, 1074, 1082, 1083, 1102, 1095, 1077, 1085, 1080, 1077, 1084, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 58, 32) & Err.Description, vbCritical
End Sub

Sub DisableAutoSave()
    On Error GoTo ErrorHandler
    DisableAutoSaveAttachments
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1074, 1099, 1087, 1086, 1083, 1085, 1080, 1090, 1100, 32, 1084, 1072, 1082, 1088, 1086, 1089, 32, 1087, 1077, 1088, 1077, 1076, 32, 1086, 1090, 1082, 1083, 1102, 1095, 1077, 1085, 1080, 1077, 1084, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 58, 32) & Err.Description, vbCritical
End Sub

Sub ToggleAutoSaveLogging()
    On Error GoTo ErrorHandler
    
    If IsAutoSaveLoggingEnabled() Then
        SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_LOGGING_ENABLED_KEY, "False"
    Else
        SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_LOGGING_ENABLED_KEY, "True"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1080, 1079, 1084, 1077, 1085, 1080, 1090, 1100, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1091, 32, 1083, 1086, 1075, 1080, 1088, 1086, 1074, 1072, 1085, 1080, 1103, 58, 32) & Err.Description, vbCritical
End Sub

Sub RunAutoSaveNow()
    On Error GoTo ErrorHandler
    
    ResetLastAutoSaveRunToLookbackStart
    
    If SaveAttachmentsFromConfiguredFoldersSinceLastRun() Then
        If IsAutoSaveEnabled() Then
            ScheduleNextAutoSave
        End If
    ElseIf Len(lastAutoSaveRunFailureMessage) > 0 Then
        MsgBox lastAutoSaveRunFailureMessage, vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Manual auto save run failed: " & Err.Description
End Sub

Sub ShowAutoSaveSettingsForm()
    On Error GoTo ErrorHandler
    EnsureAutoSaveIsRunning
    AutoSaveSettingsForm.Show
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1086, 1090, 1082, 1088, 1099, 1090, 1100, 32, 1092, 1086, 1088, 1084, 1091, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1077, 1082, 58, 32) & Err.Description, vbCritical
End Sub

Sub ShowAutoSaveStatus()
    Dim statusText As String
    
    statusText = GetAutoSaveStatusText()
    
    MsgBox statusText, vbInformation, RuText(1057, 1090, 1072, 1090, 1091, 1089, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103)
End Sub

Sub AutoSaveStartup()
    On Error GoTo ErrorHandler
    
    If IsAutoSaveEnabled() Then
        ScheduleNextAutoSave
        Debug.Print "Auto save startup: next run scheduled."
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Auto save startup failed: " & Err.Description
End Sub

Sub EnsureAutoSaveIsRunning()
    On Error GoTo ErrorHandler
    
    If IsAutoSaveEnabled() And Not IsAutoSaveTimerActive() Then
        Debug.Print "Auto save timer is inactive while enabled; restoring scheduled run."
        ScheduleNextAutoSave
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Auto save restore failed: " & Err.Description
End Sub

Function GetAutoSaveStatusText() As String
    Dim rowRules As String
    Dim activeRuleSummary As String
    
    rowRules = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_SORT_RULE_ROWS_KEY, "")
    activeRuleSummary = GetActiveSortRuleSummaryText()
    
    GetAutoSaveStatusText = RuText(1042, 1077, 1088, 1089, 1080, 1103) & ": " & MACRO_VERSION & vbCrLf & _
                            RuText(1040, 1074, 1090, 1086, 1088) & ": " & GetMacroAuthor() & vbCrLf & _
                            RuText(1040, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 58, 32) & RuEnabledDisabled(IsAutoSaveEnabled()) & vbCrLf & _
                            RuText(1048, 1085, 1090, 1077, 1088, 1074, 1072, 1083, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 58, 32) & FormatIntervalSeconds(GetAutoSaveIntervalSeconds()) & vbCrLf & _
                            RuText(1055, 1077, 1088, 1080, 1086, 1076, 32, 1086, 1073, 1088, 1072, 1073, 1086, 1090, 1082, 1080, 58, 32) & FormatIntervalSeconds(GetAutoSaveLookbackSeconds()) & vbCrLf & _
                            RuText(1051, 1086, 1075, 1080, 1088, 1086, 1074, 1072, 1085, 1080, 1077, 58, 32) & RuEnabledDisabled(IsAutoSaveLoggingEnabled()) & vbCrLf & _
                            RuText(1058, 1072, 1081, 1084, 1077, 1088, 58, 32) & RuActiveDisabled(IsAutoSaveTimerActive()) & vbCrLf & _
                            RuText(1055, 1086, 1089, 1083, 1077, 1076, 1085, 1080, 1081, 32, 1079, 1072, 1087, 1091, 1089, 1082, 58, 32) & GetLastAutoSaveRunText() & vbCrLf & _
                            RuText(1057, 1083, 1077, 1076, 1091, 1102, 1097, 1080, 1081, 32, 1079, 1072, 1087, 1091, 1089, 1082, 58, 32) & GetNextAutoSaveRunText()
    
    If Len(activeRuleSummary) > 0 Then
        GetAutoSaveStatusText = GetAutoSaveStatusText & vbCrLf & activeRuleSummary
    End If
    
    If IsAutoSaveEnabled() And CountSourceLimitedSortRuleRows(rowRules) = 0 Then
        GetAutoSaveStatusText = GetAutoSaveStatusText & vbCrLf & RuText(1040, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 58, 32, 1085, 1077, 1090, 32, 1072, 1082, 1090, 1080, 1074, 1085, 1099, 1093, 32, 1087, 1088, 1072, 1074, 1080, 1083, 32, 1089, 32, 1080, 1089, 1093, 1086, 1076, 1085, 1086, 1081, 32, 1087, 1072, 1087, 1082, 1086, 1081)
    End If
End Function

Sub SaveAttachmentsFromSelectedEmails()
    On Error GoTo ErrorHandler
    
    Dim objSelection As Outlook.Selection
    Dim objItem As Object
    Dim objMail As Outlook.mailItem
    Dim savedFilePaths As Collection
    Dim manualFolderPath As String
    Dim savedCount As Long
    Dim checkedCount As Long
    
    manualFolderPath = GetManualSaveFolder()
    
    If Len(manualFolderPath) = 0 Or Not IsAbsoluteFolderPath(manualFolderPath) Then
        MsgBox RuText(1055, 1072, 1087, 1082, 1072, 32, 1076, 1086, 1083, 1078, 1085, 1072, 32, 1073, 1099, 1090, 1100, 32, 1072, 1073, 1089, 1086, 1083, 1102, 1090, 1085, 1099, 1084, 32, 1087, 1091, 1090, 1077, 1084, 46), vbCritical
        ShowSelectedSaveSettingsForm
        Exit Sub
    End If
    
    Set savedFilePaths = New Collection
    
    ' Get the selected items
    Set objSelection = Outlook.ActiveExplorer.Selection
    
    ' Loop through each selected item
    For Each objItem In objSelection
        If TypeOf objItem Is mailItem Then
            Set objMail = objItem
            checkedCount = checkedCount + 1
            savedCount = savedCount + SaveAllAttachmentsFromMailItemToFolder(objMail, manualFolderPath, savedFilePaths)
        End If
    Next objItem
    
    If checkedCount = 0 Then
        MsgBox RuText(1053, 1077, 1090, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1093, 32, 1087, 1080, 1089, 1077, 1084, 32, 79, 117, 116, 108, 111, 111, 107, 46), vbInformation
        Exit Sub
    End If
    
    ' Display a message when done
    MsgBox BuildSavedAttachmentsMessage(savedCount, savedFilePaths), vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1055, 1088, 1086, 1080, 1079, 1086, 1096, 1083, 1072, 32, 1086, 1096, 1080, 1073, 1082, 1072, 58, 32) & Err.Description, vbCritical
End Sub

Sub ShowSelectedSaveSettingsForm()
    On Error GoTo ErrorHandler
    
    SelectedSaveSettingsForm.Show
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1086, 1090, 1082, 1088, 1099, 1090, 1100, 32, 1092, 1086, 1088, 1084, 1091, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1077, 1082, 58, 32) & Err.Description, vbCritical
End Sub

Function GetManualSaveFolder() As String
    Dim folderPath As String
    
    folderPath = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_MANUAL_SAVE_FOLDER_KEY, GetDefaultSaveFolder())
    folderPath = Trim(folderPath)
    
    If Len(folderPath) = 0 Then
        folderPath = GetDefaultSaveFolder()
    End If
    
    GetManualSaveFolder = EnsureTrailingBackslash(folderPath)
End Function

Function GetSelectedSaveAllowedExtensions() As String
    GetSelectedSaveAllowedExtensions = NormalizeAllowedExtensions(GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_SELECTED_ALLOWED_EXTENSIONS_KEY, ""))
End Function

Function GetSelectedSaveExcludedExtensions() As String
    GetSelectedSaveExcludedExtensions = NormalizeAllowedExtensions(GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_SELECTED_EXCLUDED_EXTENSIONS_KEY, ""))
End Function

Sub SaveSelectedSaveSettings(ByVal folderPath As String, ByVal allowedExtensions As String, ByVal excludedExtensions As String)
    folderPath = EnsureTrailingBackslash(Trim(folderPath))
    
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_MANUAL_SAVE_FOLDER_KEY, folderPath
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_SELECTED_ALLOWED_EXTENSIONS_KEY, NormalizeAllowedExtensions(allowedExtensions)
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_SELECTED_EXCLUDED_EXTENSIONS_KEY, NormalizeAllowedExtensions(excludedExtensions)
End Sub

Private Function BuildSavedAttachmentsMessage(ByVal savedCount As Long, ByVal savedFilePaths As Collection) As String
    Const MAX_SAVED_PATHS_IN_MESSAGE As Long = 20
    
    Dim message As String
    Dim i As Long
    Dim displayCount As Long
    
    message = RuText(1057, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1086, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 58, 32) & CStr(savedCount)
    
    If savedFilePaths Is Nothing Then
        BuildSavedAttachmentsMessage = message
        Exit Function
    End If
    
    If savedFilePaths.Count = 0 Then
        BuildSavedAttachmentsMessage = message
        Exit Function
    End If
    
    message = message & vbCrLf & vbCrLf & RuText(1060, 1072, 1081, 1083, 1099, 58)
    displayCount = savedFilePaths.Count
    
    If displayCount > MAX_SAVED_PATHS_IN_MESSAGE Then
        displayCount = MAX_SAVED_PATHS_IN_MESSAGE
    End If
    
    For i = 1 To displayCount
        message = message & vbCrLf & CStr(i) & ". " & CStr(savedFilePaths(i))
    Next i
    
    If savedFilePaths.Count > displayCount Then
        message = message & vbCrLf & RuText(46, 46, 46, 1080, 32, 1077, 1097, 1077, 32) & CStr(savedFilePaths.Count - displayCount) & RuText(32, 1092, 1072, 1081, 1083, 40, 1086, 1074, 41, 46)
    End If
    
    BuildSavedAttachmentsMessage = message
End Function

Private Sub SetAttachmentsSaveFolder()
    Dim strFolderPath As String
    
    strFolderPath = BrowseForFolderEx(RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1087, 1072, 1087, 1082, 1091, 32, 1076, 1083, 1103, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081), GetAttachmentsSaveFolder())
    
    If Len(strFolderPath) = 0 Then
        strFolderPath = InputBox(RuText(1042, 1099, 1073, 1086, 1088, 32, 1087, 1072, 1087, 1082, 1080, 32, 1086, 1090, 1084, 1077, 1085, 1077, 1085, 32, 1080, 1083, 1080, 32, 1085, 1077, 1076, 1086, 1089, 1090, 1091, 1087, 1077, 1085, 46) & vbCrLf & RuText(1042, 1074, 1077, 1076, 1080, 1090, 1077, 32, 1087, 1091, 1090, 1100, 32, 1082, 32, 1087, 1072, 1087, 1082, 1077, 32, 1076, 1083, 1103, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 46) & vbCrLf & RuText(1055, 1088, 1080, 1084, 1077, 1088, 58, 32) & GetDefaultSaveFolder(), RuText(1055, 1072, 1087, 1082, 1072, 32, 1076, 1083, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081), GetAttachmentsSaveFolder())
    End If
    
    If Len(strFolderPath) = 0 Then
        MsgBox RuText(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1072, 32, 1087, 1072, 1087, 1082, 1080, 32, 1086, 1090, 1084, 1077, 1085, 1077, 1085, 1072, 46), vbInformation
        Exit Sub
    End If
    
    strFolderPath = EnsureTrailingBackslash(strFolderPath)
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_SAVE_FOLDER_KEY, strFolderPath
    
    MsgBox RuText(1055, 1072, 1087, 1082, 1072, 32, 1076, 1083, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 32, 1091, 1089, 1090, 1072, 1085, 1086, 1074, 1083, 1077, 1085, 1072, 58) & vbCrLf & strFolderPath, vbInformation
End Sub

Private Sub SetAutoSaveMailFolders()
    Dim folders As Collection
    Dim prompt As String
    Dim inputValue As String
    Dim savedFolders As String
    
    Set folders = New Collection
    BuildMailFolderSelectionList folders, prompt
    
    If folders.Count = 0 Then
        MsgBox RuText(1055, 1072, 1087, 1082, 1080, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1085, 1077, 32, 1085, 1072, 1081, 1076, 1077, 1085, 1099, 46), vbInformation
        Exit Sub
    End If
    
    inputValue = InputBox(prompt, RuText(1042, 1099, 1073, 1086, 1088, 32, 1087, 1072, 1087, 1086, 1082, 32, 79, 117, 116, 108, 111, 111, 107), GetSelectedFolderNumbers(folders))
    
    If Len(inputValue) = 0 Then
        MsgBox RuText(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1072, 32, 1087, 1072, 1087, 1086, 1082, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1086, 1090, 1084, 1077, 1085, 1077, 1085, 1072, 46, 32, 1040, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 32, 1073, 1091, 1076, 1077, 1090, 32, 1080, 1089, 1087, 1086, 1083, 1100, 1079, 1086, 1074, 1072, 1090, 1100, 32, 1042, 1093, 1086, 1076, 1103, 1097, 1080, 1077, 44, 32, 1087, 1086, 1082, 1072, 32, 1087, 1072, 1087, 1082, 1080, 32, 1085, 1077, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1077, 1085, 1099, 46), vbInformation
        Exit Sub
    End If
    
    savedFolders = ApplyMailFolderSelection(folders, inputValue)
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_FOLDERS_KEY, savedFolders
    
    MsgBox RuText(1053, 1072, 1089, 1090, 1088, 1086, 1077, 1085, 1086, 32, 1087, 1072, 1087, 1086, 1082, 32, 79, 117, 116, 108, 111, 111, 107, 58, 32) & CountSelectedFolderRecords(savedFolders), vbInformation
End Sub

Private Sub BuildMailFolderSelectionList(ByRef folders As Collection, ByRef prompt As String)
    Dim ns As Outlook.NameSpace
    Dim store As Outlook.Store
    Dim selectedKeys As Object
    
    Set selectedKeys = CreateObject("Scripting.Dictionary")
    LoadSelectedFolderKeys selectedKeys
    
    prompt = RuText(1042, 1074, 1077, 1076, 1080, 1090, 1077, 32, 1085, 1086, 1084, 1077, 1088, 1072, 32, 1087, 1072, 1087, 1086, 1082, 32, 1095, 1077, 1088, 1077, 1079, 32, 1079, 1072, 1087, 1103, 1090, 1091, 1102, 46) & vbCrLf & _
             RuText(1055, 1088, 1080, 1084, 1077, 1088, 58, 32) & "1, 3, 7" & vbCrLf & _
             RuText(91, 120, 93, 32, 1086, 1079, 1085, 1072, 1095, 1072, 1077, 1090, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1091, 1102, 32, 1089, 1077, 1081, 1095, 1072, 1089, 46) & vbCrLf & vbCrLf
    
    Set ns = Outlook.Application.GetNamespace("MAPI")
    
    For Each store In ns.Stores
        AddFolderChoice store.GetRootFolder, 0, folders, prompt, selectedKeys
        
        If folders.Count >= MAX_FOLDER_CHOICES Then
            prompt = prompt & vbCrLf & RuText(1055, 1086, 1082, 1072, 1079, 1072, 1085, 1099, 32, 1090, 1086, 1083, 1100, 1082, 1086, 32, 1087, 1077, 1088, 1074, 1099, 1077, 32) & MAX_FOLDER_CHOICES & RuText(32, 1087, 1072, 1087, 1086, 1082, 46)
            Exit For
        End If
    Next store
End Sub

Private Sub AddFolderChoice(ByVal folder As Outlook.MAPIFolder, ByVal level As Long, ByRef folders As Collection, ByRef prompt As String, ByVal selectedKeys As Object)
    Dim childFolder As Outlook.MAPIFolder
    Dim record As String
    Dim key As String
    Dim marker As String
    
    If folders.Count >= MAX_FOLDER_CHOICES Then
        Exit Sub
    End If
    
    record = folder.StoreID & "|" & folder.EntryID & "|" & folder.FolderPath
    key = FolderRecordKey(record)
    folders.Add record
    
    If selectedKeys.Exists(key) Then
        marker = "x"
    Else
        marker = " "
    End If
    
    prompt = prompt & folders.Count & ". [" & marker & "] " & GetFolderDisplayName(folder, level) & vbCrLf
    Debug.Print folders.Count & ". " & folder.FolderPath
    
    For Each childFolder In folder.Folders
        AddFolderChoice childFolder, level + 1, folders, prompt, selectedKeys
        
        If folders.Count >= MAX_FOLDER_CHOICES Then
            Exit For
        End If
    Next childFolder
End Sub

Private Sub LoadSelectedFolderKeys(ByVal selectedKeys As Object)
    Dim records() As String
    Dim i As Long
    Dim savedFolders As String
    Dim key As String
    
    savedFolders = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_FOLDERS_KEY, "")
    
    If Len(Trim(savedFolders)) = 0 Then
        Exit Sub
    End If
    
    records = Split(savedFolders, vbLf)
    
    For i = LBound(records) To UBound(records)
        key = FolderRecordKey(records(i))
        
        If Len(key) > 0 Then
            selectedKeys(key) = True
        End If
    Next i
End Sub

Private Function GetFolderDisplayName(ByVal folder As Outlook.MAPIFolder, ByVal level As Long) As String
    GetFolderDisplayName = String(level * 2, " ") & folder.Name
End Function

Private Function GetSelectedFolderNumbers(ByVal folders As Collection) As String
    Dim selectedKeys As Object
    Dim i As Long
    Dim key As String
    Dim result As String
    
    Set selectedKeys = CreateObject("Scripting.Dictionary")
    LoadSelectedFolderKeys selectedKeys
    
    For i = 1 To folders.Count
        key = FolderRecordKey(CStr(folders.Item(i)))
        
        If selectedKeys.Exists(key) Then
            If Len(result) > 0 Then
                result = result & ", "
            End If
            
            result = result & CStr(i)
        End If
    Next i
    
    GetSelectedFolderNumbers = result
End Function

Private Function ApplyMailFolderSelection(ByVal folders As Collection, ByVal inputValue As String) As String
    Dim selectedIndexes As Object
    Dim parts() As String
    Dim part As Variant
    Dim indexValue As Long
    Dim result As String
    
    Set selectedIndexes = CreateObject("Scripting.Dictionary")
    inputValue = Replace(inputValue, ";", ",")
    inputValue = Replace(inputValue, " ", ",")
    parts = Split(inputValue, ",")
    
    For Each part In parts
        If IsNumeric(part) Then
            indexValue = CLng(part)
            
            If indexValue >= 1 And indexValue <= folders.Count Then
                selectedIndexes(CStr(indexValue)) = True
            End If
        End If
    Next part
    
    For Each part In selectedIndexes.Keys
        result = result & CStr(folders.Item(CLng(part))) & vbLf
    Next part
    
    ApplyMailFolderSelection = result
End Function

Function CountSelectedFolderRecords(ByVal savedFolders As String) As Long
    Dim records() As String
    Dim i As Long
    Dim result As Long
    
    If Len(Trim(savedFolders)) = 0 Then
        CountSelectedFolderRecords = 0
        Exit Function
    End If
    
    records = Split(savedFolders, vbLf)
    
    For i = LBound(records) To UBound(records)
        If Len(Trim(records(i))) > 0 Then
            result = result + 1
        End If
    Next i
    
    CountSelectedFolderRecords = result
End Function

Private Function FolderRecordKey(ByVal record As String) As String
    Dim parts() As String
    
    If Len(Trim(record)) = 0 Then
        FolderRecordKey = ""
        Exit Function
    End If
    
    parts = Split(record, "|")
    
    If UBound(parts) >= 1 Then
        FolderRecordKey = parts(0) & "|" & parts(1)
    Else
        FolderRecordKey = ""
    End If
End Function

Function GetSavedAutoSaveFolderRecords() As String
    GetSavedAutoSaveFolderRecords = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_FOLDERS_KEY, "")
End Function

Sub SaveAutoSaveFolderRecords(ByVal savedFolders As String)
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_FOLDERS_KEY, savedFolders
End Sub

Function GetAutoSaveIntervalSeconds() As Long
    Dim rawValue As String
    
    rawValue = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_INTERVAL_SECONDS_KEY, CStr(DEFAULT_AUTO_SAVE_INTERVAL_SECONDS))
    
    If IsNumeric(rawValue) Then
        GetAutoSaveIntervalSeconds = CLng(rawValue)
    Else
        GetAutoSaveIntervalSeconds = DEFAULT_AUTO_SAVE_INTERVAL_SECONDS
    End If
    
    If GetAutoSaveIntervalSeconds <= 0 Then
        GetAutoSaveIntervalSeconds = DEFAULT_AUTO_SAVE_INTERVAL_SECONDS
    End If
End Function

Function GetAutoSaveLookbackSeconds() As Long
    Dim rawValue As String
    
    rawValue = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_LOOKBACK_SECONDS_KEY, CStr(DEFAULT_AUTO_SAVE_LOOKBACK_SECONDS))
    
    If IsNumeric(rawValue) Then
        GetAutoSaveLookbackSeconds = CLng(rawValue)
    Else
        GetAutoSaveLookbackSeconds = DEFAULT_AUTO_SAVE_LOOKBACK_SECONDS
    End If
    
    If GetAutoSaveLookbackSeconds <= 0 Then
        GetAutoSaveLookbackSeconds = DEFAULT_AUTO_SAVE_LOOKBACK_SECONDS
    End If
End Function

Function ParseIntervalSeconds(ByVal inputValue As String) As Long
    Dim unitValue As String
    Dim numberValue As String
    Dim i As Long
    Dim ch As String
    Dim amount As Double
    
    inputValue = LCase(Trim(inputValue))
    inputValue = Replace(inputValue, " ", "")
    
    For i = 1 To Len(inputValue)
        ch = Mid(inputValue, i, 1)
        
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "," Then
            numberValue = numberValue & Replace(ch, ",", ".")
        Else
            unitValue = unitValue & ch
        End If
    Next i
    
    If Len(numberValue) = 0 Or Not IsNumeric(numberValue) Then
        ParseIntervalSeconds = 0
        Exit Function
    End If
    
    amount = CDbl(numberValue)
    
    Select Case unitValue
        Case "", "s", "sec", "secs", "second", "seconds"
            ParseIntervalSeconds = CLng(amount)
        Case "m", "min", "mins", "minute", "minutes"
            ParseIntervalSeconds = CLng(amount * 60)
        Case "h", "hr", "hrs", "hour", "hours"
            ParseIntervalSeconds = CLng(amount * 3600)
        Case "d", "day", "days"
            ParseIntervalSeconds = CLng(amount * 86400)
        Case Else
            ParseIntervalSeconds = 0
    End Select
End Function

Function FormatIntervalSeconds(ByVal intervalSeconds As Long) As String
    If intervalSeconds > 0 And intervalSeconds Mod 86400 = 0 Then
        FormatIntervalSeconds = CStr(intervalSeconds \ 86400) & "d"
    ElseIf intervalSeconds > 0 And intervalSeconds Mod 3600 = 0 Then
        FormatIntervalSeconds = CStr(intervalSeconds \ 3600) & "h"
    ElseIf intervalSeconds > 0 And intervalSeconds Mod 60 = 0 Then
        FormatIntervalSeconds = CStr(intervalSeconds \ 60) & "m"
    Else
        FormatIntervalSeconds = CStr(intervalSeconds) & "s"
    End If
End Function

Private Sub SetAllowedAttachmentExtensions()
    Dim currentValue As String
    Dim inputValue As String
    Dim normalizedValue As String
    
    currentValue = FormatAllowedExtensionsForInput(GetAllowedAttachmentExtensions())
    inputValue = InputBox(RuText(1042, 1074, 1077, 1076, 1080, 1090, 1077, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 32, 1095, 1077, 1088, 1077, 1079, 32, 1079, 1072, 1087, 1103, 1090, 1091, 1102, 46, 32, 1055, 1088, 1080, 1084, 1077, 1088, 58, 32, 112, 100, 102, 44, 32, 100, 111, 99, 120, 44, 32, 120, 108, 115, 120) & vbCrLf & RuText(1054, 1089, 1090, 1072, 1074, 1100, 1090, 1077, 32, 1087, 1091, 1089, 1090, 1099, 1084, 44, 32, 1095, 1090, 1086, 1073, 1099, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1103, 1090, 1100, 32, 1074, 1089, 1077, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1103, 46), RuText(1056, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081), currentValue)
    normalizedValue = NormalizeAllowedExtensions(inputValue)
    
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_ALLOWED_EXTENSIONS_KEY, normalizedValue
    
    If Len(normalizedValue) = 0 Then
        MsgBox RuText(1060, 1080, 1083, 1100, 1090, 1088, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1081, 32, 1086, 1090, 1082, 1083, 1102, 1095, 1077, 1085, 46, 32, 1041, 1091, 1076, 1091, 1090, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1103, 1090, 1100, 1089, 1103, 32, 1074, 1089, 1077, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1103, 46), vbInformation
    Else
        MsgBox RuText(1041, 1091, 1076, 1091, 1090, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1103, 1090, 1100, 1089, 1103, 32, 1090, 1086, 1083, 1100, 1082, 1086, 32, 1101, 1090, 1080, 32, 1088, 1072, 1089, 1096, 1080, 1088, 1077, 1085, 1080, 1103, 58, 32) & FormatAllowedExtensionsForInput(normalizedValue), vbInformation
    End If
End Sub

Private Sub SetAttachmentSortRules()
    Dim currentValue As String
    Dim inputValue As String
    Dim normalizedValue As String
    
    currentValue = FormatAttachmentSortRulesForInput(GetAttachmentSortRules())
    inputValue = InputBox(RuText(1042, 1074, 1077, 1076, 1080, 1090, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1072, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 32, 1095, 1077, 1088, 1077, 1079, 32, 1090, 1086, 1095, 1082, 1091, 32, 1089, 32, 1079, 1072, 1087, 1103, 1090, 1086, 1081, 46) & vbCrLf & _
                          RuText(1060, 1086, 1088, 1084, 1072, 1090, 58, 32, 1096, 1072, 1073, 1083, 1086, 1085, 61, 1087, 1072, 1087, 1082, 1072, 32, 1080, 1083, 1080, 32, 1096, 1072, 1073, 1083, 1086, 1085, 61, 67, 58, 92, 1055, 1086, 1083, 1085, 1099, 1081, 92, 1055, 1091, 1090, 1100) & vbCrLf & _
                          RuText(1055, 1088, 1080, 1084, 1077, 1088, 58, 32, 52, 48, 55, 48, 50, 56, 49, 48, 48, 48, 48, 48, 48, 48, 48, 49, 52, 51, 51, 48, 61, 1054, 1087, 1077, 1088, 1072, 1090, 1086, 1088), _
                          RuText(1055, 1088, 1072, 1074, 1080, 1083, 1072, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081), currentValue)
    
    normalizedValue = NormalizeAttachmentSortRules(inputValue)
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_SORT_RULES_KEY, normalizedValue
    
    If Len(normalizedValue) = 0 Then
        MsgBox RuText(1055, 1088, 1072, 1074, 1080, 1083, 1072, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080, 32, 1086, 1090, 1082, 1083, 1102, 1095, 1077, 1085, 1099, 46, 32, 1060, 1072, 1081, 1083, 1099, 32, 1073, 1091, 1076, 1091, 1090, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1103, 1090, 1100, 1089, 1103, 32, 1087, 1086, 32, 1083, 1086, 1075, 1080, 1082, 1077, 32, 1087, 1072, 1087, 1082, 1080, 32, 1087, 1086, 32, 1091, 1084, 1086, 1083, 1095, 1072, 1085, 1080, 1102, 46), vbInformation
    Else
        MsgBox RuText(1053, 1072, 1089, 1090, 1088, 1086, 1077, 1085, 1086, 32, 1087, 1088, 1072, 1074, 1080, 1083, 32, 1089, 1086, 1088, 1090, 1080, 1088, 1086, 1074, 1082, 1080, 58, 32) & CountAttachmentSortRules(normalizedValue), vbInformation
    End If
End Sub

Private Sub SetAutoSaveIntervalSetting()
    Dim currentValue As String
    Dim inputValue As String
    Dim intervalSeconds As Long
    
    currentValue = FormatIntervalSeconds(GetAutoSaveIntervalSeconds())
    inputValue = InputBox(RuText(1042, 1074, 1077, 1076, 1080, 1090, 1077, 32, 1080, 1085, 1090, 1077, 1088, 1074, 1072, 1083, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 46) & vbCrLf & _
                          RuText(1048, 1089, 1087, 1086, 1083, 1100, 1079, 1091, 1081, 1090, 1077, 32, 115, 44, 32, 109, 32, 1080, 1083, 1080, 32, 104, 46, 32, 1055, 1088, 1080, 1084, 1077, 1088, 1099, 58, 32, 51, 48, 115, 44, 32, 49, 53, 109, 44, 32, 49, 50, 104) & vbCrLf & _
                          RuText(1053, 1072, 1087, 1086, 1084, 1080, 1085, 1072, 1085, 1080, 1103, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1084, 1086, 1075, 1091, 1090, 32, 1079, 1072, 1087, 1091, 1089, 1082, 1072, 1090, 1100, 1089, 1103, 32, 1085, 1077, 32, 1090, 1086, 1095, 1085, 1086, 32, 1076, 1086, 32, 1089, 1077, 1082, 1091, 1085, 1076, 1099, 46), _
                          RuText(1048, 1085, 1090, 1077, 1088, 1074, 1072, 1083, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103), currentValue)
    
    If Len(inputValue) = 0 Then
        MsgBox RuText(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1072, 32, 1080, 1085, 1090, 1077, 1088, 1074, 1072, 1083, 1072, 32, 1086, 1090, 1084, 1077, 1085, 1077, 1085, 1072, 46), vbInformation
        Exit Sub
    End If
    
    intervalSeconds = ParseIntervalSeconds(inputValue)
    
    If intervalSeconds <= 0 Then
        MsgBox RuText(1053, 1077, 1082, 1086, 1088, 1088, 1077, 1082, 1090, 1085, 1099, 1081, 32, 1080, 1085, 1090, 1077, 1088, 1074, 1072, 1083, 46, 32, 1048, 1089, 1087, 1086, 1083, 1100, 1079, 1091, 1081, 1090, 1077, 32, 1079, 1085, 1072, 1095, 1077, 1085, 1080, 1103, 32, 1074, 1088, 1086, 1076, 1077, 32, 51, 48, 115, 44, 32, 49, 53, 109, 32, 1080, 1083, 1080, 32, 49, 50, 104, 46), vbCritical
        Exit Sub
    End If
    
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_INTERVAL_SECONDS_KEY, CStr(intervalSeconds)
    
    If IsAutoSaveEnabled() Then
        ScheduleNextAutoSave
    End If
End Sub

Private Sub SetAutoSaveLookbackSetting()
    Dim currentValue As String
    Dim inputValue As String
    Dim lookbackSeconds As Long
    
    currentValue = FormatIntervalSeconds(GetAutoSaveLookbackSeconds())
    inputValue = InputBox(RuText(1042, 1074, 1077, 1076, 1080, 1090, 1077, 32, 1087, 1077, 1088, 1080, 1086, 1076, 32, 1086, 1073, 1088, 1072, 1073, 1086, 1090, 1082, 1080, 32, 1087, 1080, 1089, 1077, 1084, 46) & vbCrLf & _
                          RuText(1048, 1089, 1087, 1086, 1083, 1100, 1079, 1091, 1081, 1090, 1077, 32, 115, 44, 32, 109, 44, 32, 104, 32, 1080, 1083, 1080, 32, 100, 46, 32, 1055, 1088, 1080, 1084, 1077, 1088, 1099, 58, 32, 49, 50, 104, 44, 32, 50, 52, 104, 44, 32, 51, 100, 44, 32, 55, 100) & vbCrLf & _
                          RuText(79, 117, 116, 108, 111, 111, 107, 32, 1073, 1091, 1076, 1077, 1090, 32, 1080, 1089, 1082, 1072, 1090, 1100, 32, 1087, 1080, 1089, 1100, 1084, 1072, 32, 1085, 1077, 32, 1089, 1090, 1072, 1088, 1096, 1077, 32, 1101, 1090, 1086, 1075, 1086, 32, 1087, 1077, 1088, 1080, 1086, 1076, 1072, 46), _
                          RuText(1055, 1077, 1088, 1080, 1086, 1076, 32, 1086, 1073, 1088, 1072, 1073, 1086, 1090, 1082, 1080), currentValue)
    
    If Len(inputValue) = 0 Then
        MsgBox RuText(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1072, 32, 1087, 1077, 1088, 1080, 1086, 1076, 1072, 32, 1086, 1090, 1084, 1077, 1085, 1077, 1085, 1072, 46), vbInformation
        Exit Sub
    End If
    
    lookbackSeconds = ParseIntervalSeconds(inputValue)
    
    If lookbackSeconds <= 0 Then
        MsgBox RuText(1053, 1077, 1082, 1086, 1088, 1088, 1077, 1082, 1090, 1085, 1099, 1081, 32, 1087, 1077, 1088, 1080, 1086, 1076, 46, 32, 1048, 1089, 1087, 1086, 1083, 1100, 1079, 1091, 1081, 1090, 1077, 32, 1079, 1085, 1072, 1095, 1077, 1085, 1080, 1103, 32, 1074, 1088, 1086, 1076, 1077, 32, 49, 50, 104, 44, 32, 50, 52, 104, 44, 32, 51, 100, 32, 1080, 1083, 1080, 32, 55, 100, 46), vbCritical
        Exit Sub
    End If
    
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_LOOKBACK_SECONDS_KEY, CStr(lookbackSeconds)
    ResetLastAutoSaveRunToLookbackStart
End Sub

Private Sub EnableAutoSaveAttachments()
    Dim previousLastRunValue As String
    
    previousLastRunValue = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_LAST_RUN_KEY, "")
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_ENABLED_KEY, "True"
    ResetLastAutoSaveRunToLookbackStart
    
    If Not SaveAttachmentsFromConfiguredFoldersSinceLastRun() Then
        SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_ENABLED_KEY, "False"
        SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_LAST_RUN_KEY, previousLastRunValue
        StopAutoSaveTimer
        DeleteLegacyAutoSaveReminders
        
        If Len(lastAutoSaveRunFailureMessage) > 0 Then
            MsgBox lastAutoSaveRunFailureMessage, vbExclamation
        Else
            MsgBox RuText(1040, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 32, 1085, 1077, 32, 1074, 1082, 1083, 1102, 1095, 1077, 1085, 1086, 58, 32, 1087, 1077, 1088, 1074, 1099, 1081, 32, 1079, 1072, 1087, 1091, 1089, 1082, 32, 1079, 1072, 1074, 1077, 1088, 1096, 1080, 1083, 1089, 1103, 32, 1089, 32, 1086, 1096, 1080, 1073, 1082, 1086, 1081, 46), vbExclamation
        End If
        
        Exit Sub
    End If
    
    ScheduleNextAutoSave
End Sub

Private Sub ResetLastAutoSaveRunToLookbackStart()
    Dim dtLookbackStart As Date
    
    dtLookbackStart = DateAdd("s", -GetAutoSaveLookbackSeconds(), Now)
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_LAST_RUN_KEY, Format(dtLookbackStart, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Auto save last run reset to lookback start: " & Format(dtLookbackStart, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Sub DisableAutoSaveAttachments()
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_ENABLED_KEY, "False"
    StopAutoSaveTimer
    DeleteLegacyAutoSaveReminders
End Sub

Private Function SaveAttachmentsFromConfiguredFoldersSinceLastRun() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ns As Outlook.NameSpace
    Dim mailFolder As Outlook.MAPIFolder
    Dim dtLastRun As Date
    Dim dtRunStarted As Date
    Dim savedCount As Long
    Dim checkedCount As Long
    Dim attachmentSortRules As Collection
    Dim ruleSourceFolderPaths As Collection
    Dim sourceFolderPath As Variant
    Dim processedFolderCount As Long
    Dim reachedItemLimit As Boolean
    
    lastAutoSaveRunFailureMessage = ""
    dtRunStarted = Now
    dtLastRun = ClampAutoSaveStartTime(GetLastAutoSaveRun(), dtRunStarted)
    Set attachmentSortRules = LoadAttachmentSortRules()
    Set ns = Outlook.Application.GetNamespace("MAPI")

    Set ruleSourceFolderPaths = GetSourceFolderPathsFromSortRules(attachmentSortRules)
    
    If ruleSourceFolderPaths.Count = 0 Then
        Debug.Print "Auto save skipped: no active sort rules with source Outlook folders."
        lastAutoSaveRunFailureMessage = RuText(1040, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 32, 1085, 1077, 32, 1074, 1082, 1083, 1102, 1095, 1077, 1085, 1086, 58, 32, 1085, 1077, 1090, 32, 1072, 1082, 1090, 1080, 1074, 1085, 1099, 1093, 32, 1087, 1088, 1072, 1074, 1080, 1083, 32, 1089, 32, 1080, 1089, 1093, 1086, 1076, 1085, 1086, 1081, 32, 1087, 1072, 1087, 1082, 1086, 1081, 46)
    Else
        For Each sourceFolderPath In ruleSourceFolderPaths
            If checkedCount >= AUTO_SAVE_MAX_ITEMS_PER_RUN Then
                reachedItemLimit = True
                Exit For
            End If
            
            Set mailFolder = FindMailFolderByPath(ns, CStr(sourceFolderPath))
            
            If Not mailFolder Is Nothing Then
                processedFolderCount = processedFolderCount + 1
                savedCount = savedCount + SaveAttachmentsFromFolderSince(mailFolder, "", dtLastRun, checkedCount, "", attachmentSortRules, CStr(sourceFolderPath))
                
                If checkedCount > AUTO_SAVE_MAX_ITEMS_PER_RUN Then
                    reachedItemLimit = True
                End If
            Else
                Debug.Print "Auto save skipped missing rule source folder: " & CStr(sourceFolderPath)
            End If
        Next sourceFolderPath
    End If
    
    If processedFolderCount > 0 And Not reachedItemLimit Then
        SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_LAST_RUN_KEY, Format(dtRunStarted, "yyyy-mm-dd hh:nn:ss")
        SaveAttachmentsFromConfiguredFoldersSinceLastRun = True
    ElseIf reachedItemLimit Then
        Debug.Print "Auto save last run not updated: item limit was reached."
        SaveAttachmentsFromConfiguredFoldersSinceLastRun = True
    Else
        Debug.Print "Auto save last run not updated: no source folders were processed."
        If Len(lastAutoSaveRunFailureMessage) = 0 Then
            lastAutoSaveRunFailureMessage = RuText(1040, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 32, 1085, 1077, 32, 1074, 1099, 1087, 1086, 1083, 1085, 1077, 1085, 1086, 58, 32, 1085, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1086, 1090, 1082, 1088, 1099, 1090, 1100, 32, 1080, 1089, 1093, 1086, 1076, 1085, 1099, 1077, 32, 1087, 1072, 1087, 1082, 1080, 32, 1072, 1082, 1090, 1080, 1074, 1085, 1099, 1093, 32, 1087, 1088, 1072, 1074, 1080, 1083, 46)
        End If
    End If
    
    Debug.Print "Auto save completed. Items checked: " & checkedCount & ". Attachments saved: " & savedCount
    Exit Function
    
ErrorHandler:
    Debug.Print "Auto save failed: " & CStr(Err.Number) & ". " & Err.Description
    lastAutoSaveRunFailureMessage = RuText(1040, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1077, 32, 1085, 1077, 32, 1074, 1082, 1083, 1102, 1095, 1077, 1085, 1086, 58, 32, 1087, 1077, 1088, 1074, 1099, 1081, 32, 1079, 1072, 1087, 1091, 1089, 1082, 32, 1079, 1072, 1074, 1077, 1088, 1096, 1080, 1083, 1089, 1103, 32, 1089, 32, 1086, 1096, 1080, 1073, 1082, 1086, 1081, 46) & " " & CStr(Err.Number) & ". " & Err.Description
    SaveAttachmentsFromConfiguredFoldersSinceLastRun = False
End Function

Private Function GetSourceFolderPathsFromSortRules(ByVal attachmentSortRules As Collection) As Collection
    Dim folderPaths As Collection
    Dim seenPaths As Object
    Dim sortRule As Variant
    Dim folderPath As String
    
    Set folderPaths = New Collection
    Set seenPaths = CreateObject("Scripting.Dictionary")
    seenPaths.CompareMode = vbTextCompare
    
    For Each sortRule In attachmentSortRules
        folderPath = NormalizeOutlookFolderPath(CStr(sortRule("SourceFolder")))
        
        If Len(folderPath) > 0 And Not seenPaths.Exists(folderPath) Then
            seenPaths(folderPath) = True
            folderPaths.Add folderPath
        End If
    Next sortRule
    
    Set GetSourceFolderPathsFromSortRules = folderPaths
End Function

Private Function FindMailFolderByPath(ByVal ns As Outlook.NameSpace, ByVal folderPath As String) As Outlook.MAPIFolder
    Dim store As Outlook.Store
    Dim rootFolder As Outlook.MAPIFolder
    Dim foundFolder As Outlook.MAPIFolder
    Dim normalizedFolderPath As String
    Dim folderPathSuffix As String
    
    normalizedFolderPath = NormalizeOutlookFolderPath(folderPath)
    
    If Len(normalizedFolderPath) = 0 Then
        Exit Function
    End If
    
    Set foundFolder = FindCurrentExplorerFolderByPath(normalizedFolderPath)
    
    If Not foundFolder Is Nothing Then
        Set FindMailFolderByPath = foundFolder
        Exit Function
    End If
    
    For Each store In ns.Stores
        Set rootFolder = Nothing
        
        On Error Resume Next
        Set rootFolder = store.GetRootFolder
        On Error GoTo 0
        
        If Not rootFolder Is Nothing Then
            Set foundFolder = FindMailFolderByPathUnder(rootFolder, normalizedFolderPath)
            
            If Not foundFolder Is Nothing Then
                Set FindMailFolderByPath = foundFolder
                Exit Function
            End If
        End If
    Next store
    
    folderPathSuffix = GetOutlookFolderPathSuffix(normalizedFolderPath)
    
    If Len(folderPathSuffix) = 0 Then
        Exit Function
    End If
    
    For Each store In ns.Stores
        Set rootFolder = Nothing
        
        On Error Resume Next
        Set rootFolder = store.GetRootFolder
        On Error GoTo 0
        
        If Not rootFolder Is Nothing Then
            Set foundFolder = FindMailFolderByPathSuffixUnder(rootFolder, folderPathSuffix)
            
            If Not foundFolder Is Nothing Then
                Set FindMailFolderByPath = foundFolder
                Exit Function
            End If
        End If
    Next store
    
    Debug.Print "Auto save source folder not found by exact path or suffix: " & normalizedFolderPath
End Function

Private Function FindCurrentExplorerFolderByPath(ByVal normalizedFolderPath As String) As Outlook.MAPIFolder
    On Error GoTo ErrorHandler
    
    Dim currentFolder As Outlook.MAPIFolder
    Dim currentFolderPath As String
    Dim folderPathSuffix As String
    
    If Outlook.Application.ActiveExplorer Is Nothing Then
        Exit Function
    End If
    
    If Not TypeOf Outlook.Application.ActiveExplorer.CurrentFolder Is Outlook.MAPIFolder Then
        Exit Function
    End If
    
    Set currentFolder = Outlook.Application.ActiveExplorer.CurrentFolder
    currentFolderPath = NormalizeOutlookFolderPath(currentFolder.FolderPath)
    
    If StrComp(currentFolderPath, normalizedFolderPath, vbTextCompare) = 0 Then
        Set FindCurrentExplorerFolderByPath = currentFolder
        Exit Function
    End If
    
    folderPathSuffix = GetOutlookFolderPathSuffix(normalizedFolderPath)
    
    If Len(folderPathSuffix) > 0 And Len(currentFolderPath) >= Len(folderPathSuffix) Then
        If StrComp(Right(currentFolderPath, Len(folderPathSuffix)), folderPathSuffix, vbTextCompare) = 0 Then
            Set FindCurrentExplorerFolderByPath = currentFolder
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Set FindCurrentExplorerFolderByPath = Nothing
End Function

Private Function GetOutlookFolderPathSuffix(ByVal normalizedFolderPath As String) As String
    Dim pathWithoutPrefix As String
    Dim pos As Long
    
    normalizedFolderPath = NormalizeOutlookFolderPath(normalizedFolderPath)
    
    If Left(normalizedFolderPath, 2) <> "\\" Then
        Exit Function
    End If
    
    pathWithoutPrefix = Mid(normalizedFolderPath, 3)
    pos = InStr(1, pathWithoutPrefix, "\", vbBinaryCompare)
    
    If pos <= 0 Or pos >= Len(pathWithoutPrefix) Then
        Exit Function
    End If
    
    GetOutlookFolderPathSuffix = "\" & Mid(pathWithoutPrefix, pos + 1)
End Function

Private Function FindMailFolderByPathSuffixUnder(ByVal folder As Outlook.MAPIFolder, ByVal normalizedFolderPathSuffix As String) As Outlook.MAPIFolder
    Dim childFolder As Outlook.MAPIFolder
    Dim foundFolder As Outlook.MAPIFolder
    Dim currentFolderPath As String
    
    On Error Resume Next
    
    currentFolderPath = NormalizeOutlookFolderPath(folder.FolderPath)
    
    If folder.DefaultItemType = olMailItem Then
        If Len(currentFolderPath) >= Len(normalizedFolderPathSuffix) Then
            If StrComp(Right(currentFolderPath, Len(normalizedFolderPathSuffix)), normalizedFolderPathSuffix, vbTextCompare) = 0 Then
                Set FindMailFolderByPathSuffixUnder = folder
                Exit Function
            End If
        End If
    End If
    
    For Each childFolder In folder.Folders
        Set foundFolder = FindMailFolderByPathSuffixUnder(childFolder, normalizedFolderPathSuffix)
        
        If Not foundFolder Is Nothing Then
            Set FindMailFolderByPathSuffixUnder = foundFolder
            Exit Function
        End If
    Next childFolder
    
    On Error GoTo 0
End Function

Private Function FindMailFolderByPathUnder(ByVal folder As Outlook.MAPIFolder, ByVal normalizedFolderPath As String) As Outlook.MAPIFolder
    Dim childFolder As Outlook.MAPIFolder
    Dim foundFolder As Outlook.MAPIFolder
    
    On Error Resume Next
    
    If folder.DefaultItemType = olMailItem Then
        If StrComp(NormalizeOutlookFolderPath(folder.FolderPath), normalizedFolderPath, vbTextCompare) = 0 Then
            Set FindMailFolderByPathUnder = folder
            Exit Function
        End If
    End If
    
    For Each childFolder In folder.Folders
        Set foundFolder = FindMailFolderByPathUnder(childFolder, normalizedFolderPath)
        
        If Not foundFolder Is Nothing Then
            Set FindMailFolderByPathUnder = foundFolder
            Exit Function
        End If
    Next childFolder
    
    On Error GoTo 0
End Function

Private Sub RunScheduledAutoSaveAttachments()
    If Not IsAutoSaveEnabled() Then
        DeleteLegacyAutoSaveReminders
        Exit Sub
    End If
    
    SaveAttachmentsFromConfiguredFoldersSinceLastRun
    ScheduleNextAutoSave
End Sub

Private Sub ScheduleNextAutoSave()
    Dim dtNextRun As Date
    
    DeleteLegacyAutoSaveReminders
    
    If Not StartAutoSaveTimer() Then
        SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_NEXT_RUN_KEY, ""
        Debug.Print "Auto save next timer run was not scheduled."
        Exit Sub
    End If
    
    dtNextRun = DateAdd("s", GetAutoSaveIntervalSeconds(), Now)
    
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_NEXT_RUN_KEY, Format(dtNextRun, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Auto save next timer run: " & Format(dtNextRun, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Sub DeleteLegacyAutoSaveReminders()
    On Error Resume Next
    
    Dim ns As Outlook.NameSpace
    Dim calendar As Outlook.MAPIFolder
    Dim items As Outlook.Items
    Dim i As Long
    
    Set ns = Outlook.Application.GetNamespace("MAPI")
    Set calendar = ns.GetDefaultFolder(olFolderCalendar)
    Set items = calendar.Items
    
    For i = items.Count To 1 Step -1
        If items.Item(i).Class = olAppointment Then
            If items.Item(i).Subject = LEGACY_AUTO_SAVE_REMINDER_SUBJECT Then
                items.Item(i).Delete
            End If
        End If
    Next i
    
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_NEXT_RUN_KEY, ""
End Sub

Private Function StartAutoSaveTimer() As Boolean
    Dim intervalMs As Long
    Dim intervalSeconds As Long
    
    StopAutoSaveTimer
    
    If Not IsAutoSaveEnabled() Then
        Exit Function
    End If
    
    intervalSeconds = GetAutoSaveIntervalSeconds()
    
    If intervalSeconds < 1 Then
        intervalSeconds = 1
    End If
    
    If intervalSeconds > 2147483 Then
        intervalSeconds = 2147483
    End If
    
    intervalMs = CLng(intervalSeconds * 1000)
    autoSaveTimerId = SetTimer(0, 0, intervalMs, AddressOf AutoSaveTimerProc)
    
    If autoSaveTimerId <> 0 Then
        StartAutoSaveTimer = True
        Debug.Print "Auto save Windows timer started. Interval: " & FormatIntervalSeconds(intervalSeconds) & ". Timer id: " & CStr(autoSaveTimerId)
    Else
        Debug.Print "Auto save Windows timer failed to start."
    End If
End Function

Private Sub StopAutoSaveTimer()
    If autoSaveTimerId <> 0 Then
        KillTimer 0, autoSaveTimerId
        Debug.Print "Auto save Windows timer stopped. Timer id: " & CStr(autoSaveTimerId)
        autoSaveTimerId = 0
    End If
End Sub

Private Function IsAutoSaveTimerActive() As Boolean
    IsAutoSaveTimerActive = (autoSaveTimerId <> 0)
End Function

#If VBA7 Then
Private Sub AutoSaveTimerProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Private Sub AutoSaveTimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    Dim timerRestarted As Boolean
    
    On Error GoTo ErrorHandler
    
    StopAutoSaveTimer
    
    If IsAutoSaveEnabled() Then
        RunScheduledAutoSaveAttachments
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Auto save Windows timer failed: " & Err.Description
    If IsAutoSaveEnabled() Then
        timerRestarted = StartAutoSaveTimer()
    End If
End Sub

Function IsAutoSaveEnabled() As Boolean
    IsAutoSaveEnabled = (GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_ENABLED_KEY, "False") = "True")
End Function

Function IsAutoSaveLoggingEnabled() As Boolean
    IsAutoSaveLoggingEnabled = (GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_AUTO_LOGGING_ENABLED_KEY, "False") = "True")
End Function

Function GetAttachmentsSaveFolder() As String
    Dim strFolderPath As String
    
    strFolderPath = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_SAVE_FOLDER_KEY, GetDefaultSaveFolder())
    strFolderPath = Trim(strFolderPath)
    
    If Len(strFolderPath) = 0 Then
        strFolderPath = GetDefaultSaveFolder()
    End If
    
    GetAttachmentsSaveFolder = EnsureTrailingBackslash(strFolderPath)
End Function

Function GetDefaultSaveFolder() As String
    Dim userProfilePath As String
    
    userProfilePath = Environ$("USERPROFILE")
    
    If Len(Trim(userProfilePath)) = 0 Then
        GetDefaultSaveFolder = "C:\EmailAttachments\"
    Else
        GetDefaultSaveFolder = EnsureTrailingBackslash(userProfilePath) & "Downloads\EmailAttachments\"
    End If
End Function

Function GetLastAutoSaveRun() As Date
    Dim rawValue As String
    
    rawValue = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_LAST_RUN_KEY, "")
    
    If IsNumeric(rawValue) Then
        GetLastAutoSaveRun = CDate(CDbl(rawValue))
    ElseIf IsDate(rawValue) Then
        GetLastAutoSaveRun = CDate(rawValue)
    Else
        GetLastAutoSaveRun = DateAdd("s", -GetAutoSaveLookbackSeconds(), Now)
    End If
End Function

Function GetLastAutoSaveRunText() As String
    Dim rawValue As String
    Dim dtValue As Date
    
    rawValue = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_LAST_RUN_KEY, "")
    
    If Len(Trim(rawValue)) = 0 Then
        GetLastAutoSaveRunText = RuText(40, 1085, 1077, 1090, 41)
    ElseIf IsNumeric(rawValue) Then
        dtValue = CDate(CDbl(rawValue))
        GetLastAutoSaveRunText = Format(dtValue, "yyyy-mm-dd hh:nn:ss")
    ElseIf IsDate(rawValue) Then
        dtValue = CDate(rawValue)
        GetLastAutoSaveRunText = Format(dtValue, "yyyy-mm-dd hh:nn:ss")
    Else
        GetLastAutoSaveRunText = rawValue
    End If
End Function

Function GetNextAutoSaveRunText() As String
    Dim rawValue As String
    Dim dtValue As Date
    
    rawValue = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_NEXT_RUN_KEY, "")
    
    If Len(Trim(rawValue)) = 0 Then
        GetNextAutoSaveRunText = RuText(40, 1085, 1077, 32, 1079, 1072, 1087, 1083, 1072, 1085, 1080, 1088, 1086, 1074, 1072, 1085, 1086, 41)
    ElseIf IsDate(rawValue) Then
        dtValue = CDate(rawValue)
        GetNextAutoSaveRunText = Format(dtValue, "yyyy-mm-dd hh:nn:ss")
    Else
        GetNextAutoSaveRunText = rawValue
    End If
End Function

Function ClampAutoSaveStartTime(ByVal dtLastRun As Date, ByVal dtRunStarted As Date) As Date
    Dim dtOldestAllowed As Date
    
    dtOldestAllowed = DateAdd("s", -GetAutoSaveLookbackSeconds(), dtRunStarted)
    
    If dtLastRun < dtOldestAllowed Then
        ClampAutoSaveStartTime = dtOldestAllowed
    Else
        ClampAutoSaveStartTime = dtLastRun
    End If
End Function

Function BrowseForFolder(ByVal strTitle As String) As String
    Dim shellApp As Object
    Dim folder As Object
    
    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, strTitle, 0, 0)
    
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If
End Function

Function BrowseForFolderEx(ByVal strTitle As String, Optional ByVal initialFolderPath As String = "") As String
    On Error GoTo Fallback
    
    Dim dialog As Object
    Dim normalizedInitialPath As String
    
    Set dialog = Outlook.Application.FileDialog(4)
    
    normalizedInitialPath = Trim(initialFolderPath)
    
    If Len(normalizedInitialPath) > 0 And Dir(normalizedInitialPath, vbDirectory) <> "" Then
        normalizedInitialPath = EnsureTrailingBackslash(normalizedInitialPath)
    Else
        normalizedInitialPath = ""
    End If
    
    With dialog
        .Title = strTitle
        .AllowMultiSelect = False
        
        If Len(normalizedInitialPath) > 0 Then
            .InitialFileName = normalizedInitialPath
        End If
        
        If .Show = -1 Then
            BrowseForFolderEx = .SelectedItems(1)
        Else
            BrowseForFolderEx = ""
        End If
    End With
    
    Exit Function
    
Fallback:
    Err.Clear
    BrowseForFolderEx = BrowseForFolder(strTitle)
End Function

Function EnsureTrailingBackslash(ByVal strFolderPath As String) As String
    If Right(strFolderPath, 1) = "\" Then
        EnsureTrailingBackslash = strFolderPath
    Else
        EnsureTrailingBackslash = strFolderPath & "\"
    End If
End Function

Function RemoveTrailingBackslash(ByVal strFolderPath As String) As String
    strFolderPath = Trim(strFolderPath)
    
    Do While Len(strFolderPath) > 3 And Right(strFolderPath, 1) = "\"
        strFolderPath = Left(strFolderPath, Len(strFolderPath) - 1)
    Loop
    
    RemoveTrailingBackslash = strFolderPath
End Function

Function SaveAttachmentsFromFolderSince(ByVal mailFolder As Outlook.MAPIFolder, ByVal strBaseFolderPath As String, ByVal dtLastRun As Date, ByRef checkedCount As Long, Optional ByVal allowedExtensionsOverride As Variant, Optional ByVal attachmentSortRulesOverride As Variant, Optional ByVal sourceFolderPathOverride As Variant) As Long
    On Error GoTo FolderError
    
    Dim items As Outlook.Items
    Dim objItem As Object
    Dim objMail As Outlook.mailItem
    Dim attachmentSortRules As Collection
    Dim savedCount As Long
    
    If IsMissing(attachmentSortRulesOverride) Then
        Set attachmentSortRules = LoadAttachmentSortRules()
    Else
        Set attachmentSortRules = attachmentSortRulesOverride
    End If
    
    Set items = mailFolder.Items
    items.Sort "[ReceivedTime]", True
    
    Debug.Print "Auto save checking Outlook folder: " & mailFolder.FolderPath & ". Since: " & Format(dtLastRun, "yyyy-mm-dd hh:nn:ss") & ". Items in folder: " & CStr(items.Count)
    
    For Each objItem In items
        On Error GoTo ItemError
        
        checkedCount = checkedCount + 1
        If checkedCount > AUTO_SAVE_MAX_ITEMS_PER_RUN Then
            Debug.Print "Auto save stopped after checking " & AUTO_SAVE_MAX_ITEMS_PER_RUN & " item(s)."
            Exit For
        End If
        
        If TypeOf objItem Is mailItem Then
            Set objMail = objItem
            
            If objMail.ReceivedTime <= dtLastRun Then
                Exit For
            End If
            
            If IsMissing(sourceFolderPathOverride) Then
                savedCount = savedCount + SaveAttachmentsFromMailItem(objMail, strBaseFolderPath, "", attachmentSortRules)
            Else
                savedCount = savedCount + SaveAttachmentsFromMailItem(objMail, strBaseFolderPath, "", attachmentSortRules, sourceFolderPathOverride:=CStr(sourceFolderPathOverride))
            End If
        End If
        
ContinueNextItem:
        On Error GoTo FolderError
    Next objItem
    
    SaveAttachmentsFromFolderSince = savedCount
    Exit Function
    
ItemError:
    Debug.Print "Auto save skipped item after error: " & CStr(Err.Number) & ". " & Err.Description
    Err.Clear
    Resume ContinueNextItem
    
FolderError:
    Debug.Print "Auto save skipped folder after error: " & CStr(Err.Number) & ". " & Err.Description
    Err.Clear
    SaveAttachmentsFromFolderSince = savedCount
End Function

Private Sub WriteAutoSaveLog(ByVal targetFolderPath As String, ByVal statusValue As String, ByVal reasonValue As String, ByVal objMail As Outlook.mailItem, ByVal attachmentName As String, ByVal targetFilePath As String, Optional ByVal ruleDetails As String = "")
    On Error GoTo ErrorHandler
    
    Dim logFilePath As String
    Dim fileNumber As Integer
    Dim mailSubject As String
    Dim mailReceived As String
    
    If Not IsAutoSaveLoggingEnabled() Then
        Exit Sub
    End If
    
    If Len(targetFolderPath) = 0 Then
        Exit Sub
    End If
    
    If Dir(targetFolderPath, vbDirectory) = "" Then
        Exit Sub
    End If
    
    If Not objMail Is Nothing Then
        mailSubject = objMail.Subject
        mailReceived = Format(objMail.ReceivedTime, "yyyy-mm-dd hh:nn:ss")
    End If
    
    logFilePath = EnsureTrailingBackslash(targetFolderPath) & "autosave.log"
    fileNumber = FreeFile
    
    Open logFilePath For Append As #fileNumber
    Print #fileNumber, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & LogField(statusValue) & " | " & LogField(reasonValue) & " | " & LogField(ruleDetails) & " | " & LogField(mailSubject) & " | " & LogField(mailReceived) & " | " & LogField(attachmentName) & " | " & LogField(targetFilePath)
    Close #fileNumber
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If fileNumber <> 0 Then
        Close #fileNumber
    End If
End Sub

Private Function LogField(ByVal value As String) As String
    value = Replace(value, vbCrLf, " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, "|", "/")
    
    LogField = value
End Function

Function SaveAttachmentsFromMailItem(ByVal objMail As Outlook.mailItem, ByVal strBaseFolderPath As String, Optional ByVal allowedExtensionsOverride As Variant, Optional ByVal attachmentSortRulesOverride As Variant, Optional ByVal savedFilePathsOverride As Variant, Optional ByVal sourceFolderPathOverride As Variant) As Long
    Dim objAttachments As Outlook.Attachments
    Dim strFolderPath As String
    Dim strFolderName As String
    Dim strFileName As String
    Dim strOriginalFileName As String
    Dim strAttachmentName As String
    Dim logStatus As String
    Dim logReason As String
    Dim autoSaveIndexKey As String
    Dim indexedFilePath As String
    Dim sourceFolderPath As String
    Dim ruleLogDetails As String
    Dim appliedSortRule As Object
    Dim attachmentSortRules As Collection
    Dim i As Long
    Dim savedCount As Long
    
    Set objAttachments = objMail.Attachments
    
    If objAttachments.Count = 0 Then
        SaveAttachmentsFromMailItem = 0
        Exit Function
    End If
    
    strFolderName = CreateValidFolderName(objMail.Subject)
    If IsMissing(sourceFolderPathOverride) Then
        sourceFolderPath = GetMailItemFolderPath(objMail)
    Else
        sourceFolderPath = CStr(sourceFolderPathOverride)
    End If
    
    If IsMissing(attachmentSortRulesOverride) Then
        Set attachmentSortRules = LoadAttachmentSortRules()
    Else
        Set attachmentSortRules = attachmentSortRulesOverride
    End If
    
    Debug.Print "Sanitized Folder Name: " & strFolderName
    
    For i = 1 To objAttachments.Count
        On Error Resume Next
        Err.Clear
        Set appliedSortRule = Nothing
        ruleLogDetails = ""
        
        strAttachmentName = objAttachments.Item(i).FileName
        
        Set appliedSortRule = FindAttachmentSortRule(strAttachmentName, sourceFolderPath, objMail.Subject, attachmentSortRules)
        
        If appliedSortRule Is Nothing Then
            Debug.Print "Skipping attachment without matching sort rule: " & strAttachmentName
            On Error GoTo 0
            GoTo NextAttachment
        End If
        
        strFolderPath = RemoveTrailingBackslash(CStr(appliedSortRule("Target")))
        ruleLogDetails = BuildSortRuleLogDetails(appliedSortRule)
        
        If Dir(strFolderPath, vbDirectory) = "" Then
            CreateFolder strFolderPath
        End If
        
        If Err.Number <> 0 Then
            Debug.Print "Failed to create target folder: " & strFolderPath & ". Error: " & Err.Description
            WriteAutoSaveLog strFolderPath, "FailedCreateFolder", Err.Description, objMail, strAttachmentName, "", ruleLogDetails
            Err.Clear
            On Error GoTo 0
            GoTo NextAttachment
        End If
        
        If Dir(strFolderPath, vbDirectory) = "" Then
            Debug.Print "Skipping attachment because target folder is not available: " & strFolderPath
            WriteAutoSaveLog strFolderPath, "SkippedNoTargetFolder", "Target folder is not available", objMail, strAttachmentName, "", ruleLogDetails
            On Error GoTo 0
            GoTo NextAttachment
        End If
        
        autoSaveIndexKey = BuildAutoSaveIndexKey(objMail, strAttachmentName, i)
        
        If AutoSaveIndexContains(strFolderPath, autoSaveIndexKey, indexedFilePath) Then
            Debug.Print "Skipping already indexed attachment: " & strAttachmentName
            WriteAutoSaveLog strFolderPath, "SkippedExists", "Attachment already recorded in index", objMail, strAttachmentName, indexedFilePath, ruleLogDetails
            On Error GoTo 0
            GoTo NextAttachment
        End If
        
        strOriginalFileName = EnsureTrailingBackslash(strFolderPath) & CreateValidName(strAttachmentName)
        strFileName = strOriginalFileName
        logStatus = "Saved"
        logReason = "Attachment saved"
        
        If Dir(strFileName) <> "" Then
            strFileName = GetDatedAutoSaveFilePath(strOriginalFileName, objMail)
            
            If Len(strFileName) = 0 Then
                Debug.Print "Skipping already saved attachment: " & strAttachmentName
                WriteAutoSaveLog strFolderPath, "SkippedExists", "File already exists", objMail, strAttachmentName, strOriginalFileName, ruleLogDetails
                On Error GoTo 0
                GoTo NextAttachment
            End If
            
            If Dir(strFileName) <> "" Then
                Debug.Print "Skipping already saved dated attachment: " & strAttachmentName
                AppendAutoSaveIndex strFolderPath, autoSaveIndexKey, strFileName
                WriteAutoSaveLog strFolderPath, "SkippedExists", "Dated file already exists", objMail, strAttachmentName, strFileName, ruleLogDetails
                On Error GoTo 0
                GoTo NextAttachment
            End If
            
            logStatus = "SavedAsDatedCopy"
            logReason = "Original file name already exists"
        End If
        
        Debug.Print "Trying to save to: " & strFileName
        
        objAttachments.Item(i).SaveAsFile strFileName
        
        If Err.Number <> 0 Then
            Debug.Print "Failed to save attachment to: " & strFileName & ". Error: " & Err.Description
            WriteAutoSaveLog strFolderPath, "FailedSave", Err.Description, objMail, strAttachmentName, strFileName, ruleLogDetails
            Err.Clear
        Else
            savedCount = savedCount + 1
            AppendAutoSaveIndex strFolderPath, autoSaveIndexKey, strFileName
            WriteAutoSaveLog strFolderPath, logStatus, logReason, objMail, strAttachmentName, strFileName, ruleLogDetails
            
            If Not IsMissing(savedFilePathsOverride) Then
                savedFilePathsOverride.Add strFileName
            End If
        End If
        On Error GoTo 0
NextAttachment:
    Next i
    
    SaveAttachmentsFromMailItem = savedCount
End Function

Private Function SaveAllAttachmentsFromMailItemToFolder(ByVal objMail As Outlook.mailItem, ByVal targetFolderPath As String, ByVal savedFilePaths As Collection) As Long
    Dim objAttachments As Outlook.Attachments
    Dim strFileName As String
    Dim strAttachmentName As String
    Dim i As Long
    Dim savedCount As Long
    
    Set objAttachments = objMail.Attachments
    
    If objAttachments.Count = 0 Then
        SaveAllAttachmentsFromMailItemToFolder = 0
        Exit Function
    End If
    
    targetFolderPath = EnsureTrailingBackslash(targetFolderPath)
    
    If Dir(targetFolderPath, vbDirectory) = "" Then
        CreateFolder targetFolderPath
    End If
    
    For i = 1 To objAttachments.Count
        On Error Resume Next
        
        strAttachmentName = objAttachments.Item(i).FileName
        
        If Not IsSelectedSaveAttachmentAllowed(strAttachmentName) Then
            Debug.Print "Skipping selected attachment by selected-save filter: " & strAttachmentName
            On Error GoTo 0
            GoTo NextAttachment
        End If
        
        strFileName = GetUniqueFilePath(targetFolderPath & CreateValidName(strAttachmentName))
        
        If Len(strFileName) = 0 Then
            Debug.Print "Skipping selected attachment because no unique file name is available: " & strAttachmentName
            On Error GoTo 0
            GoTo NextAttachment
        End If
        
        Debug.Print "Trying to save selected attachment to: " & strFileName
        
        objAttachments.Item(i).SaveAsFile strFileName
        
        If Err.Number <> 0 Then
            Debug.Print "Failed to save selected attachment to: " & strFileName & ". Error: " & Err.Description
            Err.Clear
        Else
            savedCount = savedCount + 1
            
            If Not savedFilePaths Is Nothing Then
                savedFilePaths.Add strFileName
            End If
        End If
        
        On Error GoTo 0
NextAttachment:
    Next i
    
    SaveAllAttachmentsFromMailItemToFolder = savedCount
End Function

Private Function IsSelectedSaveAttachmentAllowed(ByVal strFileName As String) As Boolean
    Dim allowedExtensions As String
    Dim excludedExtensions As String
    Dim fileExtension As String
    
    allowedExtensions = GetSelectedSaveAllowedExtensions()
    excludedExtensions = GetSelectedSaveExcludedExtensions()
    fileExtension = GetFileExtension(strFileName)
    
    If Len(fileExtension) > 0 And Len(excludedExtensions) > 0 Then
        If IsExtensionInDelimitedList(fileExtension, excludedExtensions) Then
            IsSelectedSaveAttachmentAllowed = False
            Exit Function
        End If
    End If
    
    If Len(allowedExtensions) > 0 Then
        If Len(fileExtension) = 0 Then
            IsSelectedSaveAttachmentAllowed = False
        Else
            IsSelectedSaveAttachmentAllowed = IsExtensionInDelimitedList(fileExtension, allowedExtensions)
        End If
    Else
        IsSelectedSaveAttachmentAllowed = True
    End If
End Function

Private Function IsExtensionInDelimitedList(ByVal fileExtension As String, ByVal extensionsValue As String) As Boolean
    fileExtension = CleanFileExtension(fileExtension)
    extensionsValue = NormalizeAllowedExtensions(extensionsValue)
    
    If Len(fileExtension) = 0 Or Len(extensionsValue) = 0 Then
        IsExtensionInDelimitedList = False
    Else
        IsExtensionInDelimitedList = (InStr(1, ";" & extensionsValue & ";", ";" & fileExtension & ";", vbTextCompare) > 0)
    End If
End Function

Private Function GetUniqueFilePath(ByVal filePath As String) As String
    Const MAX_UNIQUE_FILE_SUFFIX As Long = 9999
    
    Dim folderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim extensionValue As String
    Dim slashPos As Long
    Dim dotPos As Long
    Dim counter As Long
    Dim candidatePath As String
    
    If Dir(filePath) = "" Then
        GetUniqueFilePath = filePath
        Exit Function
    End If
    
    slashPos = InStrRev(filePath, "\")
    
    If slashPos > 0 Then
        folderPath = Left(filePath, slashPos)
        fileName = Mid(filePath, slashPos + 1)
    Else
        folderPath = ""
        fileName = filePath
    End If
    
    dotPos = InStrRev(fileName, ".")
    
    If dotPos > 1 Then
        baseName = Left(fileName, dotPos - 1)
        extensionValue = Mid(fileName, dotPos)
    Else
        baseName = fileName
        extensionValue = ""
    End If
    
    counter = 2
    
    Do While counter <= MAX_UNIQUE_FILE_SUFFIX
        candidatePath = folderPath & baseName & " (" & CStr(counter) & ")" & extensionValue
        
        If Dir(candidatePath) = "" Then
            GetUniqueFilePath = candidatePath
            Exit Function
        End If
        
        counter = counter + 1
    Loop
    
    GetUniqueFilePath = ""
End Function

Private Function GetDatedAutoSaveFilePath(ByVal filePath As String, ByVal objMail As Outlook.mailItem) As String
    Dim datedFilePath As String
    
    datedFilePath = BuildDatedFilePath(filePath, GetMailDateForFileName(objMail))
    
    If Len(datedFilePath) = 0 Then
        GetDatedAutoSaveFilePath = ""
    Else
        GetDatedAutoSaveFilePath = datedFilePath
    End If
End Function

Private Function BuildAutoSaveIndexKey(ByVal objMail As Outlook.mailItem, ByVal attachmentName As String, ByVal attachmentIndex As Long) As String
    On Error Resume Next
    
    Dim mailKey As String
    
    If Not objMail Is Nothing Then
        mailKey = Format(objMail.ReceivedTime, "yyyy-mm-dd hh:nn:ss") & "|" & objMail.Subject
    End If
    
    BuildAutoSaveIndexKey = CleanAutoSaveIndexField(mailKey) & "|" & CStr(attachmentIndex) & "|" & CleanAutoSaveIndexField(attachmentName)
End Function

Private Function CleanAutoSaveIndexField(ByVal value As String) As String
    value = Replace(value, vbCrLf, " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, "|", "/")
    
    CleanAutoSaveIndexField = value
End Function

Private Function AutoSaveIndexContains(ByVal targetFolderPath As String, ByVal indexKey As String, ByRef savedFilePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim indexFilePath As String
    Dim fileNumber As Integer
    Dim lineValue As String
    Dim separatorPos As Long
    
    savedFilePath = ""
    
    If Len(indexKey) = 0 Or Len(targetFolderPath) = 0 Then
        Exit Function
    End If
    
    indexFilePath = EnsureTrailingBackslash(targetFolderPath) & AUTO_SAVE_INDEX_FILE_NAME
    
    If Not FileExistsIncludingHidden(indexFilePath) Then
        Exit Function
    End If
    
    fileNumber = FreeFile
    Open indexFilePath For Input As #fileNumber
    
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineValue
        separatorPos = InStr(1, lineValue, vbTab)
        
        If separatorPos > 0 Then
            If Left(lineValue, separatorPos - 1) = indexKey Then
                savedFilePath = Mid(lineValue, separatorPos + 1)
                AutoSaveIndexContains = True
                Exit Do
            End If
        End If
    Loop
    
    Close #fileNumber
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If fileNumber <> 0 Then
        Close #fileNumber
    End If
    AutoSaveIndexContains = False
End Function

Private Function FileExistsIncludingHidden(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Len(filePath) = 0 Then
        FileExistsIncludingHidden = False
    Else
        FileExistsIncludingHidden = (Len(Dir(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
    End If
    
    Exit Function
    
ErrorHandler:
    FileExistsIncludingHidden = False
End Function

Private Sub AppendAutoSaveIndex(ByVal targetFolderPath As String, ByVal indexKey As String, ByVal savedFilePath As String)
    On Error GoTo ErrorHandler
    
    Dim indexFilePath As String
    Dim existingPath As String
    Dim fileNumber As Integer
    
    If Len(indexKey) = 0 Or Len(targetFolderPath) = 0 Then
        Exit Sub
    End If
    
    If Dir(targetFolderPath, vbDirectory) = "" Then
        Exit Sub
    End If
    
    If AutoSaveIndexContains(targetFolderPath, indexKey, existingPath) Then
        Exit Sub
    End If
    
    indexFilePath = EnsureTrailingBackslash(targetFolderPath) & AUTO_SAVE_INDEX_FILE_NAME
    fileNumber = FreeFile
    
    Open indexFilePath For Append As #fileNumber
    Print #fileNumber, indexKey & vbTab & savedFilePath
    Close #fileNumber
    
    On Error Resume Next
    SetAttr indexFilePath, vbHidden
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If fileNumber <> 0 Then
        Close #fileNumber
    End If
End Sub

Private Function BuildDatedFilePath(ByVal filePath As String, ByVal mailDate As Date) As String
    Dim folderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim extensionValue As String
    Dim slashPos As Long
    Dim dotPos As Long
    
    slashPos = InStrRev(filePath, "\")
    
    If slashPos > 0 Then
        folderPath = Left(filePath, slashPos)
        fileName = Mid(filePath, slashPos + 1)
    Else
        folderPath = ""
        fileName = filePath
    End If
    
    dotPos = InStrRev(fileName, ".")
    
    If dotPos > 1 Then
        baseName = Left(fileName, dotPos - 1)
        extensionValue = Mid(fileName, dotPos)
    Else
        baseName = fileName
        extensionValue = ""
    End If
    
    BuildDatedFilePath = folderPath & baseName & "_" & Format(mailDate, "yyyy-mm-dd_hh-nn-ss") & extensionValue
End Function

Private Function GetMailDateForFileName(ByVal objMail As Outlook.mailItem) As Date
    On Error Resume Next
    
    If Not objMail Is Nothing Then
        If Year(objMail.ReceivedTime) > 1900 Then
            GetMailDateForFileName = objMail.ReceivedTime
            Exit Function
        End If
    End If
    
    GetMailDateForFileName = Now
End Function

Function CreateValidName(ByVal strName As String) As String
    Dim strValidName As String
    Dim ext As String
    Dim pos As Integer
    
    strValidName = strName
    
    ' Find the last period (file extension separator)
    pos = InStrRev(strValidName, ".")
    
    ' Extract the extension and the name separately if a period is found
    If pos > 0 Then
        ext = Mid(strValidName, pos)
        strValidName = Left(strValidName, pos - 1)
    Else
        ext = ""
    End If
    
    ' Remove trailing spaces from the name part
    strValidName = RemoveTrailingSpaces(strValidName)
    
    ' Replace invalid characters in the name part
    strValidName = ReplaceInvalidCharacters(strValidName)
    
    ' Combine the sanitized name with the extension
    CreateValidName = strValidName & ext
End Function

Function IsAttachmentExtensionAllowed(ByVal strFileName As String, Optional ByVal allowedExtensionsOverride As Variant) As Boolean
    Dim allowedExtensions As String
    Dim fileExtension As String
    
    If IsMissing(allowedExtensionsOverride) Then
        allowedExtensions = GetAllowedAttachmentExtensions()
    Else
        allowedExtensions = CStr(allowedExtensionsOverride)
    End If
    
    If Len(allowedExtensions) = 0 Then
        IsAttachmentExtensionAllowed = True
        Exit Function
    End If
    
    fileExtension = GetFileExtension(strFileName)
    
    If Len(fileExtension) = 0 Then
        IsAttachmentExtensionAllowed = False
    Else
        IsAttachmentExtensionAllowed = (InStr(1, ";" & allowedExtensions & ";", ";" & fileExtension & ";", vbTextCompare) > 0)
    End If
End Function

Function GetAllowedAttachmentExtensions() As String
    Dim rules As String
    
    rules = GetFileTypeRules()
    
    If Len(Trim(rules)) > 0 Then
        GetAllowedAttachmentExtensions = GetActiveExtensionsFromRules(rules)
    Else
        GetAllowedAttachmentExtensions = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_ALLOWED_EXTENSIONS_KEY, "")
    End If
End Function

Function GetFileTypeRules() As String
    GetFileTypeRules = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_FILE_TYPE_RULES_KEY, "")
End Function

Sub SaveFileTypeRules(ByVal rules As String)
    Dim normalizedRules As String
    
    normalizedRules = NormalizeFileTypeRules(rules)
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_FILE_TYPE_RULES_KEY, normalizedRules
    SaveSetting SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_ALLOWED_EXTENSIONS_KEY, GetActiveExtensionsFromRules(normalizedRules)
End Sub

Function BuildFileTypeRulesFromAllowedExtensions(ByVal allowedExtensions As String) As String
    Dim parts() As String
    Dim part As Variant
    Dim result As String
    
    allowedExtensions = NormalizeAllowedExtensions(allowedExtensions)
    
    If Len(allowedExtensions) = 0 Then
        BuildFileTypeRulesFromAllowedExtensions = ""
        Exit Function
    End If
    
    parts = Split(allowedExtensions, ";")
    
    For Each part In parts
        If Len(Trim(CStr(part))) > 0 Then
            result = result & "1|" & Trim(CStr(part)) & vbLf
        End If
    Next part
    
    BuildFileTypeRulesFromAllowedExtensions = result
End Function

Function NormalizeFileTypeRules(ByVal rules As String) As String
    Dim lines() As String
    Dim line As Variant
    Dim parts() As String
    Dim enabledValue As String
    Dim extensionValue As String
    Dim result As String
    Dim seenExtensions As Object
    
    Set seenExtensions = CreateObject("Scripting.Dictionary")
    seenExtensions.CompareMode = vbTextCompare
    
    rules = Replace(rules, vbCrLf, vbLf)
    rules = Replace(rules, vbCr, vbLf)
    
    If Len(Trim(rules)) = 0 Then
        NormalizeFileTypeRules = ""
        Exit Function
    End If
    
    lines = Split(rules, vbLf)
    
    For Each line In lines
        If Len(Trim(CStr(line))) > 0 Then
            parts = Split(CStr(line), "|")
            
            If UBound(parts) >= 1 Then
                enabledValue = Trim(parts(0))
                extensionValue = CleanFileExtension(parts(1))
                
                If Len(extensionValue) > 0 And Not seenExtensions.Exists(extensionValue) Then
                    seenExtensions(extensionValue) = True
                    
                    If enabledValue <> "0" Then
                        enabledValue = "1"
                    End If
                    
                    result = result & enabledValue & "|" & extensionValue & vbLf
                End If
            End If
        End If
    Next line
    
    NormalizeFileTypeRules = result
End Function

Function GetActiveExtensionsFromRules(ByVal rules As String) As String
    Dim lines() As String
    Dim line As Variant
    Dim parts() As String
    Dim extensionValue As String
    Dim result As String
    
    rules = NormalizeFileTypeRules(rules)
    
    If Len(Trim(rules)) = 0 Then
        GetActiveExtensionsFromRules = ""
        Exit Function
    End If
    
    lines = Split(rules, vbLf)
    
    For Each line In lines
        If Len(Trim(CStr(line))) > 0 Then
            parts = Split(CStr(line), "|")
            
            If UBound(parts) >= 1 And Trim(parts(0)) = "1" Then
                extensionValue = CleanFileExtension(parts(1))
                
                If Len(extensionValue) > 0 Then
                    If Len(result) > 0 Then
                        result = result & ";"
                    End If
                    
                    result = result & extensionValue
                End If
            End If
        End If
    Next line
    
    GetActiveExtensionsFromRules = result
End Function

Function GetActiveFileTypeStatusText() As String
    Dim activeExtensions As String
    Dim activeCount As Long
    
    activeExtensions = GetAllowedAttachmentExtensions()
    
    If Len(Trim(activeExtensions)) = 0 Then
        GetActiveFileTypeStatusText = RuText(1042, 1089, 1077, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1103)
        Exit Function
    End If
    
    activeCount = CountDelimitedValues(activeExtensions, ";")
    GetActiveFileTypeStatusText = CStr(activeCount) & " (" & FormatAllowedExtensionsForInput(activeExtensions) & ")"
End Function

Function CountDelimitedValues(ByVal value As String, ByVal delimiter As String) As Long
    Dim parts() As String
    Dim part As Variant
    Dim result As Long
    
    If Len(Trim(value)) = 0 Then
        CountDelimitedValues = 0
        Exit Function
    End If
    
    parts = Split(value, delimiter)
    
    For Each part In parts
        If Len(Trim(CStr(part))) > 0 Then
            result = result + 1
        End If
    Next part
    
    CountDelimitedValues = result
End Function

Function CleanFileExtension(ByVal extensionValue As String) As String
    extensionValue = LCase(Trim(extensionValue))
    extensionValue = Replace(extensionValue, ".", "")
    extensionValue = Replace(extensionValue, " ", "")
    CleanFileExtension = ReplaceInvalidCharacters(extensionValue)
End Function

Function GetAttachmentSortRules() As String
    GetAttachmentSortRules = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_SORT_RULES_KEY, "")
End Function

Function NormalizeAttachmentSortRules(ByVal strRules As String) As String
    Dim lines() As String
    Dim line As Variant
    Dim cleanedLine As String
    Dim pos As Long
    Dim patternValue As String
    Dim targetValue As String
    Dim result As String
    Dim seenPatterns As Object
    
    Set seenPatterns = CreateObject("Scripting.Dictionary")
    seenPatterns.CompareMode = vbTextCompare
    
    strRules = Replace(strRules, vbCrLf, vbLf)
    strRules = Replace(strRules, vbCr, vbLf)
    strRules = Replace(strRules, ";", vbLf)
    
    If Len(Trim(strRules)) = 0 Then
        NormalizeAttachmentSortRules = ""
        Exit Function
    End If
    
    lines = Split(strRules, vbLf)
    
    For Each line In lines
        cleanedLine = Trim(CStr(line))
        pos = InStr(1, cleanedLine, "=", vbBinaryCompare)
        
        If pos > 1 Then
            patternValue = Trim(Left(cleanedLine, pos - 1))
            targetValue = Trim(Mid(cleanedLine, pos + 1))
            targetValue = TrimQuotes(targetValue)
            targetValue = Replace(targetValue, "/", "\")
            
            If Len(patternValue) > 0 And Len(targetValue) > 0 Then
                If Not seenPatterns.Exists(patternValue) Then
                    seenPatterns.Add patternValue, True
                    
                    If Len(result) > 0 Then
                        result = result & vbLf
                    End If
                    
                    result = result & patternValue & "=" & targetValue
                End If
            End If
        End If
    Next line
    
    NormalizeAttachmentSortRules = result
End Function

Function FormatAttachmentSortRulesForInput(ByVal strRules As String) As String
    FormatAttachmentSortRulesForInput = Replace(strRules, vbLf, "; ")
End Function

Function CountAttachmentSortRules(ByVal strRules As String) As Long
    Dim lines() As String
    Dim i As Long
    Dim result As Long
    
    If Len(Trim(strRules)) = 0 Then
        CountAttachmentSortRules = 0
        Exit Function
    End If
    
    lines = Split(strRules, vbLf)
    
    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            result = result + 1
        End If
    Next i
    
    CountAttachmentSortRules = result
End Function

Function CountActiveSortRuleRows(ByVal rowRules As String) As Long
    Dim lines() As String
    Dim i As Long
    Dim parts() As String
    Dim result As Long
    Dim activeRules As Collection
    
    rowRules = Replace(rowRules, vbCrLf, vbLf)
    rowRules = Replace(rowRules, vbCr, vbLf)
    
    If Len(Trim(rowRules)) = 0 Then
        Set activeRules = LoadAttachmentSortRules()
        CountActiveSortRuleRows = activeRules.Count
        Exit Function
    End If
    
    rowRules = SR_NormalizeSortRuleRows(rowRules)
    
    If Len(Trim(rowRules)) = 0 Then
        CountActiveSortRuleRows = 0
        Exit Function
    End If
    
    lines = Split(rowRules, vbLf)
    
    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            parts = Split(lines(i), "|")
            
            If UBound(parts) >= 2 And Trim(parts(0)) = "1" Then
                result = result + 1
            End If
        End If
    Next i
    
    CountActiveSortRuleRows = result
End Function

Function GetActiveSortRuleSummaryText() As String
    Const MAX_RULES_IN_STATUS As Long = 5
    
    Dim sortRules As Collection
    Dim sortRule As Variant
    Dim result As String
    Dim displayedCount As Long
    Dim totalCount As Long
    
    Set sortRules = LoadAttachmentSortRules()
    
    totalCount = sortRules.Count
    
    If totalCount = 0 Then
        GetActiveSortRuleSummaryText = ""
        Exit Function
    End If
    
    For Each sortRule In sortRules
        displayedCount = displayedCount + 1
        
        If Len(result) > 0 Then
            result = result & "; "
        End If
        
        result = result & BuildSortRuleStatusLabel(sortRule)
        
        If displayedCount >= MAX_RULES_IN_STATUS Then
            Exit For
        End If
    Next sortRule
    
    If totalCount > displayedCount Then
        result = result & "; " & RuText(1077, 1097, 1077, 32) & CStr(totalCount - displayedCount)
    End If
    
    GetActiveSortRuleSummaryText = RuText(1040, 1082, 1090, 1080, 1074, 1085, 1099, 1077, 32, 1087, 1088, 1072, 1074, 1080, 1083, 1072, 58, 32) & result
End Function

Private Function BuildSortRuleStatusLabel(ByVal sortRule As Object) As String
    On Error GoTo ErrorHandler
    
    Dim ruleNumber As String
    Dim labelText As String
    
    ruleNumber = CStr(sortRule("Number"))
    labelText = Trim(CStr(sortRule("Name")))
    
    If Len(labelText) = 0 Then
        labelText = BuildUnnamedSortRuleStatusLabel(sortRule)
    End If
    
    BuildSortRuleStatusLabel = ruleNumber & ". " & ShortenStatusText(labelText, 36)
    Exit Function
    
ErrorHandler:
    BuildSortRuleStatusLabel = RuText(1041, 1077, 1079, 32, 1080, 1084, 1077, 1085, 1080)
End Function

Private Function BuildUnnamedSortRuleStatusLabel(ByVal sortRule As Object) As String
    Dim patternValue As String
    Dim extensionsValue As String
    Dim subjectValue As String
    
    patternValue = Trim(CStr(sortRule("Pattern")))
    extensionsValue = Trim(CStr(sortRule("Extensions")))
    subjectValue = Trim(CStr(sortRule("Subject")))
    
    If Len(patternValue) > 0 Then
        BuildUnnamedSortRuleStatusLabel = patternValue
    ElseIf Len(extensionsValue) > 0 Then
        BuildUnnamedSortRuleStatusLabel = "[" & extensionsValue & "]"
    ElseIf Len(subjectValue) > 0 Then
        BuildUnnamedSortRuleStatusLabel = subjectValue
    Else
        BuildUnnamedSortRuleStatusLabel = RuText(1042, 1089, 1077, 32, 1092, 1072, 1081, 1083, 1099)
    End If
End Function

Private Function ShortenStatusText(ByVal value As String, ByVal maxLength As Long) As String
    value = Trim(value)
    
    If maxLength <= 0 Or Len(value) <= maxLength Then
        ShortenStatusText = value
    Else
        ShortenStatusText = Left(value, maxLength - 3) & "..."
    End If
End Function

Function CountSourceLimitedSortRuleRows(ByVal rowRules As String) As Long
    Dim lines() As String
    Dim i As Long
    Dim parts() As String
    Dim result As Long
    
    rowRules = Replace(rowRules, vbCrLf, vbLf)
    rowRules = Replace(rowRules, vbCr, vbLf)
    
    If Len(Trim(rowRules)) = 0 Then
        CountSourceLimitedSortRuleRows = 0
        Exit Function
    End If
    
    rowRules = SR_NormalizeSortRuleRows(rowRules)
    
    If Len(Trim(rowRules)) = 0 Then
        CountSourceLimitedSortRuleRows = 0
        Exit Function
    End If
    
    lines = Split(rowRules, vbLf)
    
    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            parts = Split(lines(i), "|")
            
            If UBound(parts) >= 4 And Trim(parts(0)) = "1" And Len(Trim(parts(4))) > 0 Then
                result = result + 1
            End If
        End If
    Next i
    
    CountSourceLimitedSortRuleRows = result
End Function

Function GetAttachmentTargetFolder(ByVal strBaseFolderPath As String, ByVal strDefaultFolderName As String, ByVal strAttachmentName As String, Optional ByVal sourceFolderPath As String = "", Optional ByVal mailSubject As String = "", Optional ByVal attachmentSortRulesOverride As Variant) As String
    Dim targetValue As String
    
    If IsMissing(attachmentSortRulesOverride) Then
        targetValue = FindAttachmentSortTarget(strAttachmentName, sourceFolderPath, mailSubject)
    Else
        targetValue = FindAttachmentSortTarget(strAttachmentName, sourceFolderPath, mailSubject, attachmentSortRulesOverride)
    End If
    
    If Len(targetValue) = 0 Then
        GetAttachmentTargetFolder = ""
    ElseIf IsAbsoluteFolderPath(targetValue) Then
        GetAttachmentTargetFolder = RemoveTrailingBackslash(targetValue)
    Else
        GetAttachmentTargetFolder = ""
    End If
End Function

Function LoadAttachmentSortRules() As Collection
    Dim sortRules As Collection
    Dim rowRules As String
    Dim rules As String
    Dim lines() As String
    Dim i As Long
    Dim pos As Long
    Dim parts() As String
    Dim patternValue As String
    Dim targetValue As String
    Dim extensionsValue As String
    Dim ruleSourceFolderValue As String
    Dim subjectValue As String
    Dim displayNameValue As String
    
    Set sortRules = New Collection
    rowRules = GetSetting(SETTINGS_APP_NAME, SETTINGS_SECTION, SETTINGS_SORT_RULE_ROWS_KEY, "")
    
    If Len(Trim(rowRules)) > 0 Then
        rowRules = Replace(rowRules, vbCrLf, vbLf)
        rowRules = Replace(rowRules, vbCr, vbLf)
        lines = Split(rowRules, vbLf)
        
        For i = LBound(lines) To UBound(lines)
            If Len(Trim(lines(i))) > 0 Then
                parts = Split(lines(i), "|")
                
                If UBound(parts) >= 2 And Trim(parts(0)) = "1" Then
                    patternValue = Trim(parts(1))
                    targetValue = Trim(parts(2))
                    
                    If UBound(parts) >= 3 Then
                        extensionsValue = Trim(parts(3))
                    Else
                        extensionsValue = ""
                    End If
                    
                    If UBound(parts) >= 4 Then
                        ruleSourceFolderValue = Trim(parts(4))
                    Else
                        ruleSourceFolderValue = ""
                    End If
                    
                    If UBound(parts) >= 5 Then
                        subjectValue = Trim(parts(5))
                    Else
                        subjectValue = ""
                    End If
                    
                    If UBound(parts) >= 6 Then
                        displayNameValue = Trim(parts(6))
                    Else
                        displayNameValue = ""
                    End If
                    
                    AddAttachmentSortRule sortRules, patternValue, targetValue, extensionsValue, ruleSourceFolderValue, subjectValue, i + 1, displayNameValue
                End If
            End If
        Next i
        
        Set LoadAttachmentSortRules = sortRules
        Exit Function
    End If
    
    rules = GetAttachmentSortRules()
    
    If Len(Trim(rules)) > 0 Then
        lines = Split(rules, vbLf)
        
        For i = LBound(lines) To UBound(lines)
            pos = InStr(1, lines(i), "=", vbBinaryCompare)
            
            If pos > 1 Then
                patternValue = Trim(Left(lines(i), pos - 1))
                targetValue = Trim(Mid(lines(i), pos + 1))
                AddAttachmentSortRule sortRules, patternValue, targetValue, "", "", "", sortRules.Count + 1, ""
            End If
        Next i
    End If
    
    Set LoadAttachmentSortRules = sortRules
End Function

Private Sub AddAttachmentSortRule(ByRef sortRules As Collection, ByVal patternValue As String, ByVal targetValue As String, ByVal extensionsValue As String, ByVal ruleSourceFolderValue As String, ByVal subjectValue As String, Optional ByVal ruleNumber As Long = 0, Optional ByVal displayNameValue As String = "")
    Dim sortRule As Object
    
    If Len(targetValue) = 0 Then
        Exit Sub
    End If
    
    If Not IsAbsoluteFolderPath(targetValue) Then
        Exit Sub
    End If
    
    If Len(patternValue) = 0 And (Len(extensionsValue) = 0 Or Len(ruleSourceFolderValue) = 0) Then
        Exit Sub
    End If
    
    Set sortRule = CreateObject("Scripting.Dictionary")
    sortRule("Pattern") = patternValue
    sortRule("Target") = targetValue
    sortRule("Extensions") = extensionsValue
    sortRule("SourceFolder") = ruleSourceFolderValue
    sortRule("Subject") = subjectValue
    sortRule("Number") = ruleNumber
    sortRule("Name") = displayNameValue
    sortRules.Add sortRule
End Sub

Function CombineRelativeFolderPath(ByVal strBaseFolderPath As String, ByVal strRelativeFolderPath As String) As String
    Dim parts() As String
    Dim part As Variant
    Dim cleanedPart As String
    Dim result As String
    
    result = EnsureTrailingBackslash(strBaseFolderPath)
    strRelativeFolderPath = Replace(strRelativeFolderPath, "/", "\")
    parts = Split(strRelativeFolderPath, "\")
    
    For Each part In parts
        cleanedPart = CreateValidFolderName(CStr(part))
        
        If Len(cleanedPart) > 0 Then
            If Right(result, 1) <> "\" Then
                result = result & "\"
            End If
            
            result = result & cleanedPart
        End If
    Next part
    
    CombineRelativeFolderPath = RemoveTrailingBackslash(result)
End Function

Function FindAttachmentSortTarget(ByVal strAttachmentName As String, Optional ByVal sourceFolderPath As String = "", Optional ByVal mailSubject As String = "", Optional ByVal attachmentSortRulesOverride As Variant) As String
    Dim sortRule As Object
    
    If IsMissing(attachmentSortRulesOverride) Then
        Set sortRule = FindAttachmentSortRule(strAttachmentName, sourceFolderPath, mailSubject)
    Else
        Set sortRule = FindAttachmentSortRule(strAttachmentName, sourceFolderPath, mailSubject, attachmentSortRulesOverride)
    End If
    
    If sortRule Is Nothing Then
        FindAttachmentSortTarget = ""
    Else
        FindAttachmentSortTarget = CStr(sortRule("Target"))
    End If
End Function

Private Function FindAttachmentSortRule(ByVal strAttachmentName As String, Optional ByVal sourceFolderPath As String = "", Optional ByVal mailSubject As String = "", Optional ByVal attachmentSortRulesOverride As Variant) As Object
    Dim attachmentSortRules As Collection
    Dim sortRule As Variant
    Dim patternValue As String
    Dim targetValue As String
    Dim extensionsValue As String
    Dim ruleSourceFolderValue As String
    Dim subjectValue As String
    
    If IsMissing(attachmentSortRulesOverride) Then
        Set attachmentSortRules = LoadAttachmentSortRules()
    Else
        Set attachmentSortRules = attachmentSortRulesOverride
    End If
    
    For Each sortRule In attachmentSortRules
        patternValue = CStr(sortRule("Pattern"))
        targetValue = CStr(sortRule("Target"))
        extensionsValue = CStr(sortRule("Extensions"))
        ruleSourceFolderValue = CStr(sortRule("SourceFolder"))
        
        If sortRule.Exists("Subject") Then
            subjectValue = CStr(sortRule("Subject"))
        Else
            subjectValue = ""
        End If
        
        If Len(patternValue) = 0 Or InStr(1, strAttachmentName, patternValue, vbTextCompare) > 0 Then
            If RuleExtensionsMatch(strAttachmentName, extensionsValue) And RuleSourceFolderMatches(sourceFolderPath, ruleSourceFolderValue) And RuleSubjectMatches(mailSubject, subjectValue) Then
                Set FindAttachmentSortRule = sortRule
                Exit Function
            End If
        End If
    Next sortRule
    
    Set FindAttachmentSortRule = Nothing
End Function

Private Function BuildSortRuleLogDetails(ByVal sortRule As Object) As String
    On Error GoTo ErrorHandler
    
    Dim result As String
    
    If sortRule Is Nothing Then
        BuildSortRuleLogDetails = ""
        Exit Function
    End If
    
    result = "RuleNo=" & CStr(sortRule("Number"))
    result = result & "; RuleName=""" & CStr(sortRule("Name")) & """"
    result = result & "; FileText=""" & CStr(sortRule("Pattern")) & """"
    result = result & "; Extensions=""" & CStr(sortRule("Extensions")) & """"
    result = result & "; Subject=""" & CStr(sortRule("Subject")) & """"
    result = result & "; SourceFolder=""" & CStr(sortRule("SourceFolder")) & """"
    result = result & "; Target=""" & CStr(sortRule("Target")) & """"
    
    BuildSortRuleLogDetails = result
    Exit Function
    
ErrorHandler:
    BuildSortRuleLogDetails = ""
End Function

Function RuleSubjectMatches(ByVal mailSubject As String, ByVal ruleSubject As String) As Boolean
    ruleSubject = Trim(ruleSubject)
    
    If Len(ruleSubject) = 0 Then
        RuleSubjectMatches = True
    Else
        RuleSubjectMatches = (InStr(1, mailSubject, ruleSubject, vbTextCompare) > 0)
    End If
End Function

Function RuleExtensionsMatch(ByVal strAttachmentName As String, ByVal extensionsValue As String) As Boolean
    Dim fileExtension As String
    
    extensionsValue = LCase(Replace(Trim(extensionsValue), " ", ""))
    extensionsValue = Replace(extensionsValue, ".", "")
    extensionsValue = Replace(extensionsValue, ";", ",")
    
    If Len(extensionsValue) = 0 Then
        RuleExtensionsMatch = True
        Exit Function
    End If
    
    fileExtension = GetFileExtension(strAttachmentName)
    
    If Len(fileExtension) = 0 Then
        RuleExtensionsMatch = False
    Else
        RuleExtensionsMatch = (InStr(1, "," & extensionsValue & ",", "," & fileExtension & ",", vbTextCompare) > 0)
    End If
End Function

Function RuleSourceFolderMatches(ByVal mailFolderPath As String, ByVal ruleFolderPath As String) As Boolean
    mailFolderPath = NormalizeOutlookFolderPath(mailFolderPath)
    ruleFolderPath = NormalizeOutlookFolderPath(ruleFolderPath)
    
    If Len(ruleFolderPath) = 0 Then
        RuleSourceFolderMatches = True
    Else
        RuleSourceFolderMatches = (StrComp(mailFolderPath, ruleFolderPath, vbTextCompare) = 0)
    End If
End Function

Function GetMailItemFolderPath(ByVal objMail As Outlook.mailItem) As String
    On Error GoTo ErrorHandler
    
    If TypeOf objMail.Parent Is Outlook.MAPIFolder Then
        GetMailItemFolderPath = objMail.Parent.FolderPath
    Else
        GetMailItemFolderPath = ""
    End If
    
    Exit Function
    
ErrorHandler:
    GetMailItemFolderPath = ""
End Function

Function NormalizeOutlookFolderPath(ByVal folderPath As String) As String
    folderPath = Trim(folderPath)
    folderPath = Replace(folderPath, "/", "\")
    
    Do While Len(folderPath) > 0 And Right(folderPath, 1) = "\"
        folderPath = Left(folderPath, Len(folderPath) - 1)
    Loop
    
    NormalizeOutlookFolderPath = folderPath
End Function

Function IsAbsoluteFolderPath(ByVal strFolderPath As String) As Boolean
    strFolderPath = Trim(strFolderPath)
    
    If Len(strFolderPath) >= 3 Then
        If Mid(strFolderPath, 2, 2) = ":\" Or Mid(strFolderPath, 2, 2) = ":/" Then
            IsAbsoluteFolderPath = True
            Exit Function
        End If
    End If
    
    IsAbsoluteFolderPath = (Left(strFolderPath, 2) = "\\")
End Function

Function NormalizeAllowedExtensions(ByVal strExtensions As String) As String
    Dim parts() As String
    Dim part As Variant
    Dim cleanedPart As String
    Dim result As String
    
    strExtensions = Replace(strExtensions, ";", ",")
    strExtensions = Replace(strExtensions, ".", "")
    strExtensions = Replace(strExtensions, " ", "")
    strExtensions = LCase(strExtensions)
    
    If Len(strExtensions) = 0 Then
        NormalizeAllowedExtensions = ""
        Exit Function
    End If
    
    parts = Split(strExtensions, ",")
    
    For Each part In parts
        cleanedPart = ReplaceInvalidCharacters(CStr(part))
        
        If Len(cleanedPart) > 0 Then
            If InStr(1, ";" & result & ";", ";" & cleanedPart & ";", vbTextCompare) = 0 Then
                If Len(result) > 0 Then
                    result = result & ";"
                End If
                
                result = result & cleanedPart
            End If
        End If
    Next part
    
    NormalizeAllowedExtensions = result
End Function

Function FormatAllowedExtensionsForInput(ByVal strExtensions As String) As String
    FormatAllowedExtensionsForInput = Replace(strExtensions, ";", ", ")
End Function

Function FormatAllowedExtensionsForStatus(ByVal strExtensions As String) As String
    If Len(Trim(strExtensions)) = 0 Then
        FormatAllowedExtensionsForStatus = RuText(1042, 1089, 1077, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1103)
    Else
        FormatAllowedExtensionsForStatus = FormatAllowedExtensionsForInput(strExtensions)
    End If
End Function

Function GetFileExtension(ByVal strFileName As String) As String
    Dim pos As Long
    
    pos = InStrRev(strFileName, ".")
    
    If pos > 0 And pos < Len(strFileName) Then
        GetFileExtension = LCase(Mid(strFileName, pos + 1))
    Else
        GetFileExtension = ""
    End If
End Function

Function CreateValidFolderName(ByVal strName As String) As String
    Dim strValidName As String
    
    strValidName = NormalizeSpaces(strName)
    strValidName = ReplaceInvalidCharacters(strValidName)
    strValidName = RemoveTrailingSpaces(strValidName)
    
    If Len(strValidName) = 0 Then
        strValidName = RuText(1041, 1077, 1079, 32, 1090, 1077, 1084, 1099)
    End If
    
    CreateValidFolderName = strValidName
End Function

Function NormalizeSpaces(ByVal str As String) As String
    str = Replace(str, "&#8201;", " ")
    str = Replace(str, ChrW(8201), " ")
    
    NormalizeSpaces = str
End Function

Function RemoveTrailingSpaces(ByVal str As String) As String
    Dim i As Integer
    For i = Len(str) To 1 Step -1
        If Mid(str, i, 1) <> " " Then
            RemoveTrailingSpaces = Left(str, i)
            Exit Function
        End If
    Next i
    RemoveTrailingSpaces = ""
End Function

Function TrimQuotes(ByVal str As String) As String
    str = Trim(str)
    
    If Len(str) >= 2 Then
        If Left(str, 1) = """" And Right(str, 1) = """" Then
            str = Mid(str, 2, Len(str) - 2)
        End If
    End If
    
    TrimQuotes = Trim(str)
End Function

Function ReplaceInvalidCharacters(ByVal str As String) As String
    ' Replace invalid characters in the name part
    str = Replace(str, "\", "")
    str = Replace(str, "/", "")
    str = Replace(str, ":", "")
    str = Replace(str, "*", "")
    str = Replace(str, "?", "")
    str = Replace(str, """", "")
    str = Replace(str, "<", "")
    str = Replace(str, ">", "")
    str = Replace(str, "|", "")
    ' Retaining #, ., (, and )
    ' str = Replace(str, "#", "")
    ' str = Replace(str, ".", "")
    ' str = Replace(str, "(", "")
    ' str = Replace(str, ")", "")
    
    ReplaceInvalidCharacters = str
End Function

Private Sub CreateFolder(ByVal strFolderPath As String)
    Dim strParentPath As String
    Dim strLastFolder As String
    Dim pos As Long
    
    strFolderPath = RemoveTrailingBackslash(strFolderPath)
    
    If Len(strFolderPath) = 0 Then Exit Sub
    If Len(strFolderPath) <= 3 And Mid(strFolderPath, 2, 2) = ":\" Then Exit Sub
    
    ' Check if folder already exists
    If Dir(strFolderPath, vbDirectory) <> "" Then Exit Sub
    
    ' Find the parent folder path and the last folder name in the path
    pos = InStrRev(strFolderPath, "\")
    If pos = 0 Then Exit Sub
    
    strParentPath = Left(strFolderPath, pos - 1)
    strLastFolder = Mid(strFolderPath, pos + 1)
    
    If Right(strParentPath, 1) = ":" Then
        strParentPath = strParentPath & "\"
    End If
    
    If Len(strLastFolder) = 0 Then Exit Sub
    
    ' Recursively create parent folders if they don't exist
    If Dir(strParentPath, vbDirectory) = "" Then
        CreateFolder strParentPath
    End If
    
    ' Create the folder
    MkDir strFolderPath
End Sub


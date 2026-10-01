Attribute VB_Name = "SortRulesSettingsCore"

Private Const SR_SETTINGS_APP_NAME As String = "OutlookAttachmentSaver"
Private Const SR_SETTINGS_SECTION As String = "Settings"
Private Const SR_SORT_RULE_ROWS_KEY As String = "AttachmentSortRuleRows"
Private Const SR_LEGACY_SORT_RULES_KEY As String = "AttachmentSortRules"

Function SR_GetSortRuleRows() As String
    SR_GetSortRuleRows = GetSetting(SR_SETTINGS_APP_NAME, SR_SETTINGS_SECTION, SR_SORT_RULE_ROWS_KEY, "")
End Function

Function SR_GetLegacySortRules() As String
    SR_GetLegacySortRules = GetSetting(SR_SETTINGS_APP_NAME, SR_SETTINGS_SECTION, SR_LEGACY_SORT_RULES_KEY, "")
End Function

Sub SR_SaveSortRuleRows(ByVal rows As String)
    Dim normalizedRows As String
    
    normalizedRows = SR_NormalizeSortRuleRows(rows)
    SaveSetting SR_SETTINGS_APP_NAME, SR_SETTINGS_SECTION, SR_SORT_RULE_ROWS_KEY, normalizedRows
    SaveSetting SR_SETTINGS_APP_NAME, SR_SETTINGS_SECTION, SR_LEGACY_SORT_RULES_KEY, SR_GetActiveLegacyRulesFromRows(normalizedRows)
End Sub

Function SR_BuildRowsFromLegacyRules(ByVal legacyRules As String) As String
    Dim lines() As String
    Dim line As Variant
    Dim cleanedLine As String
    Dim pos As Long
    Dim patternValue As String
    Dim targetValue As String
    Dim extensionsValue As String
    Dim sourceFolderValue As String
    Dim subjectValue As String
    Dim result As String
    
    legacyRules = Replace(legacyRules, vbCrLf, vbLf)
    legacyRules = Replace(legacyRules, vbCr, vbLf)
    legacyRules = Replace(legacyRules, ";", vbLf)
    
    If Len(Trim(legacyRules)) = 0 Then
        SR_BuildRowsFromLegacyRules = ""
        Exit Function
    End If
    
    lines = Split(legacyRules, vbLf)
    
    For Each line In lines
        cleanedLine = Trim(CStr(line))
        pos = InStr(1, cleanedLine, "=", vbBinaryCompare)
        
        If pos > 1 Then
            patternValue = SR_CleanRulePart(Left(cleanedLine, pos - 1))
            targetValue = SR_CleanTargetValue(Mid(cleanedLine, pos + 1))
            
            If Len(patternValue) > 0 And Len(targetValue) > 0 Then
                result = result & "1|" & patternValue & "|" & targetValue & "|" & extensionsValue & "|" & sourceFolderValue & vbLf
            End If
        End If
    Next line
    
    SR_BuildRowsFromLegacyRules = SR_NormalizeSortRuleRows(result)
End Function

Function SR_NormalizeSortRuleRows(ByVal rows As String) As String
    Dim lines() As String
    Dim line As Variant
    Dim parts() As String
    Dim enabledValue As String
    Dim patternValue As String
    Dim targetValue As String
    Dim extensionsValue As String
    Dim sourceFolderValue As String
    Dim subjectValue As String
    Dim displayNameValue As String
    Dim result As String
    Dim seenPatterns As Object
    Dim ruleKey As String
    
    Set seenPatterns = CreateObject("Scripting.Dictionary")
    seenPatterns.CompareMode = vbTextCompare
    
    rows = Replace(rows, vbCrLf, vbLf)
    rows = Replace(rows, vbCr, vbLf)
    
    If Len(Trim(rows)) = 0 Then
        SR_NormalizeSortRuleRows = ""
        Exit Function
    End If
    
    lines = Split(rows, vbLf)
    
    For Each line In lines
        If Len(Trim(CStr(line))) > 0 Then
            parts = Split(CStr(line), "|")
            
            If UBound(parts) >= 2 Then
                enabledValue = Trim(parts(0))
                patternValue = SR_CleanRulePart(parts(1))
                targetValue = SR_CleanTargetValue(parts(2))
                
                If UBound(parts) >= 3 Then
                    extensionsValue = SR_CleanExtensions(parts(3))
                Else
                    extensionsValue = ""
                End If
                
                If UBound(parts) >= 4 Then
                    sourceFolderValue = SR_CleanFolderPath(parts(4))
                Else
                    sourceFolderValue = ""
                End If
                
                If UBound(parts) >= 5 Then
                    subjectValue = SR_CleanRulePart(parts(5))
                Else
                    subjectValue = ""
                End If
                
                If UBound(parts) >= 6 Then
                    displayNameValue = SR_CleanRulePart(parts(6))
                Else
                    displayNameValue = ""
                End If
                
                If Len(targetValue) > 0 And IsAbsoluteFolderPath(targetValue) And (Len(patternValue) > 0 Or (Len(extensionsValue) > 0 And Len(sourceFolderValue) > 0)) Then
                    ruleKey = patternValue & "|" & extensionsValue & "|" & sourceFolderValue & "|" & subjectValue
                    
                    If Not seenPatterns.Exists(ruleKey) Then
                        seenPatterns(ruleKey) = True
                        
                        If enabledValue <> "0" Then
                            enabledValue = "1"
                        End If
                        
                        result = result & enabledValue & "|" & patternValue & "|" & targetValue & "|" & extensionsValue & "|" & sourceFolderValue & "|" & subjectValue & "|" & displayNameValue & vbLf
                    End If
                End If
            End If
        End If
    Next line
    
    SR_NormalizeSortRuleRows = result
End Function

Function SR_GetActiveLegacyRulesFromRows(ByVal rows As String) As String
    Dim lines() As String
    Dim line As Variant
    Dim parts() As String
    Dim patternValue As String
    Dim targetValue As String
    Dim result As String
    
    rows = SR_NormalizeSortRuleRows(rows)
    
    If Len(Trim(rows)) = 0 Then
        SR_GetActiveLegacyRulesFromRows = ""
        Exit Function
    End If
    
    lines = Split(rows, vbLf)
    
    For Each line In lines
        If Len(Trim(CStr(line))) > 0 Then
            parts = Split(CStr(line), "|")
            
            If UBound(parts) >= 2 And Trim(parts(0)) = "1" Then
                patternValue = SR_CleanRulePart(parts(1))
                targetValue = SR_CleanTargetValue(parts(2))
                
                If Len(patternValue) > 0 And Len(targetValue) > 0 Then
                    If Len(result) > 0 Then
                        result = result & vbLf
                    End If
                    
                    result = result & patternValue & "=" & targetValue
                End If
            End If
        End If
    Next line
    
    SR_GetActiveLegacyRulesFromRows = result
End Function

Function SR_CleanRulePart(ByVal value As String) As String
    value = Trim(value)
    value = Replace(value, "|", "")
    SR_CleanRulePart = value
End Function

Function SR_CleanTargetValue(ByVal value As String) As String
    value = Trim(value)
    value = SR_TrimQuotes(value)
    value = Replace(value, "/", "\")
    value = Replace(value, "|", "")
    SR_CleanTargetValue = value
End Function

Function SR_CleanFolderPath(ByVal value As String) As String
    value = Trim(value)
    value = SR_TrimQuotes(value)
    value = Replace(value, "/", "\")
    value = Replace(value, "|", "")
    SR_CleanFolderPath = value
End Function

Function SR_CleanExtensions(ByVal value As String) As String
    Dim parts() As String
    Dim part As Variant
    Dim cleanedPart As String
    Dim result As String
    
    value = LCase(Trim(value))
    value = Replace(value, ".", "")
    value = Replace(value, " ", "")
    value = Replace(value, ";", ",")
    value = Replace(value, "|", "")
    
    If Len(value) = 0 Then
        SR_CleanExtensions = ""
        Exit Function
    End If
    
    parts = Split(value, ",")
    
    For Each part In parts
        cleanedPart = SR_CleanRulePart(CStr(part))
        cleanedPart = Replace(cleanedPart, "\", "")
        cleanedPart = Replace(cleanedPart, "/", "")
        cleanedPart = Replace(cleanedPart, ":", "")
        cleanedPart = Replace(cleanedPart, "*", "")
        cleanedPart = Replace(cleanedPart, "?", "")
        cleanedPart = Replace(cleanedPart, """", "")
        cleanedPart = Replace(cleanedPart, "<", "")
        cleanedPart = Replace(cleanedPart, ">", "")
        
        If Len(cleanedPart) > 0 Then
            If InStr(1, "," & result & ",", "," & cleanedPart & ",", vbTextCompare) = 0 Then
                If Len(result) > 0 Then
                    result = result & ","
                End If
                
                result = result & cleanedPart
            End If
        End If
    Next part
    
    SR_CleanExtensions = result
End Function

Function SR_BrowseForFolder(ByVal titleText As String, Optional ByVal initialFolderPath As String = "") As String
    On Error GoTo Fallback
    
    Dim dialog As Object
    Dim normalizedInitialPath As String
    
    Set dialog = Outlook.Application.FileDialog(4)
    
    normalizedInitialPath = Trim(initialFolderPath)
    
    If Len(normalizedInitialPath) > 0 And Dir(normalizedInitialPath, vbDirectory) <> "" Then
        If Right(normalizedInitialPath, 1) <> "\" Then
            normalizedInitialPath = normalizedInitialPath & "\"
        End If
    Else
        normalizedInitialPath = ""
    End If
    
    With dialog
        .Title = titleText
        .AllowMultiSelect = False
        
        If Len(normalizedInitialPath) > 0 Then
            .InitialFileName = normalizedInitialPath
        End If
        
        If .Show = -1 Then
            SR_BrowseForFolder = .SelectedItems(1)
        Else
            SR_BrowseForFolder = ""
        End If
    End With
    
    Exit Function
    
Fallback:
    Err.Clear
    SR_BrowseForFolder = SR_BrowseForFolderLegacy(titleText)
End Function

Private Function SR_BrowseForFolderLegacy(ByVal titleText As String) As String
    Dim shellApp As Object
    Dim folder As Object
    
    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, titleText, 0, 0)
    
    If Not folder Is Nothing Then
        SR_BrowseForFolderLegacy = folder.Self.Path
    Else
        SR_BrowseForFolderLegacy = ""
    End If
End Function

Private Function SR_TrimQuotes(ByVal value As String) As String
    value = Trim(value)
    
    If Len(value) >= 2 Then
        If Left(value, 1) = """" And Right(value, 1) = """" Then
            value = Mid(value, 2, Len(value) - 2)
        End If
    End If
    
    SR_TrimQuotes = Trim(value)
End Function

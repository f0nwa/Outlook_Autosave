Attribute VB_Name = "FileTypesSettingsCore"

Private Const FT_SETTINGS_APP_NAME As String = "OutlookAttachmentSaver"
Private Const FT_SETTINGS_SECTION As String = "Settings"
Private Const FT_ALLOWED_EXTENSIONS_KEY As String = "AllowedExtensions"
Private Const FT_FILE_TYPE_RULES_KEY As String = "FileTypeRules"

Function FT_GetFileTypeRules() As String
    FT_GetFileTypeRules = GetSetting(FT_SETTINGS_APP_NAME, FT_SETTINGS_SECTION, FT_FILE_TYPE_RULES_KEY, "")
End Function

Sub FT_SaveFileTypeRules(ByVal rules As String)
    Dim normalizedRules As String
    
    normalizedRules = FT_NormalizeFileTypeRules(rules)
    SaveSetting FT_SETTINGS_APP_NAME, FT_SETTINGS_SECTION, FT_FILE_TYPE_RULES_KEY, normalizedRules
    SaveSetting FT_SETTINGS_APP_NAME, FT_SETTINGS_SECTION, FT_ALLOWED_EXTENSIONS_KEY, FT_GetActiveExtensionsFromRules(normalizedRules)
End Sub

Function FT_BuildFileTypeRulesFromAllowedExtensions(ByVal allowedExtensions As String) As String
    Dim parts() As String
    Dim part As Variant
    Dim result As String
    
    allowedExtensions = FT_NormalizeAllowedExtensions(allowedExtensions)
    
    If Len(allowedExtensions) = 0 Then
        FT_BuildFileTypeRulesFromAllowedExtensions = ""
        Exit Function
    End If
    
    parts = Split(allowedExtensions, ";")
    
    For Each part In parts
        If Len(Trim(CStr(part))) > 0 Then
            result = result & "1|" & Trim(CStr(part)) & vbLf
        End If
    Next part
    
    FT_BuildFileTypeRulesFromAllowedExtensions = result
End Function

Function FT_NormalizeFileTypeRules(ByVal rules As String) As String
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
        FT_NormalizeFileTypeRules = ""
        Exit Function
    End If
    
    lines = Split(rules, vbLf)
    
    For Each line In lines
        If Len(Trim(CStr(line))) > 0 Then
            parts = Split(CStr(line), "|")
            
            If UBound(parts) >= 1 Then
                enabledValue = Trim(parts(0))
                extensionValue = FT_CleanFileExtension(parts(1))
                
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
    
    FT_NormalizeFileTypeRules = result
End Function

Function FT_GetActiveExtensionsFromRules(ByVal rules As String) As String
    Dim lines() As String
    Dim line As Variant
    Dim parts() As String
    Dim extensionValue As String
    Dim result As String
    
    rules = FT_NormalizeFileTypeRules(rules)
    
    If Len(Trim(rules)) = 0 Then
        FT_GetActiveExtensionsFromRules = ""
        Exit Function
    End If
    
    lines = Split(rules, vbLf)
    
    For Each line In lines
        If Len(Trim(CStr(line))) > 0 Then
            parts = Split(CStr(line), "|")
            
            If UBound(parts) >= 1 And Trim(parts(0)) = "1" Then
                extensionValue = FT_CleanFileExtension(parts(1))
                
                If Len(extensionValue) > 0 Then
                    If Len(result) > 0 Then
                        result = result & ";"
                    End If
                    
                    result = result & extensionValue
                End If
            End If
        End If
    Next line
    
    FT_GetActiveExtensionsFromRules = result
End Function

Function FT_CleanFileExtension(ByVal extensionValue As String) As String
    extensionValue = LCase(Trim(extensionValue))
    extensionValue = Replace(extensionValue, ".", "")
    extensionValue = Replace(extensionValue, " ", "")
    FT_CleanFileExtension = FT_ReplaceInvalidFileTypeCharacters(extensionValue)
End Function

Function FT_NormalizeAllowedExtensions(ByVal strExtensions As String) As String
    Dim parts() As String
    Dim part As Variant
    Dim cleanedPart As String
    Dim result As String
    
    strExtensions = Replace(strExtensions, ";", ",")
    strExtensions = Replace(strExtensions, ".", "")
    strExtensions = Replace(strExtensions, " ", "")
    strExtensions = LCase(strExtensions)
    
    If Len(strExtensions) = 0 Then
        FT_NormalizeAllowedExtensions = ""
        Exit Function
    End If
    
    parts = Split(strExtensions, ",")
    
    For Each part In parts
        cleanedPart = FT_CleanFileExtension(CStr(part))
        
        If Len(cleanedPart) > 0 Then
            If InStr(1, ";" & result & ";", ";" & cleanedPart & ";", vbTextCompare) = 0 Then
                If Len(result) > 0 Then
                    result = result & ";"
                End If
                
                result = result & cleanedPart
            End If
        End If
    Next part
    
    FT_NormalizeAllowedExtensions = result
End Function

Function FT_GetAllowedAttachmentExtensions() As String
    Dim rules As String
    
    rules = FT_GetFileTypeRules()
    
    If Len(Trim(rules)) > 0 Then
        FT_GetAllowedAttachmentExtensions = FT_GetActiveExtensionsFromRules(rules)
    Else
        FT_GetAllowedAttachmentExtensions = GetSetting(FT_SETTINGS_APP_NAME, FT_SETTINGS_SECTION, FT_ALLOWED_EXTENSIONS_KEY, "")
    End If
End Function

Private Function FT_ReplaceInvalidFileTypeCharacters(ByVal str As String) As String
    str = Replace(str, "\", "")
    str = Replace(str, "/", "")
    str = Replace(str, ":", "")
    str = Replace(str, "*", "")
    str = Replace(str, "?", "")
    str = Replace(str, """", "")
    str = Replace(str, "<", "")
    str = Replace(str, ">", "")
    str = Replace(str, "|", "")
    
    FT_ReplaceInvalidFileTypeCharacters = str
End Function

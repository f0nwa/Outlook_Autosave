Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Win32

Namespace Core

    ''' <summary>
    ''' Хранилище настроек. Использует тот же раздел реестра, что и VBA-версия
    ''' (GetSetting/SaveSetting "OutlookAttachmentSaver", "Settings"), поэтому
    ''' правила и параметры макроса подхватываются надстройкой без переноса.
    ''' </summary>
    Friend Module SettingsStore

        Friend Const RegistryPath As String = "Software\VB and VBA Program Settings\OutlookAttachmentSaver\Settings"

        Friend Const KeyManualSaveFolder As String = "ManualSaveFolder"
        Friend Const KeySelectedAllowedExtensions As String = "SelectedSaveAllowedExtensions"
        Friend Const KeySelectedExcludedExtensions As String = "SelectedSaveExcludedExtensions"
        Friend Const KeyAutoEnabled As String = "AutoSaveEnabled"
        Friend Const KeyAutoLoggingEnabled As String = "AutoSaveLoggingEnabled"
        Friend Const KeyLastRun As String = "LastAutoSaveRun"
        Friend Const KeyNextRun As String = "NextAutoSaveRun"
        ''' <summary>Прогресс по исходным папкам (только надстройка): строки "путь&lt;TAB&gt;дата".</summary>
        Friend Const KeyFolderProgress As String = "AutoSaveFolderProgress"
        Friend Const KeyRunIncomplete As String = "AutoSaveRunIncomplete"
        Friend Const KeyIntervalSeconds As String = "AutoSaveIntervalSeconds"
        Friend Const KeyLookbackSeconds As String = "AutoSaveLookbackSeconds"
        Friend Const KeyLegacySortRules As String = "AttachmentSortRules"
        Friend Const KeySortRuleRows As String = "AttachmentSortRuleRows"

        Friend Function GetValue(key As String, Optional defaultValue As String = "") As String
            Try
                Using regKey = Registry.CurrentUser.OpenSubKey(RegistryPath, False)
                    If regKey Is Nothing Then
                        Return defaultValue
                    End If

                    Dim value = regKey.GetValue(key, Nothing)

                    If value Is Nothing Then
                        Return defaultValue
                    End If

                    Return Convert.ToString(value, Globalization.CultureInfo.InvariantCulture)
                End Using
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: settings read failed: " & ex.Message)
                Return defaultValue
            End Try
        End Function

        Friend Sub SetValue(key As String, value As String)
            Using regKey = Registry.CurrentUser.CreateSubKey(RegistryPath)
                regKey.SetValue(key, If(value, String.Empty), RegistryValueKind.String)
            End Using
        End Sub

        Friend Function GetBool(key As String) As Boolean
            Return String.Equals(GetValue(key, "False"), "True", StringComparison.Ordinal)
        End Function

        Friend Sub SetBool(key As String, value As Boolean)
            SetValue(key, If(value, "True", "False"))
        End Sub

    End Module

End Namespace

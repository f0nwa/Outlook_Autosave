Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Text
Imports OutlookAutosave.Core

Namespace UI

    ''' <summary>Действия, доступные с ленты и из главной формы.</summary>
    Friend Module Commands

        Friend Sub ShowSettings()
            Dim owner = OutlookHost.GetOwnerWindow()

            Try
                AutoSaveEngine.EnsureRunning()

                Using form As New MainForm()
                    form.ShowDialog(owner)
                End Using
            Catch ex As Exception
                UiHelper.ErrorBox(owner, "Не удалось открыть форму настроек: " & ex.Message)
            End Try
        End Sub

        Friend Sub SaveSelected(owner As IWin32Window)
            Try
                Dim folderPath = AutoSaveEngine.GetManualSaveFolder()

                If folderPath.Length = 0 OrElse Not TextUtil.IsAbsoluteFolderPath(folderPath) Then
                    UiHelper.ErrorBox(owner, "Папка должна быть абсолютным путем.")
                    ShowSelectedSaveSettings(owner)
                    Return
                End If

                Dim savedPaths As New List(Of String)()
                Dim checkedCount As Integer
                Dim savedCount = AutoSaveEngine.SaveSelected(folderPath, savedPaths, checkedCount)

                If checkedCount = 0 Then
                    UiHelper.Info(owner, "Нет выбранных писем Outlook.")
                    Return
                End If

                UiHelper.Info(owner, BuildSavedMessage(savedCount, savedPaths))
            Catch ex As Exception
                UiHelper.ErrorBox(owner, "Произошла ошибка: " & ex.Message)
            End Try
        End Sub

        Private Function BuildSavedMessage(savedCount As Integer, savedPaths As List(Of String)) As String
            Const maxPaths As Integer = 20
            Dim sb As New StringBuilder("Сохранено вложений: " & savedCount.ToString(CultureInfo.InvariantCulture))

            If savedPaths.Count = 0 Then
                Return sb.ToString()
            End If

            sb.AppendLine().AppendLine().Append("Файлы:")

            For i As Integer = 0 To Math.Min(savedPaths.Count, maxPaths) - 1
                sb.AppendLine().Append((i + 1).ToString(CultureInfo.InvariantCulture)).Append(". ").Append(savedPaths(i))
            Next

            If savedPaths.Count > maxPaths Then
                sb.AppendLine().Append("...и еще " & (savedPaths.Count - maxPaths).ToString(CultureInfo.InvariantCulture) & " файл(ов).")
            End If

            Return sb.ToString()
        End Function

        Friend Sub ShowSelectedSaveSettings(owner As IWin32Window)
            Try
                Using form As New SelectedSaveSettingsForm()
                    form.ShowDialog(owner)
                End Using
            Catch ex As Exception
                UiHelper.ErrorBox(owner, "Не удалось открыть форму настроек: " & ex.Message)
            End Try
        End Sub

        Friend Sub ShowSortRules(owner As IWin32Window)
            Try
                Using form As New SortRulesForm()
                    If form.ShowDialog(owner) = DialogResult.OK Then
                        ' Правила могли поменять исходные папки — переподписываемся на события.
                        FolderWatcher.Refresh()
                    End If
                End Using
            Catch ex As Exception
                UiHelper.ErrorBox(owner, "Не удалось открыть форму правил сортировки: " & ex.Message)
            End Try
        End Sub

        Friend Sub ToggleAutoSave(owner As IWin32Window)
            Try
                If AutoSaveEngine.IsAutoSaveEnabled() Then
                    AutoSaveEngine.Disable()
                ElseIf Not AutoSaveEngine.Enable() Then
                    UiHelper.Warn(owner, AutoSaveEngine.LastFailureMessage)
                Else
                    ShowIncompleteNotice(owner)
                End If
            Catch ex As Exception
                UiHelper.ErrorBox(owner, "Не удалось изменить режим автосохранения: " & ex.Message)
            End Try
        End Sub

        Friend Sub RunNow(owner As IWin32Window)
            Try
                If AutoSaveEngine.RunNow() Then
                    ShowIncompleteNotice(owner)
                ElseIf AutoSaveEngine.LastFailureMessage.Length > 0 Then
                    UiHelper.Warn(owner, AutoSaveEngine.LastFailureMessage)
                End If
            Catch ex As Exception
                UiHelper.ErrorBox(owner, "Не удалось запустить автосохранение: " & ex.Message)
            End Try
        End Sub

        ''' <summary>Сообщает, что прогон остановился на лимите и будет продолжен автоматически.</summary>
        Private Sub ShowIncompleteNotice(owner As IWin32Window)
            If Not AutoSaveEngine.LastRunIncomplete Then
                Return
            End If

            UiHelper.Info(owner,
                "За один запуск проверяется не более " &
                AutoSaveEngine.ItemsPerRunLimit.ToString("N0", CultureInfo.CurrentCulture) & " писем." & vbCrLf &
                "Обработка продолжится автоматически примерно через минуту и будет повторяться, пока не дойдет до новых писем." & vbCrLf &
                "Outlook должен оставаться открытым.")
        End Sub

        Friend Sub EditInterval(owner As IWin32Window)
            Try
                Dim inputValue = UiHelper.Prompt(owner,
                    "Введите интервал автосохранения." & vbCrLf & "Используйте s, m или h. Примеры: 30s, 15m, 12h",
                    "Интервал автосохранения",
                    TextUtil.FormatIntervalSeconds(AutoSaveEngine.GetIntervalSeconds()))

                If inputValue.Trim().Length = 0 Then
                    Return
                End If

                Dim seconds = TextUtil.ParseIntervalSeconds(inputValue)

                If seconds <= 0 Then
                    UiHelper.ErrorBox(owner, "Некорректный интервал. Используйте значения вроде 30s, 15m или 12h.")
                    Return
                End If

                AutoSaveEngine.SetIntervalSeconds(seconds)
            Catch ex As Exception
                UiHelper.ErrorBox(owner, "Не удалось изменить интервал: " & ex.Message)
            End Try
        End Sub

        Friend Sub EditLookback(owner As IWin32Window)
            Try
                Dim inputValue = UiHelper.Prompt(owner,
                    "Введите период: s, m, h или d. Рекомендуется 1d–7d, например 3d." & vbCrLf &
                    "Каждая проверка заново пересматривает письма за этот период." & vbCrLf &
                    "Большой период (700d) — только для разового запуска: иначе" & vbCrLf &
                    "Outlook будет тормозить на каждой проверке.",
                    "Период обработки",
                    TextUtil.FormatIntervalSeconds(AutoSaveEngine.GetLookbackSeconds()))

                If inputValue.Trim().Length = 0 Then
                    Return
                End If

                Dim seconds = TextUtil.ParseIntervalSeconds(inputValue)

                If seconds <= 0 Then
                    UiHelper.ErrorBox(owner, "Некорректный период. Используйте значения вроде 12h, 24h, 3d или 7d.")
                    Return
                End If

                AutoSaveEngine.SetLookbackSeconds(seconds)
            Catch ex As Exception
                UiHelper.ErrorBox(owner, "Не удалось изменить период: " & ex.Message)
            End Try
        End Sub

        Friend Sub ToggleLogging(owner As IWin32Window)
            Try
                AutoSaveEngine.ToggleLogging()
            Catch ex As Exception
                UiHelper.ErrorBox(owner, "Не удалось изменить настройку логирования: " & ex.Message)
            End Try
        End Sub

    End Module

End Namespace

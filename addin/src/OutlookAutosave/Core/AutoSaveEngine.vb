Option Strict Off

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Text

Namespace Core

    ''' <summary>
    ''' Логика автосохранения и ручного сохранения вложений (перенос MainModule.bas).
    ''' Все методы вызываются из главного потока Outlook.
    ''' </summary>
    Friend Module AutoSaveEngine

        Friend Const Author As String = "Погуца Владислав"

        Private Const DefaultIntervalSeconds As Integer = 43200
        Private Const DefaultLookbackSeconds As Integer = 86400
        Private Const MaxItemsPerRun As Integer = 10000
        ''' <summary>Через сколько секунд продолжить, если прогон остановился на лимите писем.</summary>
        Private Const ContinuationDelaySeconds As Integer = 60
        ''' <summary>Пауза перед первым прогоном после запуска Outlook: время на синхронизацию почты с сервером.</summary>
        Private Const StartupDelaySeconds As Integer = 120
        Private Const MaxTimerSeconds As Integer = 2147483
        Private Const IndexFileName As String = "autosave.index"
        Private Const LogFileName As String = "autosave.log"

        ' VBA-версия писала индекс и лог в кодировке ANSI системы. В ANSI теряются символы вне кодовой страницы
        ' (например, эмодзи в теме письма), и ключи индекса перестают совпадать с записанными — вложение
        ' сохраняется повторно. Поэтому файлы пишутся в UTF-8 с BOM, а старые ANSI-файлы один раз переводятся
        ' в UTF-8 (см. EnsureUtf8File).
        Private ReadOnly LegacyFileEncoding As Encoding = Encoding.Default
        Private ReadOnly FileEncoding As Encoding = New UTF8Encoding(True)
        ''' <summary>Файлы индекса и лога, которые уже проверены и записаны в UTF-8.</summary>
        Private ReadOnly _utf8Files As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Private _timer As System.Windows.Forms.Timer
        Private _isRunning As Boolean
        ''' <summary>Первый прогон после запуска Outlook еще не выполнен.</summary>
        Private _startupRunPending As Boolean
        Private _lastFailureMessage As String = String.Empty
        Private _indexCache As Dictionary(Of String, Dictionary(Of String, String))

#Region "Жизненный цикл"

        Friend Sub Initialize()
            If _timer Is Nothing Then
                _timer = New System.Windows.Forms.Timer()
                AddHandler _timer.Tick, AddressOf OnTimerTick
            End If
        End Sub

        Friend Sub Shutdown()
            FolderWatcher.Shutdown()
            StopTimer()

            If _timer IsNot Nothing Then
                RemoveHandler _timer.Tick, AddressOf OnTimerTick
                _timer.Dispose()
                _timer = Nothing
            End If
        End Sub

        ''' <summary>Запуск при старте Outlook (аналог Application_Startup -> AutoSaveStartup).</summary>
        Friend Sub Startup()
            Try
                If IsAutoSaveEnabled() Then
                    ' Первый прогон — вскоре после запуска, а не через полный интервал: письма,
                    ' пришедшие пока Outlook был закрыт, обрабатываются с отметки прошлого сеанса.
                    _startupRunPending = True
                    ScheduleNext()
                    FolderWatcher.Refresh()
                    Trace.WriteLine("OutlookAutosave: startup, catch-up run scheduled.")
                Else
                    ' Незавершенная обработка после ручного запуска не продолжается в новом сеансе,
                    ' если автосохранение выключено.
                    RunIncompleteFlag = False
                    SettingsStore.SetValue(SettingsStore.KeyNextRun, String.Empty)
                End If
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: startup failed: " & ex.Message)
            End Try
        End Sub

        Friend Sub EnsureRunning()
            Try
                If IsAutoSaveEnabled() AndAlso Not IsTimerActive() Then
                    Trace.WriteLine("OutlookAutosave: timer inactive while enabled; restoring scheduled run.")
                    ScheduleNext()
                End If

                If IsAutoSaveEnabled() AndAlso FolderWatcher.WatchedFolderCount = 0 Then
                    FolderWatcher.Refresh()
                End If
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: restore failed: " & ex.Message)
            End Try
        End Sub

#End Region

#Region "Настройки"

        Friend ReadOnly Property LastFailureMessage As String
            Get
                Return _lastFailureMessage
            End Get
        End Property

        Friend ReadOnly Property LastRunIncomplete As Boolean
            Get
                Return RunIncompleteFlag
            End Get
        End Property

        ''' <summary>
        ''' Последний прогон остановился на лимите писем; следующий продолжит с места остановки.
        ''' Хранится в реестре, чтобы после перезапуска Outlook продолжение не ждало полный интервал.
        ''' </summary>
        Private Property RunIncompleteFlag As Boolean
            Get
                Return SettingsStore.GetBool(SettingsStore.KeyRunIncomplete)
            End Get
            Set(value As Boolean)
                SettingsStore.SetBool(SettingsStore.KeyRunIncomplete, value)
            End Set
        End Property

        ''' <summary>Идет прогон автосохранения.</summary>
        Friend ReadOnly Property IsBusy As Boolean
            Get
                Return _isRunning
            End Get
        End Property

        Friend ReadOnly Property ItemsPerRunLimit As Integer
            Get
                Return MaxItemsPerRun
            End Get
        End Property

        Friend Function IsAutoSaveEnabled() As Boolean
            Return SettingsStore.GetBool(SettingsStore.KeyAutoEnabled)
        End Function

        Friend Function IsLoggingEnabled() As Boolean
            Return SettingsStore.GetBool(SettingsStore.KeyAutoLoggingEnabled)
        End Function

        Friend Sub ToggleLogging()
            SettingsStore.SetBool(SettingsStore.KeyAutoLoggingEnabled, Not IsLoggingEnabled())
        End Sub

        Private Function GetPositiveInt(key As String, defaultValue As Integer) As Integer
            Dim value As Integer

            If Integer.TryParse(SettingsStore.GetValue(key, defaultValue.ToString(CultureInfo.InvariantCulture)),
                                NumberStyles.Integer, CultureInfo.InvariantCulture, value) AndAlso value > 0 Then
                Return value
            End If

            Return defaultValue
        End Function

        Friend Function GetIntervalSeconds() As Integer
            Return GetPositiveInt(SettingsStore.KeyIntervalSeconds, DefaultIntervalSeconds)
        End Function

        Friend Function GetLookbackSeconds() As Integer
            Return GetPositiveInt(SettingsStore.KeyLookbackSeconds, DefaultLookbackSeconds)
        End Function

        Friend Sub SetIntervalSeconds(seconds As Integer)
            SettingsStore.SetValue(SettingsStore.KeyIntervalSeconds, seconds.ToString(CultureInfo.InvariantCulture))

            If IsAutoSaveEnabled() Then
                ScheduleNext()
            End If
        End Sub

        Friend Sub SetLookbackSeconds(seconds As Integer)
            ' Сбрасывать отметки не нужно: каждый прогон и так перепроверяет весь период.
            SettingsStore.SetValue(SettingsStore.KeyLookbackSeconds, seconds.ToString(CultureInfo.InvariantCulture))
        End Sub

        Friend Function GetDefaultSaveFolder() As String
            Dim profile = Environment.GetEnvironmentVariable("USERPROFILE")

            If String.IsNullOrWhiteSpace(profile) Then
                Return "C:\EmailAttachments\"
            End If

            Return TextUtil.EnsureTrailingBackslash(profile) & "Downloads\EmailAttachments\"
        End Function

        Friend Function GetManualSaveFolder() As String
            Dim folderPath = SettingsStore.GetValue(SettingsStore.KeyManualSaveFolder, GetDefaultSaveFolder()).Trim()

            If folderPath.Length = 0 Then
                folderPath = GetDefaultSaveFolder()
            End If

            Return TextUtil.EnsureTrailingBackslash(folderPath)
        End Function

        Friend Function GetSelectedAllowedExtensions() As String
            Return TextUtil.NormalizeAllowedExtensions(SettingsStore.GetValue(SettingsStore.KeySelectedAllowedExtensions))
        End Function

        Friend Function GetSelectedExcludedExtensions() As String
            Return TextUtil.NormalizeAllowedExtensions(SettingsStore.GetValue(SettingsStore.KeySelectedExcludedExtensions))
        End Function

        Friend Sub SaveSelectedSaveSettings(folderPath As String, allowedExtensions As String, excludedExtensions As String)
            SettingsStore.SetValue(SettingsStore.KeyManualSaveFolder, TextUtil.EnsureTrailingBackslash(If(folderPath, String.Empty).Trim()))
            SettingsStore.SetValue(SettingsStore.KeySelectedAllowedExtensions, TextUtil.NormalizeAllowedExtensions(allowedExtensions))
            SettingsStore.SetValue(SettingsStore.KeySelectedExcludedExtensions, TextUtil.NormalizeAllowedExtensions(excludedExtensions))
        End Sub

#End Region

#Region "Включение, отключение и запуск"

        ''' <summary>Включает автосохранение и сразу выполняет первый прогон. False — если первый прогон не удался.</summary>
        Friend Function Enable() As Boolean
            Dim previousLastRun = SettingsStore.GetValue(SettingsStore.KeyLastRun)

            SettingsStore.SetBool(SettingsStore.KeyAutoEnabled, True)
            ResetProgressToLookbackStart()

            If Not RunConfiguredFolders() Then
                SettingsStore.SetBool(SettingsStore.KeyAutoEnabled, False)
                SettingsStore.SetValue(SettingsStore.KeyLastRun, previousLastRun)
                SettingsStore.SetValue(SettingsStore.KeyNextRun, String.Empty)
                StopTimer()

                If _lastFailureMessage.Length = 0 Then
                    _lastFailureMessage = "Автосохранение не включено: первый запуск завершился с ошибкой."
                End If

                Return False
            End If

            ScheduleNext()
            FolderWatcher.Refresh()
            Return True
        End Function

        Friend Sub Disable()
            SettingsStore.SetBool(SettingsStore.KeyAutoEnabled, False)
            FolderWatcher.StopAll()
            RunIncompleteFlag = False
            StopTimer()
            SettingsStore.SetValue(SettingsStore.KeyNextRun, String.Empty)
        End Sub

        ''' <summary>Один прогон за весь период сейчас. Постоянный режим не включает.</summary>
        Friend Function RunNow() As Boolean
            ' При включенном автосохранении отметки не трогаем: прогон и так проверит весь период,
            ' а незавершенная догрузка после простоя Outlook продолжится, а не начнется заново.
            ' При выключенном отметки могли устареть, поэтому начинаем с начала периода.
            If Not IsAutoSaveEnabled() Then
                ResetProgressToLookbackStart()
            End If

            If RunConfiguredFolders() Then
                If IsAutoSaveEnabled() OrElse RunIncompleteFlag Then
                    ScheduleNext()
                End If

                Return True
            End If

            Return False
        End Function

        ''' <summary>Начать с начала периода: сбрасывает отметки всех папок.</summary>
        Private Sub ResetProgressToLookbackStart()
            Dim lookbackStart = DateTime.Now.AddSeconds(-GetLookbackSeconds())
            SettingsStore.SetValue(SettingsStore.KeyLastRun, TextUtil.FormatDate(lookbackStart))
            ClearFolderProgress()
            RunIncompleteFlag = False
            Trace.WriteLine("OutlookAutosave: progress reset to lookback start: " & TextUtil.FormatDate(lookbackStart))
        End Sub

#End Region

#Region "Таймер"

        Friend Function IsTimerActive() As Boolean
            Return _timer IsNot Nothing AndAlso _timer.Enabled
        End Function

        Private Sub ScheduleNext()
            If Not StartTimer() Then
                SettingsStore.SetValue(SettingsStore.KeyNextRun, String.Empty)
                Trace.WriteLine("OutlookAutosave: next timer run was not scheduled.")
                Return
            End If

            Dim nextRun = DateTime.Now.AddSeconds(GetTimerDelaySeconds())
            SettingsStore.SetValue(SettingsStore.KeyNextRun, TextUtil.FormatDate(nextRun))
            Trace.WriteLine("OutlookAutosave: next timer run: " & TextUtil.FormatDate(nextRun))
        End Sub

        Private Function StartTimer() As Boolean
            StopTimer()

            If _timer Is Nothing OrElse Not (IsAutoSaveEnabled() OrElse RunIncompleteFlag) Then
                Return False
            End If

            Dim seconds = Math.Max(1, Math.Min(GetTimerDelaySeconds(), MaxTimerSeconds))
            _timer.Interval = seconds * 1000
            _timer.Start()
            Return True
        End Function

        ''' <summary>Обычный интервал или короткая пауза перед продолжением незавершенной обработки.</summary>
        Private Function GetTimerDelaySeconds() As Integer
            Dim interval = GetIntervalSeconds()

            If RunIncompleteFlag Then
                Return Math.Min(interval, ContinuationDelaySeconds)
            ElseIf _startupRunPending Then
                Return Math.Min(interval, StartupDelaySeconds)
            End If

            Return interval
        End Function

        Private Sub StopTimer()
            If _timer IsNot Nothing AndAlso _timer.Enabled Then
                _timer.Stop()
            End If
        End Sub

        Private Sub OnTimerTick(sender As Object, e As EventArgs)
            StopTimer()
            _startupRunPending = False

            Try
                ' Таймер работает и при выключенном автосохранении, если нужно дообработать
                ' письма после ручного запуска, остановившегося на лимите.
                If IsAutoSaveEnabled() OrElse RunIncompleteFlag Then
                    RunConfiguredFolders()
                End If

                If IsAutoSaveEnabled() OrElse RunIncompleteFlag Then
                    ScheduleNext()
                Else
                    SettingsStore.SetValue(SettingsStore.KeyNextRun, String.Empty)
                End If
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: timer run failed: " & ex.Message)

                If IsAutoSaveEnabled() OrElse RunIncompleteFlag Then
                    StartTimer()
                End If
            End Try
        End Sub

#End Region

#Region "Статус"

        Friend Function GetVersionText() As String
            Dim version = GetType(AutoSaveEngine).Assembly.GetName().Version
            Return String.Format(CultureInfo.InvariantCulture, "{0}.{1}.{2}", version.Major, version.Minor, version.Build)
        End Function

        Private Function EnabledText(value As Boolean) As String
            Return If(value, "Включено", "Отключено")
        End Function

        Friend Function GetLastRunText() As String
            Dim raw = SettingsStore.GetValue(SettingsStore.KeyLastRun)
            Dim value As DateTime

            If raw.Trim().Length = 0 Then
                Return "(нет)"
            ElseIf TextUtil.TryParseStoredDate(raw, value) Then
                Return TextUtil.FormatDate(value)
            Else
                Return raw
            End If
        End Function

        Friend Function GetNextRunText() As String
            Dim raw = SettingsStore.GetValue(SettingsStore.KeyNextRun)
            Dim value As DateTime

            If raw.Trim().Length = 0 Then
                Return "(не запланировано)"
            ElseIf TextUtil.TryParseStoredDate(raw, value) Then
                Return TextUtil.FormatDate(value)
            Else
                Return raw
            End If
        End Function

        Friend Function GetStatusText() As String
            Dim activeRules = SortRuleStore.LoadActiveRules()
            Dim sb As New StringBuilder()

            sb.AppendLine("Версия: " & GetVersionText())
            sb.AppendLine("Автор: " & Author)
            sb.AppendLine("Автосохранение: " & EnabledText(IsAutoSaveEnabled()))
            sb.AppendLine("Интервал автосохранения: " & TextUtil.FormatIntervalSeconds(GetIntervalSeconds()))
            sb.AppendLine("Период обработки: " & TextUtil.FormatIntervalSeconds(GetLookbackSeconds()))
            sb.AppendLine("Логирование: " & EnabledText(IsLoggingEnabled()))
            sb.AppendLine("Таймер: " & If(IsTimerActive(), "Активен", "Отключен"))
            sb.AppendLine("Последний запуск: " & GetLastRunText())
            sb.Append("Следующий запуск: " & GetNextRunText())

            If RunIncompleteFlag Then
                sb.AppendLine()
                sb.Append("Обработка не завершена: за один запуск проверяется не более " &
                          MaxItemsPerRun.ToString("N0", CultureInfo.CurrentCulture) &
                          " писем, продолжение по таймеру")
            End If

            If activeRules.Count > 0 Then
                Const maxRulesInStatus As Integer = 5
                Dim labels = activeRules.Take(maxRulesInStatus).Select(Function(r) r.StatusLabel()).ToList()

                If activeRules.Count > maxRulesInStatus Then
                    labels.Add("еще " & (activeRules.Count - maxRulesInStatus).ToString(CultureInfo.InvariantCulture))
                End If

                sb.AppendLine()
                sb.Append("Активные правила: " & String.Join("; ", labels))
            End If

            If IsAutoSaveEnabled() AndAlso Not activeRules.Any(Function(r) r.SourceFolder.Length > 0) Then
                sb.AppendLine()
                sb.Append("Автосохранение: нет активных правил с исходной папкой")
            End If

            Return sb.ToString()
        End Function

#End Region

#Region "Автосохранение по правилам"

        ''' <summary>
        ''' Сохраняет вложения из исходных папок активных правил. Для каждой папки хранится своя отметка
        ''' прогресса: письма перебираются от старых к новым, и если за прогон достигнут лимит писем,
        ''' следующий прогон продолжает с места остановки (см. ContinuationDelaySeconds).
        ''' </summary>
        Private Function RunConfiguredFolders() As Boolean
            _lastFailureMessage = String.Empty

            If _isRunning Then
                _lastFailureMessage = "Автосохранение уже выполняется."
                Return False
            End If

            _isRunning = True
            _startupRunPending = False
            RunIncompleteFlag = False
            _indexCache = New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.OrdinalIgnoreCase)

            Try
                Dim runStarted = DateTime.Now
                Dim fallbackSince = GetLastRun()
                Dim progress = LoadFolderProgress()
                Dim rules = SortRuleStore.LoadActiveRules()
                Dim sourcePaths = rules.
                    Select(Function(r) TextUtil.NormalizeOutlookFolderPath(r.SourceFolder)).
                    Where(Function(p) p.Length > 0).
                    Distinct(StringComparer.CurrentCultureIgnoreCase).
                    ToList()

                Dim windowStart = runStarted.AddSeconds(-GetLookbackSeconds())

                ' Папки, отставшие сильнее всего, обрабатываются первыми.
                Dim ordered = sourcePaths.
                    Select(Function(p) New KeyValuePair(Of String, DateTime)(p, GetFolderSince(progress, p, fallbackSince, windowStart))).
                    OrderBy(Function(kv) kv.Value).
                    ToList()

                Dim checkedCount As Integer = 0
                Dim savedCount As Integer = 0
                Dim processedFolders As Integer = 0
                Dim reachedLimit As Boolean = False

                If sourcePaths.Count = 0 Then
                    Trace.WriteLine("OutlookAutosave: skipped, no active sort rules with source Outlook folders.")
                    _lastFailureMessage = "Автосохранение не включено: нет активных правил с исходной папкой."
                Else
                    For Each entry In ordered
                        If checkedCount >= MaxItemsPerRun Then
                            reachedLimit = True
                            Exit For
                        End If

                        Dim folder = OutlookHost.FindMailFolderByPath(entry.Key)

                        If folder Is Nothing Then
                            ' Отметка папки не меняется: когда папка снова станет доступна, она догонит пропущенное.
                            Trace.WriteLine("OutlookAutosave: skipped missing rule source folder: " & entry.Key)
                            KeepFolderProgress(progress, entry.Key, fallbackSince)
                            Continue For
                        End If

                        Dim folderProgress As DateTime = entry.Value
                        Dim completed As Boolean = True
                        Dim failed As Boolean = False
                        savedCount += SaveFromFolderSince(folder, entry.Value, runStarted, checkedCount, rules, entry.Key, folderProgress, completed, failed)

                        If failed Then
                            ' Папку не удалось просмотреть: отметка остается прежней, чтобы следующий прогон
                            ' проверил ее с того же места, а не с текущего времени.
                            KeepFolderProgress(progress, entry.Key, fallbackSince)
                            Continue For
                        End If

                        processedFolders += 1

                        If completed Then
                            progress(entry.Key) = New FolderProgress With {.DoneUntil = runStarted}
                        Else
                            progress(entry.Key) = New FolderProgress With {
                                .DoneUntil = GetDoneUntil(progress, entry.Key, fallbackSince),
                                .ResumeAfter = folderProgress}
                        End If

                        If Not completed Then
                            reachedLimit = True
                            Exit For
                        End If
                    Next
                End If

                ' Отметки храним только для папок из текущих правил.
                For Each key In progress.Keys.ToList()
                    If Not sourcePaths.Contains(key, StringComparer.CurrentCultureIgnoreCase) Then
                        progress.Remove(key)
                    End If
                Next

                SaveFolderProgress(progress)

                Trace.WriteLine(String.Format(CultureInfo.InvariantCulture,
                                              "OutlookAutosave: completed. Items checked: {0}. Attachments saved: {1}. Limit reached: {2}",
                                              checkedCount, savedCount, reachedLimit))

                If processedFolders = 0 Then
                    Trace.WriteLine("OutlookAutosave: last run not updated, no source folders were processed.")

                    If _lastFailureMessage.Length = 0 Then
                        _lastFailureMessage = "Автосохранение не выполнено: не удалось открыть исходные папки активных правил."
                    End If

                    Return False
                End If

                If reachedLimit Then
                    ' Общая отметка (ее же читает VBA-версия) сдвигается только после полной обработки.
                    RunIncompleteFlag = True
                    Trace.WriteLine("OutlookAutosave: item limit reached, processing will continue in the next run.")
                Else
                    SettingsStore.SetValue(SettingsStore.KeyLastRun, TextUtil.FormatDate(runStarted))
                End If

                Return True
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: run failed: " & ex.ToString())
                _lastFailureMessage = "Автосохранение не включено: первый запуск завершился с ошибкой. " & ex.Message
                Return False
            Finally
                _indexCache = Nothing
                _isRunning = False
            End Try
        End Function

        ''' <summary>Общая отметка последнего полного прогона (ее же читает VBA-версия).</summary>
        Private Function GetLastRun() As DateTime
            Dim value As DateTime

            If TextUtil.TryParseStoredDate(SettingsStore.GetValue(SettingsStore.KeyLastRun), value) Then
                Return value
            End If

            Return DateTime.Now.AddSeconds(-GetLookbackSeconds())
        End Function

#Region "Прогресс по папкам"

        ''' <summary>Состояние исходной папки.</summary>
        Private NotInheritable Class FolderProgress
            ''' <summary>Время последнего полного прогона папки: все письма до этого момента проверены.</summary>
            Public Property DoneUntil As DateTime
            ''' <summary>Незавершенный прогон (остановился на лимите): продолжить с писем новее этого времени.</summary>
            Public Property ResumeAfter As DateTime?
        End Class

        ''' <summary>Строки "путь TAB DoneUntil [TAB ResumeAfter]".</summary>
        Private Function LoadFolderProgress() As Dictionary(Of String, FolderProgress)
            Dim result As New Dictionary(Of String, FolderProgress)(StringComparer.CurrentCultureIgnoreCase)
            Dim raw = SettingsStore.GetValue(SettingsStore.KeyFolderProgress)

            For Each line In raw.Replace(vbCr, String.Empty).Split(New Char() {ChrW(10)}, StringSplitOptions.RemoveEmptyEntries)
                Dim parts = line.Split(ControlChars.Tab)
                Dim doneUntil As DateTime

                If parts.Length < 2 OrElse Not TextUtil.TryParseStoredDate(parts(1), doneUntil) Then
                    Continue For
                End If

                Dim state As New FolderProgress With {.DoneUntil = doneUntil}
                Dim resumeAfter As DateTime

                If parts.Length >= 3 AndAlso TextUtil.TryParseStoredDate(parts(2), resumeAfter) Then
                    state.ResumeAfter = resumeAfter
                End If

                result(TextUtil.NormalizeOutlookFolderPath(parts(0))) = state
            Next

            Return result
        End Function

        Private Sub SaveFolderProgress(progress As Dictionary(Of String, FolderProgress))
            Dim lines As New List(Of String)()

            For Each kv In progress
                Dim line = kv.Key & ControlChars.Tab & TextUtil.FormatDate(kv.Value.DoneUntil)

                If kv.Value.ResumeAfter.HasValue Then
                    line &= ControlChars.Tab & TextUtil.FormatDate(kv.Value.ResumeAfter.Value)
                End If

                lines.Add(line)
            Next

            SettingsStore.SetValue(SettingsStore.KeyFolderProgress, String.Join(vbLf, lines))
        End Sub

        Private Sub ClearFolderProgress()
            SettingsStore.SetValue(SettingsStore.KeyFolderProgress, String.Empty)
        End Sub

        ''' <summary>
        ''' Сохраняет текущую отметку папки, которую в этом прогоне не удалось обработать. Если отметки еще нет,
        ''' записывает общую отметку прошлого прогона: иначе после успешного прогона других папок общая отметка
        ''' сдвинется, и письма этой папки за пропущенное время не будут проверены.
        ''' </summary>
        Private Sub KeepFolderProgress(progress As Dictionary(Of String, FolderProgress), folderPath As String, fallback As DateTime)
            If Not progress.ContainsKey(folderPath) Then
                progress(folderPath) = New FolderProgress With {.DoneUntil = fallback}
            End If
        End Sub

        Private Function GetDoneUntil(progress As Dictionary(Of String, FolderProgress), folderPath As String, fallback As DateTime) As DateTime
            Dim state As FolderProgress = Nothing
            Return If(progress.TryGetValue(folderPath, state), state.DoneUntil, fallback)
        End Function

        ''' <summary>
        ''' С какого времени проверять папку:
        ''' - незавершенный прогон продолжается с места остановки;
        ''' - иначе берется более раннее из двух: последний полный прогон или начало периода.
        '''   Обычно это начало периода — каждый прогон перепроверяет весь период и ловит письма,
        '''   которые попали в папку позже (докачались с сервера, перенесены вручную).
        '''   После простоя Outlook (отпуск) раньше оказывается последний прогон — тогда
        '''   обрабатываются все письма за время простоя, даже если простой длиннее периода.
        ''' </summary>
        Private Function GetFolderSince(progress As Dictionary(Of String, FolderProgress), folderPath As String,
                                        fallback As DateTime, windowStart As DateTime) As DateTime
            Dim state As FolderProgress = Nothing

            If progress.TryGetValue(folderPath, state) AndAlso state.ResumeAfter.HasValue Then
                Return state.ResumeAfter.Value
            End If

            Dim doneUntil = If(state IsNot Nothing, state.DoneUntil, fallback)
            Return If(doneUntil < windowStart, doneUntil, windowStart)
        End Function

#End Region

        ''' <summary>
        ''' Фильтр Items.Restrict по времени получения. Outlook разбирает дату по региональным настройкам
        ''' и без секунд, поэтому граница округляется вниз до минуты, а точная проверка делается в цикле.
        ''' </summary>
        Private Function BuildReceivedFilter(since As DateTime) As String
            Dim bound = New DateTime(since.Year, since.Month, since.Day, since.Hour, since.Minute, 0).AddMinutes(-1)
            Return "[ReceivedTime] >= '" & bound.ToString("g", CultureInfo.CurrentCulture) & "'"
        End Function

        ''' <summary>
        ''' Обрабатывает письма папки с временем получения в (since, runStarted], от старых к новым.
        ''' completed = False, если достигнут лимит; тогда progress — до какого времени письма обработаны.
        ''' failed = True, если папку не удалось просмотреть (ошибка Outlook); отметку папки тогда менять нельзя.
        ''' </summary>
        Private Function SaveFromFolderSince(folder As Object, since As DateTime, runStarted As DateTime, ByRef checkedCount As Integer,
                                             rules As List(Of SortRule), sourcePath As String,
                                             ByRef progress As DateTime, ByRef completed As Boolean, ByRef failed As Boolean) As Integer
            Dim savedCount As Integer = 0
            Dim allItems As Object = Nothing
            Dim items As Object = Nothing
            Dim lastProcessed As DateTime = DateTime.MinValue

            progress = since
            completed = True
            failed = False

            Try
                allItems = folder.Items

                Try
                    items = allItems.Restrict(BuildReceivedFilter(since))
                Catch ex As Exception
                    ' Без фильтра работает медленнее, но результат тот же: старые письма отсеиваются в цикле.
                    Trace.WriteLine("OutlookAutosave: Restrict failed, scanning whole folder: " & ex.Message)
                    items = allItems
                End Try

                items.Sort("[ReceivedTime]", False)

                Trace.WriteLine("OutlookAutosave: checking folder " & OutlookHost.GetFolderPath(folder) &
                                ". Since: " & TextUtil.FormatDate(since) & ". Candidates: " & CStr(items.Count))

                Dim item As Object = items.GetFirst()

                Do While item IsNot Nothing
                    Dim stopScan As Boolean = False

                    Try
                        Dim received As DateTime = CDate(item.ReceivedTime)

                        If received > runStarted Then
                            ' Письма, пришедшие во время прогона, обработает следующий прогон.
                            stopScan = True
                        ElseIf received > since Then
                            If checkedCount >= MaxItemsPerRun Then
                                completed = False
                                stopScan = True
                            Else
                                checkedCount += 1

                                If OutlookHost.IsMailItem(item) Then
                                    savedCount += SaveFromMailByRules(item, rules, sourcePath)
                                End If

                                lastProcessed = received
                            End If
                        End If
                    Catch ex As Exception
                        Trace.WriteLine("OutlookAutosave: skipped item after error: " & ex.Message)
                    End Try

                    OutlookHost.Release(item)

                    If stopScan Then
                        Exit Do
                    End If

                    item = items.GetNext()
                Loop
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: skipped folder after error: " & ex.Message)
                failed = True
            Finally
                If items IsNot allItems Then
                    OutlookHost.Release(items)
                End If

                OutlookHost.Release(allItems)
            End Try

            If Not completed AndAlso Not failed Then
                ' Письма с той же секундой получения, что и последнее обработанное, могли остаться
                ' необработанными, поэтому отметка ставится на секунду раньше. Повторно обработанные
                ' вложения отсеиваются индексом.
                Dim resumeFrom = lastProcessed.AddSeconds(-1)
                progress = If(resumeFrom > since, resumeFrom, If(lastProcessed > since, lastProcessed, since))
                Trace.WriteLine("OutlookAutosave: folder " & sourcePath & " not finished, resume after " & TextUtil.FormatDate(progress))
            End If

            Return savedCount
        End Function

        ''' <summary>
        ''' Сохраняет вложения писем, которые только что появились в исходных папках (событие ItemAdd).
        ''' Отметки прогресса не меняются: это дополнение к проверке по таймеру.
        ''' </summary>
        Friend Sub ProcessAddedItems(addedItems As List(Of FolderWatcher.AddedItem))
            If _isRunning OrElse Not IsAutoSaveEnabled() Then
                Return
            End If

            _isRunning = True
            _indexCache = New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.OrdinalIgnoreCase)

            Try
                Dim rules = SortRuleStore.LoadActiveRules()
                Dim ns = OutlookHost.GetNamespace()
                Dim savedCount As Integer = 0

                For Each added In addedItems
                    Dim item As Object = Nothing

                    Try
                        item = If(String.IsNullOrEmpty(added.StoreId),
                                  ns.GetItemFromID(added.EntryId),
                                  ns.GetItemFromID(added.EntryId, added.StoreId))

                        If OutlookHost.IsMailItem(item) Then
                            savedCount += SaveFromMailByRules(item, rules, added.SourcePath)
                        End If
                    Catch ex As Exception
                        Trace.WriteLine("OutlookAutosave: added item skipped after error: " & ex.Message)
                    Finally
                        OutlookHost.Release(item)
                    End Try
                Next

                Trace.WriteLine(String.Format(CultureInfo.InvariantCulture,
                                              "OutlookAutosave: added items processed: {0}. Attachments saved: {1}", addedItems.Count, savedCount))
            Finally
                _indexCache = Nothing
                _isRunning = False
            End Try
        End Sub

        Private Function SafeSubject(mail As Object) As String
            Try
                Return If(CStr(mail.Subject), String.Empty)
            Catch
                Return String.Empty
            End Try
        End Function

        Private Function SafeReceived(mail As Object) As DateTime
            Try
                Return CDate(mail.ReceivedTime)
            Catch
                Return DateTime.MinValue
            End Try
        End Function

        Private Function SaveFromMailByRules(mail As Object, rules As List(Of SortRule), sourcePath As String) As Integer
            Dim attachments As Object = mail.Attachments
            Dim count As Integer = CInt(attachments.Count)

            If count = 0 Then
                OutlookHost.Release(attachments)
                Return 0
            End If

            Dim subject = SafeSubject(mail)
            Dim received = SafeReceived(mail)
            Dim savedCount As Integer = 0

            For i As Integer = 1 To count
                Dim attachment As Object = Nothing

                Try
                    attachment = attachments.Item(i)
                    Dim attachmentName As String = CStr(attachment.FileName)
                    Dim matched As Boolean = False

                    ' Правила проверяются сверху вниз. Проверка вложения заканчивается на первом подходящем
                    ' правиле, если у него не включено "Продолжить проверку следующих правил".
                    For Each rule In rules
                        If Not rule.Matches(attachmentName, sourcePath, subject) Then
                            Continue For
                        End If

                        matched = True

                        If SaveAttachmentByRule(attachment, attachmentName, i, rule, subject, received) Then
                            savedCount += 1
                        End If

                        If Not rule.ContinueToNextRules Then
                            Exit For
                        End If
                    Next

                    If Not matched Then
                        Trace.WriteLine("OutlookAutosave: no matching sort rule for " & attachmentName)
                    End If
                Catch ex As Exception
                    Trace.WriteLine("OutlookAutosave: attachment skipped after error: " & ex.Message)
                Finally
                    OutlookHost.Release(attachment)
                End Try
            Next

            OutlookHost.Release(attachments)
            Return savedCount
        End Function

        ''' <summary>Сохраняет вложение в целевую папку правила. True — если файл был записан.</summary>
        Private Function SaveAttachmentByRule(attachment As Object, attachmentName As String, attachmentIndex As Integer,
                                              rule As SortRule, subject As String, received As DateTime) As Boolean
            Dim targetFolder = TextUtil.RemoveTrailingBackslash(rule.Target)
            Dim details = rule.LogDetails()

            Try
                If Not Directory.Exists(targetFolder) Then
                    Try
                        Directory.CreateDirectory(targetFolder)
                    Catch ex As Exception
                        Trace.WriteLine("OutlookAutosave: failed to create target folder " & targetFolder & ": " & ex.Message)
                        WriteLog(targetFolder, "FailedCreateFolder", ex.Message, subject, received, attachmentName, String.Empty, details)
                        Return False
                    End Try
                End If

                If Not Directory.Exists(targetFolder) Then
                    WriteLog(targetFolder, "SkippedNoTargetFolder", "Target folder is not available", subject, received, attachmentName, String.Empty, details)
                    Return False
                End If

                Dim indexKey = BuildIndexKey(received, subject, attachmentName, attachmentIndex)
                Dim indexedPath As String = Nothing

                If IndexContains(targetFolder, indexKey, indexedPath) Then
                    ' Уже обработанные вложения встречаются при каждой перепроверке периода — в лог их не пишем.
                    Return False
                End If

                Dim originalPath = TextUtil.EnsureTrailingBackslash(targetFolder) & TextUtil.CreateValidName(attachmentName)
                Dim filePath = originalPath
                Dim status = "Saved"
                Dim reason = "Attachment saved"

                If File.Exists(filePath) Then
                    filePath = BuildDatedFilePath(originalPath, GetMailDateForFileName(received))

                    If File.Exists(filePath) Then
                        AppendIndex(targetFolder, indexKey, filePath)
                        WriteLog(targetFolder, "SkippedExists", "Dated file already exists", subject, received, attachmentName, filePath, details)
                        Return False
                    End If

                    status = "SavedAsDatedCopy"
                    reason = "Original file name already exists"
                End If

                Try
                    attachment.SaveAsFile(filePath)
                Catch ex As Exception
                    Trace.WriteLine("OutlookAutosave: failed to save " & filePath & ": " & ex.Message)
                    WriteLog(targetFolder, "FailedSave", ex.Message, subject, received, attachmentName, filePath, details)
                    Return False
                End Try

                AppendIndex(targetFolder, indexKey, filePath)
                WriteLog(targetFolder, status, reason, subject, received, attachmentName, filePath, details)
                Return True
            Catch ex As Exception
                ' Ошибка по одному правилу не должна мешать следующим правилам для этого вложения.
                Trace.WriteLine("OutlookAutosave: rule " & rule.Number.ToString(CultureInfo.InvariantCulture) &
                                " skipped attachment after error: " & ex.Message)
                Return False
            End Try
        End Function

        Private Function GetMailDateForFileName(received As DateTime) As DateTime
            Return If(received.Year > 1900, received, DateTime.Now)
        End Function

        Private Function BuildDatedFilePath(filePath As String, mailDate As DateTime) As String
            Dim folderPart = String.Empty
            Dim fileName = filePath
            Dim slashPos = filePath.LastIndexOf("\"c)

            If slashPos >= 0 Then
                folderPart = filePath.Substring(0, slashPos + 1)
                fileName = filePath.Substring(slashPos + 1)
            End If

            Dim baseName = fileName
            Dim extension = String.Empty
            Dim dotPos = fileName.LastIndexOf("."c)

            If dotPos > 0 Then
                baseName = fileName.Substring(0, dotPos)
                extension = fileName.Substring(dotPos)
            End If

            Return folderPart & baseName & "_" & mailDate.ToString(TextUtil.FileNameDateFormat, CultureInfo.InvariantCulture) & extension
        End Function

        ''' <summary>Ключ индекса в том же формате, что и у VBA-версии (совместим с существующими autosave.index).</summary>
        Private Function BuildIndexKey(received As DateTime, subject As String, attachmentName As String, attachmentIndex As Integer) As String
            Dim mailKey = TextUtil.FormatDate(received) & "|" & subject
            Return TextUtil.SingleLineField(mailKey) & "|" & attachmentIndex.ToString(CultureInfo.InvariantCulture) & "|" & TextUtil.SingleLineField(attachmentName)
        End Function

        Private Function LoadIndex(targetFolder As String) As Dictionary(Of String, String)
            Dim cacheKey = TextUtil.RemoveTrailingBackslash(targetFolder)
            Dim entries As Dictionary(Of String, String) = Nothing

            If _indexCache IsNot Nothing AndAlso _indexCache.TryGetValue(cacheKey, entries) Then
                Return entries
            End If

            entries = New Dictionary(Of String, String)(StringComparer.Ordinal)
            Dim indexPath = TextUtil.EnsureTrailingBackslash(targetFolder) & IndexFileName

            Try
                If File.Exists(indexPath) Then
                    For Each line In File.ReadAllLines(indexPath, EnsureUtf8File(indexPath))
                        Dim tabPos = line.IndexOf(ControlChars.Tab)

                        If tabPos > 0 Then
                            Dim key = line.Substring(0, tabPos)

                            If Not entries.ContainsKey(key) Then
                                entries.Add(key, line.Substring(tabPos + 1))
                            End If
                        End If
                    Next
                End If
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: index read failed: " & ex.Message)
            End Try

            If _indexCache IsNot Nothing Then
                _indexCache(cacheKey) = entries
            End If

            Return entries
        End Function

        Private Function IndexContains(targetFolder As String, indexKey As String, ByRef savedFilePath As String) As Boolean
            savedFilePath = String.Empty

            If String.IsNullOrEmpty(indexKey) OrElse String.IsNullOrEmpty(targetFolder) Then
                Return False
            End If

            Dim path As String = Nothing

            If LoadIndex(targetFolder).TryGetValue(indexKey, path) Then
                savedFilePath = path
                Return True
            End If

            Return False
        End Function

        Private Sub AppendIndex(targetFolder As String, indexKey As String, savedFilePath As String)
            Try
                If String.IsNullOrEmpty(indexKey) OrElse Not Directory.Exists(targetFolder) Then
                    Return
                End If

                Dim entries = LoadIndex(targetFolder)

                If entries.ContainsKey(indexKey) Then
                    Return
                End If

                Dim indexPath = TextUtil.EnsureTrailingBackslash(targetFolder) & IndexFileName
                File.AppendAllText(indexPath, indexKey & ControlChars.Tab & savedFilePath & vbCrLf, EnsureUtf8File(indexPath))
                entries(indexKey) = savedFilePath

                Try
                    File.SetAttributes(indexPath, File.GetAttributes(indexPath) Or FileAttributes.Hidden)
                Catch
                End Try
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: index write failed: " & ex.Message)
            End Try
        End Sub

        Private Sub WriteLog(targetFolder As String, status As String, reason As String, subject As String, received As DateTime,
                             attachmentName As String, targetFilePath As String, ruleDetails As String)
            Try
                If Not IsLoggingEnabled() OrElse String.IsNullOrEmpty(targetFolder) OrElse Not Directory.Exists(targetFolder) Then
                    Return
                End If

                Dim receivedText = If(received = DateTime.MinValue, String.Empty, TextUtil.FormatDate(received))
                Dim fields = New String() {
                    TextUtil.FormatDate(DateTime.Now),
                    TextUtil.SingleLineField(status),
                    TextUtil.SingleLineField(reason),
                    TextUtil.SingleLineField(ruleDetails),
                    TextUtil.SingleLineField(subject),
                    TextUtil.SingleLineField(receivedText),
                    TextUtil.SingleLineField(attachmentName),
                    TextUtil.SingleLineField(targetFilePath)
                }

                Dim logPath = TextUtil.EnsureTrailingBackslash(targetFolder) & LogFileName
                File.AppendAllText(logPath, String.Join(" | ", fields) & vbCrLf, EnsureUtf8File(logPath))
            Catch ex As Exception
                ' Ошибки записи лога не должны останавливать автосохранение.
                Trace.WriteLine("OutlookAutosave: log write failed: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Кодировка для чтения и дописывания индекса или лога. Файл в ANSI (от VBA-версии или прежних версий
        ''' надстройки) один раз переписывается в UTF-8 с BOM. Если перевести не удалось, файл остается в ANSI
        ''' и дописывается в ANSI, чтобы не смешивать кодировки в одном файле.
        ''' </summary>
        Private Function EnsureUtf8File(filePath As String) As Encoding
            If _utf8Files.Contains(filePath) Then
                Return FileEncoding
            End If

            Try
                If File.Exists(filePath) AndAlso Not HasUtf8Bom(filePath) Then
                    Dim text = File.ReadAllText(filePath, LegacyFileEncoding)
                    Dim attributes = File.GetAttributes(filePath)

                    ' Скрытый файл нельзя перезаписать, поэтому атрибут на время снимается.
                    File.SetAttributes(filePath, attributes And Not FileAttributes.Hidden)

                    Try
                        File.WriteAllText(filePath, text, FileEncoding)
                    Finally
                        Try
                            File.SetAttributes(filePath, attributes)
                        Catch
                        End Try
                    End Try

                    Trace.WriteLine("OutlookAutosave: converted to UTF-8: " & filePath)
                End If

                _utf8Files.Add(filePath)
                Return FileEncoding
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: UTF-8 conversion failed, keeping ANSI: " & filePath & ": " & ex.Message)
                Return LegacyFileEncoding
            End Try
        End Function

        ''' <summary>True, если файл начинается с UTF-8 BOM или пустой (в пустой файл BOM запишется при дописывании).</summary>
        Private Function HasUtf8Bom(filePath As String) As Boolean
            Using stream As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                If stream.Length = 0 Then
                    Return True
                End If

                Dim preamble(2) As Byte
                Return stream.Read(preamble, 0, 3) = 3 AndAlso preamble(0) = &HEF AndAlso preamble(1) = &HBB AndAlso preamble(2) = &HBF
            End Using
        End Function

#End Region

#Region "Сохранение выбранных писем"

        ''' <summary>
        ''' Сохраняет все подходящие вложения из выделенных писем в папку. Возвращает число сохраненных файлов.
        ''' checkedCount — количество выделенных писем.
        ''' </summary>
        Friend Function SaveSelected(targetFolder As String, savedPaths As List(Of String), ByRef checkedCount As Integer) As Integer
            checkedCount = 0
            Dim explorer = OutlookHost.GetActiveExplorer()

            If explorer Is Nothing Then
                Return 0
            End If

            Dim selection As Object = explorer.Selection
            Dim savedCount As Integer = 0

            For i As Integer = 1 To CInt(selection.Count)
                Dim item As Object = Nothing

                Try
                    item = selection.Item(i)

                    If OutlookHost.IsMailItem(item) Then
                        checkedCount += 1
                        savedCount += SaveAllAttachmentsToFolder(item, targetFolder, savedPaths)
                    End If
                Finally
                    OutlookHost.Release(item)
                End Try
            Next

            Return savedCount
        End Function

        Private Function SaveAllAttachmentsToFolder(mail As Object, targetFolder As String, savedPaths As List(Of String)) As Integer
            Dim attachments As Object = mail.Attachments
            Dim count As Integer = CInt(attachments.Count)

            If count = 0 Then
                OutlookHost.Release(attachments)
                Return 0
            End If

            targetFolder = TextUtil.EnsureTrailingBackslash(targetFolder)

            If Not Directory.Exists(targetFolder) Then
                Directory.CreateDirectory(targetFolder)
            End If

            Dim allowed = GetSelectedAllowedExtensions()
            Dim excluded = GetSelectedExcludedExtensions()
            Dim savedCount As Integer = 0

            For i As Integer = 1 To count
                Dim attachment As Object = Nothing

                Try
                    attachment = attachments.Item(i)
                    Dim attachmentName As String = CStr(attachment.FileName)

                    If Not IsSelectedSaveAllowed(attachmentName, allowed, excluded) Then
                        Trace.WriteLine("OutlookAutosave: selected attachment skipped by filter: " & attachmentName)
                        Continue For
                    End If

                    Dim filePath = GetUniqueFilePath(targetFolder & TextUtil.CreateValidName(attachmentName))

                    If filePath.Length = 0 Then
                        Trace.WriteLine("OutlookAutosave: no unique file name for " & attachmentName)
                        Continue For
                    End If

                    attachment.SaveAsFile(filePath)
                    savedCount += 1
                    savedPaths?.Add(filePath)
                Catch ex As Exception
                    Trace.WriteLine("OutlookAutosave: failed to save selected attachment: " & ex.Message)
                Finally
                    OutlookHost.Release(attachment)
                End Try
            Next

            OutlookHost.Release(attachments)
            Return savedCount
        End Function

        Private Function IsSelectedSaveAllowed(fileName As String, allowed As String, excluded As String) As Boolean
            Dim extension = TextUtil.GetFileExtension(fileName)

            If extension.Length > 0 AndAlso excluded.Length > 0 AndAlso TextUtil.IsExtensionInList(extension, excluded) Then
                Return False
            End If

            If allowed.Length > 0 Then
                Return extension.Length > 0 AndAlso TextUtil.IsExtensionInList(extension, allowed)
            End If

            Return True
        End Function

        Private Function GetUniqueFilePath(filePath As String) As String
            If Not File.Exists(filePath) Then
                Return filePath
            End If

            Dim folderPart = String.Empty
            Dim fileName = filePath
            Dim slashPos = filePath.LastIndexOf("\"c)

            If slashPos >= 0 Then
                folderPart = filePath.Substring(0, slashPos + 1)
                fileName = filePath.Substring(slashPos + 1)
            End If

            Dim baseName = fileName
            Dim extension = String.Empty
            Dim dotPos = fileName.LastIndexOf("."c)

            If dotPos > 0 Then
                baseName = fileName.Substring(0, dotPos)
                extension = fileName.Substring(dotPos)
            End If

            For counter As Integer = 2 To 9999
                Dim candidate = folderPart & baseName & " (" & counter.ToString(CultureInfo.InvariantCulture) & ")" & extension

                If Not File.Exists(candidate) Then
                    Return candidate
                End If
            Next

            Return String.Empty
        End Function

#End Region

    End Module

End Namespace

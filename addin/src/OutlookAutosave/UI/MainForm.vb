Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports OutlookAutosave.Core

Namespace UI

    ''' <summary>Главная форма: статус, действия и настройки (перенос AutoSaveSettingsForm).</summary>
    Friend NotInheritable Class MainForm
        Inherits Form

        Private ReadOnly _toolTip As New ToolTip()
        Private ReadOnly _txtStatus As New TextBox()
        Private ReadOnly _btnRefresh As Button
        Private ReadOnly _btnSaveSelected As Button
        Private ReadOnly _btnSelectedSettings As Button
        Private ReadOnly _btnToggleAutoSave As Button
        Private ReadOnly _btnRunNow As Button
        Private ReadOnly _btnSortRules As Button
        Private ReadOnly _btnInterval As Button
        Private ReadOnly _btnLookback As Button
        Private ReadOnly _btnToggleLogging As Button

        Public Sub New()
            Text = "Автосохранение вложений Outlook"
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            StartPosition = FormStartPosition.CenterParent
            MaximizeBox = False
            MinimizeBox = False
            ShowInTaskbar = False
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Font = SystemFonts.MessageBoxFont
            ClientSize = New Size(820, 520)
            Icon = UiHelper.AppIcon()

            AddLabel("Настройки автосохранения вложений", 16, 12, 600, True)
            AddLabel("Статус", 16, 44, 200, True)

            With _txtStatus
                .Location = New Point(16, 66)
                .Size = New Size(788, 220)
                .Multiline = True
                .ReadOnly = True
                .ScrollBars = ScrollBars.Vertical
                .BackColor = SystemColors.Window
            End With

            Controls.Add(_txtStatus)

            AddLabel("Действия", 16, 300, 200, True)
            _btnRefresh = AddButton("Обновить статус", 16, 324, 170, "Обновить текущий статус автосохранения.", AddressOf OnRefresh)
            _btnSaveSelected = AddButton("Сохранить выбранные", 196, 324, 170, "Сохранить вложения из выбранных писем Outlook.", AddressOf OnSaveSelected)
            _btnSelectedSettings = AddButton("Настройки выбранных", 376, 324, 190, "Открыть настройки сохранения вложений из выбранных писем.", AddressOf OnSelectedSettings)
            _btnToggleAutoSave = AddButton("Включить", 16, 362, 170, String.Empty, AddressOf OnToggleAutoSave)
            _btnRunNow = AddButton("Запустить", 196, 362, 170, "Запустить автосохранение один раз сейчас. Не включает постоянный автоматический режим.", AddressOf OnRunNow)

            AddLabel("Настройки", 16, 414, 200, True)
            _btnSortRules = AddButton("Правила", 16, 438, 150, "Настроить правила сортировки вложений.", AddressOf OnSortRules)
            _btnInterval = AddButton("Интервал", 176, 438, 150, "Задать частоту проверки настроенных папок.", AddressOf OnInterval)
            _btnLookback = AddButton("Период", 336, 438, 150, "Задать период, за который Outlook будет искать письма после простоя.", AddressOf OnLookback)
            _btnToggleLogging = AddButton("Логирование: Включить", 496, 438, 230, String.Empty, AddressOf OnToggleLogging)

            Dim btnClose = AddButton("Закрыть", 704, 480, 100, String.Empty, Sub(s, e) Close())
            CancelButton = btnClose

            RefreshStatus()
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                _toolTip.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

        Private Sub AddLabel(text As String, x As Integer, y As Integer, width As Integer, bold As Boolean)
            Dim label As New Label With {.Text = text, .Location = New Point(x, y), .Size = New Size(width, 20), .AutoSize = False}

            If bold Then
                label.Font = New Font(Font, FontStyle.Bold)
            End If

            Controls.Add(label)
        End Sub

        Private Function AddButton(text As String, x As Integer, y As Integer, width As Integer, tip As String, handler As EventHandler) As Button
            Dim button As New Button With {.Text = text, .Location = New Point(x, y), .Size = New Size(width, 30)}
            AddHandler button.Click, handler

            If tip.Length > 0 Then
                _toolTip.SetToolTip(button, tip)
            End If

            Controls.Add(button)
            Return button
        End Function

        Private Sub RefreshStatus()
            _txtStatus.Text = AutoSaveEngine.GetStatusText()
            _txtStatus.SelectionStart = 0
            _txtStatus.SelectionLength = 0

            If AutoSaveEngine.IsAutoSaveEnabled() Then
                _btnToggleAutoSave.Text = "Отключить"
                _toolTip.SetToolTip(_btnToggleAutoSave, "Отключить автоматическое сохранение вложений.")
            Else
                _btnToggleAutoSave.Text = "Включить"
                _toolTip.SetToolTip(_btnToggleAutoSave, "Включить автосохранение. Outlook сразу проверит настроенные папки и продолжит по интервалу.")
            End If

            If AutoSaveEngine.IsLoggingEnabled() Then
                _btnToggleLogging.Text = "Логирование: Отключить"
                _toolTip.SetToolTip(_btnToggleLogging, "Отключить запись лога автосохранения.")
            Else
                _btnToggleLogging.Text = "Логирование: Включить"
                _toolTip.SetToolTip(_btnToggleLogging, "Включить запись лога автосохранения.")
            End If
        End Sub

        Private Sub RunWithWaitCursor(action As Action)
            Dim previous = System.Windows.Forms.Cursor.Current
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            UseWaitCursor = True

            Try
                action()
            Finally
                UseWaitCursor = False
                System.Windows.Forms.Cursor.Current = previous
                RefreshStatus()
            End Try
        End Sub

        Private Sub OnRefresh(sender As Object, e As EventArgs)
            RefreshStatus()
        End Sub

        Private Sub OnSaveSelected(sender As Object, e As EventArgs)
            RunWithWaitCursor(Sub() Commands.SaveSelected(Me))
        End Sub

        Private Sub OnSelectedSettings(sender As Object, e As EventArgs)
            Commands.ShowSelectedSaveSettings(Me)
            RefreshStatus()
        End Sub

        Private Sub OnToggleAutoSave(sender As Object, e As EventArgs)
            RunWithWaitCursor(Sub() Commands.ToggleAutoSave(Me))
        End Sub

        Private Sub OnRunNow(sender As Object, e As EventArgs)
            RunWithWaitCursor(Sub() Commands.RunNow(Me))
        End Sub

        Private Sub OnSortRules(sender As Object, e As EventArgs)
            Commands.ShowSortRules(Me)
            RefreshStatus()
        End Sub

        Private Sub OnInterval(sender As Object, e As EventArgs)
            Commands.EditInterval(Me)
            RefreshStatus()
        End Sub

        Private Sub OnLookback(sender As Object, e As EventArgs)
            Commands.EditLookback(Me)
            RefreshStatus()
        End Sub

        Private Sub OnToggleLogging(sender As Object, e As EventArgs)
            Commands.ToggleLogging(Me)
            RefreshStatus()
        End Sub

    End Class

End Namespace

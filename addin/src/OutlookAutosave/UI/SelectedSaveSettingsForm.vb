Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports OutlookAutosave.Core

Namespace UI

    ''' <summary>Настройки кнопки "Сохранить выбранные" (перенос SelectedSaveSettingsForm).</summary>
    Friend NotInheritable Class SelectedSaveSettingsForm
        Inherits Form

        Private ReadOnly _txtFolder As New TextBox()
        Private ReadOnly _txtAllowed As New TextBox()
        Private ReadOnly _txtExcluded As New TextBox()

        Public Sub New()
            Text = "Настройки выбранных писем"
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            StartPosition = FormStartPosition.CenterParent
            MaximizeBox = False
            MinimizeBox = False
            ShowInTaskbar = False
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Font = SystemFonts.MessageBoxFont
            ClientSize = New Size(700, 390)
            Icon = UiHelper.AppIcon()

            Dim title As New Label With {.Text = "Настройки выбранных писем", .Location = New Point(16, 12), .Size = New Size(560, 22)}
            title.Font = New Font(Font, FontStyle.Bold)
            Controls.Add(title)

            Controls.Add(New Label With {.Text = "Папка по умолчанию", .Location = New Point(16, 50), .Size = New Size(400, 20)})
            _txtFolder.Location = New Point(16, 72)
            _txtFolder.Size = New Size(556, 23)
            Controls.Add(_txtFolder)

            Dim btnBrowse As New Button With {.Text = "Обзор", .Location = New Point(582, 70), .Size = New Size(102, 28)}
            AddHandler btnBrowse.Click, AddressOf OnBrowse
            Controls.Add(btnBrowse)

            Controls.Add(New Label With {.Text = "Разрешенные расширения", .Location = New Point(16, 118), .Size = New Size(400, 20)})
            _txtAllowed.Location = New Point(16, 140)
            _txtAllowed.Size = New Size(668, 23)
            Controls.Add(_txtAllowed)
            Controls.Add(New Label With {.Text = "Например: pdf, zip, xlsx. Пусто - без ограничения.", .Location = New Point(16, 166), .Size = New Size(668, 20), .ForeColor = SystemColors.GrayText})

            Controls.Add(New Label With {.Text = "Исключенные расширения", .Location = New Point(16, 206), .Size = New Size(400, 20)})
            _txtExcluded.Location = New Point(16, 228)
            _txtExcluded.Size = New Size(668, 23)
            Controls.Add(_txtExcluded)
            Controls.Add(New Label With {.Text = "Исключения имеют приоритет над разрешенными расширениями.", .Location = New Point(16, 254), .Size = New Size(668, 20), .ForeColor = SystemColors.GrayText})

            Dim btnSave As New Button With {.Text = "Сохранить", .Location = New Point(16, 340), .Size = New Size(120, 30)}
            AddHandler btnSave.Click, AddressOf OnSave
            Controls.Add(btnSave)

            Dim btnCancel As New Button With {.Text = "Отмена", .Location = New Point(564, 340), .Size = New Size(120, 30), .DialogResult = System.Windows.Forms.DialogResult.Cancel}
            Controls.Add(btnCancel)

            AcceptButton = btnSave
            CancelButton = btnCancel

            _txtFolder.Text = AutoSaveEngine.GetManualSaveFolder()
            _txtAllowed.Text = TextUtil.FormatAllowedExtensionsForInput(AutoSaveEngine.GetSelectedAllowedExtensions())
            _txtExcluded.Text = TextUtil.FormatAllowedExtensionsForInput(AutoSaveEngine.GetSelectedExcludedExtensions())
        End Sub

        Private Sub OnBrowse(sender As Object, e As EventArgs)
            Dim folderPath = UiHelper.BrowseForFolder(Me, "Выберите папку для сохранения вложений из выбранных писем", _txtFolder.Text)

            If folderPath.Length > 0 Then
                _txtFolder.Text = TextUtil.EnsureTrailingBackslash(folderPath)
            End If
        End Sub

        Private Sub OnSave(sender As Object, e As EventArgs)
            Dim folderPath = _txtFolder.Text.Trim()

            If folderPath.Length = 0 OrElse Not TextUtil.IsAbsoluteFolderPath(folderPath) Then
                UiHelper.Info(Me, "Папка должна быть абсолютным путем.")
                Return
            End If

            Try
                AutoSaveEngine.SaveSelectedSaveSettings(folderPath, _txtAllowed.Text, _txtExcluded.Text)
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Close()
            Catch ex As Exception
                UiHelper.ErrorBox(Me, "Не удалось сохранить настройки: " & ex.Message)
            End Try
        End Sub

    End Class

End Namespace

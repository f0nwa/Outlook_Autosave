Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports OutlookAutosave.Core

Namespace UI

    ''' <summary>Редактор правил сортировки вложений (перенос SortRulesForm).</summary>
    Friend NotInheritable Class SortRulesForm
        Inherits Form

        Private ReadOnly _toolTip As New ToolTip()
        Private ReadOnly _rules As List(Of SortRule)
        Private ReadOnly _lstRules As New ListView()
        Private ReadOnly _txtRuleName As New TextBox()
        Private ReadOnly _txtPattern As New TextBox()
        Private ReadOnly _txtExtensions As New TextBox()
        Private ReadOnly _txtSubject As New TextBox()
        Private ReadOnly _txtTarget As New TextBox()
        Private ReadOnly _txtSourceFolder As New TextBox()
        Private ReadOnly _chkContinue As New CheckBox()
        Private ReadOnly _btnToggleActive As Button
        Private _suppressSelectionEvents As Boolean

        Public Sub New()
            Text = "Правила сортировки вложений"
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            StartPosition = FormStartPosition.CenterParent
            MaximizeBox = False
            MinimizeBox = False
            ShowInTaskbar = False
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Font = SystemFonts.MessageBoxFont
            ClientSize = New Size(780, 724)
            Icon = UiHelper.AppIcon()
            KeyPreview = True

            Dim title As New Label With {.Text = "Правила сортировки вложений", .Location = New Point(16, 12), .Size = New Size(560, 22)}
            title.Font = New Font(Font, FontStyle.Bold)
            Controls.Add(title)
            Controls.Add(New Label With {
                .Text = "Галка значит активное правило. Выделение только выбирает правило. Двойной клик включает или отключает. Ctrl+A выделяет все.",
                .Location = New Point(16, 38), .Size = New Size(748, 36)})

            With _lstRules
                .Location = New Point(16, 76)
                .Size = New Size(748, 190)
                .View = View.Details
                .FullRowSelect = True
                .MultiSelect = True
                .HideSelection = False
                .HeaderStyle = ColumnHeaderStyle.Nonclickable
                .Columns.Add("", 30)
                .Columns.Add("Правило", 700)
            End With

            AddHandler _lstRules.SelectedIndexChanged, AddressOf OnRuleSelectionChanged
            AddHandler _lstRules.DoubleClick, AddressOf OnRuleDoubleClick
            AddHandler _lstRules.KeyDown, AddressOf OnRulesKeyDown
            Controls.Add(_lstRules)

            AddField("Название правила", _txtRuleName, 16, 278, 600)
            AddField("Текст в имени файла", _txtPattern, 16, 334, 290)
            AddField("Расширения", _txtExtensions, 326, 334, 290)
            AddField("Тема письма", _txtSubject, 16, 390, 600)
            AddField("Целевая папка", _txtTarget, 16, 446, 600)
            AddButton("Обзор", 626, 467, 138, "Выбрать целевую папку для этого правила.", AddressOf OnBrowseTarget)
            AddField("Исходная папка Outlook", _txtSourceFolder, 16, 502, 452)
            AddButton("Выбрать", 478, 523, 138, "Выбрать исходную папку Outlook из списка.", AddressOf OnBrowseSource)
            AddButton("Текущая", 626, 523, 138, "Взять текущую открытую папку Outlook как исходную для правила.", AddressOf OnCurrentSource)

            With _chkContinue
                .Text = "Продолжить проверку следующих правил (сохранить вложение и по другим подходящим правилам)"
                .Location = New Point(16, 564)
                .Size = New Size(748, 24)
            End With

            _toolTip.SetToolTip(_chkContinue,
                "Если включено, после сохранения по этому правилу вложение проверяется следующими правилами списка" & vbCrLf &
                "и может сохраниться еще в другие папки. Если выключено, проверка вложения на этом правиле заканчивается.")
            Controls.Add(_chkContinue)

            AddButton("Добавить", 16, 614, 120, "Добавить правило из заполненных полей.", AddressOf OnAdd)
            AddButton("Обновить", 146, 614, 120, "Обновить выбранное правило текущими значениями.", AddressOf OnUpdate)
            AddButton("Удалить", 276, 614, 120, "Удалить выбранные правила.", AddressOf OnRemove)
            _btnToggleActive = AddButton("Включить", 406, 614, 140, "Включить выбранные правила.", AddressOf OnToggleActive)

            Dim btnSave = AddButton("Сохранить", 16, 670, 120, "Сохранить все правила и закрыть окно.", AddressOf OnSave)
            Dim btnCancel = AddButton("Отмена", 644, 670, 120, "Закрыть без сохранения изменений.", Sub(s, e) Close())
            CancelButton = btnCancel

            _rules = SortRuleStore.LoadAllRules()
            RefreshList(New Integer() {})
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                _toolTip.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

        Private Sub AddField(caption As String, box As TextBox, x As Integer, y As Integer, width As Integer)
            Controls.Add(New Label With {.Text = caption, .Location = New Point(x, y), .Size = New Size(width, 20)})
            box.Location = New Point(x, y + 22)
            box.Size = New Size(width, 23)
            Controls.Add(box)
        End Sub

        Private Function AddButton(caption As String, x As Integer, y As Integer, width As Integer, tip As String, handler As EventHandler) As Button
            Dim btn As New Button With {.Text = caption, .Location = New Point(x, y), .Size = New Size(width, 30)}
            AddHandler btn.Click, handler
            _toolTip.SetToolTip(btn, tip)
            Controls.Add(btn)
            Return btn
        End Function

#Region "Список"

        Private Sub RefreshList(selectedIndexes As IEnumerable(Of Integer))
            Dim selected = New HashSet(Of Integer)(selectedIndexes)
            _suppressSelectionEvents = True

            Try
                _lstRules.BeginUpdate()
                _lstRules.Items.Clear()

                For i As Integer = 0 To _rules.Count - 1
                    Dim rule = _rules(i)
                    rule.Number = i + 1
                    Dim item As New ListViewItem(If(rule.Enabled, ChrW(&H2713).ToString(), String.Empty))
                    item.SubItems.Add(rule.DisplayText(i + 1))
                    item.Selected = selected.Contains(i)
                    _lstRules.Items.Add(item)
                Next
            Finally
                _lstRules.EndUpdate()
                _suppressSelectionEvents = False
            End Try

            If selected.Count > 0 Then
                Dim first = selected.Min()

                If first < _lstRules.Items.Count Then
                    _lstRules.Items(first).EnsureVisible()
                    _lstRules.FocusedItem = _lstRules.Items(first)
                End If
            End If

            RefreshToggleActiveButton()
        End Sub

        Private Function SelectedIndexes() As List(Of Integer)
            Return _lstRules.SelectedIndices.Cast(Of Integer)().OrderBy(Function(i) i).ToList()
        End Function

        Private Function SingleSelectedIndex() As Integer
            Dim indexes = SelectedIndexes()
            Return If(indexes.Count = 1, indexes(0), -1)
        End Function

        Private Sub RefreshToggleActiveButton()
            Dim indexes = SelectedIndexes()

            If indexes.Count > 0 AndAlso indexes.All(Function(i) _rules(i).Enabled) Then
                _btnToggleActive.Text = "Отключить"
                _toolTip.SetToolTip(_btnToggleActive, "Отключить выбранные правила.")
            Else
                _btnToggleActive.Text = "Включить"
                _toolTip.SetToolTip(_btnToggleActive, "Включить выбранные правила.")
            End If
        End Sub

        Private Sub OnRuleSelectionChanged(sender As Object, e As EventArgs)
            If _suppressSelectionEvents Then
                Return
            End If

            Dim index = SingleSelectedIndex()

            If index >= 0 Then
                FillFields(_rules(index))
            Else
                ClearFields()
            End If

            RefreshToggleActiveButton()
        End Sub

        Private Sub OnRuleDoubleClick(sender As Object, e As EventArgs)
            ToggleSelectedActive()
        End Sub

        Private Sub OnRulesKeyDown(sender As Object, e As KeyEventArgs)
            If e.Control AndAlso e.KeyCode = Keys.A Then
                RefreshList(Enumerable.Range(0, _rules.Count))
                ClearFields()
                e.Handled = True
                e.SuppressKeyPress = True
            End If
        End Sub

#End Region

#Region "Поля"

        Private Sub FillFields(rule As SortRule)
            _txtRuleName.Text = rule.Name
            _txtPattern.Text = rule.Pattern
            _txtExtensions.Text = rule.Extensions
            _txtSubject.Text = rule.Subject
            _txtTarget.Text = rule.Target
            _txtSourceFolder.Text = rule.SourceFolder
            _chkContinue.Checked = rule.ContinueToNextRules
        End Sub

        Private Sub ClearFields()
            For Each box In New TextBox() {_txtRuleName, _txtPattern, _txtExtensions, _txtSubject, _txtTarget, _txtSourceFolder}
                box.Text = String.Empty
            Next

            _chkContinue.Checked = False
        End Sub

        Private Function HasFieldValues() As Boolean
            Return New TextBox() {_txtRuleName, _txtPattern, _txtExtensions, _txtSubject, _txtTarget, _txtSourceFolder}.
                Any(Function(b) b.Text.Trim().Length > 0)
        End Function

        Private Function ReadRuleFromFields() As SortRule
            Return New SortRule With {
                .Name = SortRuleStore.CleanRulePart(_txtRuleName.Text),
                .Pattern = SortRuleStore.CleanRulePart(_txtPattern.Text),
                .Extensions = SortRuleStore.CleanExtensions(_txtExtensions.Text),
                .Subject = SortRuleStore.CleanRulePart(_txtSubject.Text),
                .Target = SortRuleStore.CleanTargetValue(_txtTarget.Text),
                .SourceFolder = SortRuleStore.CleanFolderPath(_txtSourceFolder.Text),
                .ContinueToNextRules = _chkContinue.Checked
            }
        End Function

        Private Function FindRuleIndex(candidate As SortRule) As Integer
            Return _rules.FindIndex(Function(r) r.SameDefinition(candidate))
        End Function

        Private Function BuildDuplicateMessage(index As Integer) As String
            Dim message = "Такое правило уже существует: №" & (index + 1).ToString(CultureInfo.InvariantCulture)
            Dim rule = _rules(index)

            If rule.Name.Length > 0 Then
                Return message & " """ & rule.Name & """"
            End If

            Return message & vbCrLf & rule.DisplayText(index + 1)
        End Function

        ''' <summary>Добавляет правило или обновляет правило с теми же условиями.</summary>
        Private Function UpsertFromFields(preserveExistingEnabled As Boolean) As Boolean
            Dim candidate = ReadRuleFromFields()

            If Not SortRuleStore.IsRuleDefinitionValid(candidate.Pattern, candidate.Target, candidate.Extensions, candidate.SourceFolder) Then
                UiHelper.Info(Me, SortRuleStore.ValidationMessage)
                Return False
            End If

            Dim index = FindRuleIndex(candidate)

            If index >= 0 Then
                candidate.Enabled = If(preserveExistingEnabled, _rules(index).Enabled, True)
                _rules(index) = candidate
            Else
                candidate.Enabled = True
                _rules.Add(candidate)
                index = _rules.Count - 1
            End If

            RefreshList(New Integer() {index})
            Return True
        End Function

#End Region

#Region "Кнопки"

        Private Sub OnAdd(sender As Object, e As EventArgs)
            UpsertFromFields(False)
        End Sub

        Private Sub OnUpdate(sender As Object, e As EventArgs)
            Dim index = SingleSelectedIndex()

            If index < 0 Then
                UiHelper.Info(Me, If(SelectedIndexes().Count > 1, "Выберите одно правило для обновления.", "Выберите правило для обновления."))
                Return
            End If

            UpdateRuleFromFields(index)
        End Sub

        ''' <summary>Заменяет правило index значениями полей. False — если поля некорректны или такое правило уже есть.</summary>
        Private Function UpdateRuleFromFields(index As Integer) As Boolean
            Dim candidate = ReadRuleFromFields()

            If Not SortRuleStore.IsRuleDefinitionValid(candidate.Pattern, candidate.Target, candidate.Extensions, candidate.SourceFolder) Then
                UiHelper.Info(Me, SortRuleStore.ValidationMessage)
                Return False
            End If

            Dim duplicateIndex = FindRuleIndex(candidate)

            If duplicateIndex >= 0 AndAlso duplicateIndex <> index Then
                UiHelper.Info(Me, BuildDuplicateMessage(duplicateIndex))
                Return False
            End If

            candidate.Enabled = _rules(index).Enabled
            _rules(index) = candidate
            RefreshList(New Integer() {index})
            Return True
        End Function

        Private Sub OnToggleActive(sender As Object, e As EventArgs)
            ToggleSelectedActive()
        End Sub

        Private Sub ToggleSelectedActive()
            Dim indexes = SelectedIndexes()

            If indexes.Count = 0 Then
                UiHelper.Info(Me, "Выберите правило для изменения активности.")
                Return
            End If

            Dim targetEnabled = indexes.Any(Function(i) Not _rules(i).Enabled)

            For Each i In indexes
                _rules(i).Enabled = targetEnabled
            Next

            RefreshList(indexes)
        End Sub

        Private Sub OnRemove(sender As Object, e As EventArgs)
            Dim indexes = SelectedIndexes()

            If indexes.Count = 0 Then
                UiHelper.Info(Me, "Выберите правило для удаления.")
                Return
            End If

            Dim question = If(indexes.Count = 1,
                              "Вы уверены, что хотите удалить правило №" & (indexes(0) + 1).ToString(CultureInfo.InvariantCulture) & "?",
                              "Вы уверены, что хотите удалить " & indexes.Count.ToString(CultureInfo.InvariantCulture) & " выбранных правил?")

            If Not UiHelper.Confirm(Me, question, "Удаление правила") Then
                Return
            End If

            For Each i In indexes.OrderByDescending(Function(x) x)
                _rules.RemoveAt(i)
            Next

            RefreshList(New Integer() {})
            ClearFields()
        End Sub

        Private Sub OnBrowseTarget(sender As Object, e As EventArgs)
            Dim folderPath = UiHelper.BrowseForFolder(Me, "Выберите целевую папку для этого правила сортировки", _txtTarget.Text)

            If folderPath.Length > 0 Then
                _txtTarget.Text = folderPath
            End If
        End Sub

        Private Sub OnBrowseSource(sender As Object, e As EventArgs)
            Try
                Using picker As New OutlookFolderPickerForm()
                    If picker.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK AndAlso picker.SelectedFolderPath.Length > 0 Then
                        _txtSourceFolder.Text = picker.SelectedFolderPath
                    End If
                End Using
            Catch ex As Exception
                UiHelper.ErrorBox(Me, "Не удалось получить список папок Outlook: " & ex.Message)
            End Try
        End Sub

        Private Sub OnCurrentSource(sender As Object, e As EventArgs)
            Try
                Dim folderPath = OutlookHost.GetFolderPath(OutlookHost.GetCurrentFolder())

                If folderPath.Length = 0 Then
                    UiHelper.Info(Me, "Текущая папка Outlook не выбрана.")
                    Return
                End If

                _txtSourceFolder.Text = folderPath
            Catch ex As Exception
                UiHelper.ErrorBox(Me, "Не удалось взять текущую папку Outlook: " & ex.Message)
            End Try
        End Sub

        Private Sub OnSave(sender As Object, e As EventArgs)
            If HasFieldValues() AndAlso Not ApplyFieldsBeforeSave() Then
                Return
            End If

            Try
                SortRuleStore.SaveRules(_rules)
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Close()
            Catch ex As Exception
                UiHelper.ErrorBox(Me, "Не удалось сохранить правила: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Переносит в список то, что набрано в полях, перед сохранением. Если выбрано одно правило и поля
        ''' отличаются от него, спрашивает: обновить это правило или добавить новое. Без вопроса измененное
        ''' правило добавлялось бы рядом со старым. False — сохранение нужно прервать.
        ''' </summary>
        Private Function ApplyFieldsBeforeSave() As Boolean
            Dim index = SingleSelectedIndex()

            If index < 0 OrElse Not FieldsDifferFrom(_rules(index)) Then
                Return UpsertFromFields(True)
            End If

            Dim answer = MessageBox.Show(Me,
                "Поля отличаются от выбранного правила №" & (index + 1).ToString(CultureInfo.InvariantCulture) & "." & vbCrLf & vbCrLf &
                "Да — обновить выбранное правило." & vbCrLf &
                "Нет — добавить как новое правило." & vbCrLf &
                "Отмена — вернуться к редактированию.",
                "Сохранение правил", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)

            Select Case answer
                Case System.Windows.Forms.DialogResult.Yes
                    Return UpdateRuleFromFields(index)
                Case System.Windows.Forms.DialogResult.No
                    Return UpsertFromFields(True)
                Case Else
                    Return False
            End Select
        End Function

        Private Function FieldsDifferFrom(rule As SortRule) As Boolean
            Dim candidate = ReadRuleFromFields()

            Return Not (String.Equals(candidate.Name, rule.Name, StringComparison.Ordinal) AndAlso
                        String.Equals(candidate.Pattern, rule.Pattern, StringComparison.Ordinal) AndAlso
                        String.Equals(candidate.Extensions, rule.Extensions, StringComparison.Ordinal) AndAlso
                        String.Equals(candidate.Subject, rule.Subject, StringComparison.Ordinal) AndAlso
                        String.Equals(candidate.Target, rule.Target, StringComparison.Ordinal) AndAlso
                        String.Equals(candidate.SourceFolder, rule.SourceFolder, StringComparison.Ordinal) AndAlso
                        candidate.ContinueToNextRules = rule.ContinueToNextRules)
        End Function

#End Region

    End Class

End Namespace

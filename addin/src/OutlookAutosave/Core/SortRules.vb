Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Text

Namespace Core

    ''' <summary>Правило сортировки вложений.</summary>
    Friend NotInheritable Class SortRule
        Public Property Enabled As Boolean = True
        Public Property Pattern As String = String.Empty
        Public Property Target As String = String.Empty
        Public Property Extensions As String = String.Empty
        Public Property SourceFolder As String = String.Empty
        Public Property Subject As String = String.Empty
        Public Property Name As String = String.Empty
        ''' <summary>
        ''' True — после этого правила вложение проверяется следующими правилами списка
        ''' (одно вложение может сохраниться в несколько папок). False — проверка на этом правиле заканчивается.
        ''' </summary>
        Public Property ContinueToNextRules As Boolean
        ''' <summary>Порядковый номер правила в списке (начиная с 1).</summary>
        Public Property Number As Integer

        Public Function Clone() As SortRule
            Return DirectCast(MemberwiseClone(), SortRule)
        End Function

        ''' <summary>
        ''' Ключ уникальности правила: условия отбора и целевая папка (без названия).
        ''' Правила с одинаковыми условиями, но разными целевыми папками допустимы.
        ''' </summary>
        Public Function DefinitionKey() As String
            Return Pattern & "|" & Extensions & "|" & SourceFolder & "|" & Subject & "|" & TextUtil.RemoveTrailingBackslash(Target)
        End Function

        Public Function SameDefinition(other As SortRule) As Boolean
            Return other IsNot Nothing AndAlso
                   String.Equals(DefinitionKey(), other.DefinitionKey(), StringComparison.CurrentCultureIgnoreCase)
        End Function

        ''' <summary>Подробное описание для списка правил (как в VBA-форме).</summary>
        Public Function DisplayText(Optional ruleNumber As Integer = 0) As String
            Dim text As String

            If Name.Length > 0 Then
                text = Name
            Else
                text = If(Pattern.Length > 0, Pattern, "Все файлы")

                If Extensions.Length > 0 Then
                    text &= " [" & Extensions & "]"
                End If

                If SourceFolder.Length > 0 Then
                    text &= " (Папка: " & SourceFolder & ")"
                End If

                If Subject.Length > 0 Then
                    text &= " (Тема: " & Subject & ")"
                End If

                text &= " -> " & Target
            End If

            If ContinueToNextRules Then
                text &= " [+ следующие правила]"
            End If

            If ruleNumber > 0 Then
                text = ruleNumber.ToString(Globalization.CultureInfo.InvariantCulture) & ". " & text
            End If

            Return text
        End Function

        ''' <summary>Короткая подпись для блока статуса.</summary>
        Public Function StatusLabel() As String
            Dim label = Name.Trim()

            If label.Length = 0 Then
                If Pattern.Trim().Length > 0 Then
                    label = Pattern.Trim()
                ElseIf Extensions.Trim().Length > 0 Then
                    label = "[" & Extensions.Trim() & "]"
                ElseIf Subject.Trim().Length > 0 Then
                    label = Subject.Trim()
                Else
                    label = "Все файлы"
                End If
            End If

            Return Number.ToString(Globalization.CultureInfo.InvariantCulture) & ". " & TextUtil.ShortenText(label, 36)
        End Function

        Public Function LogDetails() As String
            Return "RuleNo=" & Number.ToString(Globalization.CultureInfo.InvariantCulture) &
                   "; RuleName=""" & Name & """" &
                   "; FileText=""" & Pattern & """" &
                   "; Extensions=""" & Extensions & """" &
                   "; Subject=""" & Subject & """" &
                   "; SourceFolder=""" & SourceFolder & """" &
                   "; Target=""" & Target & """" &
                   "; ContinueToNextRules=" & If(ContinueToNextRules, "True", "False")
        End Function

        ''' <summary>Проверяет, подходит ли вложение под правило.</summary>
        Public Function Matches(attachmentName As String, mailFolderPath As String, mailSubject As String) As Boolean
            If Pattern.Length > 0 AndAlso Not TextUtil.ContainsIgnoreCase(attachmentName, Pattern) Then
                Return False
            End If

            Return ExtensionsMatch(attachmentName) AndAlso SourceFolderMatches(mailFolderPath) AndAlso SubjectMatches(mailSubject)
        End Function

        Private Function ExtensionsMatch(attachmentName As String) As Boolean
            Dim value = Extensions.Trim().Replace(" ", "").Replace(".", "").Replace(";", ",").ToLowerInvariant()

            If value.Length = 0 Then
                Return True
            End If

            Dim fileExtension = TextUtil.GetFileExtension(attachmentName)

            If fileExtension.Length = 0 Then
                Return False
            End If

            Return value.Split(","c).Contains(fileExtension, StringComparer.OrdinalIgnoreCase)
        End Function

        Private Function SourceFolderMatches(mailFolderPath As String) As Boolean
            Dim ruleFolder = TextUtil.NormalizeOutlookFolderPath(SourceFolder)

            If ruleFolder.Length = 0 Then
                Return True
            End If

            Return String.Equals(TextUtil.NormalizeOutlookFolderPath(mailFolderPath), ruleFolder, StringComparison.CurrentCultureIgnoreCase)
        End Function

        Private Function SubjectMatches(mailSubject As String) As Boolean
            Dim ruleSubject = Subject.Trim()
            Return ruleSubject.Length = 0 OrElse TextUtil.ContainsIgnoreCase(mailSubject, ruleSubject)
        End Function
    End Class

    ''' <summary>Чтение, нормализация и сохранение правил (перенос SortRulesSettingsCore.bas).</summary>
    Friend Module SortRuleStore

        Friend Const ValidationMessage As String =
            "Целевая папка должна быть абсолютной. Если текст имени файла пустой, заполните расширения и исходную папку Outlook."

        Friend Function IsRuleDefinitionValid(pattern As String, target As String, extensions As String, sourceFolder As String) As Boolean
            Return target.Length > 0 AndAlso TextUtil.IsAbsoluteFolderPath(target) AndAlso
                   (pattern.Length > 0 OrElse (extensions.Length > 0 AndAlso sourceFolder.Length > 0))
        End Function

        Friend Function CleanRulePart(value As String) As String
            Return If(value, String.Empty).Trim().Replace("|", "")
        End Function

        Friend Function CleanTargetValue(value As String) As String
            Return TrimQuotes(If(value, String.Empty).Trim()).Replace("/", "\").Replace("|", "")
        End Function

        Friend Function CleanFolderPath(value As String) As String
            Return CleanTargetValue(value)
        End Function

        Friend Function CleanExtensions(value As String) As String
            value = If(value, String.Empty).Trim().ToLowerInvariant().Replace(".", "").Replace(" ", "").Replace(";", ",").Replace("|", "")

            If value.Length = 0 Then
                Return String.Empty
            End If

            Dim result As New List(Of String)()

            For Each part In value.Split(","c)
                Dim cleaned = TextUtil.ReplaceInvalidCharacters(CleanRulePart(part))

                If cleaned.Length > 0 AndAlso Not result.Contains(cleaned, StringComparer.OrdinalIgnoreCase) Then
                    result.Add(cleaned)
                End If
            Next

            Return String.Join(",", result)
        End Function

        Private Function TrimQuotes(value As String) As String
            value = If(value, String.Empty).Trim()

            If value.Length >= 2 AndAlso value.StartsWith("""", StringComparison.Ordinal) AndAlso value.EndsWith("""", StringComparison.Ordinal) Then
                value = value.Substring(1, value.Length - 2)
            End If

            Return value.Trim()
        End Function

        Private Function SplitLines(value As String) As String()
            Return If(value, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(New Char() {ChrW(10)}, StringSplitOptions.None)
        End Function

        Private Function PartAt(parts As String(), index As Integer) As String
            Return If(index < parts.Length, parts(index), String.Empty)
        End Function

        ''' <summary>Разбирает и нормализует сохраненные строки правил (включая отключенные).</summary>
        Friend Function ParseRows(rows As String) As List(Of SortRule)
            Dim result As New List(Of SortRule)()
            Dim seen As New HashSet(Of String)(StringComparer.CurrentCultureIgnoreCase)

            For Each line In SplitLines(rows)
                If line.Trim().Length = 0 Then
                    Continue For
                End If

                Dim parts = line.Split("|"c)

                If parts.Length < 3 Then
                    Continue For
                End If

                Dim rule As New SortRule With {
                    .Enabled = (parts(0).Trim() <> "0"),
                    .Pattern = CleanRulePart(parts(1)),
                    .Target = CleanTargetValue(parts(2)),
                    .Extensions = CleanExtensions(PartAt(parts, 3)),
                    .SourceFolder = CleanFolderPath(PartAt(parts, 4)),
                    .Subject = CleanRulePart(PartAt(parts, 5)),
                    .Name = CleanRulePart(PartAt(parts, 6)),
                    .ContinueToNextRules = (PartAt(parts, 7).Trim() = "1")
                }

                If Not IsRuleDefinitionValid(rule.Pattern, rule.Target, rule.Extensions, rule.SourceFolder) Then
                    Continue For
                End If

                If seen.Add(rule.DefinitionKey()) Then
                    rule.Number = result.Count + 1
                    result.Add(rule)
                End If
            Next

            Return result
        End Function

        Friend Function SerializeRows(rules As IEnumerable(Of SortRule)) As String
            Dim sb As New StringBuilder()

            For Each rule In rules
                sb.Append(If(rule.Enabled, "1", "0")).Append("|"c)
                sb.Append(rule.Pattern).Append("|"c)
                sb.Append(rule.Target).Append("|"c)
                sb.Append(rule.Extensions).Append("|"c)
                sb.Append(rule.SourceFolder).Append("|"c)
                sb.Append(rule.Subject).Append("|"c)
                sb.Append(rule.Name).Append("|"c)
                ' Поле 7 — продолжить проверку следующих правил. VBA-версия его не читает.
                sb.Append(If(rule.ContinueToNextRules, "1", "0")).Append(vbLf)
            Next

            Return sb.ToString()
        End Function

        ''' <summary>Правила из старого формата "шаблон=папка".</summary>
        Friend Function ParseLegacyRules(legacyRules As String) As List(Of SortRule)
            Dim rows As New StringBuilder()

            For Each line In SplitLines(If(legacyRules, String.Empty).Replace(";", vbLf))
                Dim cleaned = line.Trim()
                Dim pos = cleaned.IndexOf("="c)

                If pos > 0 Then
                    Dim pattern = CleanRulePart(cleaned.Substring(0, pos))
                    Dim target = CleanTargetValue(cleaned.Substring(pos + 1))

                    If pattern.Length > 0 AndAlso target.Length > 0 Then
                        rows.Append("1|").Append(pattern).Append("|"c).Append(target).Append("||").Append(vbLf)
                    End If
                End If
            Next

            Return ParseRows(rows.ToString())
        End Function

        ''' <summary>Все правила для формы редактирования.</summary>
        Friend Function LoadAllRules() As List(Of SortRule)
            Dim rows = SettingsStore.GetValue(SettingsStore.KeySortRuleRows)

            If rows.Trim().Length = 0 Then
                Return ParseLegacyRules(SettingsStore.GetValue(SettingsStore.KeyLegacySortRules))
            End If

            Return ParseRows(rows)
        End Function

        ''' <summary>Только активные правила (для автосохранения). Номер правила сохраняется как в списке.</summary>
        Friend Function LoadActiveRules() As List(Of SortRule)
            Return LoadAllRules().Where(Function(r) r.Enabled).ToList()
        End Function

        Friend Sub SaveRules(rules As IEnumerable(Of SortRule))
            Dim normalized = ParseRows(SerializeRows(rules))
            SettingsStore.SetValue(SettingsStore.KeySortRuleRows, SerializeRows(normalized))

            ' Старый ключ поддерживается для совместимости с VBA-версией.
            Dim legacy = normalized.
                Where(Function(r) r.Enabled AndAlso r.Pattern.Length > 0).
                Select(Function(r) r.Pattern & "=" & r.Target)

            SettingsStore.SetValue(SettingsStore.KeyLegacySortRules, String.Join(vbLf, legacy))
        End Sub

    End Module

End Namespace

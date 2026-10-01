Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Text

Namespace Core

    ''' <summary>Строковые и файловые утилиты, перенесенные из MainModule.bas.</summary>
    Friend Module TextUtil

        Friend Const StoredDateFormat As String = "yyyy-MM-dd HH:mm:ss"
        Friend Const FileNameDateFormat As String = "yyyy-MM-dd_HH-mm-ss"

        Friend Function FormatDate(value As DateTime) As String
            Return value.ToString(StoredDateFormat, CultureInfo.InvariantCulture)
        End Function

        ''' <summary>Разбирает дату, сохраненную VBA-версией или надстройкой.</summary>
        Friend Function TryParseStoredDate(rawValue As String, ByRef result As DateTime) As Boolean
            rawValue = If(rawValue, String.Empty).Trim()

            If rawValue.Length = 0 Then
                Return False
            End If

            If DateTime.TryParseExact(rawValue, StoredDateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, result) Then
                Return True
            End If

            ' VBA мог сохранить дату числом (OLE Automation date).
            Dim oaDate As Double

            If Double.TryParse(rawValue.Replace(","c, "."c), NumberStyles.Float, CultureInfo.InvariantCulture, oaDate) Then
                Try
                    result = DateTime.FromOADate(oaDate)
                    Return True
                Catch ex As ArgumentException
                    Return False
                End Try
            End If

            Return DateTime.TryParse(rawValue, CultureInfo.CurrentCulture, DateTimeStyles.None, result) OrElse
                   DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.None, result)
        End Function

        ''' <summary>Разбирает интервал вида 30s, 15m, 12h, 3d. Возвращает 0, если значение некорректно.</summary>
        Friend Function ParseIntervalSeconds(inputValue As String) As Integer
            inputValue = If(inputValue, String.Empty).Trim().ToLowerInvariant().Replace(" ", "")

            Dim numberValue As New StringBuilder()
            Dim unitValue As New StringBuilder()

            For Each ch As Char In inputValue
                If (ch >= "0"c AndAlso ch <= "9"c) OrElse ch = "."c OrElse ch = ","c Then
                    numberValue.Append(If(ch = ","c, "."c, ch))
                Else
                    unitValue.Append(ch)
                End If
            Next

            Dim amount As Double

            If numberValue.Length = 0 OrElse
               Not Double.TryParse(numberValue.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, amount) Then
                Return 0
            End If

            Dim multiplier As Double

            Select Case unitValue.ToString()
                Case "", "s", "sec", "secs", "second", "seconds"
                    multiplier = 1
                Case "m", "min", "mins", "minute", "minutes"
                    multiplier = 60
                Case "h", "hr", "hrs", "hour", "hours"
                    multiplier = 3600
                Case "d", "day", "days"
                    multiplier = 86400
                Case Else
                    Return 0
            End Select

            Dim seconds = Math.Round(amount * multiplier, MidpointRounding.ToEven)

            If seconds <= 0 OrElse seconds > Integer.MaxValue Then
                Return 0
            End If

            Return CInt(seconds)
        End Function

        Friend Function FormatIntervalSeconds(intervalSeconds As Integer) As String
            If intervalSeconds > 0 AndAlso intervalSeconds Mod 86400 = 0 Then
                Return (intervalSeconds \ 86400).ToString(CultureInfo.InvariantCulture) & "d"
            ElseIf intervalSeconds > 0 AndAlso intervalSeconds Mod 3600 = 0 Then
                Return (intervalSeconds \ 3600).ToString(CultureInfo.InvariantCulture) & "h"
            ElseIf intervalSeconds > 0 AndAlso intervalSeconds Mod 60 = 0 Then
                Return (intervalSeconds \ 60).ToString(CultureInfo.InvariantCulture) & "m"
            Else
                Return intervalSeconds.ToString(CultureInfo.InvariantCulture) & "s"
            End If
        End Function

        Friend Function EnsureTrailingBackslash(folderPath As String) As String
            folderPath = If(folderPath, String.Empty)
            Return If(folderPath.EndsWith("\", StringComparison.Ordinal), folderPath, folderPath & "\")
        End Function

        Friend Function RemoveTrailingBackslash(folderPath As String) As String
            folderPath = If(folderPath, String.Empty).Trim()

            Do While folderPath.Length > 3 AndAlso folderPath.EndsWith("\", StringComparison.Ordinal)
                folderPath = folderPath.Substring(0, folderPath.Length - 1)
            Loop

            Return folderPath
        End Function

        Friend Function IsAbsoluteFolderPath(folderPath As String) As Boolean
            folderPath = If(folderPath, String.Empty).Trim()

            If folderPath.Length >= 3 Then
                Dim part = folderPath.Substring(1, 2)

                If part = ":\" OrElse part = ":/" Then
                    Return True
                End If
            End If

            Return folderPath.StartsWith("\\", StringComparison.Ordinal)
        End Function

        ''' <summary>Приводит список расширений к виду "pdf;zip;xlsx".</summary>
        Friend Function NormalizeAllowedExtensions(extensions As String) As String
            extensions = If(extensions, String.Empty).Replace(";", ",").Replace(".", "").Replace(" ", "").ToLowerInvariant()

            If extensions.Length = 0 Then
                Return String.Empty
            End If

            Dim result As New List(Of String)()

            For Each part In extensions.Split(","c)
                Dim cleaned = ReplaceInvalidCharacters(part)

                If cleaned.Length > 0 AndAlso Not result.Contains(cleaned, StringComparer.OrdinalIgnoreCase) Then
                    result.Add(cleaned)
                End If
            Next

            Return String.Join(";", result)
        End Function

        Friend Function FormatAllowedExtensionsForInput(extensions As String) As String
            Return If(extensions, String.Empty).Replace(";", ", ")
        End Function

        Friend Function IsExtensionInList(fileExtension As String, normalizedExtensions As String) As Boolean
            fileExtension = CleanFileExtension(fileExtension)
            normalizedExtensions = NormalizeAllowedExtensions(normalizedExtensions)

            If fileExtension.Length = 0 OrElse normalizedExtensions.Length = 0 Then
                Return False
            End If

            Return normalizedExtensions.Split(";"c).Contains(fileExtension, StringComparer.OrdinalIgnoreCase)
        End Function

        Friend Function CleanFileExtension(extensionValue As String) As String
            extensionValue = If(extensionValue, String.Empty).Trim().ToLowerInvariant().Replace(".", "").Replace(" ", "")
            Return ReplaceInvalidCharacters(extensionValue)
        End Function

        Friend Function GetFileExtension(fileName As String) As String
            fileName = If(fileName, String.Empty)
            Dim pos = fileName.LastIndexOf("."c)

            If pos >= 0 AndAlso pos < fileName.Length - 1 Then
                Return fileName.Substring(pos + 1).ToLowerInvariant()
            End If

            Return String.Empty
        End Function

        ''' <summary>Очищает имя файла вложения, сохраняя расширение.</summary>
        Friend Function CreateValidName(name As String) As String
            name = If(name, String.Empty)
            Dim pos = name.LastIndexOf("."c)
            Dim ext As String = String.Empty

            If pos >= 0 Then
                ext = name.Substring(pos)
                name = name.Substring(0, pos)
            End If

            name = ReplaceInvalidCharacters(name.TrimEnd(" "c))
            ext = ReplaceInvalidCharacters(ext)

            Return name & ext
        End Function

        Friend Function ReplaceInvalidCharacters(value As String) As String
            value = If(value, String.Empty)

            For Each ch In New Char() {"\"c, "/"c, ":"c, "*"c, "?"c, """"c, "<"c, ">"c, "|"c}
                value = value.Replace(ch.ToString(), String.Empty)
            Next

            ' Управляющие символы тоже недопустимы в именах файлов Windows.
            Dim sb As New StringBuilder(value.Length)

            For Each ch As Char In value
                If Not Char.IsControl(ch) Then
                    sb.Append(ch)
                End If
            Next

            Return sb.ToString()
        End Function

        Friend Function NormalizeOutlookFolderPath(folderPath As String) As String
            folderPath = If(folderPath, String.Empty).Trim().Replace("/", "\")

            Do While folderPath.Length > 0 AndAlso folderPath.EndsWith("\", StringComparison.Ordinal)
                folderPath = folderPath.Substring(0, folderPath.Length - 1)
            Loop

            Return folderPath
        End Function

        ''' <summary>Для "\\mailbox\Входящие\Отчеты" возвращает "\Входящие\Отчеты".</summary>
        Friend Function GetOutlookFolderPathSuffix(folderPath As String) As String
            folderPath = NormalizeOutlookFolderPath(folderPath)

            If Not folderPath.StartsWith("\\", StringComparison.Ordinal) Then
                Return String.Empty
            End If

            Dim withoutPrefix = folderPath.Substring(2)
            Dim pos = withoutPrefix.IndexOf("\"c)

            If pos < 0 OrElse pos >= withoutPrefix.Length - 1 Then
                Return String.Empty
            End If

            Return "\" & withoutPrefix.Substring(pos + 1)
        End Function

        Friend Function ShortenText(value As String, maxLength As Integer) As String
            value = If(value, String.Empty).Trim()

            If maxLength <= 3 OrElse value.Length <= maxLength Then
                Return value
            End If

            Return value.Substring(0, maxLength - 3) & "..."
        End Function

        ''' <summary>Однострочное значение для лога и индекса.</summary>
        Friend Function SingleLineField(value As String) As String
            Return If(value, String.Empty).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ").Replace("|", "/")
        End Function

        Friend Function StartsWithIgnoreCase(value As String, prefix As String) As Boolean
            Return If(value, String.Empty).StartsWith(If(prefix, String.Empty), StringComparison.OrdinalIgnoreCase)
        End Function

        Friend Function EndsWithIgnoreCase(value As String, suffix As String) As Boolean
            Return If(value, String.Empty).EndsWith(If(suffix, String.Empty), StringComparison.OrdinalIgnoreCase)
        End Function

        Friend Function ContainsIgnoreCase(value As String, fragment As String) As Boolean
            Return If(value, String.Empty).IndexOf(If(fragment, String.Empty), StringComparison.CurrentCultureIgnoreCase) >= 0
        End Function

    End Module

End Namespace

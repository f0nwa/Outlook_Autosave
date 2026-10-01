Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports OutlookAutosave.Interop

Namespace UI

    ''' <summary>Сообщения, ввод текста и выбор папки Windows.</summary>
    Friend Module UiHelper

        Friend Const AppTitle As String = "Outlook Autosave"

        Private _appIcon As Icon

        ''' <summary>Значок надстройки для заголовков окон.</summary>
        Friend Function AppIcon() As Icon
            If _appIcon Is Nothing Then
                Try
                    Using stream = GetType(UiHelper).Assembly.GetManifestResourceStream("OutlookAutosave.app.ico")
                        If stream IsNot Nothing Then
                            _appIcon = New Icon(stream)
                        End If
                    End Using
                Catch ex As Exception
                    Trace.WriteLine("OutlookAutosave: app icon not loaded: " & ex.Message)
                End Try
            End If

            Return _appIcon
        End Function

        Friend Sub Info(owner As IWin32Window, text As String)
            MessageBox.Show(owner, text, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Friend Sub Warn(owner As IWin32Window, text As String)
            MessageBox.Show(owner, text, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Sub

        Friend Sub ErrorBox(owner As IWin32Window, text As String)
            MessageBox.Show(owner, text, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Sub

        Friend Function Confirm(owner As IWin32Window, text As String, caption As String) As Boolean
            Return MessageBox.Show(owner, text, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = System.Windows.Forms.DialogResult.Yes
        End Function

        ''' <summary>Аналог InputBox. Возвращает пустую строку при отмене.</summary>
        Friend Function Prompt(owner As IWin32Window, text As String, caption As String, defaultValue As String) As String
            Using form As New Form()
                form.Text = caption
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
                form.StartPosition = If(owner Is Nothing, FormStartPosition.CenterScreen, FormStartPosition.CenterParent)
                form.MinimizeBox = False
                form.MaximizeBox = False
                form.ShowInTaskbar = False
                form.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
                form.Font = SystemFonts.MessageBoxFont
                form.ClientSize = New Size(460, 170)
                form.Icon = AppIcon()

                Dim label As New Label With {.Text = text, .Location = New Point(12, 12), .Size = New Size(436, 80)}
                Dim input As New TextBox With {.Text = defaultValue, .Location = New Point(12, 98), .Size = New Size(436, 23)}
                Dim ok As New Button With {.Text = "ОК", .DialogResult = System.Windows.Forms.DialogResult.OK, .Location = New Point(292, 134), .Size = New Size(75, 26)}
                Dim cancel As New Button With {.Text = "Отмена", .DialogResult = System.Windows.Forms.DialogResult.Cancel, .Location = New Point(373, 134), .Size = New Size(75, 26)}

                form.Controls.AddRange(New Control() {label, input, ok, cancel})
                form.AcceptButton = ok
                form.CancelButton = cancel

                If form.ShowDialog(owner) = System.Windows.Forms.DialogResult.OK Then
                    Return input.Text
                End If

                Return String.Empty
            End Using
        End Function

        ''' <summary>Выбор папки Windows: современный диалог, при ошибке — классический.</summary>
        Friend Function BrowseForFolder(owner As IWin32Window, title As String, initialPath As String) As String
            Try
                Return BrowseModern(owner, title, initialPath)
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: modern folder dialog failed: " & ex.Message)
            End Try

            Using dialog As New FolderBrowserDialog()
                dialog.Description = title
                dialog.ShowNewFolderButton = True

                If Not String.IsNullOrWhiteSpace(initialPath) AndAlso Directory.Exists(initialPath.Trim()) Then
                    dialog.SelectedPath = initialPath.Trim()
                End If

                If dialog.ShowDialog(owner) = System.Windows.Forms.DialogResult.OK Then
                    Return dialog.SelectedPath
                End If
            End Using

            Return String.Empty
        End Function

        Private Function BrowseModern(owner As IWin32Window, title As String, initialPath As String) As String
            Const FOS_NOCHANGEDIR As UInteger = &H8UI
            Const FOS_PICKFOLDERS As UInteger = &H20UI
            Const FOS_FORCEFILESYSTEM As UInteger = &H40UI
            Const FOS_PATHMUSTEXIST As UInteger = &H800UI
            Const SIGDN_FILESYSPATH As UInteger = &H80058000UI
            Const ERROR_CANCELLED As Integer = &H800704C7

            Dim dialogType = Type.GetTypeFromCLSID(New Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7"), True)
            Dim dialogObject = Activator.CreateInstance(dialogType)

            Try
                Dim dialog = DirectCast(dialogObject, IFileDialog)
                Dim options As UInteger
                dialog.GetOptions(options)
                dialog.SetOptions(options Or FOS_PICKFOLDERS Or FOS_FORCEFILESYSTEM Or FOS_PATHMUSTEXIST Or FOS_NOCHANGEDIR)
                dialog.SetTitle(title)

                If Not String.IsNullOrWhiteSpace(initialPath) AndAlso Directory.Exists(initialPath.Trim()) Then
                    Dim shellItemGuid As New Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")
                    Dim folderItem As IShellItem = Nothing
                    NativeMethods.SHCreateItemFromParsingName(initialPath.Trim(), IntPtr.Zero, shellItemGuid, folderItem)

                    If folderItem IsNot Nothing Then
                        dialog.SetFolder(folderItem)
                    End If
                End If

                Dim hwnd = If(owner Is Nothing, IntPtr.Zero, owner.Handle)
                Dim hr = dialog.Show(hwnd)

                If hr = ERROR_CANCELLED Then
                    Return String.Empty
                End If

                If hr <> 0 Then
                    Marshal.ThrowExceptionForHR(hr)
                End If

                Dim resultItem As IShellItem = Nothing
                dialog.GetResult(resultItem)

                Dim path As String = Nothing
                resultItem.GetDisplayName(SIGDN_FILESYSPATH, path)
                Return If(path, String.Empty)
            Finally
                Marshal.FinalReleaseComObject(dialogObject)
            End Try
        End Function

    End Module

End Namespace
